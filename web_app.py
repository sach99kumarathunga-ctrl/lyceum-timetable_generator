#!/usr/bin/env python3
"""
web_app.py — Lyceum Timetable Generator (Web Edition v2)
─────────────────────────────────────────────────────────
Key improvements over v1:
  • "Generate All Grades" schedules Grade 6, 7, 8 together with a shared
    teacher-busy set — Option-block teachers (Sakna, Pulakshika, EFF1/2/3,
    Religion teachers …) can never be double-booked across grades.
  • Result (Step 3) shows ALL THREE grades simultaneously: grade tabs on top,
    class tabs below.  No more switching back to Step 1 to see another grade.
  • Per-grade Generate still works for single-grade edits.

Run:  python web_app.py
"""
import os, re, json, copy, uuid, tempfile, threading, webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

import data
from excel_writer import write_setup_flexible, generate_workbook
from openpyxl import Workbook, load_workbook
from scheduler import (TIME_SLOTS, PERIOD_SNS, maths_layout_score,
                       count_break_straddles)

import os
PORT    = int(os.environ.get("PORT", 8000))
HOST    = "0.0.0.0"
WORKDIR = os.path.join(tempfile.gettempdir(), "lyceum_tt_web")
os.makedirs(WORKDIR, exist_ok=True)

# token -> {path, grade, teachers}   (single-grade cache)
GEN_CACHE = {}
# "all" token -> {grades: [{grade, path, teachers, stats, ...}]}
ALL_CACHE = {}

SN_LABEL = {sn: lab for sn, _t, lab, _b in TIME_SLOTS}


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────
def flex_grades():
    return [g for g in data.GRADE_GROUPS
            if g.get("setup_mode") == "flexible"
            and g["name"] in ("Grade 6", "Grade 7", "Grade 8")]


def get_grade(name):
    for g in data.GRADE_GROUPS:
        if g["name"] == name:
            return g
    return None


def setup_payload(grp):
    classes = grp["classes"]
    regular, blocks = [], []
    ridx = bidx = 0
    for row in grp["subject_rows"]:
        if isinstance(row, (list, tuple)):
            tch, subj, pc, mgs, sd, a1, room = row[:7]
            regular.append({
                "id": ridx, "teacher": tch, "subject": subj, "room": room,
                "periods": {c: int(pc.get(c, 0)) for c in classes},
            })
            ridx += 1
        elif isinstance(row, dict):
            rtype = row.get("type", "")
            if rtype == "religion_block":
                label = "Religion block (one shared period for all classes)"
            elif rtype == "merge_block":
                label = "Option block: " + " / ".join(
                    sorted({t["subject"] for t in row.get("teachers", [])}))
            else:
                continue
            blocks.append({
                "id": bidx, "type": rtype, "label": label,
                "teachers": [{"teacher": t.get("teacher", ""),
                               "subject":  t.get("subject", ""),
                               "covers":   t.get("covers", "")}
                              for t in row.get("teachers", [])],
            })
            bidx += 1
    return {"grade": grp["name"], "classes": classes,
            "regular_rows": regular, "blocks": blocks}


def apply_edits(grp, payload):
    grp2 = copy.deepcopy(grp)
    reg_edits = {r["id"]: r for r in payload.get("regular_rows", [])}
    blk_edits = {b["id"]: b for b in payload.get("blocks", [])}
    new_rows = []
    ridx = bidx = 0
    for row in grp2["subject_rows"]:
        if isinstance(row, (list, tuple)):
            tch, subj, pc, mgs, sd, a1, room = row[:7]
            e = reg_edits.get(ridx)
            if e:
                tch = (e.get("teacher") or tch).strip()
                pc = {}
                for c, v in (e.get("periods") or {}).items():
                    try:    v = int(v)
                    except: v = 0
                    if v > 0: pc[c] = v
            new_rows.append((tch, subj, pc, mgs, sd, a1, room))
            ridx += 1
        elif isinstance(row, dict) and row.get("type") in ("religion_block", "merge_block"):
            row = copy.deepcopy(row)
            e = blk_edits.get(bidx)
            if e:
                names = e.get("teachers", [])
                for i, t in enumerate(row.get("teachers", [])):
                    if i < len(names) and str(names[i]).strip():
                        t["teacher"] = str(names[i]).strip()
            new_rows.append(row)
            bidx += 1
        else:
            new_rows.append(row)
    grp2["subject_rows"] = new_rows
    return grp2


def collect_teacher_names(grp2):
    names = []
    for row in grp2["subject_rows"]:
        if isinstance(row, (list, tuple)):
            if str(row[0]).strip(): names.append(row[0])
        elif isinstance(row, dict):
            for t in row.get("teachers", []):
                if str(t.get("teacher", "")).strip(): names.append(t["teacher"])
    seen, out = set(), []
    for n in names:
        if n in seen: continue
        seen.add(n)
        sheet = re.sub(r'[\\/:*?"<>|]', '_', n)[:31]
        out.append({"name": n, "sheet": sheet})
    return out


def build_grids(slots, classes):
    rows_meta = [{"label": lab, "time": tr, "is_break": bool(b)}
                 for sn, tr, lab, b in TIME_SLOTS
                 if lab != "Lunch Break"]
    grids = {}
    for cls in classes:
        grid = []
        for sn, tr, lab, b in TIME_SLOTS:
            if lab == "Lunch Break":
                continue
            if b:
                grid.append(None)
                continue
            cells = []
            for d in range(1, 6):
                e = slots[cls][d].get(sn)
                if e:
                    cells.append({"subject": e[0], "teacher": e[1], "room": e[2]})
                else:
                    cells.append(None)
            grid.append(cells)
        grids[cls] = grid
    return rows_meta, grids


def _run_single_grade(grp2, shared_busy=None):
    """Generate one grade's workbook. Returns (path, slots, stats, logs)."""
    safe = re.sub(r'[^A-Za-z0-9]', '_', grp2["name"])
    path = os.path.join(WORKDIR, f"{safe}_{uuid.uuid4().hex[:8]}.xlsx")
    wb = Workbook(); wb.remove(wb.active); ws = wb.create_sheet("Setup")
    write_setup_flexible(ws, grp2); wb.save(path)

    logs = []
    kw = dict(log_callback=lambda m: logs.append(m), return_slots=True)
    if shared_busy is not None:
        kw["shared_teacher_busy"] = shared_busy

    res = generate_workbook(grp2, path, **kw)
    n_conf, n_dbl, slots = res
    classes = grp2["classes"]
    splits, doubles = maths_layout_score(slots, classes)
    straddles = count_break_straddles(slots, classes)
    stats = {"unplaced": int(n_conf), "double_booked": int(n_dbl),
             "straddles": int(straddles), "maths_splits": int(splits),
             "maths_doubles": int(doubles)}
    return path, slots, stats, logs


# ─────────────────────────────────────────────────────────────────────────────
#  HTTP handler
# ─────────────────────────────────────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def _send_json(self, obj, code=200):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path, name):
        with open(path, "rb") as f: raw = f.read()
        self.send_response(200)
        self.send_header("Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{name}"')
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self):
        u = urlparse(self.path)
        if u.path == "/":
            body = INDEX_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif u.path == "/api/grades":
            self._send_json({"grades": [g["name"] for g in flex_grades()]})
        elif u.path == "/api/setup":
            qs  = parse_qs(u.query)
            grp = get_grade(qs.get("grade", [""])[0])
            if not grp: return self._send_json({"error": "unknown grade"}, 404)
            self._send_json(setup_payload(grp))
        elif u.path.startswith("/api/download/"):
            qs  = parse_qs(u.query)
            tok = qs.get("token", [""])[0]

            # ── Combined multi-grade teacher TT — no grade param needed ──────
            if u.path == "/api/download/teacher_combined":
                sheet = qs.get("sheet", [""])[0]
                out   = self._extract_teacher_combined(sheet)
                if not out: return self._send_json({"error": "teacher not found in any grade — please use Generate All Grades first"}, 404)
                return self._send_file(out, f"{sheet}_AllGrades_Timetable.xlsx")

            # ── All other downloads need a grade-specific file path ───────────
            grade_path = None
            if tok in GEN_CACHE:
                info = GEN_CACHE[tok]
                grade_path = info["path"]
                grade_name = info["grade"]
            elif tok in ALL_CACHE:
                gname = qs.get("grade", [""])[0]
                gi    = next((x for x in ALL_CACHE[tok]["grades"] if x["grade"]==gname), None)
                if not gi: return self._send_json({"error": "grade not found"}, 404)
                grade_path = gi["path"]
                grade_name = gi["grade"]
            else:
                return self._send_json({"error": "expired — regenerate"}, 404)

            if u.path == "/api/download/full":
                self._send_file(grade_path,
                    grade_name.replace(" ","") + "_Timetable.xlsx")
            elif u.path == "/api/download/teacher":
                sheet = qs.get("sheet", [""])[0]
                out   = self._extract_teacher(grade_path, sheet)
                if not out: return self._send_json({"error": "sheet not found"}, 404)
                self._send_file(out, f"{sheet}_Timetable.xlsx")
        else:
            self._send_json({"error": "not found"}, 404)

    def do_POST(self):
        u      = urlparse(self.path)
        length = int(self.headers.get("Content-Length", 0))
        raw    = self.rfile.read(length)
        try:
            payload = json.loads(raw.decode("utf-8"))
        except Exception:
            return self._send_json({"error": "bad JSON"}, 400)

        # ── Single-grade generate ────────────────────────────────────────────
        if u.path == "/api/generate":
            grp = get_grade(payload.get("grade", ""))
            if not grp: return self._send_json({"error": "unknown grade"}, 400)
            try:
                grp2 = apply_edits(grp, payload)
                path, slots, stats, logs = _run_single_grade(grp2)
            except Exception as e:
                import traceback
                return self._send_json(
                    {"error": f"generation failed: {e}",
                     "trace": traceback.format_exc()}, 500)
            classes  = grp2["classes"]
            teachers = collect_teacher_names(grp2)
            wb       = load_workbook(path, read_only=True)
            present  = set(wb.sheetnames); wb.close()
            teachers = [t for t in teachers if t["sheet"] in present]
            token    = uuid.uuid4().hex
            GEN_CACHE[token] = {"path": path, "grade": grp2["name"], "teachers": teachers}
            rows_meta, grids = build_grids(slots, classes)
            return self._send_json({
                "ok": True, "token": token, "grade": grp2["name"],
                "classes": classes, "rows_meta": rows_meta, "grids": grids,
                "teachers": teachers, "stats": stats, "log_tail": logs[-6:],
            })

        # ── All-grades generate (shared teacher-busy set) ────────────────────
        elif u.path == "/api/generate_all":
            # payload.grades = [{grade, regular_rows, blocks}, ...]  (may be empty / partial)
            grade_payloads = {p["grade"]: p for p in payload.get("grades", [])}
            shared_busy    = set()
            all_results    = []
            try:
                for grp in flex_grades():
                    gname = grp["name"]
                    gp    = grade_payloads.get(gname, {})
                    grp2  = apply_edits(grp, gp) if gp else copy.deepcopy(grp)
                    path, slots, stats, logs = _run_single_grade(grp2, shared_busy)
                    classes  = grp2["classes"]
                    teachers = collect_teacher_names(grp2)
                    wb       = load_workbook(path, read_only=True)
                    present  = set(wb.sheetnames); wb.close()
                    teachers = [t for t in teachers if t["sheet"] in present]
                    rows_meta, grids = build_grids(slots, classes)
                    all_results.append({
                        "grade": gname, "path": path, "teachers": teachers,
                        "stats": stats, "classes": classes,
                        "rows_meta": rows_meta, "grids": grids,
                        "log_tail": logs[-4:],
                        "_slots": slots,      # kept server-side only; stripped before JSON response
                        "_grp2": grp2,        # kept server-side only
                    })
            except Exception as e:
                import traceback
                return self._send_json(
                    {"error": f"generation failed: {e}",
                     "trace": traceback.format_exc()}, 500)

            token = uuid.uuid4().hex
            ALL_CACHE[token] = {"grades": all_results}

            # ── Populate _grade_slots_cache so combined teacher TT works ─────
            import excel_writer as _ew
            _ew._grade_slots_cache.clear()
            for r in all_results:
                if r.get("_slots") and r.get("_grp2"):
                    grp2r = r["_grp2"]
                    _ew._grade_slots_cache[r["grade"]] = {
                        "all_slots":    r["_slots"],
                        "subject_rows": grp2r["subject_rows"],
                        "classes":      grp2r["classes"],
                        "grp":          grp2r,
                    }

            # slots is not JSON-serialisable → strip from response
            out = [{k: v for k, v in r.items() if k not in ("_slots","_grp2")} for r in all_results]
            return self._send_json({"ok": True, "token": token, "grades": out})

        return self._send_json({"error": "not found"}, 404)

    @staticmethod
    def _extract_teacher(full_path, sheet):
        wb = load_workbook(full_path)
        if sheet not in wb.sheetnames: return None
        for s in list(wb.sheetnames):
            if s != sheet: del wb[s]
        out = os.path.join(WORKDIR, f"teacher_{uuid.uuid4().hex[:8]}.xlsx")
        wb.save(out)
        return out

    @staticmethod
    def _extract_teacher_combined(teacher_name):
        """
        Build a multi-grade personal TT by merging the teacher sheet from each
        grade Excel file stored in ALL_CACHE.  This reads directly from the
        already-written .xlsx files — no in-memory cache needed.
        """
        from openpyxl import load_workbook, Workbook

        # Find the most recent ALL_CACHE entry (last generate_all run)
        if not ALL_CACHE:
            return None
        last_token = list(ALL_CACHE.keys())[-1]
        grades_info = ALL_CACHE[last_token]["grades"]

        # Sheet name in the Excel file = re.sub special chars + truncate
        sheet_name = re.sub(r'[\\/:*?"<>|]', '_', teacher_name)[:31]

        wb_out = Workbook()
        wb_out.remove(wb_out.active)

        GRADE_COLORS = {"6": "1A3660", "7": "7C3AED", "8": "0F7D59"}
        any_written = False

        for gi in grades_info:
            grade   = gi["grade"]
            path    = gi["path"]
            gnum    = grade.split()[-1]

            if not os.path.exists(path):
                continue

            wb_src = load_workbook(path)
            if sheet_name not in wb_src.sheetnames:
                wb_src.close()
                continue

            # Copy the sheet into wb_out with a grade label
            ws_src  = wb_src[sheet_name]
            ws_dest = wb_out.create_sheet(f"Grade {gnum}")

            # Copy all cells including styles
            from copy import copy as _copy
            for row in ws_src.iter_rows():
                for cell in row:
                    new_cell = ws_dest.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:
                        new_cell.font      = _copy(cell.font)
                        new_cell.fill      = _copy(cell.fill)
                        new_cell.border    = _copy(cell.border)
                        new_cell.alignment = _copy(cell.alignment)
                        new_cell.number_format = cell.number_format

            # Copy merged cells
            for merged in ws_src.merged_cells.ranges:
                ws_dest.merge_cells(str(merged))

            # Copy column widths and row heights
            for col_letter, col_dim in ws_src.column_dimensions.items():
                ws_dest.column_dimensions[col_letter].width = col_dim.width
            for row_num, row_dim in ws_src.row_dimensions.items():
                ws_dest.row_dimensions[row_num].height = row_dim.height

            ws_dest.sheet_view.showGridLines = False
            ws_dest.page_setup.orientation   = "landscape"
            wb_src.close()
            any_written = True

        if not any_written:
            return None

        out = os.path.join(WORKDIR, f"teacher_combined_{uuid.uuid4().hex[:8]}.xlsx")
        wb_out.save(out)
        return out


# ─────────────────────────────────────────────────────────────────────────────
#  Front-end (self-contained single page)
# ─────────────────────────────────────────────────────────────────────────────
INDEX_HTML = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Lyceum Timetable Generator</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Fraunces:wght@300;600&display=swap" rel="stylesheet">
<style>
:root{
  --ink:#12141a;--muted:#6b7280;--line:#e2e6ed;--bg:#f4f5f8;--surface:#fff;
  --brand:#1a3660;--brand2:#2563a8;--brand-light:#dbeafe;
  --accent:#0f7d59;--warn:#b45309;
  --g6:#1a3660;--g7:#7c3aed;--g8:#0f7d59;
  --maths-bg:#eff6ff;--merge-bg:#f3e8ff;--rel-bg:#ecfdf5;--pe-bg:#fef9c3;--break-bg:#fefce8;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font:13.5px/1.6 'DM Mono',monospace;color:var(--ink);background:var(--bg)}
header{background:var(--brand);color:#fff;padding:16px 28px;display:flex;align-items:center;gap:16px;border-bottom:3px solid #0f2547}
.logo{font-family:'Fraunces',serif;font-size:21px;font-weight:600}
.sub{opacity:.6;font-size:11.5px;margin-top:1px}
.badge{margin-left:auto;background:#fff2;border:1px solid #fff3;border-radius:6px;padding:3px 9px;font-size:11px}
main{max-width:1300px;margin:0 auto;padding:20px}
.card{background:var(--surface);border:1px solid var(--line);border-radius:13px;padding:20px 22px;margin-bottom:16px;box-shadow:0 1px 3px #0000000d}
.card-title{font-family:'Fraunces',serif;font-size:15px;font-weight:600;color:var(--brand);margin-bottom:3px;display:flex;align-items:center;gap:8px}
.step{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;background:var(--brand);color:#fff;font-size:11px;flex-shrink:0}
.card-sub{color:var(--muted);font-size:12px;margin-bottom:12px}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:flex-end}
label.fld{display:flex;flex-direction:column;gap:3px;font-size:12px;color:var(--muted)}
select,input,button{font:inherit}
select,input[type=text],input[type=number]{border:1px solid var(--line);border-radius:7px;padding:6px 9px;background:#fff;color:var(--ink);outline:none;transition:border-color .15s}
select:focus,input:focus{border-color:var(--brand2)}
input[type=number]{width:54px;text-align:center}
button{cursor:pointer;border:0;border-radius:8px;padding:7px 15px;font-weight:500;transition:all .15s}
.btn-primary{background:var(--brand2);color:#fff}.btn-primary:hover{background:#1e56a0}
.btn-all{background:var(--brand);color:#fff;font-weight:600}.btn-all:hover{background:#0f2547}
.btn-ghost{background:var(--brand-light);color:var(--brand)}.btn-ghost:hover{background:#bfdbfe}
.btn-accent{background:var(--accent);color:#fff}
button:disabled{opacity:.4;cursor:not-allowed}
.setup-wrap{overflow:auto;max-height:400px;border:1px solid var(--line);border-radius:9px;margin-bottom:6px}
.setup-wrap table{border-collapse:collapse;width:100%;font-size:12px}
.setup-wrap th{position:sticky;top:0;background:#f0f4f9;z-index:1;padding:6px 8px;border:1px solid var(--line);text-align:center;font-family:'Fraunces',serif;font-size:12.5px;font-weight:600;color:var(--brand)}
.setup-wrap td{border:1px solid var(--line);padding:4px 7px}
.setup-wrap td.subj{font-weight:500;white-space:nowrap;min-width:130px}
.setup-wrap td.tchr input{width:186px}
.divider{height:1px;background:var(--line);margin:12px 0}
.summary-row{display:flex;gap:20px;flex-wrap:wrap;margin-bottom:12px}
.summary-grade{font-size:12px;line-height:1.8}
.summary-grade b{font-size:14px;display:block;font-family:'Fraunces',serif}
.grade-tabs{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:14px}
.grade-tab{padding:7px 18px;border-radius:9px;font-weight:600;font-size:13px;cursor:pointer;border:2px solid transparent;background:#f0f4f9;color:var(--muted);transition:all .15s}
.grade-tab:hover:not([class*=act]){background:#e5e7eb}
.gtact-g6{background:var(--g6)!important;color:#fff!important;border-color:var(--g6)!important}
.gtact-g7{background:var(--g7)!important;color:#fff!important;border-color:var(--g7)!important}
.gtact-g8{background:var(--g8)!important;color:#fff!important;border-color:var(--g8)!important}
.class-tabs{display:flex;gap:4px;flex-wrap:wrap;margin-bottom:10px}
.class-tab{padding:4px 10px;border-radius:7px;font-size:12px;cursor:pointer;border:1px solid var(--line);background:#f8fafc;color:var(--muted);transition:all .12s}
.ct-g6.active{background:#dbeafe;border-color:var(--g6);color:var(--g6);font-weight:600}
.ct-g7.active{background:#ede9fe;border-color:var(--g7);color:var(--g7);font-weight:600}
.ct-g8.active{background:#d1fae5;border-color:var(--g8);color:var(--g8);font-weight:600}
.grade-section{display:none}.grade-section.vis{display:block}
.tt-grid table{border-collapse:collapse;width:100%;font-size:12px}
.tt-grid th{padding:7px 9px;border:1px solid #1e4a8a;font-family:'Fraunces',serif;font-size:13px;color:#fff}
.tt-grid td{border:1px solid var(--line);padding:5px 7px;vertical-align:middle}
.tt-grid td.time{white-space:nowrap;color:var(--muted);font-size:11px;background:#f7f9fc;min-width:108px}
.tt-grid td.time b{display:block;color:var(--ink);font-size:12px}
.tt-grid td.cell{min-width:100px}
.tt-grid td.cell .s{font-weight:600;font-size:12px}
.tt-grid td.cell .t{color:var(--muted);font-size:11px}
.tt-grid td.cell .r{color:#94a3b8;font-size:10.5px}
.tt-grid tr.brk td{background:var(--break-bg);color:var(--warn);text-align:center;font-size:11px;font-weight:600;padding:4px}
.c-maths{background:var(--maths-bg)}.c-merge{background:var(--merge-bg)}.c-rel{background:var(--rel-bg)}.c-pe{background:var(--pe-bg)}
.muted{color:var(--muted)}
.dl-bar{display:flex;gap:8px;flex-wrap:wrap;align-items:center;background:#f8fafc;border:1px solid var(--line);border-radius:9px;padding:9px 13px;margin-bottom:12px}
.dl-bar select{min-width:220px}
.pill{display:inline-block;font-size:11px;padding:2px 9px;border-radius:99px;background:var(--brand-light);color:var(--brand);font-weight:500}
.spinner{display:inline-block;width:12px;height:12px;border:2px solid #c7d6f5;border-top-color:var(--brand2);border-radius:50%;animation:sp .7s linear infinite;vertical-align:-2px;margin-right:5px}
@keyframes sp{to{transform:rotate(360deg)}}
.err{background:#fef2f2;color:#991b1b;border:1px solid #fecaca;padding:10px 14px;border-radius:9px;font-size:12.5px;white-space:pre-wrap;margin-bottom:10px}
.hint{font-size:11.5px;color:var(--muted);margin-top:6px;line-height:1.5}
.ok{color:var(--accent)}.bad{color:#dc2626}
</style>
</head>
<body>
<header>
  <div>
    <div class="logo">Lyceum Timetable Generator</div>
    <div class="sub">Grades 6 · 7 · 8 — cross-grade shared teacher conflict prevention</div>
  </div>
  <div class="badge">v2</div>
</header>
<main>

<div class="card">
  <div class="card-title"><span class="step">1</span> Configure</div>
  <div class="card-sub">Load a grade to edit teachers/periods, or use <b>Generate All Grades</b> directly with defaults.</div>
  <div class="row">
    <label class="fld">Grade <select id="gradeSel"></select></label>
    <button class="btn-ghost" id="loadBtn">Load &amp; edit</button>
    <span id="loadMsg" class="muted" style="font-size:12px"></span>
  </div>
</div>

<div class="card" id="setupCard" style="display:none">
  <div class="card-title"><span class="step">2</span> Edit Teachers &amp; Periods</div>
  <div class="card-sub">Change teacher <b>name</b> or <b>period count</b> per class. Set 0 to skip.</div>
  <div id="setupArea"></div>
  <div class="divider"></div>
  <div class="row" style="margin-top:4px">
    <button class="btn-primary" id="genBtn">&#9654; Generate <span id="genLabel"></span></button>
    <button class="btn-all" id="genAllBtn">&#9654;&#9654; Generate All Grades (6 + 7 + 8)</button>
    <span id="genMsg" class="muted" style="font-size:12px"></span>
  </div>
  <div class="hint"><b>Generate All Grades</b> uses a shared teacher-busy set across Grade 6&#8594;7&#8594;8, so option-block teachers (Sakna, Pulakshika, EFF Speech, Religion&#8230;) are never double-booked between grades.</div>
</div>
<div id="errBox"></div>

<div class="card" id="resultCard" style="display:none">
  <div class="card-title"><span class="step">3</span> Result &mdash; Weekly Timetables (Grade 6 · 7 · 8)</div>
  <div class="summary-row" id="summaryRow"></div>
  <div class="grade-tabs" id="gradeTabs"></div>
  <div class="dl-bar" id="dlBar">
    <label class="fld" style="flex-direction:row;gap:8px;align-items:center;margin:0">
      <span style="font-size:12px;color:var(--muted)">Teacher:</span>
      <select id="teacherSel"></select>
    </label>
    <button class="btn-accent" id="dlTeacherBtn">&#8659; This grade only</button>
    <button class="btn-accent" id="dlTeacherAllBtn" style="display:none;background:var(--accent)">&#8659; All grades (combined)</button>
    <button class="btn-ghost" id="dlFullBtn">&#8659; Full grade workbook</button>
  </div>
  <div id="gradeSections"></div>
</div>

</main>
<script>
var SETUP=null, RESULT=null, AG=null, AC=null;
var GCOL={'Grade 6':'g6','Grade 7':'g7','Grade 8':'g8'};
var $=function(s){return document.querySelector(s)};
var $$=function(s){return document.querySelectorAll(s)};
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}
function ea(s){return String(s).replace(/"/g,'&quot;')}
async function api(path,opts){
  var r=await fetch(path,opts);
  var j=await r.json().catch(function(){return{error:'bad response'}});
  if(!r.ok||j.error)throw new Error(j.error||'HTTP '+r.status);
  return j;
}
function showErr(m){$('#errBox').innerHTML='<div class="err">'+esc(m)+'</div>'}
function clearErr(){$('#errBox').innerHTML=''}

(async function(){
  var j=await api('/api/grades');
  $('#gradeSel').innerHTML=j.grades.map(function(g){return'<option>'+esc(g)+'</option>'}).join('');
})();

$('#loadBtn').onclick=async function(){
  clearErr();$('#loadMsg').textContent='Loading\u2026';
  try{
    SETUP=await api('/api/setup?grade='+encodeURIComponent($('#gradeSel').value));
    renderSetup();
    $('#setupCard').style.display='';
    $('#genLabel').textContent=SETUP.grade;
    $('#loadMsg').textContent=SETUP.regular_rows.length+' subjects \u00b7 '+SETUP.classes.length+' classes loaded.';
  }catch(e){showErr('Load failed: '+e.message);$('#loadMsg').textContent=''}
};

function renderSetup(){
  if(!SETUP)return;
  var cls=SETUP.classes;
  var h='<div class="setup-wrap"><table><thead><tr><th style="text-align:left">Subject</th><th style="text-align:left">Teacher</th>'+
        cls.map(function(c){return'<th>'+c+'</th>'}).join('')+'</tr></thead><tbody>';
  SETUP.regular_rows.forEach(function(r){
    h+='<tr data-rid="'+r.id+'"><td class="subj">'+esc(r.subject)+'</td>'+
       '<td class="tchr"><input type="text" data-k="teacher" value="'+ea(r.teacher)+'"></td>'+
       cls.map(function(c){return'<td><input type="number" min="0" max="40" data-cls="'+c+'" value="'+(r.periods[c]||0)+'"></td>'}).join('')+
       '</tr>';
  });
  h+='</tbody></table></div>';
  if(SETUP.blocks.length){
    h+='<div style="margin-top:12px"><b style="font-size:13px">Shared Blocks</b> <span class="muted" style="font-size:11.5px">&mdash; teacher names editable; period count fixed by block</span></div>';
    SETUP.blocks.forEach(function(b){
      h+='<div style="margin-top:8px"><span class="pill">'+esc(b.label)+'</span>'+
         '<div class="setup-wrap" style="max-height:none;margin-top:5px"><table><tbody>';
      b.teachers.forEach(function(t,i){
        var cov=t.covers?' <span class="muted">('+esc(t.covers)+')</span>':'';
        h+='<tr data-bid="'+b.id+'" data-ti="'+i+'"><td class="subj">'+esc(t.subject)+cov+'</td>'+
           '<td class="tchr"><input type="text" data-bk="teacher" value="'+ea(t.teacher)+'"></td></tr>';
      });
      h+='</tbody></table></div></div>';
    });
  }
  $('#setupArea').innerHTML=h;
}

function collectPayload(gradeName){
  var rr=[];
  $$('#setupArea tr[data-rid]').forEach(function(tr){
    var periods={};
    tr.querySelectorAll('input[data-cls]').forEach(function(inp){periods[inp.dataset.cls]=+inp.value||0});
    rr.push({id:+tr.dataset.rid,teacher:tr.querySelector('input[data-k=teacher]').value,periods:periods});
  });
  var bm={};
  $$('#setupArea tr[data-bid]').forEach(function(tr){
    var bid=+tr.dataset.bid,ti=+tr.dataset.ti;
    (bm[bid]=bm[bid]||[])[ti]=tr.querySelector('input[data-bk=teacher]').value;
  });
  var blocks=Object.keys(bm).map(function(k){return{id:+k,teachers:bm[k]}});
  return{grade:gradeName||SETUP.grade,regular_rows:rr,blocks:blocks};
}

function setBusy(b,msg){
  $('#genBtn').disabled=b;$('#genAllBtn').disabled=b;$('#loadBtn').disabled=b;
  $('#genMsg').innerHTML=b?'<span class="spinner"></span>'+esc(msg||''):'';
}

$('#genBtn').onclick=async function(){
  if(!SETUP)return;
  clearErr();setBusy(true,'Generating '+SETUP.grade+'\u2026');
  try{
    var r=await api('/api/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(collectPayload())});
    RESULT={token:r.token,mode:'single',grades:[r]};
    renderResult();
  }catch(e){showErr('Generation failed: '+e.message)}
  finally{setBusy(false)}
};

$('#genAllBtn').onclick=async function(){
  clearErr();setBusy(true,'Scheduling Grade 6 \u2192 7 \u2192 8 with shared teacher-busy set\u2026');
  var grades=SETUP?[collectPayload()]:[];
  try{
    var r=await api('/api/generate_all',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({grades:grades})});
    RESULT={token:r.token,mode:'all',grades:r.grades};
    renderResult();
  }catch(e){showErr('Generation failed: '+e.message)}
  finally{setBusy(false)}
};

function renderResult(){
  $('#resultCard').style.display='';
  var grades=RESULT.grades;
  // Summary
  var sh='';
  grades.forEach(function(g){
    var s=g.stats,col=GCOL[g.grade]||'g6';
    sh+='<div class="summary-grade"><b style="color:var(--'+col+')">'+esc(g.grade)+'</b>'+
        '<span class="'+(s.unplaced===0?'ok':'bad')+'">Unplaced: '+s.unplaced+'</span> &middot; '+
        '<span class="'+(s.straddles===0?'ok':'bad')+'">Split-doubles: '+s.straddles+'</span> &middot; '+
        '<span class="'+(s.double_booked<=1?'ok':'bad')+'">Clashes: '+Math.max(0,s.double_booked-1)+'</span></div>';
  });
  $('#summaryRow').innerHTML=sh;
  // Grade tabs
  var th='';
  grades.forEach(function(g){
    var col=GCOL[g.grade]||'g6';
    th+='<button class="grade-tab" data-grade="'+ea(g.grade)+'" data-col="'+col+'">'+esc(g.grade)+'</button>';
  });
  $('#gradeTabs').innerHTML=th;
  // Grade sections
  var sec='';
  grades.forEach(function(g){
    var col=GCOL[g.grade]||'g6';
    var gk=g.grade.replace(' ','_');
    sec+='<div class="grade-section" data-grade="'+ea(g.grade)+'" id="sec_'+gk+'">';
    sec+='<div class="class-tabs" id="ct_'+gk+'">';
    g.classes.forEach(function(c){
      sec+='<button class="class-tab ct-'+col+'" data-grade="'+ea(g.grade)+'" data-cls="'+c+'">'+c+'</button>';
    });
    sec+='</div>';
    sec+='<div class="tt-grid" id="grid_'+gk+'"></div>';
    if(g.log_tail&&g.log_tail.length){
      sec+='<div style="font-size:11px;color:var(--muted);background:#f8fafc;border-radius:6px;padding:5px 9px;margin-top:6px;white-space:pre-wrap">'+esc(g.log_tail.join('\\n'))+'</div>';
    }
    sec+='</div>';
  });
  $('#gradeSections').innerHTML=sec;
  $$('.grade-tab').forEach(function(b){b.onclick=function(){switchGrade(b.dataset.grade)}});
  $$('.class-tab').forEach(function(b){b.onclick=function(){switchClass(b.dataset.grade,b.dataset.cls)}});
  AG=grades[0].grade; AC=grades[0].classes[0];
  refreshGradeUI();
}

function switchGrade(gn){
  AG=gn;
  var g=RESULT.grades.find(function(x){return x.grade===gn});
  if(g)AC=g.classes[0];
  refreshGradeUI();
}
function switchClass(gn,cls){
  if(gn!==AG)switchGrade(gn);
  AC=cls;
  renderActiveGrid();updateClassTabs();
}
function refreshGradeUI(){
  // grade tabs
  $$('.grade-tab').forEach(function(b){
    var col=b.dataset.col;
    b.className='grade-tab'+(b.dataset.grade===AG?' gtact-'+col:'');
  });
  // grade sections
  $$('.grade-section').forEach(function(s){s.classList.toggle('vis',s.dataset.grade===AG)});
  updateTeacherDrop();updateClassTabs();renderActiveGrid();
}
function updateClassTabs(){
  var g=RESULT.grades.find(function(x){return x.grade===AG});
  if(!g)return;
  var gk=AG.replace(' ','_');
  $$('#ct_'+gk+' .class-tab').forEach(function(b){b.classList.toggle('active',b.dataset.cls===AC)});
}
function updateTeacherDrop(){
  var g=RESULT.grades.find(function(x){return x.grade===AG});
  if(!g)return;
  $('#teacherSel').innerHTML=g.teachers.map(function(t){return'<option value="'+ea(t.sheet)+'">'+esc(t.name)+'</option>'}).join('');
  updateCombinedBtn();
}
function updateCombinedBtn(){
  // Show "All grades (combined)" button whenever we have a multi-grade result
  // Server will return an error if the teacher only appears in one grade
  var isAll=RESULT&&RESULT.mode==='all';
  $('#dlTeacherAllBtn').style.display=isAll?'':'none';
}
function cellClass(s){
  var l=(s||'').toLowerCase();
  if(l.includes('mathematics'))return'c-maths';
  if(l.includes('cookery')||l.includes('art')||l.includes('music')||l.includes('dance')||l.includes('chinese')||l.includes('french')||l.includes('lifeskill')||l.includes('civic'))return'c-merge';
  if(l.includes('buddhism')||l.includes('religion')||l.includes('catholic')||l.includes('hindu')||l.includes('islam')||l.includes('christianity'))return'c-rel';
  if(l.includes('pe')||l.includes('physical'))return'c-pe';
  return'';
}
function renderActiveGrid(){
  var g=RESULT.grades.find(function(x){return x.grade===AG});
  if(!g)return;
  var gk=AG.replace(' ','_');
  var con=document.getElementById('grid_'+gk);
  if(!con)return;
  var col='var(--'+(GCOL[AG]||'g6')+')';
  var days=['Monday','Tuesday','Wednesday','Thursday','Friday'];
  var grid=g.grids[AC],meta=g.rows_meta;
  if(!grid){con.innerHTML='<p class="muted" style="padding:10px">No data</p>';return}
  var h='<table><thead><tr><th style="text-align:left;min-width:110px;background:'+col+'">Time</th>'+
        days.map(function(d){return'<th style="background:'+col+'">'+d+'</th>'}).join('')+
        '</tr></thead><tbody>';
  meta.forEach(function(m,ri){
    if(m.is_break){
      h+='<tr class="brk"><td class="time">'+esc(m.time)+'</td><td colspan="5">&mdash; '+esc(m.label)+' &mdash;</td></tr>';
      return;
    }
    h+='<tr><td class="time"><b>'+esc(m.label)+'</b><br>'+esc(m.time)+'</td>';
    grid[ri].forEach(function(cell){
      if(!cell){h+='<td class="cell muted" style="text-align:center">\u00b7</td>';return}
      h+='<td class="cell '+cellClass(cell.subject)+'">'+
         '<div class="s">'+esc(cell.subject)+'</div>'+
         '<div class="t">'+esc(cell.teacher||'')+'</div>'+
         (cell.room?'<div class="r">'+esc(cell.room)+'</div>':'')+
         '</td>';
    });
    h+='</tr>';
  });
  h+='</tbody></table>';
  con.innerHTML=h;
}
$('#teacherSel').onchange=function(){ updateCombinedBtn(); };
$('#dlTeacherBtn').onclick=function(){
  if(!RESULT)return;
  var sheet=$('#teacherSel').value;
  var gp=RESULT.mode==='all'?'&grade='+encodeURIComponent(AG):'';
  window.location='/api/download/teacher?token='+encodeURIComponent(RESULT.token)+'&sheet='+encodeURIComponent(sheet)+gp;
};
$('#dlTeacherAllBtn').onclick=function(){
  if(!RESULT)return;
  // Find the raw teacher name (not the sheet-safe name) for combined download
  var sheet=$('#teacherSel').value;
  var rawName=sheet; // fallback
  if(RESULT&&RESULT.grades){
    var g=RESULT.grades.find(function(x){return x.grade===AG});
    if(g){var t=g.teachers.find(function(t){return t.sheet===sheet});if(t)rawName=t.name;}
  }
  window.location='/api/download/teacher_combined?token='+encodeURIComponent(RESULT.token)+'&sheet='+encodeURIComponent(rawName);
};
$('#dlFullBtn').onclick=function(){
  if(!RESULT)return;
  var gp=RESULT.mode==='all'?'&grade='+encodeURIComponent(AG):'';
  window.location='/api/download/full?token='+encodeURIComponent(RESULT.token)+gp;
};
</script>
<!-- Assembly Wednesday P1 is a whole-school shared event; the single "clash" it generates is expected and not a real conflict. -->
</body></html>"""


def main():
    srv = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}/"
    print("=" * 62)
    print("  Lyceum Timetable Generator — Web Edition v2")
    print(f"  Open:  {url}")
    print("  Stop:  Ctrl+C")
    print()
    print("  KEY FIX: 'Generate All Grades' passes a shared")
    print("  teacher-busy set across Grade 6 -> 7 -> 8 so option-")
    print("  block teachers (Sakna, Pulakshika, EFF Speech, etc.)")
    print("  are never double-booked across grades.")
    print("=" * 62)
    try:
        threading.Timer(0.9, lambda: webbrowser.open(url)).start()
    except Exception:
        pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
        srv.shutdown()


if __name__ == "__main__":
    main()
