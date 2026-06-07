"""
scheduler.py  –  Enhanced scheduling engine v5.1
New: religion_block support — multiple teachers share ONE slot, each teaching
     their own religion to their assigned classes simultaneously.
"""
from collections import defaultdict, Counter

TIME_SLOTS = [
    (1,  "7:40 - 8:15",   "Period 1",         False),
    (2,  "8:15 - 8:30",   "Register Marking", True),
    (3,  "8:30 - 9:10",   "Period 2",         False),
    (4,  "9:10 - 9:50",   "Period 3",         False),
    (5,  "9:50 - 10:25",  "Period 4",         False),
    (6,  "10:25 - 10:50", "Interval",         True),
    (7,  "10:50 - 10:55", "Seiri Time",       True),
    (8,  "10:55 - 11:35", "Period 5",         False),
    (9,  "11:35 - 12:15", "Period 6",         False),
    (10, "12:15 - 1:00",  "Period 7",         False),
    (11, "12:45 - 1:00",  "Lunch Break",      True),
    (12, "1:00 - 1:45",   "Period 8",         False),
]
PERIOD_SNS = [s[0] for s in TIME_SLOTS if not s[3]]
DAYS       = ["Monday","Tuesday","Wednesday","Thursday","Friday"]
SLOT_ROW   = {sn: sn + 4 for sn, *_ in TIME_SLOTS}

# Clock-contiguous teaching-slot pairs (for double periods): two teaching
# periods form a valid double ONLY when genuinely back-to-back in clock time,
# i.e. adjacent in TIME_SLOTS with no break (Register/Interval/Seiri/Lunch)
# between them. Derived from TIME_SLOTS so it stays correct if the bell changes.
CONTIGUOUS_PAIRS = set()
for _i in range(len(TIME_SLOTS) - 1):
    _sn_a, _, _, _brk_a = TIME_SLOTS[_i]
    _sn_b, _, _, _brk_b = TIME_SLOTS[_i + 1]
    if not _brk_a and not _brk_b:
        CONTIGUOUS_PAIRS.add(frozenset((_sn_a, _sn_b)))


def _is_straddle_pair(a, b):
    """True if teaching periods a and b sit immediately on either side of a
    break (adjacent in the teaching-period list but NOT clock-contiguous).
    These are the 'one period before the interval + one period after it'
    pairs that look like a broken double — to be avoided for any subject."""
    ia, ib = PERIOD_SNS.index(a), PERIOD_SNS.index(b)
    return abs(ia - ib) == 1 and frozenset((a, b)) not in CONTIGUOUS_PAIRS


def _contiguous_runs(slots_avail):
    """Group the available teaching periods into maximal clock-contiguous runs
    (genuinely back-to-back periods with no break between)."""
    s = sorted(slots_avail, key=PERIOD_SNS.index)
    runs, cur = [], []
    for p in s:
        if cur and frozenset((cur[-1], p)) in CONTIGUOUS_PAIRS:
            cur.append(p)
        else:
            if cur:
                runs.append(cur)
            cur = [p]
    if cur:
        runs.append(cur)
    return runs


def _pick_block_slots(slots_avail, n, rng=None):
    """Choose `n` teaching periods for a same-day block (e.g. a double period
    or option block), preferring genuinely adjacent (back-to-back) periods.

    Priority:
      1. A clock-contiguous run of length >= n  → a real adjacent double/block.
      2. Otherwise `n` slots that form NO break-straddling pair (so the block
         never reads as 'one period before the break + one after it').
      3. Last resort: the first `n` free slots (keeps the grid fully placed).
    `rng` (optional) lets best-of-N attempts explore equally-valid choices.
    Returns a list of period serial numbers, or None if fewer than n are free.
    """
    av = sorted(slots_avail, key=PERIOD_SNS.index)
    if len(av) < n:
        return None
    if n <= 1:
        return av[:n]
    runs = _contiguous_runs(av)
    long_runs = [r for r in runs if len(r) >= n]
    if long_runs:
        long_runs.sort(key=len, reverse=True)
        if rng is not None:
            best_len = len(long_runs[0])
            top = [r for r in long_runs if len(r) == best_len]
            r = rng.choice(top)
        else:
            r = long_runs[0]
        return r[:n]                          # genuine adjacent block — preferred
    import itertools
    combos = [c for c in itertools.combinations(av, n)
              if not any(_is_straddle_pair(c[i], c[j])
                         for i in range(n) for j in range(i + 1, n))]
    if combos:
        return list(rng.choice(combos) if rng is not None else combos[0])
    return av[:n]                             # last resort: keep it placed


def _is_maths(subject_name):
    return "mathematics" in str(subject_name).lower()


def maths_layout_score(slots, class_list):
    """Quality of the Maths layout: (#split_days, #contiguous_double_days).
    A 'split day' is a day with two Maths periods that are NOT a genuine
    back-to-back double (e.g. straddling a break, or far apart). Lower split is
    better; more contiguous doubles is better."""
    splits = doubles = 0
    for cls in class_list:
        for d in range(1, 6):
            ms = sorted(p for p in PERIOD_SNS
                        if slots[cls][d].get(p) and _is_maths(slots[cls][d][p][0]))
            if len(ms) < 2:
                continue
            contig = any(frozenset((ms[i], ms[i + 1])) in CONTIGUOUS_PAIRS
                         for i in range(len(ms) - 1))
            if len(ms) == 2 and frozenset((ms[0], ms[1])) in CONTIGUOUS_PAIRS:
                doubles += 1
            elif contig:
                doubles += 1
            else:
                splits += 1
    return splits, doubles


def count_break_straddles(slots, class_list):
    """Count same-subject, same-teacher, same-day pairs that sit immediately on
    either side of a break (one period before it + one period after it). These
    read as a 'broken double' and are what we want to drive to zero. Used to
    rank candidate layouts so the scheduler prefers genuinely adjacent doubles."""
    n = 0
    for cls in class_list:
        for d in range(1, 6):
            bysub = {}
            for p in PERIOD_SNS:
                e = slots[cls][d].get(p)
                if e:
                    bysub.setdefault((e[0], e[1]), []).append(p)
            for ps in bysub.values():
                ps.sort()
                for i in range(len(ps) - 1):
                    if _is_straddle_pair(ps[i], ps[i + 1]):
                        n += 1
    return n

# ── Fixed slot constants ──────────────────────────────────────
ASSEMBLY_DAY    = 3   # Wednesday
ASSEMBLY_PERIOD = 1   # Period 1 (7:40–8:15)
ASSEMBLY_KEYWORDS = {"assembly", "assemble", "asm"}

def _is_assembly(subject_name):
    return any(kw in subject_name.lower() for kw in ASSEMBLY_KEYWORDS)


def schedule_flexible(class_list, subject_rows, shared_teacher_busy=None, attempt=0):
    """
    Flexible scheduler for Gr 6/7/8.
    shared_teacher_busy: set of (teacher, day, period) already used by OTHER grades.
    Pass this in when scheduling multiple grades with shared teachers so they
    don't double-book the same teacher across grades.
    After scheduling, the set is updated in-place with this grade's teacher slots.
    """
    # ── Normalise all rows to dicts ─────────────────────────────
    import random as _random
    _rng = _random.Random(attempt * 7919 + len(class_list))  # deterministic per attempt
    def _jit():
        return _rng.random() if attempt else 0.0
    _normalised = []
    for row in subject_rows:
        if isinstance(row, (list, tuple)):
            tch_,subj_,pc_,mgs_,sd_,a1_,room_ = row[:7]
            bid_ = row[7] if len(row) > 7 else ""
            _normalised.append({"teacher":tch_,"subject":subj_,"room":room_,
                                 "per_class":pc_,"merge_groups":mgs_,
                                 "same_day":sd_,"all_1period":a1_,"basket_id":bid_})
        else:
            _normalised.append(row)
    subject_rows = _normalised

    slots = {cls: {d: {p: None for p in PERIOD_SNS} for d in range(1, 6)}
             for cls in class_list}

    teacher_busy     = set()
    class_day_load   = defaultdict(int)
    teacher_day_load = defaultdict(int)
    subj_day_placed  = defaultdict(int)

    # Seed teacher_busy with slots already used by other grades (shared teachers)
    if shared_teacher_busy:
        for entry in shared_teacher_busy:
            teacher_busy.add(entry)

    MAX_PER_DAY = len(PERIOD_SNS)   # 8 teaching periods per class per day
    MAX_TCH_DAY = len(PERIOD_SNS)   # teacher can teach all 8 periods in a day if needed

    # ─────────────────────────────────────────────────────────
    # STEP 1: Pre-place Assembly → Wednesday Period 1
    # Assembly is ALWAYS Wednesday P1 for ALL classes.
    # No subject row needed — placed automatically.
    # ─────────────────────────────────────────────────────────
    d, p = ASSEMBLY_DAY, ASSEMBLY_PERIOD
    for cls in class_list:
        if slots[cls][d][p] is None:
            slots[cls][d][p] = ("Assembly", "ASSEMBLY", "R-HLL")
            # Assembly is a fixed school event — does NOT consume a teaching slot
            # Leave class_day_load unchanged so Wednesday keeps all 8 teaching slots
            subj_day_placed[(cls, "ASSEMBLY", d)]  += 1
    # Also handle any explicit Assembly rows in subject_rows (legacy support)
    for row in subject_rows:
        if row.get("type") == "religion_block":
            continue
        if not _is_assembly(row.get("subject", "")):
            continue
        tch  = row["teacher"]
        subj = row["subject"]
        room = row.get("room", "R-HLL")
        skey = f"{tch}|{subj}"
        pc   = row.get("per_class", {})
        a1   = row.get("all_1period", False)
        for cls in class_list:
            has_cls = a1 or (cls in pc and pc[cls] > 0)
            if not has_cls:
                continue
            # Update the slot with teacher info if provided
            slots[cls][d][p] = (subj, tch, room)
            teacher_busy.add((tch, d, p))
            teacher_day_load[(tch, d)] += 1
            subj_day_placed[(cls, skey, d)] += 1

    # ─────────────────────────────────────────────────────────
    # STEP 2: Pre-place Religion Block
    #
    # The WHOLE GRADE merges into one shared period once per week.
    # All religion teachers are active SIMULTANEOUSLY in that same slot:
    #   - Buddhism: may have MULTIPLE teachers (high student count),
    #     each in their own room taking a subset of Buddhist students
    #   - Other religions (Islam/Hindu/RC/Christianity): ONE teacher each,
    #     takes ALL students of that religion from the merged grade pool
    #
    # Every class timetable shows: "Religion Period" at that slot
    # Teacher personal TT shows:  subject + covers description
    # ─────────────────────────────────────────────────────────
    religion_blocks = [r for r in subject_rows if r.get("type") == "religion_block"]

    for rb in religion_blocks:
        teachers_in_block  = rb.get("teachers", [])
        all_classes_in_block = [c for c in rb.get("all_classes", class_list)
                                 if c in class_list]
        if not teachers_in_block or not all_classes_in_block:
            continue

        # Find ONE slot where ALL classes AND ALL religion teachers are free
        placed = False
        for d in sorted(range(1, 6), key=lambda d: sum(
                class_day_load[(c, d)] for c in all_classes_in_block)):
            if placed: break
            for p in PERIOD_SNS:
                if d == ASSEMBLY_DAY and p == ASSEMBLY_PERIOD:
                    continue
                # Every class in the grade must be free
                if any(slots[cls][d][p] is not None
                       for cls in all_classes_in_block):
                    continue
                # No class at day load limit
                if any(class_day_load[(cls, d)] >= MAX_PER_DAY
                       for cls in all_classes_in_block):
                    continue
                # Every religion teacher must be free
                if any((t["teacher"], d, p) in teacher_busy
                       for t in teachers_in_block):
                    continue

                # ── Place "Religion Period" in every class slot ──────
                for cls in all_classes_in_block:
                    slots[cls][d][p] = ("Religion Period", "RELIGION_BLOCK", "R-REL")
                    class_day_load[(cls, d)]                    += 1
                    subj_day_placed[(cls, "RELIGION_BLOCK", d)] += 1

                # ── All religion teachers busy this slot ─────────────
                for t in teachers_in_block:
                    teacher_busy.add((t["teacher"], d, p))
                    teacher_day_load[(t["teacher"], d)] += 1

                # ── Save placed slot for use in teacher TT writer ────
                rb["_placed_day"]    = d
                rb["_placed_period"] = p
                placed = True
                break

    # ─────────────────────────────────────────────────────────
    # STEP 2c: Pre-place Merge Blocks
    #
    # Each merge_block: multiple teachers active SIMULTANEOUSLY per merge pair.
    # Students in the pair choose any 1 subject.
    # ALL pairs of the block fall on the SAME weekday (same_day=True).
    # Each pair gets `periods` consecutive-or-spread slots on that day.
    # Standalone classes (e.g. 6G) scheduled normally via regular rows below.
    # ─────────────────────────────────────────────────────────
    merge_blocks  = [r for r in subject_rows if r.get("type") == "merge_block"]
    mb_day_chosen = {}   # block_id -> chosen weekday

    for mb in merge_blocks:
        bid        = mb.get("block_id", "mb")
        periods    = mb.get("periods", 2)
        same_day   = mb.get("same_day", True)
        pairs      = mb.get("merge_pairs", [])
        teachers   = mb.get("teachers", [])

        if not pairs or not teachers:
            continue

        # Collect all classes that appear in any pair
        all_pair_cls = [c for pair in pairs for c in pair if c in class_list]

        # Determine candidate days (same day across all pairs in this block)
        if same_day and bid in mb_day_chosen:
            candidate_days = [mb_day_chosen[bid]]
        else:
            candidate_days = sorted(range(1, 6), key=lambda d: (
                sum(class_day_load[(c, d)] for c in all_pair_cls),
                sum(teacher_day_load[(t["teacher"], d)] for t in teachers),
                _jit(),
            ))

        # Within a day, vary the order pairs claim slots so best-of-N can find a
        # layout where every option block is a clean adjacent double (or at least
        # a well-separated pair) rather than one straddling a break.
        pair_iter_base = list(pairs)

        # Try to place ALL pairs on the same day — each pair must use DIFFERENT slots
        for day in candidate_days:
            pair_placements = {}   # pair_key -> [period_sn, ...]
            used_slots_this_day = set()  # slots already assigned to earlier pairs

            pair_order = list(pair_iter_base)
            if attempt:
                _rng.shuffle(pair_order)

            # For each pair find `periods` free slots on this day (avoiding slots used by other pairs)
            for pair in pair_order:
                pair_cls = [c for c in pair if c in class_list]
                if not pair_cls:
                    continue

                # Collect ALL free teaching slots for this pair on this day,
                # then choose them as a genuine back-to-back block when possible
                # (a 2-period option block should be a real double, never a
                # 'before the interval + after the interval' split).
                avail = []
                for p in PERIOD_SNS:
                    if day == ASSEMBLY_DAY and p == ASSEMBLY_PERIOD:
                        continue
                    if p in used_slots_this_day:
                        continue  # another pair already uses this slot today
                    if any(slots[c][day][p] is not None for c in pair_cls):
                        continue
                    if any(class_day_load[(c, day)] >= MAX_PER_DAY for c in pair_cls):
                        continue
                    if any((t["teacher"], day, p) in teacher_busy for t in teachers):
                        continue
                    avail.append(p)

                free_slots = _pick_block_slots(avail, periods,
                                               rng=(_rng if attempt else None))
                if not free_slots or len(free_slots) < periods:
                    pair_placements = {}; break

                pair_placements[tuple(pair)] = free_slots
                used_slots_this_day.update(free_slots)

            if len(pair_placements) != len([p for p in pairs
                                            if any(c in class_list for c in p)]):
                continue  # try next day

            # Place all pairs on this day
            for pair, slot_list in pair_placements.items():
                pair_cls = [c for c in pair if c in class_list]
                pair_label = "+".join(sorted(pair_cls))
                for p in slot_list:
                    # Mark all pair classes with "Merge Block" marker
                    for c in pair_cls:
                        slots[c][day][p] = (f"Merge:{pair_label}", f"MERGE_{bid}", "R-MRG")
                        class_day_load[(c, day)]              += 1
                        subj_day_placed[(c, f"MERGE_{bid}", day)] += 1
                    # Mark all teachers busy
                    for t in teachers:
                        teacher_busy.add((t["teacher"], day, p))
                        teacher_day_load[(t["teacher"], day)] += 1

            mb_day_chosen[bid] = day
            break   # placed all pairs on this day

    # ─────────────────────────────────────────────────────────
    # STEP 3: Build normal task list
    # ─────────────────────────────────────────────────────────
    tasks     = []
    remaining = {}
    merge_constraints = {}
    same_day_map      = {}
    assembly_skeys    = set()

    for row in subject_rows:
        if row.get("type") in ("religion_block","merge_block"):
            continue  # already placed
        tch  = row["teacher"]
        subj = row["subject"]
        room = row.get("room", "R??")
        pc   = row.get("per_class", {})
        mgs  = row.get("merge_groups", [])
        sd   = row.get("same_day", False)
        a1   = row.get("all_1period", False)
        skey = f"{tch}|{subj}"

        merge_constraints[skey] = mgs
        same_day_map[skey]      = sd or a1

        if _is_assembly(subj):
            assembly_skeys.add(skey)
            for cls in class_list:
                remaining[(cls, skey)] = 0
            continue

        if a1:
            for cls in class_list:
                tasks.append((cls, skey, tch, subj, room, 1))
                remaining[(cls, skey)] = 1
        else:
            for cls, pw in pc.items():
                if pw > 0 and cls in class_list:
                    tasks.append((cls, skey, tch, subj, room, pw))
                    remaining[(cls, skey)] = pw

    # Priority tiers:
    # 0 = core academic (Mathematics, English, Science, Sinhala, ICT, History)
    # 1 = semi-core (Buddhism/PE, Assembly, Effective Speech, PE, Drama)
    # 2 = merge/aesthetic (Cookery, Art, Music, Dancing, Chinese, French, Civic, Lifeskill)
    # Within each tier: sort by total periods desc (hardest to place first)
    TIER0 = {"mathematics","english","english sdr","science","science/rel",
             "science/budd","sinhala","sinhala/tamil","ict","ict/computing",
             "history","history/geo"}
    TIER1 = {"buddhism","buddhism/pe","assembly","effective speech","pe",
             "speech and drama","religion"}

    def _task_pri(t):
        sl = t[3].lower()
        if any(c in sl for c in TIER0): tier = 0
        elif any(c in sl for c in TIER1): tier = 1
        else: tier = 2
        return (tier, -t[5])
    tasks.sort(key=_task_pri)

    def get_merge_partners(cls, skey):
        for mg in merge_constraints.get(skey, []):
            if cls in mg:
                return [c for c in mg if c != cls]
        return []

    same_day_chosen = {}

    # ─────────────────────────────────────────────────────────
    # STEP 3.5: Pre-place MATHEMATICS as double + single periods
    #
    # Each grade now has only 3 Maths teachers, so a single teacher may carry
    # 2-3 classes (e.g. Grade 6 = 9 periods x 3 classes = 27 periods). To make
    # that fit cleanly - and because Maths is taught best in blocks - every
    # class's weekly Maths is spread evenly across the 5 days, and any day that
    # carries 2 periods is placed as a genuine back-to-back DOUBLE period
    # (clock-contiguous); days with 1 period are SINGLES.
    #   9/week -> [2,2,2,2,1] = 4 doubles + 1 single
    #   7/week -> [2,2,1,1,1] = 2 doubles + 3 singles
    # Triples are never created. Anything not pre-placed here falls through to
    # STEP 4 and is mopped up - so this only ever improves layout.
    # ─────────────────────────────────────────────────────────
    def _free(cls_, d_, p_):
        return (p_ in slots[cls_][d_] and slots[cls_][d_][p_] is None
                and not (d_ == ASSEMBLY_DAY and p_ == ASSEMBLY_PERIOD))

    def _place_one(cls_, d_, p_, tch_, subj_, room_, skey_):
        slots[cls_][d_][p_] = (subj_, tch_, room_)
        teacher_busy.add((tch_, d_, p_))
        class_day_load[(cls_, d_)]        += 1
        teacher_day_load[(tch_, d_)]      += 1
        subj_day_placed[(cls_, skey_, d_)] += 1
        remaining[(cls_, skey_)]          -= 1

    def _try_double(cls_, d_, tch_, subj_, room_, skey_):
        """Place a clock-contiguous double on day d_. Returns True on success."""
        if class_day_load[(cls_, d_)] > MAX_PER_DAY - 2:
            return False
        pair_order = [tuple(sorted(fp)) for fp in CONTIGUOUS_PAIRS]
        if attempt:
            _rng.shuffle(pair_order)
        else:
            pair_order.sort()
        for a, b in pair_order:
            if not (_free(cls_, d_, a) and _free(cls_, d_, b)):
                continue
            if (tch_, d_, a) in teacher_busy or (tch_, d_, b) in teacher_busy:
                continue
            _place_one(cls_, d_, a, tch_, subj_, room_, skey_)
            _place_one(cls_, d_, b, tch_, subj_, room_, skey_)
            return True
        return False

    def _try_single(cls_, d_, tch_, subj_, room_, skey_):
        if class_day_load[(cls_, d_)] >= MAX_PER_DAY:
            return False
        for p_ in PERIOD_SNS:
            if not _free(cls_, d_, p_):
                continue
            if (tch_, d_, p_) in teacher_busy:
                continue
            _place_one(cls_, d_, p_, tch_, subj_, room_, skey_)
            return True
        return False

    maths_rows = [r for r in subject_rows
                  if r.get("type") not in ("religion_block", "merge_block")
                  and _is_maths(r.get("subject", ""))]

    def _maths_adjacency_ok(cls_, d_, p_, tch_, subj_):
        """True if placing this Maths period keeps same-day Maths clean: any
        same-day neighbour in the teaching-period list must be a GENUINE
        contiguous double (no break between), and never a triple. This forbids
        a 'double' made of one period before a break and one after it (e.g. the
        period before the interval + the period after it)."""
        pidx = PERIOD_SNS.index(p_)
        for off in (-1, 1):
            nb = pidx + off
            if not (0 <= nb < len(PERIOD_SNS)):
                continue
            nb_p = PERIOD_SNS[nb]
            nb_e = slots[cls_][d_][nb_p]
            if not (nb_e and nb_e[1] == tch_ and nb_e[0] == subj_):
                continue
            if frozenset((p_, nb_p)) not in CONTIGUOUS_PAIRS:
                return False                      # break-straddling pair
            far = nb + off
            if 0 <= far < len(PERIOD_SNS):
                far_p = PERIOD_SNS[far]
                far_e = slots[cls_][d_][far_p]
                if (far_e and far_e[1] == tch_ and far_e[0] == subj_
                        and frozenset((nb_p, far_p)) in CONTIGUOUS_PAIRS):
                    return False                  # would make a triple
        return True


    for row in maths_rows:
        tch  = row["teacher"]
        subj = row["subject"]
        room = row.get("room", "R-MAT")
        skey = f"{tch}|{subj}"
        pc   = row.get("per_class", {})
        for cls, n in pc.items():
            if cls not in class_list or n <= 0:
                continue
            base, extra = divmod(n, 5)
            day_counts = sorted(
                [base + (1 if i < extra else 0) for i in range(5)], reverse=True)
            used_days = set()
            for cnt in day_counts:
                if cnt <= 0 or remaining.get((cls, skey), 0) <= 0:
                    continue
                cand = sorted(
                    [d for d in range(1, 6) if d not in used_days],
                    key=lambda d: (class_day_load[(cls, d)], teacher_day_load[(tch, d)], _jit()))
                done = False
                for d in cand:
                    if cnt >= 2:
                        if _try_double(cls, d, tch, subj, room, skey):
                            used_days.add(d); done = True; break
                    else:
                        if _try_single(cls, d, tch, subj, room, skey):
                            used_days.add(d); done = True; break
                if not done and cnt >= 2:
                    for d in cand:
                        if _try_single(cls, d, tch, subj, room, skey):
                            used_days.add(d); break

    # ─────────────────────────────────────────────────────────
    # STEP 4: Place all remaining subjects
    # ─────────────────────────────────────────────────────────
    for _pass in range(100000):
        placed_any = False

        for cls, skey, tch, subj, room, _pw in tasks:
            if remaining.get((cls, skey), 0) <= 0:
                continue

            partners = get_merge_partners(cls, skey)
            is_sd    = same_day_map.get(skey, False)

            if is_sd and skey in same_day_chosen:
                days_ord = [same_day_chosen[skey]]
            else:
                days_ord = sorted(range(1, 6), key=lambda d: (
                    subj_day_placed[(cls, skey, d)],
                    class_day_load[(cls, d)],
                    teacher_day_load[(tch, d)],
                    _jit(),
                ))

            for d in days_ord:
                if class_day_load[(cls, d)]   >= MAX_PER_DAY:  continue
                if teacher_day_load[(tch, d)] >= MAX_TCH_DAY:  continue
                # Core subjects: no per-day cap (can place up to 8/day if needed)
                # Merge/aesthetic subjects: max 2/day to spread across week
                subj_lower = subj.lower()
                is_core = any(c in subj_lower for c in {
                    "mathematics","english","science","sinhala","ict","history"})
                if not is_core and subj_day_placed[(cls, skey, d)] >= 2: continue

                for p in PERIOD_SNS:
                    if d == ASSEMBLY_DAY and p == ASSEMBLY_PERIOD:
                        continue
                    if slots[cls][d][p] is not None:
                        continue
                    if (tch, d, p) in teacher_busy:
                        continue

                    pidx = PERIOD_SNS.index(p)
                    if _is_maths(subj):
                        # Maths: allow a genuine contiguous double, but never a
                        # break-straddling pair and never a triple.
                        if not _maths_adjacency_ok(cls, d, p, tch, subj):
                            continue
                    else:
                        # Non-Maths singles stay spread out: never place the same
                        # subject in a teaching period adjacent to itself. Checking
                        # BOTH neighbours (not just the previous one) also stops an
                        # accidental 'before the break + after the break' pair from
                        # forming regardless of the order slots get filled.
                        adj_same = False
                        for off in (-1, 1):
                            nb = pidx + off
                            if 0 <= nb < len(PERIOD_SNS):
                                nb_e = slots[cls][d][PERIOD_SNS[nb]]
                                if nb_e and nb_e[1] == tch and nb_e[0] == subj:
                                    adj_same = True; break
                        if adj_same:
                            continue

                    partner_ok = True
                    for pc2 in partners:
                        if slots[pc2][d][p] is not None:
                            partner_ok = False; break
                        if class_day_load[(pc2, d)] >= MAX_PER_DAY:
                            partner_ok = False; break
                    if not partner_ok:
                        continue

                    # Place
                    slots[cls][d][p] = (subj, tch, room)
                    teacher_busy.add((tch, d, p))
                    class_day_load[(cls, d)]        += 1
                    teacher_day_load[(tch, d)]      += 1
                    subj_day_placed[(cls, skey, d)] += 1
                    remaining[(cls, skey)]          -= 1

                    for pc2 in partners:
                        slots[pc2][d][p] = (subj, tch, room)
                        class_day_load[(pc2, d)]        += 1
                        subj_day_placed[(pc2, skey, d)] += 1
                        remaining[(pc2, skey)] = max(0, remaining.get((pc2, skey), 0) - 1)

                    if is_sd and skey not in same_day_chosen:
                        same_day_chosen[skey] = d

                    placed_any = True
                    break
                else:
                    continue
                break

        if not placed_any:
            break

    # ── Final cleanup pass: try once more for any remaining core subjects ──────
    # This handles cases where greedy ordering left valid slots unused
    core_keywords = {"mathematics","english","science","sinhala","ict","history"}
    for (cls, skey), rem in list(remaining.items()):
        if rem <= 0: continue
        # Check if this is a core subject
        subj_check = skey.split("|")[1].lower() if "|" in skey else ""
        if not any(c in subj_check for c in core_keywords): continue
        tch = skey.split("|")[0]
        # Try every slot
        for d in range(1,6):
            if remaining[(cls,skey)] <= 0: break
            if class_day_load[(cls,d)] >= MAX_PER_DAY: continue
            if teacher_day_load[(tch,d)] >= MAX_TCH_DAY: continue
            for p in PERIOD_SNS:
                if remaining[(cls,skey)] <= 0: break
                if d == ASSEMBLY_DAY and p == ASSEMBLY_PERIOD: continue
                if slots[cls][d][p] is not None: continue
                if (tch,d,p) in teacher_busy: continue
                slots[cls][d][p] = (subj_check.capitalize(), tch, "R-???")
                teacher_busy.add((tch,d,p))
                class_day_load[(cls,d)] += 1
                teacher_day_load[(tch,d)] += 1
                subj_day_placed[(cls,skey,d)] += 1
                remaining[(cls,skey)] -= 1

    # ── Local-repair pass (bounded recursive eviction) ────────────────────────
    # On a 100%-full grid a core lesson can be stranded: the class has a free
    # slot, but its teacher is busy there teaching ANOTHER class, and that
    # teacher (e.g. a Science teacher carrying 3 classes) is loaded enough that
    # the blocking lesson has nowhere obvious to go. We free the slot by
    # relocating the blocking lesson, recursively shuffling the same teacher's
    # other single lessons if needed (depth-limited, cycle-safe). Purely
    # additive: it can only resolve conflicts, never create them.
    def _block_marker(e):
        k = str(e[1])
        return (k == "RELIGION_BLOCK" or k == "ASSEMBLY"
                or k.startswith("MERGE_") or k.startswith("AESTHETIC_"))

    def _evict(tch, d, p, depth, touched):
        """Make (tch,d,p) free by relocating the single lesson the teacher has there."""
        if (tch, d, p) not in teacher_busy:
            return True
        if depth <= 0 or (tch, d, p) in touched:
            return False
        touched = touched | {(tch, d, p)}
        bc = be = None
        for c2 in class_list:
            e = slots[c2][d][p]
            if e and not _block_marker(e) and e[1] == tch:
                bc, be = c2, e; break
        if bc is None:
            return False
        bsubj, broom = be[0], be[2]
        bskey = f"{tch}|{bsubj}"
        for d2 in range(1, 6):
            if class_day_load[(bc, d2)] >= MAX_PER_DAY:
                continue
            for p2 in PERIOD_SNS:
                if (d2, p2) == (d, p):                            continue
                if d2 == ASSEMBLY_DAY and p2 == ASSEMBLY_PERIOD:  continue
                if slots[bc][d2][p2] is not None:                 continue
                i2 = PERIOD_SNS.index(p2); bad = False
                for nb in (i2 - 1, i2 + 1):
                    if 0 <= nb < len(PERIOD_SNS):
                        e_nb = slots[bc][d2][PERIOD_SNS[nb]]
                        if e_nb and e_nb[1] == tch and e_nb[0] == bsubj:
                            bad = True; break
                if bad:
                    continue
                if (tch, d2, p2) in teacher_busy:
                    if not _evict(tch, d2, p2, depth - 1, touched):
                        continue
                slots[bc][d][p] = None
                teacher_busy.discard((tch, d, p))
                class_day_load[(bc, d)]        -= 1
                teacher_day_load[(tch, d)]     -= 1
                subj_day_placed[(bc, bskey, d)] -= 1
                slots[bc][d2][p2] = (bsubj, tch, broom)
                teacher_busy.add((tch, d2, p2))
                class_day_load[(bc, d2)]        += 1
                teacher_day_load[(tch, d2)]     += 1
                subj_day_placed[(bc, bskey, d2)] += 1
                return True
        return False

    for (cls, skey), rem in list(remaining.items()):
        if rem <= 0: continue
        subj_chk = skey.split("|")[1] if "|" in skey else ""
        if not any(c in subj_chk.lower() for c in core_keywords): continue
        tch = skey.split("|")[0]
        room = "R-???"
        for c0 in class_list:
            hit = False
            for d0 in range(1, 6):
                for p0 in PERIOD_SNS:
                    e0 = slots[c0][d0][p0]
                    if e0 and e0[1] == tch and e0[0] == subj_chk:
                        room = e0[2]; hit = True; break
                if hit: break
            if hit: break
        guard = 0
        while remaining[(cls, skey)] > 0 and guard < 60:
            guard += 1
            placed = False
            for d in range(1, 6):
                if class_day_load[(cls, d)] >= MAX_PER_DAY: continue
                for p in PERIOD_SNS:
                    if d == ASSEMBLY_DAY and p == ASSEMBLY_PERIOD: continue
                    if slots[cls][d][p] is not None: continue
                    if not _evict(tch, d, p, 8, set()): continue
                    slots[cls][d][p] = (subj_chk, tch, room)
                    teacher_busy.add((tch, d, p))
                    class_day_load[(cls, d)]   += 1
                    teacher_day_load[(tch, d)] += 1
                    subj_day_placed[(cls, skey, d)] += 1
                    remaining[(cls, skey)]     -= 1
                    placed = True; break
                if placed: break
            if not placed:
                break

    # ── Maths compaction ──────────────────────────────────────────────────────
    # Pull any "split" same-day Maths (two periods on one day that are NOT a
    # genuine back-to-back double) into a real contiguous double, so doubles sit
    # on adjacent periods as much as possible. Tries a free contiguous slot first,
    # then a same-class/same-day swap with a non-Maths lesson. Never creates a
    # conflict, a triple, or a break-straddling pair.
    def _teacher_free(t, d, p):
        return (t, d, p) not in teacher_busy

    for row in maths_rows:
        tch = row["teacher"]; subj = row["subject"]
        for cls in [c for c in row.get("per_class", {}) if c in class_list]:
            for d in range(1, 6):
                ms = [p for p in PERIOD_SNS
                      if slots[cls][d][p] and slots[cls][d][p][1] == tch
                      and slots[cls][d][p][0] == subj]
                if len(ms) != 2:
                    continue
                a, b = ms
                if frozenset((a, b)) in CONTIGUOUS_PAIRS:
                    continue                      # already a clean double
                room = slots[cls][d][a][2]
                fixed = False
                # try: relocate one period to a free slot contiguous to the other
                for keep, move in ((a, b), (b, a)):
                    ki = PERIOD_SNS.index(keep)
                    for off in (-1, 1):
                        ni = ki + off
                        if not (0 <= ni < len(PERIOD_SNS)):
                            continue
                        tgt = PERIOD_SNS[ni]
                        if tgt == move or frozenset((keep, tgt)) not in CONTIGUOUS_PAIRS:
                            continue
                        if d == ASSEMBLY_DAY and tgt == ASSEMBLY_PERIOD:
                            continue
                        if slots[cls][d][tgt] is None and _teacher_free(tch, d, tgt):
                            slots[cls][d][move] = None; teacher_busy.discard((tch, d, move))
                            slots[cls][d][tgt] = (subj, tch, room); teacher_busy.add((tch, d, tgt))
                            fixed = True; break
                    if fixed:
                        break
                if fixed:
                    continue
                # try: swap with a non-Maths single sitting in a slot contiguous to one period
                for keep, move in ((a, b), (b, a)):
                    ki = PERIOD_SNS.index(keep)
                    for off in (-1, 1):
                        ni = ki + off
                        if not (0 <= ni < len(PERIOD_SNS)):
                            continue
                        tgt = PERIOD_SNS[ni]
                        if tgt == move or frozenset((keep, tgt)) not in CONTIGUOUS_PAIRS:
                            continue
                        e = slots[cls][d][tgt]
                        if not e or _block_marker(e) or _is_maths(e[0]):
                            continue
                        xs, xt, xr = e[0], e[1], e[2]
                        # X must be able to move to 'move' slot: teacher free there
                        # (the 'move' slot's current occupant is the Maths teacher,
                        # being vacated) and no same-subject neighbour for X there.
                        if (xt, d, move) in teacher_busy and xt != tch:
                            continue
                        mi = PERIOD_SNS.index(move); xbad = False
                        for o2 in (-1, 1):
                            j = mi + o2
                            if 0 <= j < len(PERIOD_SNS):
                                en = slots[cls][d][PERIOD_SNS[j]]
                                if en and en[1] == xt and en[0] == xs and PERIOD_SNS[j] != tgt:
                                    xbad = True; break
                        if xbad:
                            continue
                        if not _teacher_free(tch, d, tgt) and tch != xt:
                            continue
                        # perform swap: Maths 'move' -> tgt ; X tgt -> 'move'
                        teacher_busy.discard((tch, d, move)); teacher_busy.discard((xt, d, tgt))
                        slots[cls][d][tgt] = (subj, tch, room)
                        slots[cls][d][move] = (xs, xt, xr)
                        teacher_busy.add((tch, d, tgt)); teacher_busy.add((xt, d, move))
                        fixed = True; break
                    if fixed:
                        break

    conflicts = {k: v for k, v in remaining.items() if v > 0}

    # Update shared teacher busy set — only add CROSS-GRADE shared teachers
    # (teachers who work across multiple grades e.g. Sakna, Pulakshika, RC teacher)
    # Do NOT add grade-specific teachers (e.g. G7_Maths_T1) as they only teach 1 grade
    if shared_teacher_busy is not None:
        for (tch, d, p) in teacher_busy:
            # Include all teachers — the shared_busy dict is keyed by teacher name
            # Grade-specific teachers (G7_Maths_T1) won't appear in other grades anyway
            # so adding them is harmless. But to be safe, only add teachers who
            # don't start with a grade prefix (G6_, G7_, G8_)
            if not (tch.startswith("G6_") or tch.startswith("G7_") or tch.startswith("G8_")):
                shared_teacher_busy.add((tch, d, p))

    # ─────────────────────────────────────────────────────────
    # STEP 5: Double-booking check
    # ─────────────────────────────────────────────────────────
    tc = Counter()
    for cls in class_list:
        for d in range(1, 6):
            for p in PERIOD_SNS:
                e = slots[cls][d][p]
                if e:
                    tch_key = e[1]
                    # Skip block markers — these legitimately appear in multiple class slots
                    if tch_key in ("RELIGION_BLOCK",) or                        str(tch_key).startswith("MERGE_") or                        str(tch_key).startswith("AESTHETIC_"):
                        continue
                    tc[(tch_key, d, p)] += 1

    shared_ok = set()
    # Merge groups
    for row in subject_rows:
        if row.get("type") == "religion_block":
            continue
        for mg in row.get("merge_groups", []):
            tch = row["teacher"]
            for d in range(1, 6):
                for p in PERIOD_SNS:
                    entries = [slots[c][d][p] for c in mg if c in class_list]
                    if entries and all(e == entries[0] for e in entries) and entries[0]:
                        shared_ok.add((tch, d, p))
    # Assembly whole-school
    for row in subject_rows:
        if row.get("type") == "religion_block":
            continue
        if _is_assembly(row.get("subject", "")):
            shared_ok.add((row["teacher"], ASSEMBLY_DAY, ASSEMBLY_PERIOD))
    # Merge blocks — all teachers share same slots per pair
    for mb in merge_blocks:
        for cls in class_list:
            for d in range(1, 6):
                for p in PERIOD_SNS:
                    e = slots[cls][d][p]
                    if e and "MERGE_" in str(e[1]):
                        for t in mb.get("teachers",[]):
                            shared_ok.add((t["teacher"], d, p))
        # Religion block — all religion teachers legitimately share one slot
    for rb in religion_blocks:
        d = rb.get("_placed_day")
        p = rb.get("_placed_period")
        if d and p:
            for t in rb.get("teachers", []):
                shared_ok.add((t["teacher"], d, p))

    # RELIGION_BLOCK is a pseudo-teacher marker — never a real double-booking
    double_booked = {k: v for k, v in tc.items()
                     if v > 1 and k not in shared_ok and k[0] != "RELIGION_BLOCK"}
    return slots, conflicts, double_booked


def schedule_all(class_subjects_map):
    """Legacy scheduler (Gr 9/10/11-12)."""
    all_classes = list(class_subjects_map.keys())
    slots = {cls: {d: {p: None for p in PERIOD_SNS} for d in range(1, 6)}
             for cls in all_classes}

    teacher_busy     = set()
    class_day_load   = defaultdict(int)
    teacher_day_load = defaultdict(int)
    subj_day_placed  = defaultdict(int)

    MAX_PER_DAY     = len(PERIOD_SNS)
    MAX_TEACHER_DAY = 9

    tasks = []
    for cls, subjs in class_subjects_map.items():
        for code, name, tid, pw, room, *_ in subjs:
            if pw > 0:
                tasks.append((cls, code, name, pw, room, tid))
    tasks.sort(key=lambda x: -x[3])
    remaining = {(t[0], t[1]): t[3] for t in tasks}

    # Pre-place Assembly
    for cls, code, name, pw, room, tid in tasks:
        if not (_is_assembly(name) or _is_assembly(code)):
            continue
        d, p = ASSEMBLY_DAY, ASSEMBLY_PERIOD
        if slots[cls][d][p] is not None:
            continue
        slots[cls][d][p] = (code, name, tid, room)
        teacher_busy.add((tid, d, p))
        class_day_load[(cls, d)]        += 1
        teacher_day_load[(tid, d)]      += 1
        subj_day_placed[(cls, code, d)] += 1
        remaining[(cls, code)]          -= 1

    for _ in range(25000):
        placed_any = False
        for cls, code, name, pw, room, tid in tasks:
            if remaining.get((cls, code), 0) <= 0:
                continue
            days_ord = sorted(range(1, 6), key=lambda d: (
                subj_day_placed[(cls, code, d)],
                class_day_load[(cls, d)],
                teacher_day_load[(tid, d)],
            ))
            for d in days_ord:
                if class_day_load[(cls, d)]   >= MAX_PER_DAY:    continue
                if teacher_day_load[(tid, d)] >= MAX_TEACHER_DAY: continue
                if subj_day_placed[(cls, code, d)] >= 3:          continue
                for p in PERIOD_SNS:
                    if d == ASSEMBLY_DAY and p == ASSEMBLY_PERIOD: continue
                    if slots[cls][d][p] is not None:      continue
                    if (tid, d, p) in teacher_busy:        continue
                    pidx = PERIOD_SNS.index(p)
                    if pidx > 0:
                        prev = PERIOD_SNS[pidx - 1]
                        if slots[cls][d][prev] and slots[cls][d][prev][0] == code:
                            continue
                    slots[cls][d][p] = (code, name, tid, room)
                    teacher_busy.add((tid, d, p))
                    class_day_load[(cls, d)]        += 1
                    teacher_day_load[(tid, d)]      += 1
                    subj_day_placed[(cls, code, d)] += 1
                    remaining[(cls, code)]          -= 1
                    placed_any = True
                    break
                else:
                    continue
                break
        if not placed_any:
            break

    conflicts     = {k: v for k, v in remaining.items() if v > 0}
    tc            = Counter()
    for cls in all_classes:
        for d in range(1, 6):
            for p in PERIOD_SNS:
                e = slots[cls][d][p]
                if e: tc[(e[2], d, p)] += 1
    double_booked = {k: v for k, v in tc.items() if v > 1}
    return slots, conflicts, double_booked
