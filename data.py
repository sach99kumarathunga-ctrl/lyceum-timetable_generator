"""
data.py — Lyceum International School
Each class: 9 Maths, 7 English, 6 Science, 3 ICT, 3 History, 2 Sinhala,
            2 Merge Block 1, 2 Merge Block 2, 1 Religion, 1 Eff Speech,
            1 Speech & Drama, 1 PE, 1 Assembly = 39 periods
Available: 40 total – 1 Assembly = 39. Zero slack — teacher loads maximised at 2 classes.

Teacher load rules (to avoid scheduling conflicts):
  Maths:   max 2 classes × 9 = 18 periods/teacher
  English: max 2 classes × 7 = 14 periods/teacher  
  Science: max 2 classes × 6 = 12 periods/teacher
  History: 1 teacher × all 7 classes × 3 = 21 (spread across days, manageable)
"""

GRADE_GROUPS = []

# ── Shared teachers across ALL grades ───────────────────────────────────────
MERGE1 = [
    {"teacher":"Sakna",  "subject":"Cookery",  "room":"R-CKR"},
    {"teacher":"Prasad", "subject":"Lifeskill","room":"R-LS"},
    {"teacher":"Nimesha","subject":"Civic",    "room":"R-CIV"},
]
MERGE2 = [
    {"teacher":"Pulakshika",        "subject":"Art",           "room":"R-ART"},
    {"teacher":"Shehan",            "subject":"Eastern Music", "room":"R-MUS"},
    {"teacher":"Savini",            "subject":"Western Music", "room":"R-MUS"},
    {"teacher":"dance1",            "subject":"Dancing",       "room":"R-DAN"},
    {"teacher":"Dulshani",          "subject":"Chinese",       "room":"R-CHN"},
    {"teacher":"New Teacher_french","subject":"French",        "room":"R-FRN"},
]
OTHER_REL = [
    {"teacher":"New Teacher_Roman catholic","subject":"Roman Catholic","room":"R-REL-H","covers":"All classes"},
    {"teacher":"New Teacher2_Hindu",        "subject":"Hindu",         "room":"R-REL-I","covers":"All classes"},
    {"teacher":"New Teacher3_Christianity", "subject":"Christianity",  "room":"R-REL-J","covers":"All classes"},
    {"teacher":"New Teacher4_Jabir",        "subject":"Islam",         "room":"R-REL-K","covers":"All classes"},
]
EFF1 = "New Teacher_Eff Speech1"
EFF2 = "New Teacher_Eff Speech2"
EFF3 = "New Teacher_Eff Speech3"


# ═══════════════════════════════════════════════════════════════════════════════
# GRADE 6  — Real data, 9 Maths per class, balanced teacher loads
# ═══════════════════════════════════════════════════════════════════════════════
GRADE_GROUPS.append({
    "name":       "Grade 6",
    "file":       "Grade6_AllClasses_Timetable.xlsx",
    "setup_mode": "flexible",
    "classes":    ["6A","6B","6C","6D","6E","6F","6G"],
    "teachers": [
        # Maths — ONLY 3 teachers for the whole grade (rename these once known)
        ("T01","Grade 6 Maths Teacher 1", "Mathematics"),
        ("T02","Grade 6 Maths Teacher 2", "Mathematics"),
        ("T03","Grade 6 Maths Teacher 3", "Mathematics"),
        # ICT
        ("T05","Ms. Prasadi ICT",           "ICT/Computing"),
        # English (4 teachers, max 2 classes × 7 = 14 each)
        ("T06","Ms. Gimhani Hettiarachchi", "English"),
        ("T07","Ms. Maheli Kulathunga",     "English"),
        ("T08","Ms. Nimnada Jayathilake",   "English SDR"),
        ("T09","Ms. Shermila Jayawickrama", "English SDR"),
        # Sinhala/Tamil
        ("T10","Dinusha Nadeeshani",        "Sinhala/Tamil"),
        # Science (3 teachers, max 2 classes × 6 = 12 each)
        ("T11","Ms. Hasini Rathnayaka",     "Science"),
        ("T12","Ms. Sankalpi",             "Science"),
        ("T13","Ms. Ruwanga",              "Science"),
        # History
        ("T14","Bagya",                    "History/Geo"),
        # Individual
        ("T15",EFF1,                        "Effective Speech"),
        ("T16",EFF2,                        "Speech and Drama"),
        ("T17",EFF3,                        "PE"),
        # Religion
        ("T18","New Teacher_Roman catholic","R.Catholicism"),
        ("T19","New Teacher2_Hindu",        "Hindu"),
        ("T20","New Teacher3_Christianity", "Christianity"),
        ("T21","New Teacher4_Jabir",        "Islam"),
        # Merge
        ("T22","Sakna",  "Cookery"),
        ("T23","Prasad", "Lifeskill"),
        ("T24","Nimesha","Civic"),
        ("T25","Pulakshika","Art"),
        ("T26","Shehan","Eastern Music"),
        ("T27","Savini","Western Music"),
        ("T28","dance1","Dancing"),
        ("T29","Dulshani","Chinese"),
        ("T30","New Teacher_french","French"),
    ],
    "subject_rows": [
        # Maths: 3 teachers cover all 7 classes (9 periods/class).
        #   T1 -> 6A,6B,6C (27p) | T2 -> 6D,6E (18p) | T3 -> 6F,6G (18p)
        # Scheduler lays each class out as double + single periods (9 -> 4 doubles + 1 single).
        ("Grade 6 Maths Teacher 1","Mathematics",{"6A":9,"6B":9,"6C":9}, [],False,False,"R-MAT"),
        ("Grade 6 Maths Teacher 2","Mathematics",{"6D":9,"6E":9},        [],False,False,"R-MAT"),
        ("Grade 6 Maths Teacher 3","Mathematics",{"6F":9,"6G":9},        [],False,False,"R-MAT"),
        # ── ICT: 3 periods/class ───────────────────────────────────────────────
        ("Ms. Prasadi ICT","ICT/Computing",{"6A":3,"6B":3,"6C":3,"6D":3,"6E":3,"6F":3,"6G":3},
                                                                                [],False,False,"R-ICT"),
        # ── English: 7 periods/class, 2 classes per teacher max ───────────────
        ("Ms. Gimhani Hettiarachchi","English",   {"6C":7,"6G":7},             [],False,False,"R-ENG"),
        ("Ms. Maheli Kulathunga",    "English",   {"6D":7},                    [],False,False,"R-ENG"),
        ("Ms. Nimnada Jayathilake",  "English SDR",{"6A":6,"6F":6},            [],False,False,"R-ENG"),
        ("Ms. Shermila Jayawickrama","English SDR",{"6B":6,"6E":6},            [],False,False,"R-ENG"),
        # ── Sinhala/Tamil: 2 periods/class ────────────────────────────────────
        ("Dinusha Nadeeshani","Sinhala/Tamil",{"6A":2,"6B":2,"6C":2,"6D":2,"6E":2,"6F":2,"6G":2},
                                                                                [],False,False,"R-SIN"),
        # ── Science: 6 periods/class, 2 classes per teacher max ───────────────
        ("Ms. Hasini Rathnayaka","Science",{"6A":6,"6C":6},                    [],False,False,"R-SCI"),
        ("Ms. Sankalpi",         "Science",{"6B":6,"6D":6},                    [],False,False,"R-SCI"),
        ("Ms. Nawoda Dilanka",          "Science",{"6E":6,"6F":6,"6G":6},             [],False,False,"R-SCI"),
        # ── History/Geo: 3 periods/class ──────────────────────────────────────
        ("Bagya","History/Geo",{"6A":3,"6B":3,"6C":3,"6D":3,"6E":3,"6F":3,"6G":3},
                                                                                [],False,False,"R-HIS"),
        # ── Individual subjects (no merge, no group) ───────────────────────────
        (EFF1,"Effective Speech",{"6A":1,"6B":1,"6C":1,"6D":1,"6E":1,"6F":1,"6G":1},
                                                                                [],False,False,"R-ENG"),
        (EFF2,"Speech and Drama", {"6A":1,"6B":1,"6C":1,"6D":1,"6E":1,"6F":1,"6G":1},
                                                                                [],False,False,"R-ENG"),
        (EFF3,"PE",               {"6A":1,"6B":1,"6C":1,"6D":1,"6E":1,"6F":1,"6G":1},
                                                                                [],False,False,"R-PE"),
        # ── Religion Block ─────────────────────────────────────────────────────
        # Buddhism: 5 named teachers + 2 Maths teachers (as an ADDITIONAL subject,
        # each teaching Buddhism to a class they already teach Maths to).
        # Other religions: all classes merged (1 period, same slot)
        {
            "type":"religion_block","grade_label":"Grade 6",
            "all_classes":["6A","6B","6C","6D","6E","6F","6G"],
            "teachers":[
                {"teacher":"Ms. Gimhani Hettiarachchi","subject":"Buddhism","room":"R-REL-A","covers":"6A"},
                {"teacher":"Ms. Nimnada Jayathilake",  "subject":"Buddhism","room":"R-REL-B","covers":"6B"},
                {"teacher":"Ms. Hasini Rathnayaka",    "subject":"Buddhism","room":"R-REL-C","covers":"6C"},
                {"teacher":"Grade 6 Maths Teacher 2",  "subject":"Buddhism","room":"R-REL-D","covers":"6D"},
                {"teacher":"Ms. Sankalpi",             "subject":"Buddhism","room":"R-REL-E","covers":"6E"},
                {"teacher":"Grade 6 Maths Teacher 3",  "subject":"Buddhism","room":"R-REL-F","covers":"6F"},
                {"teacher":"Bagya",                    "subject":"Buddhism","room":"R-REL-G","covers":"6G"},
            ] + OTHER_REL
        },
        # ── Merge Block 1: Cookery+Lifeskill+Civic ─────────────────────────────
        {
            "type":"merge_block","block_id":"block1_6","periods":2,"same_day":True,
            "merge_pairs":[["6A","6B"],["6C","6D"],["6E","6F"],["6G"]],
            "standalone":{},"teachers":MERGE1,
        },
        # ── Merge Block 2: Art+Music+Dance+Chinese+French ──────────────────────
        {
            "type":"merge_block","block_id":"block2_6","periods":2,"same_day":True,
            "merge_pairs":[["6A","6B"],["6C","6D"],["6E","6F","6G"]],
            "standalone":{},"teachers":MERGE2,
        },
    ],
})


# ═══════════════════════════════════════════════════════════════════════════════
# GRADE 7  — Identical structure. 9 Maths. Same shared teachers.
# ═══════════════════════════════════════════════════════════════════════════════
GRADE_GROUPS.append({
    "name":       "Grade 7",
    "file":       "Grade7_AllClasses_Timetable.xlsx",
    "setup_mode": "flexible",
    "classes":    ["7A","7B","7C","7D","7E","7F"],
    "teachers": [
        ("T01","Grade 7 Maths Teacher 1",   "Mathematics"),
        ("T02","Grade 7 Maths Teacher 2",   "Mathematics"),
        ("T03","Grade 7 Maths Teacher 3",   "Mathematics"),
        ("T04","G7_ICT_T1",     "ICT/Computing"),
        ("T05","G7_English_T1", "English"),
        ("T06","G7_English_T2", "English SDR"),
        ("T07","G7_English_T3", "English SDR"),
        ("T08","G7_Sinhala_T1", "Sinhala/Tamil"),
        ("T09","G7_Science_T1", "Science"),
        ("T10","G7_Science_T2", "Science"),
        ("T11","G7_Science_T3", "Science"),
        ("T12","G7_History_T1", "History/Geo"),
        ("T13",EFF1,"Effective Speech"),
        ("T14",EFF2,"Speech and Drama"),
        ("T15",EFF3,"PE"),
        ("T16","New Teacher_Roman catholic","R.Catholicism"),
        ("T17","New Teacher2_Hindu","Hindu"),
        ("T18","New Teacher3_Christianity","Christianity"),
        ("T19","New Teacher4_Jabir","Islam"),
        ("T20","Sakna","Cookery"),("T21","Prasad","Lifeskill"),("T22","Nimesha","Civic"),
        ("T23","Pulakshika","Art"),("T24","Shehan","Eastern Music"),("T25","Savini","Western Music"),
        ("T26","dance1","Dancing"),("T27","Dulshani","Chinese"),("T28","New Teacher_french","French"),
    ],
    "subject_rows": [
        # Maths: 3 teachers cover all 6 classes, 7 periods/class (double + single).
        #   T1 -> 7A,7B | T2 -> 7C,7D | T3 -> 7E,7F  (14 periods each)
        ("Grade 7 Maths Teacher 1","Mathematics",  {"7A":7,"7B":7},     [],False,False,"R-MAT"),
        ("Grade 7 Maths Teacher 2","Mathematics",  {"7C":7,"7D":7},     [],False,False,"R-MAT"),
        ("Grade 7 Maths Teacher 3","Mathematics",  {"7E":7,"7F":7},     [],False,False,"R-MAT"),
        ("G7_ICT_T1",  "ICT/Computing",{"7A":3,"7B":3},                  [],False,False,"R-ICT"),
        ("G7_ICT_T2",  "ICT/Computing",{"7C":3,"7D":3},                  [],False,False,"R-ICT"),
        ("G7_ICT_T3",  "ICT/Computing",{"7E":3,"7F":3},                  [],False,False,"R-ICT"),
        ("G7_English_T1","English",    {"7A":7},                         [],False,False,"R-ENG"),
        ("G7_English_T2","English",    {"7B":7},                         [],False,False,"R-ENG"),
        ("G7_English_T3","English SDR",{"7C":7},                         [],False,False,"R-ENG"),
        ("G7_English_T4","English SDR",{"7D":7},                         [],False,False,"R-ENG"),
        ("G7_English_T5","English SDR",{"7E":7},                         [],False,False,"R-ENG"),
        ("G7_English_T6","English SDR",{"7F":7},                         [],False,False,"R-ENG"),
        ("G7_Sinhala_T1","Sinhala/Tamil",{"7A":2,"7B":2,"7C":2},          [],False,False,"R-SIN"),
        ("G7_Sinhala_T2","Sinhala/Tamil",{"7D":2,"7E":2,"7F":2},          [],False,False,"R-SIN"),
        ("G7_Science_T1","Science",    {"7A":6,"7B":6},                  [],False,False,"R-SCI"),
        ("G7_Science_T2","Science",    {"7C":6,"7D":6},                  [],False,False,"R-SCI"),
        ("G7_Science_T3","Science",    {"7E":6,"7F":6},                  [],False,False,"R-SCI"),
        ("G7_History_T1","History/Geo", {"7A":3,"7B":3},                  [],False,False,"R-HIS"),
        ("G7_History_T2","History/Geo", {"7C":3,"7D":3},                  [],False,False,"R-HIS"),
        ("G7_History_T3","History/Geo", {"7E":3,"7F":3},                  [],False,False,"R-HIS"),
        (EFF1,"Effective Speech",{"7A":1,"7B":1,"7C":1,"7D":1,"7E":1,"7F":1},
                                                                          [],False,False,"R-ENG"),
        (EFF2,"Speech and Drama", {"7A":1,"7B":1,"7C":1,"7D":1,"7E":1,"7F":1},
                                                                          [],False,False,"R-ENG"),
        (EFF3,"PE",               {"7A":1,"7B":1,"7C":1,"7D":1,"7E":1,"7F":1},
                                                                          [],False,False,"R-PE"),
        {
            "type":"religion_block","grade_label":"Grade 7",
            "all_classes":["7A","7B","7C","7D","7E","7F"],
            "teachers":[
                {"teacher":"Grade 7 Maths Teacher 1", "subject":"Buddhism","room":"R-REL-A","covers":"7A"},
                {"teacher":"G7_Science_T1", "subject":"Buddhism","room":"R-REL-B","covers":"7B"},
                {"teacher":"Grade 7 Maths Teacher 2", "subject":"Buddhism","room":"R-REL-C","covers":"7C"},
                {"teacher":"G7_Science_T2", "subject":"Buddhism","room":"R-REL-D","covers":"7D"},
                {"teacher":"Grade 7 Maths Teacher 3", "subject":"Buddhism","room":"R-REL-E","covers":"7E"},
                {"teacher":"G7_History_T1", "subject":"Buddhism","room":"R-REL-F","covers":"7F"},
            ] + OTHER_REL
        },
        {"type":"merge_block","block_id":"block1_7","periods":2,"same_day":True,
         "merge_pairs":[["7A","7B"],["7C","7D"],["7E","7F"]],"standalone":{},"teachers":MERGE1},
        {"type":"merge_block","block_id":"block2_7","periods":2,"same_day":True,
         "merge_pairs":[["7A","7B"],["7C","7D"],["7E","7F"]],"standalone":{},"teachers":MERGE2},
    ],
})


# ═══════════════════════════════════════════════════════════════════════════════
# GRADE 8  — 9 Maths. Same shared teachers.
# ═══════════════════════════════════════════════════════════════════════════════
GRADE_GROUPS.append({
    "name":       "Grade 8",
    "file":       "Grade8_AllClasses_Timetable.xlsx",
    "setup_mode": "flexible",
    "classes":    ["8A","8B","8C","8D","8E"],
    "teachers": [
        ("T01","Grade 8 Maths Teacher 1",   "Mathematics"),
        ("T02","Grade 8 Maths Teacher 2",   "Mathematics"),
        ("T03","Grade 8 Maths Teacher 3",   "Mathematics"),
        ("T04","G8_ICT_T1",     "ICT/Computing"),
        ("T05","G8_English_T1", "English"),
        ("T06","G8_English_T2", "English"),
        ("T07","G8_English_T3", "English SDR"),
        ("T08","G8_Sinhala_T1", "Sinhala/Tamil"),
        ("T09","G8_Science_T1", "Science"),
        ("T10","G8_Science_T2", "Science"),
        ("T11","G8_Science_T3", "Science"),
        ("T12","G8_History_T1", "History/Geo"),
        ("T13",EFF1,"Effective Speech"),
        ("T14",EFF2,"Speech and Drama"),
        ("T15",EFF3,"PE"),
        ("T16","New Teacher_Roman catholic","R.Catholicism"),
        ("T17","New Teacher2_Hindu","Hindu"),
        ("T18","New Teacher3_Christianity","Christianity"),
        ("T19","New Teacher4_Jabir","Islam"),
        ("T20","Sakna","Cookery"),("T21","Prasad","Lifeskill"),("T22","Nimesha","Civic"),
        ("T23","Pulakshika","Art"),("T24","Shehan","Eastern Music"),("T25","Savini","Western Music"),
        ("T26","dance1","Dancing"),("T27","Dulshani","Chinese"),("T28","New Teacher_french","French"),
    ],
    "subject_rows": [
        # Maths: 3 teachers cover all 5 classes, 7 periods/class (double + single).
        #   T1 -> 8A,8B | T2 -> 8C,8D | T3 -> 8E  (14 / 14 / 7 periods)
        ("Grade 8 Maths Teacher 1","Mathematics",  {"8A":7,"8B":7},     [],False,False,"R-MAT"),
        ("Grade 8 Maths Teacher 2","Mathematics",  {"8C":7,"8D":7},     [],False,False,"R-MAT"),
        ("Grade 8 Maths Teacher 3","Mathematics",  {"8E":7},            [],False,False,"R-MAT"),
        ("G8_ICT_T1",  "ICT/Computing",{"8A":3,"8B":3},                  [],False,False,"R-ICT"),
        ("G8_ICT_T2",  "ICT/Computing",{"8C":3,"8D":3},                  [],False,False,"R-ICT"),
        ("G8_ICT_T3",  "ICT/Computing",{"8E":3},                         [],False,False,"R-ICT"),
        ("G8_English_T1","English",    {"8A":7},                         [],False,False,"R-ENG"),
        ("G8_English_T2","English",    {"8B":7},                         [],False,False,"R-ENG"),
        ("G8_English_T3","English",    {"8C":7},                         [],False,False,"R-ENG"),
        ("G8_English_T4","English SDR",{"8D":7},                         [],False,False,"R-ENG"),
        ("G8_English_T5","English SDR",{"8E":7},                         [],False,False,"R-ENG"),
        ("G8_Sinhala_T1","Sinhala/Tamil",{"8A":2,"8B":2,"8C":2},          [],False,False,"R-SIN"),
        ("G8_Sinhala_T2","Sinhala/Tamil",{"8D":2,"8E":2},                  [],False,False,"R-SIN"),
        ("G8_Science_T1","Science",    {"8A":6,"8B":6},                  [],False,False,"R-SCI"),
        ("G8_Science_T2","Science",    {"8C":6,"8D":6},                  [],False,False,"R-SCI"),
        ("G8_Science_T3","Science",    {"8E":6},                         [],False,False,"R-SCI"),
        ("G8_History_T1","History/Geo", {"8A":3,"8B":3},                  [],False,False,"R-HIS"),
        ("G8_History_T2","History/Geo", {"8C":3,"8D":3},                  [],False,False,"R-HIS"),
        ("G8_History_T3","History/Geo", {"8E":3},                         [],False,False,"R-HIS"),
        (EFF1,"Effective Speech",{"8A":1,"8B":1,"8C":1,"8D":1,"8E":1},  [],False,False,"R-ENG"),
        (EFF2,"Speech and Drama", {"8A":1,"8B":1,"8C":1,"8D":1,"8E":1}, [],False,False,"R-ENG"),
        (EFF3,"PE",               {"8A":1,"8B":1,"8C":1,"8D":1,"8E":1}, [],False,False,"R-PE"),
        {
            "type":"religion_block","grade_label":"Grade 8",
            "all_classes":["8A","8B","8C","8D","8E"],
            "teachers":[
                {"teacher":"Grade 8 Maths Teacher 1", "subject":"Buddhism","room":"R-REL-A","covers":"8A"},
                {"teacher":"G8_Science_T1", "subject":"Buddhism","room":"R-REL-B","covers":"8B"},
                {"teacher":"Grade 8 Maths Teacher 2", "subject":"Buddhism","room":"R-REL-B","covers":"8C"},
                {"teacher":"G8_Science_T2", "subject":"Buddhism","room":"R-REL-D","covers":"8D"},
                {"teacher":"Grade 8 Maths Teacher 3", "subject":"Buddhism","room":"R-REL-C","covers":"8E"},
            ] + OTHER_REL
        },
        {"type":"merge_block","block_id":"block1_8","periods":2,"same_day":True,
         "merge_pairs":[["8A","8B"],["8C","8D"],["8E"]],"standalone":{},"teachers":MERGE1},
        {"type":"merge_block","block_id":"block2_8","periods":2,"same_day":True,
         "merge_pairs":[["8A","8B"],["8C","8D","8E"]],"standalone":{},"teachers":MERGE2},
    ],
})


# ── Legacy grades (extracted from the official Setup sheets) ──
LEGACY_GRADES = [{'name': 'Grade 9',
  'file': 'Grade9_AllClasses_Timetable.xlsx',
  'setup_mode': 'legacy',
  'streams': ['PED', 'NAT'],
  'stream_colors': {'PED': '2E75B6', 'NAT': '548235'},
  'classes': {'9 Nat A': 'NAT', '9 Nat B': 'NAT', '9 PED Sci A': 'PED', '9 PED Sci B': 'PED', '9 PED Commerce': 'PED'},
  'teachers': [('ET01', 'Ms. Achini (Chemistry)', 'Chemistry', 'PED'),
               ('NT01', 'Ms. Vishwani', 'Mathematics', 'NAT'),
               ('ET02', 'Ms. Asha', 'Biology', 'PED'),
               ('NT02', 'Mr. Nalinda', 'Mathematics', 'NAT'),
               ('ET03', 'Mr. Radun Wickramage', 'Physics', 'PED'),
               ('NT03', 'Ms. Harshi', 'Science', 'NAT'),
               ('ET04', 'Ms. Gayani', 'Science/Chem', 'PED'),
               ('NT04', 'Mr. Clive', 'English', 'NAT'),
               ('ET05', 'Ms. Naslina', 'Science', 'PED'),
               ('NT05', 'Ms. Apeksha', 'English', 'NAT'),
               ('ET06', 'Mr. Amila', 'Mathematics', 'PED'),
               ('NT06', 'Ms. Sewwandi', 'Sinhala/Tamil', 'NAT'),
               ('ET07', 'Ms. Sachinifff', 'CS/ICT', 'PED'),
               ('NT07', 'Ms. Disna', 'Sinhala/Tamil', 'NAT'),
               ('ET08', 'Mr. Clive', 'English', 'PED'),
               ('NT08', 'Ms. Manoja', 'History/Geo', 'NAT'),
               ('ET09', 'Ms. Shermila', 'Eff. Speech', 'PED'),
               ('NT09', 'Ms. Virani ICT', 'ICT', 'NAT'),
               ('ET10', 'Ms. Apeksha', 'English', 'PED'),
               ('NT10', 'Ms. Jabir', 'Islam', 'NAT'),
               ('ET11', 'Ms. Sewwandi', 'Sinhala/Tamil', 'PED'),
               ('NT11', 'Ms. Hasini Rathnayaka', 'Buddhism', 'NAT'),
               ('ET12', 'Ms. Disna', 'Sinhala/Tamil', 'PED'),
               ('NT12', 'Ms. Jancey', 'R.Catholicism', 'NAT'),
               ('ET13', 'Ms. Sonali single awards', 'Science Award', 'PED'),
               ('ET14', 'Ms. Sanduni Pinnawala', 'Science', 'PED')],
  'subjects': {'PED': [('MAT', 'Mathematics', 'ET06', 7, 'R-MAT', 'PDF IOL: 7'),
                       ('ENG', 'English', 'ET10', 7, 'R-ENG', 'PDF IOL: 7'),
                       ('BIO', 'Biology', 'ET02', 4, 'R-BIO', 'PDF IOL: 4 | Sci A&B only'),
                       ('CHE', 'Chemistry', 'ET01', 4, 'R-CHE', 'PDF IOL: 4'),
                       ('PHY', 'Physics', 'ET03', 4, 'R-PHY', 'PDF IOL: 4 | Sci A&B only'),
                       ('CS', 'CS/ICT', 'ET07', 5, 'R-ICT', 'PDF IOL: 5 ✓'),
                       ('COM', 'Commerce/Lit', 'ET13', 3, 'R-COM', 'PDF IOL: 3 | Commerce only'),
                       ('SIN', 'Sinhala/Tamil', 'ET11', 2, 'R-SIN', 'PDF IOL: 2'),
                       ('SPE', 'Effective Speech', 'ET09', 1, 'R-ENG', 'PDF IOL: 1'),
                       ('PE', 'Physical Education', 'ET04', 2, 'R-PE', 'PDF IOL: 2'),
                       ('DAN', 'Cultural Dancing', 'ET04', 2, 'R-AES', 'PDF IOL: 2'),
                       ('ASM', 'Assembly', 'ET10', 1, 'R-HLL', 'PDF: 1')],
               'NAT': [('MAT', 'Mathematics', 'NT01', 7, 'R-MAT', 'PDF NOL: 7'),
                       ('SCI', 'Science', 'NT03', 5, 'R-SCI', 'PDF NOL: 5'),
                       ('ENG', 'English', 'NT04', 7, 'R-ENG', 'PDF NOL: 7'),
                       ('SIN', 'Sinhala/Tamil', 'NT06', 4, 'R-SIN', 'PDF NOL: 4'),
                       ('HIS', 'History', 'NT08', 2, 'R-HIS', 'PDF NOL: 2'),
                       ('COM', 'Commerce', 'NT05', 3, 'R-COM', 'PDF NOL: 3'),
                       ('ICT', 'ICT', 'NT09', 3, 'R-ICT', 'PDF NOL: 3'),
                       ('REL', 'Religion', 'NT11', 1, 'R-REL', 'PDF NOL: 1'),
                       ('PE', 'Physical Education', 'NT08', 2, 'R-PE', 'PDF NOL: 2'),
                       ('DAN', 'Cultural Dancing', 'NT08', 2, 'R-AES', 'PDF NOL: 2'),
                       ('ASM', 'Assembly', 'NT04', 1, 'R-HLL', 'PDF: 1')]}},
 {'name': 'Grade 10',
  'file': 'Grade10_AllClasses_Timetable.xlsx',
  'setup_mode': 'legacy',
  'streams': ['PED', 'NAT'],
  'stream_colors': {'PED': '2E75B6', 'NAT': '548235'},
  'classes': {'10 Nat': 'NAT', '10 Int X': 'NAT', '10 PED Sci A': 'PED', '10 PED Sci B': 'PED', '10 PED Com': 'PED'},
  'teachers': [('ET01', 'Ms. Achini (Chemistry)', 'Chemistry', 'PED'),
               ('NT01', 'Ms. Vishwani', 'Mathematics', 'NAT'),
               ('ET02', 'Ms. Vidya (Bio)', 'Biology', 'PED'),
               ('NT02', 'Ms. Naslina', 'Mathematics', 'NAT'),
               ('ET03', 'Ms. Asha', 'Biology', 'PED'),
               ('NT03', 'Mr. Amila', 'Mathematics', 'NAT'),
               ('ET04', 'Mr. Radun Wickramage', 'Physics', 'PED'),
               ('NT04', 'Ms. Harshi', 'Science', 'NAT'),
               ('ET05', 'Ms. Sanduni Premathilake', 'Science', 'PED'),
               ('NT05', 'Mr. Clive', 'English', 'NAT'),
               ('ET06', 'Ms. Gayani', 'Science/Chem', 'PED'),
               ('NT06', 'Ms. Sewwandi', 'Sinhala/Tamil', 'NAT'),
               ('ET07', 'Mr. Amila', 'Mathematics', 'PED'),
               ('NT07', 'Ms. Disna', 'Sinhala/Tamil', 'NAT'),
               ('ET08', 'Ms. Sachini', 'CS/ICT', 'PED'),
               ('NT08', 'Ms. Manoja', 'History/Geo', 'NAT'),
               ('ET09', 'Ms. Vindya', 'English/Lit', 'PED'),
               ('NT09', 'Ms. Virani ICT', 'ICT', 'NAT'),
               ('ET10', 'Ms. Shermila', 'Eff. Speech', 'PED'),
               ('NT10', 'Ms. Jabir', 'Islam', 'NAT'),
               ('ET11', 'Ms. Sewwandi', 'Sinhala/Tamil', 'PED'),
               ('NT11', 'Ms. Hasini Rathnayaka', 'Buddhism', 'NAT'),
               ('ET12', 'Ms. Sanduni Pinnawala', 'Science', 'PED'),
               ('NT12', 'Ms. Jancey', 'R.Catholicism', 'NAT'),
               ('ET13', 'Ms. Sonali single awards', 'Science Award', 'PED')],
  'subjects': {'PED': [('MAT', 'Mathematics', 'ET07', 7, 'R-MAT', 'PDF IOL: 7'),
                       ('ENG', 'English', 'ET09', 7, 'R-ENG', 'PDF IOL: 7'),
                       ('BIO', 'Biology', 'ET02', 4, 'R-BIO', 'PDF IOL: 4 | Sci A&B only'),
                       ('CHE', 'Chemistry', 'ET01', 4, 'R-CHE', 'PDF IOL: 4'),
                       ('PHY', 'Physics', 'ET04', 4, 'R-PHY', 'PDF IOL: 4 | Sci A&B only'),
                       ('CS', 'CS/ICT', 'ET08', 5, 'R-ICT', 'PDF IOL: 5 ✓'),
                       ('COM', 'Commerce/Lit', 'ET13', 3, 'R-COM', 'PDF IOL: 3 | Commerce only'),
                       ('SIN', 'Sinhala/Tamil', 'ET11', 2, 'R-SIN', 'PDF IOL: 2'),
                       ('SPE', 'Effective Speech', 'ET10', 1, 'R-ENG', 'PDF IOL: 1'),
                       ('PE', 'Physical Education', 'ET04', 2, 'R-PE', 'PDF IOL: 2'),
                       ('DAN', 'Cultural Dancing', 'ET04', 2, 'R-AES', 'PDF IOL: 2'),
                       ('ASM', 'Assembly', 'ET09', 1, 'R-HLL', 'PDF: 1')],
               'NAT': [('MAT', 'Mathematics', 'NT01', 7, 'R-MAT', 'PDF NOL: 7'),
                       ('SCI', 'Science', 'NT04', 5, 'R-SCI', 'PDF NOL: 5'),
                       ('ENG', 'English', 'NT05', 7, 'R-ENG', 'PDF NOL: 7'),
                       ('SIN', 'Sinhala/Tamil', 'NT06', 4, 'R-SIN', 'PDF NOL: 4'),
                       ('HIS', 'History', 'NT08', 2, 'R-HIS', 'PDF NOL: 2'),
                       ('COM', 'Commerce', 'NT06', 3, 'R-COM', 'PDF NOL: 3'),
                       ('ICT', 'ICT', 'NT09', 3, 'R-ICT', 'PDF NOL: 3'),
                       ('REL', 'Religion', 'NT11', 1, 'R-REL', 'PDF NOL: 1'),
                       ('PE', 'Physical Education', 'NT08', 2, 'R-PE', 'PDF NOL: 2'),
                       ('DAN', 'Cultural Dancing', 'NT08', 2, 'R-AES', 'PDF NOL: 2'),
                       ('ASM', 'Assembly', 'NT05', 1, 'R-HLL', 'PDF: 1')]}},
 {'name': 'Grade 11-12 AL',
  'file': 'Grade11_12_AllClasses_Timetable.xlsx',
  'setup_mode': 'legacy',
  'streams': ['PED', 'NAT'],
  'stream_colors': {'PED': '2E75B6', 'NAT': '548235'},
  'classes': {'11 Nat Sci': 'NAT',
              '11 PED Sci': 'PED',
              '11 PED Com': 'PED',
              '12 Nat Sci': 'NAT',
              '12 Nat Com': 'NAT',
              '12 PED Sci': 'PED',
              '12 PED Com': 'PED'},
  'teachers': [('ET01', 'Mr. Tharindu (Chem)', 'Chemistry', 'PED'),
               ('ET04', 'Ms. Charuni (Bio)', 'Biology', 'NAT'),
               ('ET02', 'Ms. Hasini Rathnathilaka', 'Chemistry', 'PED'),
               ('ET09', 'Ms. Nimeshika', 'Combined Maths', 'NAT'),
               ('ET03', 'Ms. Vidya (Bio)', 'Biology', 'PED'),
               ('ET10', 'Mr. Nipun', 'Combined Maths', 'NAT'),
               ('ET05', 'Ms. Imesha (Physics)', 'Physics', 'PED'),
               ('ET13', 'Ms. Rusiri', 'Biology/Sci', 'NAT'),
               ('ET06', 'Mr. Radun Wickramage', 'Physics', 'PED'),
               ('ET17', 'Ms. Apeksha', 'General English', 'NAT'),
               ('ET07', 'Ms. Kavindi', 'Physics', 'PED'),
               ('ET18', 'Ms. Thakshila', 'Combined Maths', 'NAT'),
               ('ET08', 'Mr. Chinthaka', 'Combined Maths', 'PED'),
               ('ET19', 'Mr. Jabir', 'Islam/Religion', 'NAT'),
               ('ET11', 'Ms. Vindya', 'General English', 'PED'),
               ('ET20', 'Ms. Hasini Rathnayaka', 'Buddhism', 'NAT'),
               ('ET12', 'Ms. Senani', 'Commerce', 'PED'),
               ('ET14', 'Mr. Sanjeewa', 'Physics/Sci', 'PED'),
               ('ET15', 'Ms. Hashini Dissanayaka', 'CS/ICT', 'PED'),
               ('ET16', 'Ms. Ashka', 'Economics', 'PED')],
  'subjects': {'PED': [('CHE', 'Chemistry', 'ET01', 5, 'R-CHE', 'PDF IAL: 5'),
                       ('BIO', 'Biology', 'ET03', 5, 'R-BIO', 'PDF IAL: 5'),
                       ('PHY', 'Physics', 'ET05', 5, 'R-PHY', 'PDF IAL: 5'),
                       ('CBM', 'Combined Maths', 'ET08', 7, 'R-MAT', 'PDF IAL: 7'),
                       ('ENG', 'General English', 'ET11', 3, 'R-ENG', 'PDF IAL: 3'),
                       ('COM', 'Commerce', 'ET12', 5, 'R-COM', 'PED Com only'),
                       ('CS', 'CS/ICT', 'ET15', 5, 'R-ICT', 'PDF IAL: 5 | PED Com only'),
                       ('ECO', 'Economics', 'ET16', 5, 'R-ECO', 'PED Com only'),
                       ('PE', 'Physical Education', 'ET14', 2, 'R-PE', 'PDF IAL: 2'),
                       ('DAN', 'Cultural Dancing', 'ET12', 2, 'R-AES', 'PDF IAL: 2'),
                       ('ASM', 'Assembly', 'ET11', 1, 'R-HLL', 'PDF: 1'),
                       ('MOT', 'Motivational Speech', 'ET11', 1, 'R-HLL', 'PDF: 1')],
               'NAT': [('CHE', 'Chemistry', 'ET01', 5, 'R-CHE', 'PDF NAL: 5 | Nat Sci only'),
                       ('BIO', 'Biology', 'ET04', 5, 'R-BIO', 'PDF NAL: 5 | Nat Sci only'),
                       ('PHY', 'Physics', 'ET06', 5, 'R-PHY', 'PDF NAL: 5 | Nat Sci only'),
                       ('CBM', 'Combined Maths', 'ET09', 7, 'R-MAT', 'PDF NAL: 7'),
                       ('ENG', 'General English', 'ET17', 3, 'R-ENG', 'PDF NAL: 3'),
                       ('COM', 'Commerce', 'ET12', 5, 'R-COM', 'Nat Com only'),
                       ('ACC', 'Accounting', 'ET08', 5, 'R-ACC', 'Nat Com only'),
                       ('BST', 'Business Studies', 'ET08', 5, 'R-BST', 'Nat Com only'),
                       ('PE', 'Physical Education', 'ET14', 2, 'R-PE', 'PDF NAL: 2'),
                       ('ASM', 'Assembly', 'ET17', 1, 'R-HLL', 'PDF: 1'),
                       ('REL', 'Religion', 'ET20', 1, 'R-REL', 'PDF NAL: 1'),
                       ('MOT', 'Motivational Speech', 'ET17', 1, 'R-HLL', 'PDF: 1')]}}]
GRADE_GROUPS.extend(LEGACY_GRADES)

