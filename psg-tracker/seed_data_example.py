"""Example (fictional) seed data for the G2 PSG Tracker.

Same shape as the real data, which lives in seed_data_private.py (git-ignored).
Names, phones and remarks here are made up; dates and duty mix mirror Oct/Nov 26.
"""
from datetime import date

D = date

# (section, name, rank, phone)
PERSONNEL = [
    ('ACE NCOIC', 'VACANT', '', ""),
    ('FUSION', 'ADAMS, ALEX', 'MAJ', ""),
    ('FUSION', 'ALLEN, JORDAN', 'CPT', ""),
    ('FUSION', 'BAKER, CASEY', 'CW3', ""),
    ('FUSION', 'BELL, RILEY', 'CW2', ""),
    ('FUSION', 'BROOKS, JAMIE', 'CW2', ""),
    ('FUSION', 'BRYANT, DREW', 'SFC', ""),
    ('FUSION', 'CAMPBELL, SAM', 'SFC', ""),
    ('FUSION', 'CARTER, AVERY', 'SGT', ""),
    ('FUSION', 'COLE, QUINN', 'CPL', ""),
    ('FUSION', 'COLLINS, REESE', 'SGT', ""),
    ('FUSION', 'DIAZ, PAT', 'SPC', ""),
    ('FUSION', 'EDWARDS, ALEX', 'SGT', ""),
    ('FUSION', 'ELLIS, JORDAN', 'SPC', ""),
    ('FUSION', 'EVANS, CASEY', 'SPC', ""),
    ('FUSION', 'FISHER, RILEY', 'PV2', ""),
    ('FUSION', 'FLORES, JAMIE', 'SPC', ""),
    ('FUSION', 'FOSTER, DREW', 'PFC', ""),
    ('GEOINT', 'GARCIA, SAM', 'SGT', ""),
    ('GEOINT', 'GRAHAM, AVERY', 'MSG', ""),
    ('GEOINT', 'GRAY, QUINN', 'SPC', ""),
    ('GEOINT', 'GREEN, REESE', 'SPC', ""),
    ('GEOINT', 'HALL, PAT', 'PFC', ""),
    ('GEOINT', 'HARRIS, ALEX', 'CW2', ""),
    ('GEOINT', 'HAYES, JORDAN', 'SGT', ""),
    ('GEOINT', 'HILL, CASEY', 'SFC', ""),
    ('GEOINT', 'HOLMES, RILEY', 'SGT', ""),
    ('GEOINT', 'HUGHES, JAMIE', 'SGT', ""),
    ('GEOINT', 'JAMES, DREW', 'PV2', ""),
    ('GEOINT', 'JENKINS, SAM', 'SSG', ""),
    ('GEOINT', 'JOHNSON, AVERY', 'PFC', ""),
    ('GEOINT', 'KELLY, QUINN', 'SPC', ""),
    ('GEOINT', 'KING, REESE', 'SFC', ""),
    ('CM&D', 'LEWIS, PAT', 'MAJ', ""),
    ('CM&D', 'LOPEZ, ALEX', 'CPL', ""),
    ('CM&D', 'MARSH, JORDAN', 'SGT', ""),
    ('CM&D', 'MARTIN, CASEY', 'CW2', ""),
    ('CM&D', 'MASON, RILEY', 'SFC', ""),
    ('CM&D', 'MITCHELL, JAMIE', 'SGT', ""),
    ('CM&D', 'MOORE, DREW', 'SPC', ""),
    ('OPS', 'MORRIS, SAM', 'SFC', ""),
    ('OPS', 'MURPHY, AVERY', 'SGT', ""),
    ('OPS', 'NELSON, QUINN', 'SGT', ""),
    ('OPS', 'OWENS, REESE', 'SPC', ""),
    ('OPS', 'PARKER, PAT', 'SSG', ""),
    ('IEW', 'PEREZ, ALEX', 'SPC', ""),
    ('IEW', 'PHILLIPS, JORDAN', 'SGT', ""),
    ('IEW', 'PRICE, CASEY', 'SPC', ""),
    ('G2X', 'REYES, RILEY', 'CW2', ""),
    ('G2X', 'RIVERA, JAMIE', 'CW4', ""),
    ('G2X', 'ROBERTS, DREW', 'SSG', ""),
    ('G2X', 'ROSS, SAM', 'SGT', ""),
    ('G2X', 'RUSSELL, AVERY', 'SPC', ""),
    ('G2X', 'SANDERS, QUINN', 'CW4', ""),
    ('G2X', 'SIMMONS, REESE', 'SFC', ""),
    ('G2X', 'STEWART, PAT', 'SFC', ""),
    ('G2X', 'SULLIVAN, ALEX', 'CPL', ""),
    ('G2X', 'TORRES, JORDAN', 'CW3', ""),
    ('SIGINT', 'TURNER, CASEY', 'SPC', ""),
    ('SIGINT', 'WALKER, RILEY', 'SPC', ""),
    ('SIGINT', 'WARD, JAMIE', 'SPC', ""),
    ('SIGINT', 'WEBB, DREW', 'SPC', ""),
    ('SIGINT', 'WELLS, SAM', 'SGT', ""),
    ('SIGINT', 'WEST, AVERY', 'SSG', ""),
    ('SIGINT', 'WHITE, QUINN', 'SSG', ""),
    ('SIGINT', 'WOOD, REESE', 'SSG', ""),
    ('SIGINT', 'YOUNG, PAT', 'CW3', ""),
    ('SIGINT', 'ABBOTT, ALEX', 'CW3', ""),
    ('SIGINT', 'BURKE, JORDAN', 'MAJ', ""),
    ('SSO', 'DUNN, CASEY', 'CPT', ""),
    ('SSO', 'FROST, RILEY', 'SSG', ""),
]

SRC_NOV = "Example exemptions"
# (name, type, start, end or None, remarks)
AVAILABILITY = [
    ('CAMPBELL, SAM', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('CARTER, AVERY', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('COLLINS, REESE', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('DIAZ, PAT', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('EDWARDS, ALEX', 'PCS / ETS / Transition', D(2026, 11, 1), D(2026, 11, 13), 'Example'),
    ('EDWARDS, ALEX', 'Leave', D(2026, 11, 16), D(2026, 12, 9), 'Example'),
    ('EVANS, CASEY', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('GARCIA, SAM', 'Pass', D(2026, 11, 13), D(2026, 11, 16), 'Example'),
    ('GARCIA, SAM', 'Leave', D(2026, 11, 18), D(2026, 11, 25), 'Example'),
    ('GARCIA, SAM', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('GRAHAM, AVERY', 'Pass', D(2026, 11, 13), D(2026, 11, 16), 'Example'),
    ('GRAHAM, AVERY', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('GRAY, QUINN', 'Other Detail', D(2026, 11, 1), D(2026, 11, 30), 'Example'),
    ('GREEN, REESE', 'School / Training', D(2026, 11, 9), D(2026, 11, 18), 'Example'),
    ('GREEN, REESE', 'Request (non-binding)', D(2026, 11, 11), D(2026, 11, 11), 'Example'),
    ('HALL, PAT', 'School / Training', D(2026, 11, 9), D(2026, 11, 18), 'Example'),
    ('HALL, PAT', 'Request (non-binding)', D(2026, 11, 14), D(2026, 11, 14), 'Example'),
    ('HAYES, JORDAN', 'School / Training', D(2026, 10, 8), D(2026, 11, 19), 'Example'),
    ('HAYES, JORDAN', 'Leave', D(2026, 11, 24), D(2026, 12, 2), 'Example'),
    ('HILL, CASEY', 'School / Training', D(2026, 11, 9), D(2026, 11, 18), 'Example'),
    ('HILL, CASEY', 'Request (non-binding)', D(2026, 11, 11), D(2026, 11, 11), 'Example'),
    ('HOLMES, RILEY', 'Pass', D(2026, 11, 13), D(2026, 11, 16), 'Example'),
    ('HOLMES, RILEY', 'Pass', D(2026, 11, 25), D(2026, 11, 28), 'Example'),
    ('HUGHES, JAMIE', 'School / Training', D(2026, 11, 9), D(2026, 11, 18), 'Example'),
    ('HUGHES, JAMIE', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('JAMES, DREW', 'School / Training', D(2026, 11, 9), D(2026, 11, 18), 'Example'),
    ('JENKINS, SAM', 'TDY / Mission Support', D(2026, 11, 1), D(2026, 11, 30), 'Example'),
    ('JOHNSON, AVERY', 'Pass', D(2026, 11, 20), D(2026, 11, 23), 'Example'),
    ('JOHNSON, AVERY', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('KELLY, QUINN', 'School / Training', D(2026, 11, 9), D(2026, 11, 18), 'Example'),
    ('KING, REESE', 'PCS / ETS / Transition', D(2026, 10, 21), D(2026, 12, 7), 'Example'),
    ('KING, REESE', 'PCS / ETS / Transition', D(2026, 12, 8), D(2027, 2, 28), 'Example'),
    ('LOPEZ, ALEX', 'Duty Exempt', D(2026, 11, 1), None, 'Example - open-ended'),
    ('MARSH, JORDAN', 'School / Training', D(2026, 10, 13), D(2026, 11, 20), 'Example'),
    ('MASON, RILEY', 'Appointment', D(2026, 11, 2), D(2026, 11, 2), 'Example'),
    ('MASON, RILEY', 'Leave', D(2026, 11, 7), D(2026, 11, 15), 'Example'),
    ('MASON, RILEY', 'Parental Leave (MPLP)', D(2026, 11, 21), D(2026, 11, 29), 'Example'),
    ('MITCHELL, JAMIE', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('MOORE, DREW', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('MORRIS, SAM', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('MURPHY, AVERY', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('NELSON, QUINN', 'Duty Exempt', D(2026, 11, 1), None, 'Example - open-ended'),
    ('OWENS, REESE', 'Appointment', D(2026, 11, 2), D(2026, 11, 2), 'Example'),
    ('OWENS, REESE', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('PARKER, PAT', 'Special Duty', D(2026, 11, 1), None, 'Example - open-ended'),
    ('PEREZ, ALEX', 'Duty Exempt', D(2026, 11, 20), D(2026, 11, 20), 'Example'),
    ('PRICE, CASEY', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('PRICE, CASEY', 'Leave', D(2026, 11, 30), D(2026, 12, 17), 'Example'),
    ('ROBERTS, DREW', 'Pass', D(2026, 11, 27), D(2026, 11, 30), 'Example'),
    ('ROSS, SAM', 'Parental Leave (MPLP)', D(2026, 11, 16), None, 'Example - open-ended'),
    ('RUSSELL, AVERY', 'Pass', D(2026, 11, 27), D(2026, 11, 30), 'Example'),
    ('SIMMONS, REESE', 'Leave', D(2026, 11, 1), D(2026, 11, 3), 'Example'),
    ('STEWART, PAT', 'TDY / Mission Support', D(2026, 11, 1), D(2026, 11, 18), 'Example'),
    ('SULLIVAN, ALEX', 'Pass', D(2026, 10, 9), D(2026, 10, 12), 'Example'),
    ('TURNER, CASEY', 'Leave', D(2026, 11, 23), D(2026, 11, 25), 'Example'),
    ('WARD, JAMIE', 'Parental Leave (MPLP)', D(2026, 9, 18), D(2026, 11, 16), 'Example'),
    ('WEBB, DREW', 'Duty Exempt', D(2026, 11, 1), None, 'Example - open-ended'),
    ('WELLS, SAM', 'Pass', D(2026, 11, 26), D(2026, 11, 29), 'Example'),
    ('WHITE, QUINN', 'Leave', D(2026, 11, 22), D(2026, 11, 29), 'Example'),
    ('WOOD, REESE', 'Leave', D(2026, 11, 23), D(2026, 11, 29), 'Example'),
    ('FROST, RILEY', 'Duty Exempt', D(2026, 11, 1), None, 'Example - open-ended'),
]

SRC_OCT = "Example BDE roster"
# (start, end or None, duty, position, soldier, remarks)
SLOTS = [
    (D(2026, 10, 1), D(2026, 10, 8), 'Flag Detail', 'Team 1 · E5-E6', 'COLLINS, REESE', ""),
    (D(2026, 10, 1), D(2026, 10, 8), 'Flag Detail', 'Team 1 · E5-E6', 'WELLS, SAM', ""),
    (D(2026, 10, 2), None, 'CQ NCO', '', 'MARSH, JORDAN', ""),
    (D(2026, 10, 2), None, 'SD OIC (Officer)', '', 'HAWKINS, JAMIE', ""),
    (D(2026, 10, 2), None, 'Road Guard', '', 'HALL, PAT', ""),
    (D(2026, 10, 10), None, 'CQ NCO', '', 'PHILLIPS, JORDAN', ""),
    (D(2026, 10, 12), None, 'CQ Runner', '', 'MOORE, DREW', ""),
    (D(2026, 10, 13), None, 'CQ Runner', '', 'WEBB, DREW', ""),
    (D(2026, 10, 13), None, 'SD Runner', '', 'JOHNSON, AVERY', ""),
    (D(2026, 10, 14), None, 'SD NCO', '', 'HILL, CASEY', ""),
    (D(2026, 10, 15), None, 'CQ NCO', '', 'MITCHELL, JAMIE', ""),
    (D(2026, 10, 16), None, 'Road Guard', '', 'GRAY, QUINN', ""),
    (D(2026, 10, 17), D(2026, 10, 24), 'Flag Detail', 'Team 3 · NCOIC (E6-E7)', 'WHITE, QUINN', ""),
    (D(2026, 10, 17), D(2026, 10, 24), 'Flag Detail', 'Team 3 · E1-E4', 'PRICE, CASEY', ""),
    (D(2026, 10, 18), None, 'CQ NCO', '', 'HUGHES, JAMIE', ""),
    (D(2026, 10, 18), None, 'SD Runner', '', 'ELLIS, JORDAN', ""),
    (D(2026, 10, 19), None, 'SD NCO', '', 'BRYANT, DREW', ""),
    (D(2026, 10, 20), None, 'SD Runner', '', 'PEREZ, ALEX', ""),
    (D(2026, 10, 23), None, 'Road Guard', '', 'PEREZ, ALEX', ""),
    (D(2026, 10, 24), None, 'SD OIC (Officer)', '', 'HARRIS, ALEX', ""),
    (D(2026, 10, 26), None, 'SD NCO', '', 'CAMPBELL, SAM', ""),
    (D(2026, 10, 26), None, 'CQ NCO', '', 'HOLMES, RILEY', ""),
    (D(2026, 10, 28), None, 'CQ NCO', '', 'WELLS, SAM', ""),
    (D(2026, 10, 30), None, 'CQ NCO', '', 'CARTER, AVERY', ""),
    (D(2026, 10, 30), None, 'Road Guard', '', 'RUSSELL, AVERY', ""),
]

APPOINTMENTS = [
    (D(2026, 11, 2), 'All day', 'MASON, RILEY', 'Other', '', 'Scheduled', "Example"),
    (D(2026, 11, 2), '', 'OWENS, REESE', 'Medical', '', 'Scheduled', "Example"),
]

# rank, grade, order, category
RANKS = [
    ("PV1", "E1", 1, "Junior Enlisted"), ("PV2", "E2", 2, "Junior Enlisted"),
    ("PFC", "E3", 3, "Junior Enlisted"), ("SPC", "E4", 4, "Junior Enlisted"),
    ("CPL", "E4", 5, "NCO"), ("SGT", "E5", 6, "NCO"), ("SSG", "E6", 7, "NCO"),
    ("SFC", "E7", 8, "Senior NCO"), ("MSG", "E8", 9, "Senior NCO"), ("1SG", "E8", 10, "Senior NCO"),
    ("SGM", "E9", 11, "Senior NCO"), ("CSM", "E9", 12, "Senior NCO"),
    ("WO1", "W1", 20, "Warrant Officer"), ("CW2", "W2", 21, "Warrant Officer"),
    ("CW3", "W3", 22, "Warrant Officer"), ("CW4", "W4", 23, "Warrant Officer"),
    ("CW5", "W5", 24, "Warrant Officer"),
    ("2LT", "O1", 30, "Officer"), ("1LT", "O2", 31, "Officer"), ("CPT", "O3", 32, "Officer"),
    ("MAJ", "O4", 33, "Officer"), ("LTC", "O5", 34, "Officer"), ("COL", "O6", 35, "Officer"),
]

# roster key, tab, min order, max order, nature of duty
ROSTERS = [
    ("CQ NCO", "DA6 CQ NCO", 5, 6, "Charge of Quarters NCO (CPL-SGT)"),
    ("CQ RUNNER", "DA6 CQ RUNNER", 1, 4, "Charge of Quarters Runner (PVT-SPC)"),
    ("SD RUNNER", "DA6 SD RUNNER", 1, 4, "Staff Duty Runner (PVT-SPC)"),
    ("SDNCO", "DA6 SDNCO", 7, 8, "Staff Duty NCO (SSG-SFC)"),
    ("FLAG", "DA6 FLAG", 1, 8, "Flag Detail (PVT-SFC)"),
    ("ROAD GUARD", "DA6 ROAD GUARD", 1, 6, "Road Guard Detail (PVT-SGT)"),
]

# duty type (what the BDE roster calls it) -> DA 6 roster key
DUTY_TYPES = [
    ("CQ NCO", "CQ NCO", "Charge of Quarters NCO"),
    ("CQ Runner", "CQ RUNNER", "Own DA 6; coded D on SD RUNNER when tasked nearby"),
    ("SD NCO", "SDNCO", "Staff Duty NCO"),
    ("SD Runner", "SD RUNNER", "Own DA 6; coded D on CQ RUNNER when tasked nearby"),
    ("Flag Detail", "FLAG", "Enter team start date + end date"),
    ("Road Guard", "ROAD GUARD", "Road guard detail"),
    ("SD OIC (Officer)", "N/A", "Logged for visibility, no NCO DA 6"),
    ("Other Tasking", "N/A", "Any other G2-labeled requirement"),
]

# type, DA 6 code, blocks duty, description
ABSENCE_TYPES = [
    ("Leave", "A", "Y", "Ordinary / emergency leave"),
    ("Pass", "A", "Y", "Regular or special pass"),
    ("TDY / Mission Support", "A", "Y", "TDY, mission support, external tasking (beyond duty hours)"),
    ("TDY - duty hours only", "", "N", "DA PAM 220-45 7c(1) note: TDY ending by close of business is not an A"),
    ("School / Training", "A", "Y", "NCOES, courses, training events (e.g. TT26)"),
    ("Appointment", "A", "Y", "Medical / dental / admin appointment"),
    ("Medical / Quarters", "A", "Y", "Sick call, quarters, profile (in line of duty)"),
    ("Parental Leave (MPLP)", "A", "Y", "Military Parental Leave Program"),
    ("PCS / ETS / Transition", "A", "Y", "Clearing, PCS leave, retirement transition"),
    ("Special Duty", "A", "Y", "Authorized special duty (e.g. watch floor)"),
    ("Duty Exempt", "A", "Y", "Command-approved exemption"),
    ("Other Detail", "D", "Y", "Other duty / detail: number continues"),
    ("Misconduct / AWOL", "U", "Y", "AWOL, arrest, confinement, illness not in LOD: number continues"),
    ("Extra duty / restriction (Art. 15)", "U", "Y", "Own misconduct (DA PAM 220-45 7c(3)); never post extra duty as a DA 6 detail"),
    ("Request (non-binding)", "", "N", "Soldier request: flagged only, never blocks"),
]

HOLIDAYS = [
    (D(2026, 10, 9), "DONSA (per OCT26 BDE roster)"),
    (D(2026, 10, 12), "Columbus Day"),
    (D(2026, 11, 11), "Veterans Day"),
    (D(2026, 11, 26), "Thanksgiving Day"),
    (D(2026, 11, 27), "DONSA (verify)"),
    (D(2026, 12, 24), "DONSA (verify)"),
    (D(2026, 12, 25), "Christmas Day"),
    (D(2026, 12, 31), "DONSA (verify)"),
    (D(2027, 1, 1), "New Year's Day"),
    (D(2027, 1, 18), "Martin Luther King Jr. Day"),
    (D(2027, 2, 15), "Washington's Birthday"),
    (D(2027, 5, 31), "Memorial Day"),
    (D(2027, 6, 18), "Juneteenth (observed)"),
    (D(2027, 7, 5), "Independence Day (observed)"),
    (D(2027, 9, 6), "Labor Day"),
]

TRACKERS = [
    ("Duty Slots", "Duty Slots", "Every G2-labeled duty slot (source for all DA 6s)"),
    ("Availability", "Availability", "Leave, pass, TDY, school, exemptions"),
    ("Appointments", "Appointments", "Soldier appointments, duty-conflict check"),
    ("Equipment", "Equipment", "Hand receipt items, status, inventory cycle"),
    ("Taskings", "Taskings", "Incoming taskings, suspense, fill status"),
]

SECTIONS = ["FUSION", "GEOINT", "CM&D", "OPS", "IEW", "G2X", "SIGINT", "SSO"]

# PSG shown on the dashboard banner and flagged on Personnel
PSG_NAME = 'GRAHAM, AVERY'
PSG_LABEL = "MSG Graham, Avery (example)"

# data-quality notes shown on the How To tab
HOWTO_AVAILABILITY = "Availability: add end dates for any open-ended exemptions (filter End = blank)."
HOWTO_VERIFY = "Check Duty Slots ▸ Check column for NOT ON PERSONNEL and rank-band flags after every import."
