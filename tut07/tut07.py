import pandas as pd

# file names
REGISTERED_COURSES_FILE = "course_registered_by_all_students.csv"
FEEDBACK_FILE = "course_feedback_submitted_by_students.csv"
COURSE_MASTER_FILE = "course_master_dont_open_in_excel.csv"
STUDENT_FILE = "studentinfo.csv"
OUTPUT_FILE = "course_feedback_remaining.xlsx"

# dicts
COURSES = {}
STUDENTS = {}
REG_FEEDBACK = {}
GIVEN_FEEDBACK = {}


"""
NOTE:
    LTP: 1-0-0 -> LTP_BITS: 1 (001)
    LTP: 0-0-0 -> LTP_BITS: 0 (000)
    LTP: 2-0-3 -> LTP_BITS: 5 (101)
    LTP: 2-2-0 -> LTP_BITS: 3 (011)
"""


def init_courses():
    """
    initialize COURSES dict with
    course code and the integer representation of its LTP
    """
    courses_file = pd.read_csv(COURSE_MASTER_FILE)
    for _, row in courses_file.iterrows():
        ltps = str(row["ltp"]).split("-")
        ltp_bits = 0
        for idx in range(3):
            if ltps[idx].strip() != '0':
                ltp_bits |= 1 << idx
        COURSES[row["subno"]] = ltp_bits


def init_students():
    """
    initialize STUDENTS dict with
    roll number and name, email, alternate email, contact number
    """
    students_file = pd.read_csv(STUDENT_FILE)
    for _, row in students_file.iterrows():
        STUDENTS[row["Roll No"]] = {
            "Name": row["Name"],
            "email": row["email"],
            "aemail": row["aemail"],
            "contact": row["contact"],
        }


def init_reqd_feedback():
    """
    initialize REG_FEEDBACK dict with
    key -> roll number + subject code
    val -> dict of register_sem and schedule_sem
    """
    reg_feedback_file = pd.read_csv(REGISTERED_COURSES_FILE)
    for _, row in reg_feedback_file.iterrows():
        REG_FEEDBACK[str(row["subno"]) + "+" + str(row["rollno"])] = {
            "register_sem": row["register_sem"],
            "schedule_sem": row["schedule_sem"],
        }


def init_given_feedback():
    """
    initialize GIVEN_FEEDBACK with
    key -> roll number + subject code
    val -> integer representation of student's feedback LTP
    """
    given_feedback_file = pd.read_csv(FEEDBACK_FILE)
    for _, row in given_feedback_file.iterrows():
        key = str(row["course_code"]) + "+" + str(row["stud_roll"]) 
        if key not in GIVEN_FEEDBACK:
            GIVEN_FEEDBACK[key] = 0
        GIVEN_FEEDBACK[key] |= 1 << (int(row["feedback_type"]) - 1)


def generate_missing_feedback():
    """
    create output excel file in pandas
    if students' LTP bits do not match with required bits from course info
    """
    # form empty dataframe with reqd columns
    remaining_feedback = pd.DataFrame(
        columns=[
            "rollno",
            "register_sem",
            "schedule_sem",
            "subno",
            "Name",
            "email",
            "aemail",
            "contact",
        ]
    )
    # iterate through given feedback
    for key, feedback_val in GIVEN_FEEDBACK.items():
        # get back course and roll from key
        course, roll = key.split("+")
        # handle case where LTP is 0-0-0
        if COURSES[course] not in (0, feedback_val):
            remaining_feedback = remaining_feedback.append(
                {
                    "rollno": roll,
                    "register_sem": REG_FEEDBACK[key]["register_sem"],
                    "schedule_sem": REG_FEEDBACK[key]["schedule_sem"],
                    "subno": course,
                    "Name": STUDENTS[roll]["Name"],
                    "email": STUDENTS[roll]["email"],
                    "aemail": STUDENTS[roll]["aemail"],
                    "contact": STUDENTS[roll]["contact"],
                },
                ignore_index=True,
            )
    # convert and save as excel
    remaining_feedback.to_excel(OUTPUT_FILE, index=False)


def feedback_not_submitted():
    """
    driver function to generate excel file of
    students who did not submit feedback
    """
    # initialize all dicts
    init_courses()
    init_students()
    init_reqd_feedback()
    init_given_feedback()
    # get details of students who didn't fill feedback
    generate_missing_feedback()


feedback_not_submitted()
