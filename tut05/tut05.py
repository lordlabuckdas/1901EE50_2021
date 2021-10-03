import csv
import os
import shutil

# to handle spreadsheets/excel files
from openpyxl import Workbook

# names of all relevant files
GRADES_FILENAME = "grades.csv"
NAME_ROLL_FILENAME = "names-roll.csv"
SUBJECTS_FILENAME = "subjects_master.csv"

# dictionary of details of all students
students = {}
# dictionary of details of all subjects
subjects = {}

# grade-value mapping
grade_map = {
    "AA": 10,
    "AB": 9,
    "BB": 8,
    "BC": 7,
    "CC": 6,
    "CD": 5,
    "DD": 4,
    "DD*": 4,
    "F": 0,
    "F*": 0,
    "I": 0,
}


def create_output_folder(folder_name="output"):
    """
    If output folder exists, delete all of its previous contents
    """
    # remove directory if it exists
    if os.path.exists(folder_name):
        shutil.rmtree(folder_name)
    # create directory again
    os.makedirs(folder_name)


def calculate_point_index(grades: list, credits: list):
    """
    Given 2 lists - grades and credits,
    obtain the point index (float value)
    by calculating the weight mean of grades
    """
    points = 0
    assert len(grades) == len(credits)
    # take weighted mean of grades and credits
    for g, c in zip(grades, credits):
        points += g * c
    points /= sum(credits)
    return points


def initialize_students():
    """
    Initialize the `students` dictionary with Roll number and Name
    from `name-roll.csv`
    """
    stud_file = open(NAME_ROLL_FILENAME, "r")
    reader = csv.DictReader(stud_file)
    # map roll num to name and sem (empty for now)
    for row in reader:
        # initialize with empty dicts if roll is not already present
        if not students.get(row["Roll"]):
            students[row["Roll"]] = {
                "name": {},
                "sem": {
                    1: {},
                    2: {},
                    3: {},
                    4: {},
                    5: {},
                    6: {},
                    7: {},
                    8: {},
                },
            }
        students[row["Roll"]]["name"] = row["Name"]
    stud_file.close()


def initialize_subjects():
    """
    Initialize the `subjects` dictionary with
    Name, LTP breakup and Credits
    from `grades.csv`
    """
    sub_file = open(SUBJECTS_FILENAME, "r")
    reader = csv.DictReader(sub_file)
    # map subject code to subject name, LTP breakup and credits
    for row in reader:
        # initialize with empty dicts and 0s if not present
        if not subjects.get(row["subno"]):
            subjects[row["subno"]] = {"name": "", "ltp": "", "credits": 0}
        subjects[row["subno"]]["name"] = row["subname"]
        subjects[row["subno"]]["ltp"] = row["ltp"]
        subjects[row["subno"]]["credits"] = int(row["crd"])
    sub_file.close()


def populate_grades():
    """
    Populate the `students` dictionary with semester-wise
    courses and grades
    """
    grade_file = open(GRADES_FILENAME, "r")
    reader = csv.DictReader(grade_file)
    # map sem data to subjects undertaken
    # map subjects to grade obtained and subject type
    for row in reader:
        # ignore outliers/incorrect data
        if int(row["Sem"]) > 8:
            continue
        # initalize subject with empty grade and subject type strings
        students[row["Roll"]]["sem"][int(row["Sem"])][row["SubCode"]] = {
            "grade": "",
            "subtype": "",
        }
        students[row["Roll"]]["sem"][int(row["Sem"])][row["SubCode"]][
            "grade"
        ] = row["Grade"].strip()
        students[row["Roll"]]["sem"][int(row["Sem"])][row["SubCode"]][
            "subtype"
        ] = row["Sub_Type"]
    grade_file.close()


def write_marksheet():
    """
    Write the marksheets of every student
    into the corresponding output file
    """
    for roll in students:
        wb = Workbook()
        overall_sheet = wb.active
        if not overall_sheet:
            raise Exception("Error fetching active sheet")
        overall_sheet.title = "Overall"
        # create empty lists for 8 semesters
        semwise_credits = [0] * 8
        total_credits = [0] * 8
        spi = [0.0] * 8
        cpi = [0.0] * 8
        # loop through sem 1 to 8
        for sem_no in range(1, 9):
            # if sem data is not preset, proceed to next iteration
            # additioanlly, fill in cpi and total credits taken to
            # accodmodate cases where the student might have had to
            # skip Semesters and ensure data consistency
            if len(students[roll]["sem"][sem_no]) == 0:
                if sem_no > 1:
                    cpi[sem_no - 1] = cpi[sem_no - 2]
                    total_credits[sem_no - 1] = total_credits[sem_no - 2]
                continue
            # track last attended semester
            last_attended_semester = sem_no
            # initialize semester grades and credits with empty lists
            sem_creds = []
            sem_grades = []
            # create a sheet for current semester
            sem_sheet = wb.create_sheet("Sem" + str(sem_no))
            sem_sheet.append(
                [
                    "Sl No.",
                    "Subject No.",
                    "Subject Name",
                    "L-T-P",
                    "Credit",
                    "Subject Type",
                    "Grade",
                ]
            )
            # append data of each subject
            for i, subj in enumerate(students[roll]["sem"][sem_no]):
                sem_sheet.append(
                    [
                        i + 1,
                        subj,
                        subjects[subj]["name"],
                        subjects[subj]["ltp"],
                        subjects[subj]["credits"],
                        students[roll]["sem"][sem_no][subj]["subtype"],
                        students[roll]["sem"][sem_no][subj]["grade"],
                    ]
                )
                sem_creds.append(int(subjects[subj]["credits"]))
                sem_grades.append(
                    grade_map[students[roll]["sem"][sem_no][subj]["grade"]]
                )
            semwise_credits[sem_no - 1] = sum(sem_creds)
            spi[sem_no - 1] = calculate_point_index(sem_grades, sem_creds)
            if sem_no == 1:
                total_credits[0] = semwise_credits[0]
                cpi[0] = spi[0]
            else:
                total_credits[sem_no - 1] = (
                    total_credits[sem_no - 2] + semwise_credits[sem_no - 1]
                )
                cpi[sem_no - 1] = calculate_point_index(
                    spi[:sem_no], semwise_credits[:sem_no]
                )
        # round of SPI and CPI to 2 decimals
        for i in range(8):
            spi[i] = round(spi[i], 2)
            cpi[i] = round(cpi[i], 2)
        overall_sheet.append(["Roll No.", roll])
        overall_sheet.append(["Name of Student", students[roll]["name"]])
        overall_sheet.append(["Discipline", roll[4:6]])
        overall_sheet.append(
            ["Semester No.", *[x for x in range(1, last_attended_semester + 1)]]
        )
        overall_sheet.append(
            [
                "Semester wise Credit Taken",
                *semwise_credits[:last_attended_semester],
            ]
        )
        overall_sheet.append(["SPI", *spi[:last_attended_semester]])
        overall_sheet.append(
            ["Total Credits Taken", *total_credits[:last_attended_semester]]
        )
        overall_sheet.append(["CPI", *cpi[:last_attended_semester]])
        output_filename = os.path.join("output", roll + ".xlsx")
        wb.save(output_filename)


def generate_marksheet():
    """
    Driver program to generate marksheets
    of every student
    """
    create_output_folder()
    initialize_students()
    initialize_subjects()
    populate_grades()
    write_marksheet()


generate_marksheet()
