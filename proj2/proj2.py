import csv
import os
from fpdf import FPDF
from datetime import datetime
import streamlit as st
import pandas as pd


UNI_PROGRAM = {
    "01": "Bachelor of Technology",
    "11": "Master of Technology",
    "12": "Master of Science",
    "21": "Doctor of Philosophy",
}

global TAKEN_CREDITS_TOTAL
global PREVIOUS_CPI
TAKEN_CREDITS_TOTAL = 0
PREVIOUS_CPI = 0
SUBJECTS = {}
NAME_ROLLNO = {}
STUDENTS = {}
ASSETS_FOLDER = "assets"
UPLOAD_FOLDER = "uploads"
ERROR_MESSAGE = "Transcript missing: "

GRADE_POINT = {
    "AA": 10,
    "AB": 9,
    "BB": 8,
    "BC": 7,
    "CC": 6,
    "CD": 5,
    "DD": 4,
    "F": 0,
}

COURSE = {
    "CS": "Computer Science and Engineering",
    "EE": "Electrical and Electronics Engineering",
    "ME": "Mechanical Engineering",
    "CE": "Civil and Environmental Engineering",
    "CB": "Chemical and Biochemcial Engineering",
    "MM": "Metallurgical and Material Engineering",
}


class semester:
    spi = 0
    credit_taken = 0
    cpi = 0
    credit_clear = 0

    def __init__(self, subjects):
        self.credit_taken = 0
        self.subjects = subjects

    def get_taken_credits(self):
        total_credit = 0
        for sub in self.subjects:
            total_credit += int(SUBJECTS[sub[0]]["crd"])
        self.credit_taken = total_credit

    def get_cleared_credits(self):
        credit_cleared = 0
        for sub in self.subjects:
            if sub[1] != "F":
                credit_cleared += int(SUBJECTS[sub[0]]["crd"])
        self.credit_clear = credit_cleared

    def get_spi(self):
        spi = 0
        for sub in self.subjects:
            sub[1] = sub[1].strip()
            curr_grade = sub[1]
            if curr_grade[-1] == "*":
                curr_grade = curr_grade[:-1]
            if curr_grade != "F":
                spi += int(GRADE_POINT[curr_grade]) * int(
                    SUBJECTS[sub[0]]["crd"]
                )
        spi /= self.credit_taken
        spi = int((spi * 100) + 0.5) / 100.0
        self.spi = spi

    def get_cpi(self):
        global TAKEN_CREDITS_TOTAL
        global PREVIOUS_CPI
        cpi = (PREVIOUS_CPI * TAKEN_CREDITS_TOTAL) + (
            self.spi * self.credit_taken
        )
        cpi /= TAKEN_CREDITS_TOTAL + self.credit_taken
        TAKEN_CREDITS_TOTAL += self.credit_taken
        cpi = int((cpi * 100) + 0.5) / 100.0
        PREVIOUS_CPI = cpi
        self.cpi = cpi

    def get_info(self):
        self.get_taken_credits()
        self.get_cleared_credits()
        self.get_spi()
        self.get_cpi()


class student:
    def __init__(self, name, roll_no, program, year, course, semesters):
        self.name = name
        self.roll_no = roll_no
        self.program = program
        self.year = year
        self.course = course
        self.semesters = semesters
        self.A3 = 0
        if self.program[0] == "B":
            self.A3 = 1

    def draw_margins(pdf, left_x, right_x, top_y, bottom_y):
        pdf.set_line_width(0.5)
        pdf.set_draw_color(r=0, g=0, b=0)
        pdf.line(x1=left_x, x2=right_x, y1=top_y, y2=top_y)
        pdf.line(x1=left_x, x2=right_x, y1=bottom_y, y2=bottom_y)
        pdf.line(x1=left_x, x2=left_x, y1=top_y, y2=bottom_y)
        pdf.line(x1=right_x, x2=right_x, y1=top_y, y2=bottom_y)

    def create_table(curr_sem, pdf, start_x, start_y, sem, a3):
        curr_x = start_x
        curr_y = start_y

        col_width = [16, 55, 8, 8, 8] if a3 else [15, 50, 8, 8, 8]
        line_height = 4 if a3 else 3
        font = "Times"
        font_size = 6.5 if a3 else 5.5
        heading_font_size = 15 if a3 else 11
        table_heading_w = 50
        table_heading_h = 10 if a3 else 7
        cols = ["Sub. Code", "Subject Name", "L-T-P", "CRD", "GRD"]

        pdf.set_xy(x=curr_x, y=curr_y)
        pdf.set_font(size=heading_font_size)
        pdf.multi_cell(
            table_heading_w,
            table_heading_h,
            sem,
            border=0,
            ln=3,
            align="L",
            markdown=True,
        )
        pdf.ln()

        if curr_sem.subjects:
            pdf.set_font(font, size=font_size, style="B")
            pdf.set_x(curr_x)
            for j in range(5):
                pdf.multi_cell(
                    col_width[j],
                    line_height,
                    cols[j],
                    border=1,
                    ln=3,
                    align="C",
                    markdown=True,
                )
            pdf.ln()
            pdf.set_x(start_x)
            pdf.set_font(font, size=font_size, style="")

            for curr_sub in curr_sem.subjects:
                row = [
                    curr_sub[0],
                    SUBJECTS[curr_sub[0]]["subname"],
                    SUBJECTS[curr_sub[0]]["ltp"],
                    SUBJECTS[curr_sub[0]]["crd"],
                    curr_sub[1],
                ]
                for j in range(len(row)):
                    pdf.multi_cell(
                        col_width[j],
                        line_height,
                        str(row[j]),
                        border=1,
                        ln=3,
                        align="C",
                        markdown=True,
                    )
                pdf.ln()
                pdf.set_x(curr_x)

            pdf.set_y(pdf.get_y() + 3)
            pdf.set_x(start_x)
            bottom_text = (
                "**"
                + "Credits Taken: "
                + str(curr_sem.credit_taken)
                + "     "
                + "Credits Cleared: "
                + str(curr_sem.credit_clear)
                + "     "
                + "SPI: "
                + str(curr_sem.spi)
                + "     "
                + "CPI: "
                + str(curr_sem.cpi)
                + "**"
            )
            pdf.multi_cell(
                col_width[0] + col_width[1],
                line_height + 1,
                bottom_text,
                border=1,
                ln=3,
                align="J",
                markdown=True,
            )

    def create_banner(pdf, left_x, right_x, top_y):
        file_name = os.path.join(ASSETS_FOLDER, "banner.png")
        header1_x = left_x + 1
        header1_y = top_y + 1
        header1_w = right_x - left_x - 2
        header1_h = 40 - 2

        pdf.set_xy(header1_x, header1_y)
        pdf.image(w=header1_w, name=file_name)

    def create_student_info(
        self, pdf, header2_x, header2_y, header2_w, header2_h
    ):
        tab = "                "
        header2_string = (
            "**Roll Number:** "
            + self.roll_no
            + tab
            + "**Name:** "
            + self.name
            + tab
            + "**Year of Admission:** "
            + self.year
            + "\n"
        )
        header2_string += (
            "**Programme:** "
            + self.program
            + tab
            + "**Course:** "
            + self.course
        )
        pdf.set_xy(header2_x, header2_y)
        pdf.multi_cell(
            header2_w,
            header2_h / 2,
            header2_string,
            border=1,
            ln=3,
            align="L",
            markdown=True,
        )

    def insert_stamp(
        pdf, stamp_pos_x, stamp_pos_y, stamp_w, stamp_h, file_name
    ):
        pdf.set_xy(stamp_pos_x, stamp_pos_y)
        pdf.image(name=file_name, h=stamp_h, w=stamp_w)

    def create_footer(pdf, left_x, right_x, bottom_y, needed_seal, a3):
        dog_h = 5
        dog_x = left_x + 5
        dog_w = 150
        dog_y = bottom_y - 10

        sign_h = 5
        sign_x = right_x - 70
        sign_w = 70
        sign_y = bottom_y - 10

        line_x1 = sign_x
        line_y1 = sign_y - 2
        line_x2 = sign_x + sign_w - 10
        line_y2 = line_y1

        font_size = 12
        font = "Times"

        pdf.set_xy(dog_x, dog_y)
        pdf.set_font(font, size=font_size, style="B")
        date_text = "Date of Generation: " + datetime.now().strftime(
            "%d/%m/%Y %H:%M"
        )
        pdf.multi_cell(
            dog_w, dog_h, date_text, border=0, ln=3, align="L", markdown=True
        )

        pdf.set_xy(sign_x, sign_y)
        pdf.multi_cell(sign_w, sign_h, "Asssistant Registrar (Academic)")

        signi_x = sign_x + 5
        signi_y = (line_y1 - 27) if a3 else (line_y1 - 18)
        signi_w = 30 if a3 else 20
        signi_h = 30 if a3 else 20

        signi_filename = os.path.join(ASSETS_FOLDER, "signature.png")
        pdf.set_xy(signi_x, signi_y)
        if needed_seal:
            pdf.image(name=signi_filename, h=signi_h, w=signi_w)

        pdf.line(x1=line_x1, x2=line_x2, y1=line_y1, y2=line_y2)

    def create_transcript(self, needed_seal):
        a3 = self.A3
        file_name = os.path.join("transcriptsIITP", self.roll_no + ".pdf")
        page_size = "A3" if a3 else "A4"
        left_m = 5
        left_x = left_m
        right_m = 5
        right_x = 420 - right_m if a3 else 297 - right_m
        top_m = 5
        top_y = top_m
        bottom_m = 5
        bottom_y = 297 - bottom_m if a3 else 210 - bottom_m
        add_stamp = needed_seal

        font = "Times"
        font_size = 15 if a3 else 10

        pdf = FPDF(orientation="landscape", format=page_size)
        pdf.add_page()
        pdf.set_font(font, size=font_size)
        pdf.set_margins(left=left_m, right=right_m, top=top_m)
        pdf.set_auto_page_break(False)
        student.draw_margins(pdf, left_x, right_x, top_y, bottom_y)

        student.create_banner(pdf, left_x, right_x, top_y)

        curr_x = pdf.get_x()
        curr_y = pdf.get_y()

        header2_w = 250 if a3 else 170
        header2_h = 15 if a3 else 12
        header2_x = 85 if a3 else 60
        header2_y = curr_y + 5

        student.create_student_info(
            self, pdf, header2_x, header2_y, header2_w, header2_h
        )

        partition_1 = header2_y + header2_h + (76.5 if a3 else 65)
        partition_2 = (partition_1 + 70) if a3 else partition_1

        pdf.set_line_width(0.5)
        pdf.set_draw_color(r=0, g=0, b=0)
        if len(self.semesters) > 4:
            pdf.line(x1=left_x, x2=right_x, y1=partition_1, y2=partition_1)
        if len(self.semesters) > 8:
            pdf.line(x1=left_x, x2=right_x, y1=partition_2, y2=partition_2)

        start_x = (
            [left_x + 5, left_x + 105, left_x + 205, left_x + 305]
            if a3
            else [left_x + 5, left_x + 99, left_x + 193]
        )
        start_y = (
            [header2_y + header2_h, partition_1, partition_2]
            if a3
            else [header2_y + header2_h + 5, partition_1 + 5, partition_2 + 5]
        )
        row_lim = len(start_x)
        for curr_sem in range(len(self.semesters)):
            sem_string = "**--Semester " + str(curr_sem + 1) + "--**"
            student.create_table(
                self.semesters[curr_sem],
                pdf,
                start_x[(curr_sem) % row_lim],
                start_y[(curr_sem) // row_lim],
                sem_string,
                a3,
            )

        stamp_pos_x = start_x[3] + 10 if a3 else start_x[1]
        stamp_pos_y = start_y[2] + 10 if a3 else start_y[2] + 45
        stamp_w = 35 if a3 else 25
        stamp_h = 35 if a3 else 25
        stamp_name = os.path.join(ASSETS_FOLDER, "stamp.png")
        student.create_footer(pdf, left_x, right_x, bottom_y, needed_seal, a3)
        if add_stamp:
            student.insert_stamp(
                pdf, stamp_pos_x, stamp_pos_y, stamp_w, stamp_h, stamp_name
            )
        pdf.output(file_name)


def read_files():
    roll_file = open(os.path.join(UPLOAD_FOLDER, "names-roll.csv"))
    name_rollno = csv.DictReader(roll_file)
    grade_file = open(os.path.join(UPLOAD_FOLDER, "grades.csv"))
    grade = csv.DictReader(grade_file)
    subj_file = open(os.path.join(UPLOAD_FOLDER, "subjects_master.csv"))
    reader = csv.DictReader(subj_file)

    global SUBJECTS
    global NAME_ROLLNO
    global STUDENTS
    for row in reader:
        SUBJECTS[row["subno"]] = {
            "subname": row["subname"],
            "ltp": row["ltp"],
            "crd": row["crd"],
        }

    for row in name_rollno:
        NAME_ROLLNO[row["Roll"]] = row["Name"]

    grade = sorted(
        grade,
        key=lambda row: (row["Roll"], int(row["Sem"]), row["SubCode"]),
        reverse=False,
    )

    student_roll = grade[0]["Roll"]
    count = len(grade)
    idx = 0

    while idx < count:
        global TAKEN_CREDITS_TOTAL
        global PREVIOUS_CPI
        TAKEN_CREDITS_TOTAL = 0
        PREVIOUS_CPI = 0
        prev_sem = 0
        if grade[idx]["Roll"] == student_roll:
            semesters = []
            while idx < count and grade[idx]["Roll"] == student_roll:
                subjects = []
                sem = grade[idx]["Sem"]
                while prev_sem + 1 != int(sem):
                    prev_sem += 1
                    subject_1 = {}
                    curr_sem_1 = semester(subject_1)
                    semesters.append(curr_sem_1)
                while idx < count and grade[idx]["Sem"] == sem:
                    subjects.append(
                        [grade[idx]["SubCode"], grade[idx]["Grade"]]
                    )
                    idx += 1
                prev_sem += 1
                curr_sem = semester(subjects)
                curr_sem.get_info()
                semesters.append(curr_sem)
                if idx < count:
                    sem = grade[idx]["Sem"]
            STUDENTS[student_roll] = student(
                NAME_ROLLNO[student_roll],
                student_roll,
                UNI_PROGRAM[student_roll[2:4]],
                "20" + student_roll[:2],
                COURSE[student_roll[4:6]],
                semesters,
            )
            if idx < count:
                student_roll = grade[idx]["Roll"]


def generate_pdf(start, end, needed_seal):
    global ERROR_MESSAGE
    if int(start[6:8]) > int(end[6:8]) or start[:6] != end[:6]:
        st.error("Invalid range entered")
    else:
        read_files()
        for rollno in range(int(start[6:8]), int(end[6:8]) + 1):
            roll_no = start[0:6]
            if rollno < 10:
                roll_no += "0"
            roll_no += str(rollno)

            if roll_no in STUDENTS:
                if needed_seal == 1 or needed_seal == "1":
                    STUDENTS[roll_no].create_transcript(1)
                else:
                    STUDENTS[roll_no].create_transcript(0)
            else:
                ERROR_MESSAGE += roll_no
        if ERROR_MESSAGE[-1] != " ":
            print(roll_no, ": missing")
            st.error(ERROR_MESSAGE)


def create_all_transcripts(needed_seal):
    read_files()
    if needed_seal == "1" or needed_seal == 1:
        for key in STUDENTS:
            STUDENTS[key].create_transcript(1)
    else:
        for key in STUDENTS:
            STUDENTS[key].create_transcript(0)


# website
st.title("Web-based Transcript Generator")

st.write("Enter roll number range (starting and ending roll numbers):")
start_roll = st.text_input("Roll number start:")
end_roll = st.text_input("Roll number end:")
my_slot1 = st.empty()

roll_no = st.file_uploader("Browse names-roll.csv")
if roll_no:
    dataframe = pd.read_csv(roll_no)
    dataframe.to_csv(os.path.join(UPLOAD_FOLDER, "names-roll.csv"))

subject = st.file_uploader("Browse subjects_master.csv")
if subject:
    dataframe = pd.read_csv(subject)
    dataframe.to_csv(os.path.join(UPLOAD_FOLDER, "subjects_master.csv"))

grades_csv = st.file_uploader("Browse grades.csv")
if grades_csv:
    dataframe = pd.read_csv(grades_csv)
    dataframe.to_csv(os.path.join(UPLOAD_FOLDER, "grades.csv"))

needed_seal = False
needed_seal = st.checkbox(
    "Stamp and signature",
    value=False,
    key=None,
    help=None,
    on_change=None,
    args=None,
    kwargs=None,
)

st.button(
    "Generate transcript",
    key=None,
    help=None,
    on_click=generate_pdf,
    args=(start_roll, end_roll, needed_seal),
    kwargs=None,
)

st.button(
    "Generate all transcripts",
    key=None,
    help=None,
    on_click=create_all_transcripts,
    args=(needed_seal,),
)
