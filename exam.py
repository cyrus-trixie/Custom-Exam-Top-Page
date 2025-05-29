import os
import random
import string
import shutil
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import streamlit as st

# === Config ===
SCHOOL_NAME = "ST. JOSEPH’S BOYS - KITALE"
EXAM_HEADER_BASE = "Kenya Certificate of Secondary Examinations"

SUBJECT_INSTRUCTIONS = {
    "Mathematics": [
        "(a) Write your details in the spaces provided above.",
        "(b) Answer all the questions in the spaces provided.",
        "(c) All working must be clearly shown.",
        "(d) Silent and non-programmable electronic calculators may be used.",
        "(e) This paper consists of 10 printed pages.",
        "(f) Candidates should check the question paper to ascertain that all the pages are printed as indicated and that no questions are missing.",
        "(g) Candidates should answer the questions in correct English."
    ],
    "English": [
        "1. Write your name and admission number clearly in the spaces provided.",
        "2. Answer ALL the questions in this exam paper.",
        "3. Mobile phones and calculators are not allowed unless stated.",
        "4. Use correct grammar and spelling throughout.",
        "5. Do not open this paper until instructed to do so."
    ],
    "Kiswahili": [
        "1. Andika jina lako na nambari ya usajili kwa nafasi zilizotolewa.",
        "2. Jibu MASWALI YOTE katika karatasi hii ya mtihani.",
        "3. Simu za rununu na kikokotoo haziruhusiwi isipokuwa kikitajwa.",
        "4. Karatasi hii ina kurasa 10 zilizochapishwa.",
        "5. Usifungue karatasi hii hadi uelekezwe kufanya hivyo."
    ],
    "Chemistry": [
        "1. Write your name and admission number clearly.",
        "2. Answer all the questions in the spaces provided.",
        "3. Scientific calculators and mathematical tables may be used.",
        "4. Read all instructions carefully before attempting any question."
    ],
    "Biology": [
        "1. Read all questions before answering.",
        "2. Use diagrams where necessary.",
        "3. Show all your working clearly.",
        "4. No mobile phones allowed."
    ],
    "Default": [
        "1. Write your name and admission number clearly.",
        "2. Answer all questions.",
        "3. No mobile phones or unauthorized materials allowed.",
        "4. Follow invigilator instructions."
    ]
}

def generate_exam_number(length=10):
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

def generate_exam_pdf(subject, form_level, term, exam_name, exam_date, duration, logo_image, instructions, include_marking_table, include_exam_number):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    y = height - 80

    if logo_image:
        logo = ImageReader(logo_image)
        c.drawImage(logo, 60, y - 10, width=60, height=60, preserveAspectRatio=True)

    # Header
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, y, SCHOOL_NAME)

    c.setFont("Helvetica-Bold", 13)
    # Line with exam board + form + subject
    header_line_2 = f"{EXAM_HEADER_BASE} - {form_level.upper()} {subject.upper()}"
    c.drawCentredString(width / 2, y - 20, header_line_2)

    # Term and exam name
    c.drawCentredString(width / 2, y - 40, f"{term} – {exam_name}")

    # Date and duration line
    c.drawCentredString(width / 2, y - 60, f"{exam_date} - {duration}")

    # Student Info
    c.setFont("Helvetica", 12)
    c.drawString(100, y - 100, f"Name: _______________________")
    c.drawString(350, y - 100, f"Adm. No: _______________________")
    c.drawString(100, y - 120, f"Class: _______________________")
    if include_exam_number:
        c.drawString(350, y - 120, f"Exam Number: {generate_exam_number()}")

    # Instructions
    c.setFont("Helvetica-Bold", 12)
    c.drawString(100, y - 160, "Instructions:")
    c.setFont("Helvetica", 11)
    iy = y - 180
    for line in instructions:
        c.drawString(110, iy, line)
        iy -= 16

    if include_marking_table:
        c.setFont("Helvetica-Bold", 12)
        iy -= 30
        c.drawString(100, iy, "For examiners use only")
        iy -= 25

        table_data = [
            ["SECTION", "QUESTION", "MAXIMUM SCORE", "CANDIDATE’S SCORE"],
            ["A", "1 – 11", "25", ""],
            ["B", "12", "11", ""],
            ["", "13", "11", ""],
            ["", "14", "11", ""],
            ["", "15", "10", ""],
            ["", "16", "12", ""],
            ["", "TOTAL SCORE", "80", ""],
        ]

        x_start = 100
        col_widths = [80, 120, 120, 150]
        row_height = 25
        c.setFont("Helvetica", 10)

        for i, row in enumerate(table_data):
            row_y = iy - (i * row_height)
            for j, cell in enumerate(row):
                x = x_start + sum(col_widths[:j])
                c.rect(x, row_y, col_widths[j], row_height, stroke=1, fill=0)
                c.drawString(x + 5, row_y + 7, str(cell))

    c.save()
    buffer.seek(0)
    return buffer

# === UI ===
st.title("Exam Top Page Generator")

logo_file = st.file_uploader("Upload School Logo (optional)", type=["png", "jpg", "jpeg"])

subject = st.text_input("Enter Subject (e.g. Physics, Mathematics)", value="Physics")
form_level = st.text_input("Enter Form Level (e.g. Form 1, Form 2, Form 3, Form 4)", value="Form 3")
term = st.text_input("Enter Term (e.g. Term 1, Term 2, Term 3)", value="Term 2")
exam_name = st.text_input("Enter Exam Name (e.g. Exam 1, Exam 2, Jesma)", value="Exam 1")

exam_date = st.date_input("Exam Date")
duration = st.text_input("Exam Duration", value="2 HOURS")

include_exam_number = st.checkbox("Include Exam Number", value=True)
include_marking_table = st.checkbox("Include Marking Table", value=True)

# Instruction Editor
st.markdown("### Edit Instructions Below (if needed)")
default_instructions = SUBJECT_INSTRUCTIONS.get(subject.title(), SUBJECT_INSTRUCTIONS["Default"])
instructions = st.text_area("Instructions", "\n".join(default_instructions), height=200)
instruction_lines = [line.strip() for line in instructions.strip().split("\n") if line.strip()]

# Page Count Input
num_pages = st.number_input("Number of Blank Pages to Generate", min_value=1, max_value=100, value=30)

# Preview Button
if st.button("Preview One Page"):
    st.info("Preview of one exam top page:")
    pdf_buffer = generate_exam_pdf(subject, form_level, term, exam_name, exam_date.strftime("%d %B %Y"), duration, logo_file, instruction_lines, include_marking_table, include_exam_number)
    st.download_button("Download Preview", data=pdf_buffer, file_name=f"{subject}_Top_Page_Preview.pdf", mime="application/pdf")

# Bulk Generate Button
if st.button("Generate Blank Top Pages"):
    with st.spinner("Generating..."):
        output_folder = "exam_top_pages"
        os.makedirs(output_folder, exist_ok=True)
        for i in range(num_pages):
            pdf = generate_exam_pdf(subject, form_level, term, exam_name, exam_date.strftime("%d %B %Y"), duration, logo_file, instruction_lines, include_marking_table, include_exam_number)
            with open(os.path.join(output_folder, f"{subject}_Page_{i+1}.pdf"), "wb") as f:
                f.write(pdf.read())

        zip_path = shutil.make_archive("ExamTopPages", 'zip', output_folder)
        st.success("Done! Download your ZIP file below:")
        st.download_button("Download All as ZIP", open(zip_path, "rb"), "ExamTopPages.zip")
        shutil.rmtree(output_folder)
