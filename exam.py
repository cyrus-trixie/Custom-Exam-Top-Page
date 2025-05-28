import os
import random
import string
import shutil
import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# Function to generate a unique exam number
def generate_exam_number(existing_ids, length=10):
    while True:
        exam_id = ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))
        if exam_id not in existing_ids:
            existing_ids.add(exam_id)
            return exam_id

# Function to generate a single PDF
def generate_exam_pdf(name, admission_number, exam_number, subject, exam_date, output_folder):
    file_path = os.path.join(output_folder, f"{admission_number}_exam_top.pdf")
    c = canvas.Canvas(file_path, pagesize=A4)
    width, height = A4

    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(width / 2, height - 100, "EXAM PAPER")

    c.setFont("Helvetica", 12)
    c.drawString(100, height - 140, f"Name: {name}")
    c.drawString(100, height - 160, f"Admission Number: {admission_number}")
    c.drawString(100, height - 180, f"Exam Number: {exam_number}")
    c.drawString(100, height - 200, f"Subject: {subject}")
    c.drawString(100, height - 220, f"Date: {exam_date}")

    c.drawString(100, height - 260, "Instructions:")
    instructions = [
        "1. Answer all questions.",
        "2. Show all your work clearly.",
        "3. No phones or calculators unless allowed.",
        "4. Do not open this paper until instructed."
    ]
    y = height - 280
    for line in instructions:
        c.drawString(120, y, line)
        y -= 20

    c.save()

# Streamlit UI
st.title("📄 Exam Top Page Generator")

uploaded_file = st.file_uploader("Upload Student Excel File (.xlsx)", type=["xlsx"])
subject = st.text_input("Subject", value="Mathematics")
exam_date = st.date_input("Exam Date")

if st.button("Generate Exam PDFs"):
    if uploaded_file:
        with st.spinner("Generating PDFs..."):
            output_folder = "exam_top_pages"
            os.makedirs(output_folder, exist_ok=True)

            df = pd.read_excel(uploaded_file)
            existing_ids = set()

            for _, row in df.iterrows():
                name = row["Name"]
                admission_number = row["Admission Number"]
                exam_number = generate_exam_number(existing_ids)
                generate_exam_pdf(name, admission_number, exam_number, subject, exam_date.strftime("%Y-%m-%d"), output_folder)

            # ZIP all PDFs
            zip_path = shutil.make_archive("ExamPapers", 'zip', output_folder)
            st.success("✅ Done! Download your files below:")
            st.download_button("📥 Download All as ZIP", open(zip_path, "rb"), "ExamPapers.zip")

            # Clean up if needed
            shutil.rmtree(output_folder)
    else:
        st.error("Please upload an Excel file.")
