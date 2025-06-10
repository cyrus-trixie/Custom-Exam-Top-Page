import os
import random
import string
import shutil
import pandas as pd
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import A4, landscape # Import landscape
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle # Import Table and TableStyle
from reportlab.lib import colors # Import colors for table borders
import streamlit as st

# === Config ===
EXAM_HEADER = "Kenya Certificate of Secondary Examinations"

DEFAULT_INSTRUCTIONS = [
    "1. Write your name and admission number clearly.",
    "2. Answer all questions.",
    "3. No mobile phones or unauthorized materials allowed.",
    "4. Follow invigilator instructions."
]

DEFAULT_MARKING_TABLE = [
    {"SECTION": "A", "QUESTION": "1 – 11", "MAXIMUM SCORE": 25, "CANDIDATE’S SCORE": ""},
    {"SECTION": "B", "QUESTION": "12", "MAXIMUM SCORE": 11, "CANDIDATE’S SCORE": ""},
    {"SECTION": "", "QUESTION": "13", "MAXIMUM SCORE": 11, "CANDIDATE’S SCORE": ""},
    {"SECTION": "", "QUESTION": "14", "MAXIMUM SCORE": 11, "CANDIDATE’S SCORE": ""},
    {"SECTION": "", "QUESTION": "15", "MAXIMUM SCORE": 10, "CANDIDATE’S SCORE": ""},
    {"SECTION": "", "QUESTION": "16", "MAXIMUM SCORE": 12, "CANDIDATE’S SCORE": ""},
    {"SECTION": "", "QUESTION": "TOTAL SCORE", "MAXIMUM SCORE": 80, "CANDIDATE’S SCORE": ""},
]

def generate_exam_number(length=10):
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

def generate_exam_pdf(student_name, adm_no, stream, form, subject, term, exam_name, exam_date, duration, logo_image, instructions, include_marking_table, include_exam_number, school_name, marking_table_data):
    buffer = BytesIO()
    # Use landscape mode for the page if the table is wide
    # For a table that fits A4 portrait, A4 is fine. If it's too wide, consider A4, landscape
    c = canvas.Canvas(buffer, pagesize=A4) # A4 is portrait by default
    width, height = A4
    y = height - 80

    if logo_image:
        logo = ImageReader(logo_image)
        c.drawImage(logo, 60, y - 10, width=60, height=60, preserveAspectRatio=True)

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, y, EXAM_HEADER)

    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(width / 2, y - 20, school_name.upper())
    c.drawCentredString(width / 2, y - 40, f"{form.upper()} {subject.upper()}")
    c.drawCentredString(width / 2, y - 60, f"{term} – {exam_name}")
    c.drawCentredString(width / 2, y - 80, f"{exam_date} - {duration}")

    c.setFont("Helvetica", 12)
    c.drawString(100, y - 110, f"Name: {student_name}")
    c.drawString(350, y - 110, f"Adm. No: {adm_no}")
    c.drawString(100, y - 130, f"Stream: {stream}")
    if include_exam_number:
        c.drawString(350, y - 130, f"Exam Number: {generate_exam_number()}")

    c.setFont("Helvetica-Bold", 12)
    c.drawString(100, y - 170, "Instructions:")
    c.setFont("Helvetica", 11)
    iy = y - 190
    for line in instructions:
        c.drawString(110, iy, line)
        iy -= 16

    if include_marking_table:
        c.setFont("Helvetica-Bold", 12)
        iy -= 30
        c.drawString(100, iy, "For examiners use only")
        iy -= 25

        # Prepare table data for ReportLab's Table
        table_data = [["SECTION", "QUESTION", "MAXIMUM SCORE", "CANDIDATE’S SCORE"]]
        for row in marking_table_data:
            table_data.append([
                row["SECTION"],
                row["QUESTION"],
                str(row["MAXIMUM SCORE"]),
                str(row["CANDIDATE’S SCORE"])
            ])

        # Create the Table object
        # The width can be adjusted based on desired fit
        # A4 width is 595.27 units. We can set a total width and let it distribute or specify column widths.
        # Let's aim for a table width that fits within the margins.
        # The default width is 450 units, which is a good starting point.
        table = Table(table_data, colWidths=[80, 120, 120, 150]) # You can adjust these widths

        # Define table style
        table_style = TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black), # All borders
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), # Header background
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), # Header font
            ('FONTSIZE', (0,0), (-1,-1), 10), # All cell font size
            ('ALIGN', (0,0), (-1,-1), 'CENTER'), # Center align all cells
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), # Middle vertical align
        ])
        table.setStyle(table_style)

        # Calculate table size and draw it
        table_width, table_height = table.wrapOn(c, width, height)
        
        # Position the table: centered horizontally, below instructions
        x_position = (width - table_width) / 2 # Center the table
        table.drawOn(c, x_position, iy - table_height) # Draw the table

    c.save()
    buffer.seek(0)
    return buffer

# === Streamlit UI ===
st.title("Student Customized Exam Top Page Generator")

school_name = st.text_input("Enter School Name", value="ST. JOSEPH’S BOYS - KITALE")

student_file = st.file_uploader("Upload Excel File with Student Data", type=["xlsx"])
logo_file = st.file_uploader("Upload School Logo (optional)", type=["png", "jpg", "jpeg"])

subject = st.text_input("Subject", placeholder="e.g. Physics")
form = st.selectbox("Form", ["Form 1", "Form 2", "Form 3", "Form 4"])
term = st.text_input("Term", placeholder="e.g. Term 2")
exam_name = st.text_input("Exam Name", placeholder="e.g. Exam 1 or JESMA")
exam_date = st.date_input("Exam Date")
duration = st.text_input("Exam Duration", value="2 HOURS")

include_exam_number = st.checkbox("Include Exam Number", value=True)
include_marking_table = st.checkbox("Include Marking Table", value=True)

st.markdown("### Edit Instructions")
instructions = st.text_area("Instructions", "\n".join(DEFAULT_INSTRUCTIONS), height=200)
instruction_lines = [line.strip() for line in instructions.strip().split("\n") if line.strip()]

st.markdown("### Edit Marking Table")
marking_table_data = st.data_editor(
    pd.DataFrame(DEFAULT_MARKING_TABLE),
    use_container_width=True,
    num_rows="dynamic",
    key="marking_editor"
)

if st.button("Generate Personalized PDFs"):
    if student_file is None:
        st.error("Please upload an Excel file with student data.")
    else:
        with st.spinner("Generating PDFs..."):
            df = pd.read_excel(student_file)
            df.columns = [col.strip().lower() for col in df.columns]

            name_col = next((col for col in df.columns if "name" in col), None)
            adm_col = next((col for col in df.columns if "admission" in col), None)
            stream_col = next((col for col in df.columns if "stream" in col), None)

            if not all([name_col, adm_col, stream_col]):
                st.error("Could not detect necessary columns (name, admission number, stream). Please check your Excel file.")
            else:
                output_folder = "student_exam_pdfs"
                os.makedirs(output_folder, exist_ok=True)

                for _, row in df.iterrows():
                    student_name = row.get(name_col, "Unknown")
                    adm_no = row.get(adm_col, "N/A")
                    stream = row.get(stream_col, "N/A")

                    pdf = generate_exam_pdf(
                        student_name,
                        adm_no,
                        stream,
                        form,
                        subject,
                        term,
                        exam_name,
                        exam_date.strftime("%d %B %Y"),
                        duration,
                        logo_file,
                        instruction_lines,
                        include_marking_table,
                        include_exam_number,
                        school_name,
                        marking_table_data.to_dict("records")
                    )

                    safe_name = student_name.replace(" ", "_").replace("/", "_")
                    with open(os.path.join(output_folder, f"{safe_name}_{adm_no}.pdf"), "wb") as f:
                        f.write(pdf.read())

                zip_path = shutil.make_archive("Personalized_Exam_Top_Pages", 'zip', output_folder)
                st.success("Done! Download below:")
                st.download_button("Download ZIP", open(zip_path, "rb"), "Personalized_Exam_Top_Pages.zip")
                shutil.rmtree(output_folder)