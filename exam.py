import os
import random
import string
import shutil
import pandas as pd
from datetime import datetime
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
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

# Default values for easier UI population
DEFAULT_COL_WIDTHS_LIST = [80, 120, 120, 150] # For direct number inputs
DEFAULT_UNIFORM_ROW_HEIGHT = 25 # For uniform height number input

def generate_exam_number(length=10):
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

def generate_exam_pdf(
    student_name, adm_no, stream, form, subject, term, exam_name, exam_date,
    duration, logo_image, instructions, include_marking_table, include_exam_number,
    school_name, marking_table_data, col_widths_list, row_heights_val
):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    y = height - 80

    # Draw logo if provided
    if logo_image:
        logo = ImageReader(logo_image)
        c.drawImage(logo, 60, y - 10, width=60, height=60, preserveAspectRatio=True)

    # Header information
    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, y, EXAM_HEADER)

    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(width / 2, y - 20, school_name.upper())
    c.drawCentredString(width / 2, y - 40, f"{form.upper()} {subject.upper()}")
    c.drawCentredString(width / 2, y - 60, f"{term} – {exam_name}")
    c.drawCentredString(width / 2, y - 80, f"{exam_date} - {duration}")

    # Student details
    c.setFont("Helvetica", 12)
    c.drawString(100, y - 110, f"Name: {student_name}")
    c.drawString(350, y - 110, f"Adm. No: {adm_no}")
    c.drawString(100, y - 130, f"Stream: {stream}")
    if include_exam_number:
        c.drawString(350, y - 130, f"Exam Number: {generate_exam_number()}")

    # Instructions
    c.setFont("Helvetica-Bold", 12)
    c.drawString(100, y - 170, "Instructions:")
    c.setFont("Helvetica", 11)
    iy = y - 190
    for line in instructions:
        c.drawString(110, iy, line)
        iy -= 16 # Adjust line spacing for instructions

    # Marking table (For examiners use only)
    if include_marking_table:
        c.setFont("Helvetica-Bold", 12)
        iy -= 30 # Space before "For examiners use only"
        c.drawString(100, iy, "For examiners use only")
        iy -= 25 # Space before table

        # Prepare table data for ReportLab's Table
        table_data = [["SECTION", "QUESTION", "MAXIMUM SCORE", "CANDIDATE’S SCORE"]]
        for row in marking_table_data:
            table_data.append([
                row["SECTION"],
                row["QUESTION"],
                str(row["MAXIMUM SCORE"]),
                str(row["CANDIDATE’S SCORE"])
            ])

        # Create the Table object with dynamic column widths AND row heights
        table = Table(table_data, colWidths=col_widths_list, rowHeights=row_heights_val)

        # Define table style
        table_style = TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black), # All borders
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), # Header row background
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), # Header font
            ('FONTSIZE', (0,0), (-1,-1), 10), # Font size for all cells
            ('ALIGN', (0,0), (-1,-1), 'CENTER'), # Center align all cells horizontally
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), # Center align all cells vertically
        ])
        table.setStyle(table_style)

        # Calculate table size and draw it
        table_width, table_height = table.wrapOn(c, width, height)
        
        # Position the table: centered horizontally, below instructions
        x_position = (width - table_width) / 2 # Calculate X to center the table
        table.drawOn(c, x_position, iy - table_height) # Draw the table on the canvas

    c.save()
    buffer.seek(0)
    return buffer

# === Streamlit UI ===
st.title("Student Customized Exam Top Page Generator")

st.markdown("""
Welcome! This tool helps you create personalized exam top pages for your students.
Just fill in the details below, upload your student list, and generate PDFs!
""")

st.header("School and Exam Details")
school_name = st.text_input("Enter School Name", value="ST. JOSEPH’S BOYS - KITALE")

st.subheader("Upload Files")
student_file = st.file_uploader("Upload an Excel File with Student Data", type=["xlsx"])
logo_file = st.file_uploader("Upload Your School Logo (Optional)", type=["png", "jpg", "jpeg"])

st.subheader("Exam Information")
subject = st.text_input("Subject", placeholder="e.g., Physics")
form = st.selectbox("Form", ["Form 1", "Form 2", "Form 3", "Form 4"])
term = st.text_input("Term", placeholder="e.g., Term 2")
exam_name = st.text_input("Exam Name", placeholder="e.g., Mid-Term Exam or JESMA")
exam_date = st.date_input("Exam Date")
duration = st.text_input("Exam Duration", value="2 HOURS")

st.subheader("Page Options")
include_exam_number = st.checkbox("Include an Exam Number on the page", value=True)
include_marking_table = st.checkbox("Include a Marking Table (for examiners)", value=True)

st.markdown("---") # Separator

st.header("Customize Instructions")
instructions = st.text_area("Type your instructions here:", "\n".join(DEFAULT_INSTRUCTIONS), height=200)
instruction_lines = [line.strip() for line in instructions.strip().split("\n") if line.strip()]

st.markdown("---") # Separator

st.header("Customize Marking Table")
st.markdown("You can edit the content of the marking table directly below. Add or remove rows as needed.")
marking_table_data = st.data_editor(
    pd.DataFrame(DEFAULT_MARKING_TABLE),
    use_container_width=True,
    num_rows="dynamic",
    key="marking_editor"
)

st.subheader("Adjust Table Size (Columns & Rows)")
st.info("Here, you can fine-tune how the marking table looks on the page.")

# Column Widths (Direct Number Inputs for each column)
st.markdown("#### Column Widths")
st.markdown("""
    Adjust the width of each column. These values are in 'units of measure' for printing. Think of them like tiny pixels.
    The total width of your table should fit within the page margins. For a standard A4 page, keep the sum of all widths under **470-500** units.
""")

col_widths_inputs = []
col_names = ["SECTION", "QUESTION", "MAXIMUM SCORE", "CANDIDATE’S SCORE"] # Use actual column names for clarity
for i, default_width in enumerate(DEFAULT_COL_WIDTHS_LIST):
    width = st.number_input(
        f"Width for the '{col_names[i]}' column:",
        min_value=10, # Minimum sensible width
        max_value=300, # Maximum sensible width
        value=default_width,
        key=f"col_width_{i}",
        help="A larger number makes the column wider. Try adjusting and generating to see the effect."
    )
    col_widths_inputs.append(width)

# Row Heights (Uniform vs. Advanced)
st.markdown("#### Row Heights")
st.markdown("You can set a single height for all rows, or use advanced mode to set individual heights for each row.")

uniform_row_height_input = st.number_input(
    "Set a **Uniform Height** for all rows:",
    min_value=10,
    max_value=100,
    value=DEFAULT_UNIFORM_ROW_HEIGHT,
    key="uniform_row_height",
    help="This height will be applied to every row in the table. If content is too tall, the row will expand automatically."
)

# Toggle for advanced row height control
advanced_row_heights_enabled = st.checkbox(
    "Enable **Advanced (Individual) Row Height** Control",
    value=False,
    key="advanced_row_heights_checkbox"
)

row_heights_val = uniform_row_height_input # Default to uniform height

if advanced_row_heights_enabled:
    total_rows_in_table = len(marking_table_data) + 1 # +1 for the header row
    st.warning(f"""
        You've enabled advanced row height control.
        Your table has **{total_rows_in_table} rows** (this includes the top header row and all data rows).
        If you enter individual heights, you *must* provide **{total_rows_in_table} numbers**, separated by commas.
    """)
    
    advanced_row_heights_str = st.text_input(
        "Enter individual row heights (e.g., '25,20,20,20,20,20,20,25'):",
        value="", # Start empty, or you could pre-fill with defaults for all rows
        help=f"Type {total_rows_in_table} numbers separated by commas. The first number is for the header, then each subsequent number is for the data rows below it."
    )

    if advanced_row_heights_str:
        try:
            parsed_heights = [float(x.strip()) for x in advanced_row_heights_str.split(',')]
            if len(parsed_heights) == total_rows_in_table:
                row_heights_val = parsed_heights
            else:
                st.error(f"**Error:** You entered {len(parsed_heights)} values, but the table has {total_rows_in_table} rows. Please provide exactly {total_rows_in_table} numbers, separated by commas.")
                # Fallback to uniform height if error
                row_heights_val = uniform_row_height_input
        except ValueError:
            st.error("**Error:** Invalid input for individual row heights. Please make sure you're entering numbers separated by commas (e.g., '25,20,20').")
            # Fallback to uniform height if error
            row_heights_val = uniform_row_height_input
    else:
        st.info("No individual row heights entered. The 'Uniform Height' will be used instead.")
        row_heights_val = uniform_row_height_input # Explicitly set to uniform if text box is empty

st.markdown("---") # Separator

if st.button("Generate Personalized PDFs"):
    if student_file is None:
        st.error("Oops! Please upload an Excel file with your student data before generating PDFs.")
    else:
        with st.spinner("Generating PDFs... This might take a moment if you have many students."):
            df = pd.read_excel(student_file)
            # Standardize column names for robust detection (e.g., "Name", "NAME", "name " all work)
            df.columns = [col.strip().lower() for col in df.columns]

            # Attempt to find common column names for student data
            name_col = next((col for col in df.columns if "name" in col), None)
            adm_col = next((col for col in df.columns if "admission" in col or "adm" in col), None)
            stream_col = next((col for col in df.columns if "stream" in col), None)

            if not all([name_col, adm_col, stream_col]):
                st.error("""
                    **Important:** We couldn't find the necessary columns in your Excel file.
                    Please make sure your file has columns named something like:
                    - 'Name' (for student names)
                    - 'Admission No.' or 'Adm' (for admission numbers)
                    - 'Stream'
                    Double-check your Excel file and try again!
                """)
            else:
                output_folder = "student_exam_pdfs"
                # Clean up any old output folder before starting fresh
                if os.path.exists(output_folder):
                    shutil.rmtree(output_folder)
                os.makedirs(output_folder, exist_ok=True)

                for index, row in df.iterrows():
                    student_name = str(row.get(name_col, "Unknown")).strip()
                    adm_no = str(row.get(adm_col, "N/A")).strip()
                    stream = str(row.get(stream_col, "N/A")).strip()

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
                        marking_table_data.to_dict("records"),
                        col_widths_inputs, # Pass the list from number inputs
                        row_heights_val # Pass the (uniform or list) row heights
                    )

                    # Create a safe filename for the PDF
                    safe_name = student_name.replace(" ", "_").replace("/", "_").replace("\\", "_")
                    safe_adm_no = str(adm_no).replace(" ", "_").replace("/", "_").replace("\\", "_")
                    filename = f"{safe_name}_{safe_adm_no}.pdf"
                    
                    with open(os.path.join(output_folder, filename), "wb") as f:
                        f.write(pdf.read())

                zip_path = shutil.make_archive("Personalized_Exam_Top_Pages", 'zip', output_folder)
                st.success("🎉 **Success!** Your personalized PDFs are ready!")
                
                # Provide download button for the zip file
                with open(zip_path, "rb") as fp:
                    st.download_button(
                        label="⬇️ Download All PDFs (ZIP File)",
                        data=fp.read(),
                        file_name="Personalized_Exam_Top_Pages.zip",
                        mime="application/zip"
                    )
                
                # Clean up temporary files
                st.info("Temporary files are being cleaned up now.")
                shutil.rmtree(output_folder)
                os.remove(zip_path)