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
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import streamlit as st

# === Config ===
EXAM_HEADER = "Kenya Certificate of Secondary Examinations"

DEFAULT_INSTRUCTIONS = [
    "1. Write your name and index number in the spaces provided above.",
    "2. Sign and write the date of examination in the spaces provided.",
    "3. This paper consists of TWO sections: Section I and Section II.",
    "4. Answer ALL the questions in Section I and only Five from Section II.",
    "5. All answers and working must be written on the question paper in the spaces provided below each question.",
    "6. Show all the steps in your calculations, giving answers at each stage in the spaces provided below each question.",
    "7. Candidates should answer all questions in English.",
    "8. Candidates should check to ascertain that all pages are printed as indicated and that no questions are missing."
]

# Default values for the new table customization
DEFAULT_SECTION_1_QNS = 16
DEFAULT_SECTION_2_QNS = 8
DEFAULT_SECTION_1_TITLE = "SECTION I"
DEFAULT_SECTION_2_TITLE = "SECTION II"
DEFAULT_INCLUDE_GRAND_TOTAL = True

# Default scale for the new complex table - Increased slightly for better default width
DEFAULT_TABLE_SCALE = 1.1 # 1.0 means original size, 1.1 means 10% larger, 0.9 means 10% smaller

# Default data for the custom marking table
DEFAULT_CUSTOM_TABLE_DATA = pd.DataFrame([
    {"Section": "A", "Question": "1 - 11", "Maximum Score": 25},
    {"Section": "B", "Question": "12", "Maximum Score": 11},
    {"Section": "B", "Question": "13", "Maximum Score": 11},
    {"Section": "B", "Question": "14", "Maximum Score": 11},
    {"Section": "B", "Question": "15", "Maximum Score": 10},
    {"Section": "B", "Question": "16", "Maximum Score": 12},
    {"Section": "TOTAL SCORE", "Question": "", "Maximum Score": 80},
])

def generate_exam_number(length=10):
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

def draw_kcse_standard_marking_table(c, examiner_text_y, table_scale, width, height,
                                       section_1_questions, section_2_questions,
                                       section_1_title, section_2_title, include_grand_total):
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']
    normal_style.fontSize = 9

    # --- Table Dimensions ---
    q_cell_width = 20 * table_scale
    total_cell_width = 30 * table_scale
    gt_label_width = 40 * table_scale
    gt_score_width = 60 * table_scale
    row_height = 25 * table_scale 
    
    # --- Table 1: Section I ---
    col_widths_s1 = [q_cell_width] * section_1_questions + [total_cell_width]
    
    table_data_s1 = [
        [section_1_title] + [''] * section_1_questions,
        [str(i) for i in range(1, section_1_questions + 1)] + ['Total'],
        [''] * (section_1_questions + 1)
    ]
    
    table_s1 = Table(table_data_s1, colWidths=col_widths_s1, rowHeights=row_height)
    table_s1_style = TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN', (0,0), (section_1_questions,0)),
        ('BACKGROUND', (0,0), (section_1_questions,0), colors.lightgrey),
        ('FONTNAME', (0,0), (section_1_questions,0), 'Helvetica-Bold'),
        ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
    ])
    table_s1.setStyle(table_s1_style)

    table_x_position = 60
    table_s1_width, table_s1_height = table_s1.wrapOn(c, width, height)
    table_s1_y_position = examiner_text_y - table_s1_height - 10 # Add 10 points space below 'For Examiner's Use Only'
    table_s1.drawOn(c, table_x_position, table_s1_y_position)

    # --- Table 2: Section II ---
    sec2_q_start = section_1_questions + 1
    sec2_q_end = sec2_q_start + section_2_questions - 1

    col_widths_s2 = [q_cell_width] * section_2_questions + [total_cell_width]

    table_data_s2 = [
        [section_2_title] + [''] * section_2_questions,
        [str(i) for i in range(sec2_q_start, sec2_q_end + 1)] + ['Total'],
        [''] * (section_2_questions + 1)
    ]

    table_s2 = Table(table_data_s2, colWidths=col_widths_s2, rowHeights=row_height)
    table_s2_style = TableStyle([
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN', (0,0), (section_2_questions,0)),
        ('BACKGROUND', (0,0), (section_2_questions,0), colors.lightgrey),
        ('FONTNAME', (0,0), (section_2_questions,0), 'Helvetica-Bold'),
        ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
    ])
    table_s2.setStyle(table_s2_style)

    table_s2_width, table_s2_height = table_s2.wrapOn(c, width, height)
    # Add 15 points space between Section I and Section II tables
    table_s2_y_position = table_s1_y_position - table_s2_height - 15 
    table_s2.drawOn(c, table_x_position, table_s2_y_position)

    # --- Table 3: Grand Total (if included) ---
    if include_grand_total:
        col_widths_gt = [gt_label_width, gt_score_width]
        
        table_data_gt = [
            ['GRAND', ''],
            ['TOTAL', ''],
            ['', '']
        ]

        table_gt = Table(table_data_gt, colWidths=col_widths_gt, rowHeights=[row_height, row_height, row_height])
        table_gt_style = TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('SPAN', (0,0), (0,1)), # 'GRAND' and 'TOTAL' label vertical span
            ('SPAN', (1,0), (1,2)), # Grand Total score box vertical span
            ('FONTNAME', (0,0), (0,1), 'Helvetica-Bold'),
        ])
        table_gt.setStyle(table_gt_style)

        table_gt_width, table_gt_height = table_gt.wrapOn(c, width, height)
        
        # Position it to the right of Section II table, aligned at its top edge
        table_gt_x_position = table_x_position + table_s2_width + 10 
        table_gt_y_position = table_s2_y_position # Align tops of Section II and Grand Total tables
        table_gt.drawOn(c, table_gt_x_position, table_gt_y_position)

    # Return the lowest Y position used by the tables
    return table_s2_y_position # or table_gt_y_position if it extends lower

def draw_custom_marking_table(c, custom_table_df, examiner_text_y, table_scale, width, height):
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']
    normal_style.fontSize = 9

    # Header for the custom table
    table_header = ["SECTION", "QUESTION", "MAXIMUM SCORE", "CANDIDATE'S SCORE"]
    
    table_rows = []
    # Add header
    table_rows.append(table_header)

    # Add data rows
    for index, row in custom_table_df.iterrows():
        section_val = str(row.get("Section", "")).strip()
        question_val = str(row.get("Question", "")).strip()
        max_score_val = str(row.get("Maximum Score", "")).strip()
        table_rows.append([section_val, question_val, max_score_val, '']) # Candidate's score is empty

    # Column widths can be dynamic based on content or fixed for better appearance
    # Adjusted widths to better fit the example image
    col_widths = [80 * table_scale, 80 * table_scale, 60 * table_scale, 80 * table_scale]
    row_height = 25 * table_scale

    table = Table(table_rows, colWidths=col_widths, rowHeights=row_height)
    
    table_style_list = [
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), # Header background
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), # Header font bold
    ]

    # Apply specific styles for "TOTAL SCORE" row
    for i, row_data in enumerate(table_rows):
        if i > 0 and row_data[0].upper() == "TOTAL SCORE": # Check the 'Section' column for 'TOTAL SCORE'
            # Span 'SECTION' and 'QUESTION' cells for 'TOTAL SCORE' row
            table_style_list.append(('SPAN', (0, i), (1, i)))
            table_style_list.append(('ALIGN', (0, i), (1, i), 'CENTER')) # Center merged cell
            table_style_list.append(('FONTNAME', (0, i), (-1, i), 'Helvetica-Bold')) # Bold total row
            table_style_list.append(('ALIGN', (2, i), (2, i), 'RIGHT')) # Align max score to right

    table.setStyle(TableStyle(table_style_list))

    table_x_position = 60
    table_width, table_height = table.wrapOn(c, width, height)
    table_y_position = examiner_text_y - table_height - 10 # Add 10 points space below 'For Examiner's Use Only'
    table.drawOn(c, table_x_position, table_y_position)

    return table_y_position # Return the lowest Y position used by the table


def generate_exam_pdf(
    student_name, adm_no, stream, form, subject, term, exam_name, exam_date,
    duration, logo_image, raw_instructions, marking_table_style, include_exam_number,
    school_name, paper_code, total_pages_count, table_scale,
    section_1_questions, section_2_questions, section_1_title, section_2_title, include_grand_total,
    custom_table_df, # New parameter for custom table data
    prefill_student_details 
):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    # Page Number
    c.setFont("Helvetica", 9)
    c.drawString(width - 70, height - 30, "1")

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
    
    # Paper Code and Subject Title
    c.setFont("Helvetica-Bold", 12)
    c.drawString(60, y - 40, paper_code)
    c.drawCentredString(width / 2, y - 40, f"{form.upper()} {subject.upper()}")
    c.drawCentredString(width / 2, y - 60, f"{term} – {exam_name}")
    
    # Time
    c.setFont("Helvetica-Bold", 11)
    c.drawString(width - 150, y - 60, f"TIME: {duration}")

    # --- Student details (Dashed Lines Under Names) ---
    c.setFont("Helvetica", 11)
    
    line_start_x_left_section = 60
    line_end_x_left_section = 320 
    line_start_x_right_section = 330
    line_end_x_right_section = width - 60 
    line_thickness = 0.5 

    c.setLineWidth(line_thickness)
    c.setDash(1, 2) # Set dashed line pattern (1 unit on, 2 units off)

    # Name
    current_y = y - 90
    c.drawString(line_start_x_left_section, current_y, "Name:")
    name_label_width = c.stringWidth("Name:", "Helvetica", 11)
    # Draw dashed line for name
    c.line(line_start_x_left_section + name_label_width + 5, current_y - 2, line_end_x_left_section, current_y - 2)
    # Draw student name over the dashed line ONLY if prefill_student_details is True
    if prefill_student_details:
        c.drawString(line_start_x_left_section + name_label_width + 5, current_y, student_name)

    # Index No.
    c.drawString(line_start_x_right_section, current_y, "Index No.:")
    index_label_width = c.stringWidth("Index No.:", "Helvetica", 11)
    # Draw dashed line for index
    c.line(line_start_x_right_section + index_label_width + 5, current_y - 2, line_end_x_right_section, current_y - 2)
    # Draw admission number over the dashed line ONLY if prefill_student_details is True
    if prefill_student_details:
        c.drawString(line_start_x_right_section + index_label_width + 5, current_y, adm_no)

    # School
    current_y -= 20 
    c.drawString(line_start_x_left_section, current_y, "School:")
    school_label_width = c.stringWidth("School:", "Helvetica", 11)
    # Draw dashed line for school
    c.line(line_start_x_left_section + school_label_width + 5, current_y - 2, line_end_x_left_section, current_y - 2)
    # School name is always pre-filled
    c.drawString(line_start_x_left_section + school_label_width + 5, current_y, school_name)

    # Candidate's Signature
    c.drawString(line_start_x_right_section, current_y, "Candidate's Signature:")
    sig_label_width = c.stringWidth("Candidate's Signature:", "Helvetica", 11)
    # Draw dashed line for signature
    c.line(line_start_x_right_section + sig_label_width + 5, current_y - 2, line_end_x_right_section, current_y - 2)
    
    # Stream
    current_y -= 20 
    c.drawString(line_start_x_left_section, current_y, "Stream:")
    stream_label_width = c.stringWidth("Stream:", "Helvetica", 11)
    # Draw dashed line for stream
    c.line(line_start_x_left_section + stream_label_width + 5, current_y - 2, line_end_x_left_section, current_y - 2)
    # Draw stream over the dashed line ONLY if prefill_student_details is True
    if prefill_student_details:
        c.drawString(line_start_x_left_section + stream_label_width + 5, current_y, stream)

    # Date
    current_y -= 20 
    c.drawString(line_start_x_left_section, current_y, "Date:")
    date_label_width = c.stringWidth("Date:", "Helvetica", 11)
    # Draw dashed line for date
    c.line(line_start_x_left_section + date_label_width + 5, current_y - 2, line_end_x_left_section, current_y - 2)
    # Date is always pre-filled
    c.drawString(line_start_x_left_section + date_label_width + 5, current_y, exam_date)

    # Exam Number (Bolded)
    if include_exam_number:
        c.setFont("Helvetica-Bold", 11) # Bold for exam number
        c.drawString(line_start_x_right_section, current_y, f"Exam Number: {generate_exam_number()}")
        c.setFont("Helvetica", 11) # Reset font

    # Reset line styles
    c.setDash() # Turn off dashes
    c.setLineWidth(1)

    # --- Instructions Handling (Dynamic First Instruction) ---
    current_instructions = list(raw_instructions) # Create a mutable copy

    if prefill_student_details:
        # Modify the first instruction if pre-filling
        if len(current_instructions) > 0 and "Write your name and index number" in current_instructions[0]:
            current_instructions[0] = "1. Verify your name and index number in the spaces provided above."
    # Else, if not pre-filling, the original instruction "Write your name..." remains.

    c.setFont("Helvetica-Bold", 11)
    iy = current_y - 40 # Adjust starting Y for instructions based on where student details ended
    c.drawString(60, iy, "INSTRUCTIONS TO CANDIDATES")
    
    c.setFont("Helvetica", 10)
    iy -= 20 # Initial vertical space before first instruction
    for line in current_instructions: # Use the modified instructions list
        p = Paragraph(line, getSampleStyleSheet()['Normal'])
        p_width = width - 120 # 60 from each side
        
        # Corrected: p.wrapOn returns (width, height), so we need the second element (height)
        _, p_height = p.wrapOn(c, p_width, height) 
        
        # Check if drawing this paragraph would go off the page
        if iy - p_height < 60: # If it's too close to the bottom margin (e.g., 60 points)
            c.showPage() # Start a new page
            # Re-draw page number for new pages (basic for now, more complex if many pages)
            c.setFont("Helvetica", 9)
            c.drawString(width - 70, height - 30, "Page X") 
            iy = height - 60 # Reset y for new page
            c.setFont("Helvetica", 10)

        p.drawOn(c, 60, iy - p_height) # Draw paragraph at current y, adjusted for its height
        iy -= (p_height + 8) # Move y down for next paragraph, adding 8 points extra space


    # --- For Examiner's Use Only Table ---
    if marking_table_style != "None": # Only draw if a marking table style is selected
        c.setFont("Helvetica-Bold", 12)
        examiner_text_y = iy - 20 # Adjusted vertical spacing from instructions
        c.drawString(60, examiner_text_y, "For Examiner's Use Only")

        if marking_table_style == "K.C.S.E. Standard (Section I, Section II, Grand Total)":
            draw_kcse_standard_marking_table(c, examiner_text_y, table_scale, width, height,
                                              section_1_questions, section_2_questions,
                                              section_1_title, section_2_title, include_grand_total)
        elif marking_table_style == "Customized Score Sheet":
            if custom_table_df is not None and not custom_table_df.empty:
                draw_custom_marking_table(c, custom_table_df, examiner_text_y, table_scale, width, height)
            else:
                st.warning("No data provided for custom marking table. Skipping table generation.")


    c.save()
    buffer.seek(0)
    return buffer

# === Streamlit UI ===
st.title("Student Customized Exam Top Page Generator")

st.markdown("""
Welcome! This tool helps you create personalized exam top pages for your students.
Just fill in the details below, upload your student list, and generate PDFs!
""")

st.header("School and Exam Details (for PDF Generation)")
school_name = st.text_input("Enter School Name", value="ST. JOSEPH’S BOYS - KITALE")

st.subheader("Upload Files")
student_file = st.file_uploader("Upload an Excel File with Student Data", type=["xlsx"], key="student_excel_upload")
logo_file = st.file_uploader("Upload Your School Logo (Optional)", type=["png", "jpg", "jpeg"], key="school_logo_upload")

st.subheader("Exam Information")
paper_code = st.text_input("Paper Code (e.g., 121/1)", value="121/1", help="This appears at the top left of the exam paper.")
subject = st.text_input("Subject", placeholder="e.g., MATHEMATICS")
form = st.selectbox("Form", ["Form 1", "Form 2", "Form 3", "Form 4"])
term = st.text_input("Term", placeholder="e.g., Term 2")
exam_name = st.text_input("Exam Name", placeholder="e.g., Paper 1 or Mid-Term Exam")
exam_date = st.date_input("Exam Date", value=datetime.now())
duration = st.text_input("Exam Duration", value="2 HOURS")

st.subheader("Page Options")
prefill_student_details = st.checkbox("Pre-fill student name, index number, and stream", value=True, help="If checked, names, index numbers, and streams from your Excel file will be printed. If unchecked, lines will be provided for students to write them in.")
include_exam_number = st.checkbox("Include an Exam Number on the page", value=True)

st.markdown("---")

st.header("Customize Instructions")
instructions = st.text_area("Type your instructions here:", "\n".join(DEFAULT_INSTRUCTIONS), height=200)
instruction_lines = [line.strip() for line in instructions.strip().split("\n") if line.strip()]

st.markdown("---")

st.header("Customize Marking Table (Examiner's Use Only)")

marking_table_style = st.radio(
    "Choose Marking Table Style:",
    ("K.C.S.E. Standard (Section I, Section II, Grand Total)", "Customized Score Sheet", "None"),
    index=0, # Default to K.C.S.E. Standard
    help="Select the layout for the 'For Examiner's Use Only' table."
)

section_1_questions = DEFAULT_SECTION_1_QNS
section_2_questions = DEFAULT_SECTION_2_QNS
section_1_title = DEFAULT_SECTION_1_TITLE
section_2_title = DEFAULT_SECTION_2_TITLE
include_grand_total = DEFAULT_INCLUDE_GRAND_TOTAL
custom_table_df = None # Initialize to None

if marking_table_style == "K.C.S.E. Standard (Section I, Section II, Grand Total)":
    st.markdown("""
        This option generates the standard K.C.S.E. style marking table with separate sections and an optional grand total.
    """)
    section_1_title = st.text_input("Section I Title:", value=DEFAULT_SECTION_1_TITLE, key="s1_title_kcse")
    section_1_questions = st.number_input(
        "Number of Questions in Section I (e.g., 16):",
        min_value=1, value=DEFAULT_SECTION_1_QNS, step=1, key="s1_qns_kcse",
        help="This defines the number of columns for questions in Section I."
    )

    section_2_title = st.text_input("Section II Title:", value=DEFAULT_SECTION_2_TITLE, key="s2_title_kcse")
    section_2_questions = st.number_input(
        "Number of Questions in Section II (e.g., 8):",
        min_value=1, value=DEFAULT_SECTION_2_QNS, step=1, key="s2_qns_kcse",
        help="This defines the number of columns for questions in Section II."
    )

    include_grand_total = st.checkbox("Include 'Grand Total' Section", value=DEFAULT_INCLUDE_GRAND_TOTAL, key="include_gt_kcse")

elif marking_table_style == "Customized Score Sheet":
    st.markdown("""
        Create a fully customized marking table by editing the rows below.
        Add/delete rows and fill in 'Section', 'Question', and 'Maximum Score'.
        The 'Candidate's Score' column will be automatically added as empty.
    """)
    
    # Use st.data_editor for interactive table editing
    custom_table_df = st.data_editor(
        DEFAULT_CUSTOM_TABLE_DATA,
        column_config={
            "Section": st.column_config.TextColumn("Section", help="e.g., A, B, or SECTION I"),
            "Question": st.column_config.TextColumn("Question", help="e.g., 1-11, 12, or TOTAL SCORE"),
            "Maximum Score": st.column_config.NumberColumn("Maximum Score", help="Points for this question/section"),
        },
        num_rows="dynamic",
        key="custom_marking_table_editor"
    )
    # Ensure custom_table_df is not None if data_editor is used
    if custom_table_df.empty:
        st.warning("Please add rows to your custom marking table.")


# Table Size Scale (applies to both types)
table_scale = st.slider(
    "Adjust Table Size (Overall Scale)",
    min_value=0.5,
    max_value=1.5,
    value=DEFAULT_TABLE_SCALE,
    step=0.05,
    help="Drag the slider to make the entire marking table larger or smaller on the page. This adjusts all its cells uniformly."
)

st.markdown("---")

if st.button("Generate Personalized PDFs", key="generate_pdfs_button"):
    if student_file is None:
        st.error("Oops! Please upload an Excel file with your student data before generating PDFs.")
    elif marking_table_style == "Customized Score Sheet" and (custom_table_df is None or custom_table_df.empty):
        st.error("Please add some data to your 'Customized Score Sheet' marking table or select another style.")
    else:
        with st.spinner("Generating PDFs... This might take a moment if you have many students."):
            df = pd.read_excel(student_file)
            df.columns = [col.strip().lower() for col in df.columns]

            name_col = next((col for col in df.columns if "name" in col), None)
            adm_col = next((col for col in df.columns if "admission" in col or "adm" in col or "index" in col), None)
            stream_col = next((col for col in df.columns if "stream" in col), None)

            if not all([name_col, adm_col]):
                st.error("""
                    **Important:** We couldn't find the necessary columns in your Excel file.
                    Please make sure your file has columns named something like:
                    - 'Name' (for student names)
                    - 'Admission No.' or 'Adm' or 'Index No.' (for index numbers)
                    - 'Stream' (for student streams)
                    Double-check your Excel file and try again!
                """)
            else:
                output_folder = "student_exam_pdfs"
                if os.path.exists(output_folder):
                    shutil.rmtree(output_folder)
                os.makedirs(output_folder, exist_ok=True)

                for index, row in df.iterrows():
                    student_name = str(row.get(name_col, "Unknown")).strip()
                    adm_no = str(row.get(adm_col, "N/A")).strip()
                    stream = str(row.get(stream_col, "N/A")).strip() if stream_col else "N/A"

                    pdf = generate_exam_pdf(
                        student_name=student_name,
                        adm_no=adm_no,
                        stream=stream,
                        form=form,
                        subject=subject,
                        term=term,
                        exam_name=exam_name,
                        exam_date=datetime.strftime(exam_date, "%d %B %Y"),
                        duration=duration,
                        logo_image=logo_file,
                        raw_instructions=instruction_lines,
                        marking_table_style=marking_table_style, # Pass the chosen style
                        include_exam_number=include_exam_number,
                        school_name=school_name,
                        paper_code=paper_code,
                        total_pages_count=1,
                        table_scale=table_scale,
                        section_1_questions=section_1_questions,
                        section_2_questions=section_2_questions,
                        section_1_title=section_1_title,
                        section_2_title=section_2_title,
                        include_grand_total=include_grand_total,
                        custom_table_df=custom_table_df, # Pass custom table data
                        prefill_student_details=prefill_student_details
                    )

                    safe_name = student_name.replace(" ", "_").replace("/", "_").replace("\\", "_")
                    safe_adm_no = str(adm_no).replace(" ", "_").replace("/", "_").replace("\\", "_")
                    filename = f"{safe_name}_{safe_adm_no}.pdf"
                    
                    with open(os.path.join(output_folder, filename), "wb") as f:
                        f.write(pdf.read())

                zip_path = shutil.make_archive("Personalized_Exam_Top_Pages", 'zip', output_folder)
                st.success("🎉 **Success!** Your personalized PDFs are ready!")
                
                with open(zip_path, "rb") as fp:
                    st.download_button(
                        label="⬇️ Download All PDFs (ZIP File)",
                        data=fp.read(),
                        file_name="Personalized_Exam_Top_Pages.zip",
                        mime="application/zip"
                    )
                
                st.info("Temporary files are being cleaned up now.")
                shutil.rmtree(output_folder)
                os.remove(zip_path)