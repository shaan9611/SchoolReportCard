import streamlit as st
import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io

st.set_page_config(page_title="Aziz Public School - Portal", layout="wide")

SUBJECTS_FILE = "subjects.xlsx"

# --- LOCAL EXCEL FILE DATABASE FUNCTIONS ---
def init_subject_file():
    """Creates a local excel sheet for subjects if it doesn't exist yet"""
    if not os.path.exists(SUBJECTS_FILE):
        defaults = [
            {"Class": "9th", "Subject": "English"}, {"Class": "9th", "Subject": "Hindi"},
            {"Class": "9th", "Subject": "Science"}, {"Class": "9th", "Subject": "Maths"},
            {"Class": "9th", "Subject": "SST"}, {"Class": "1st", "Subject": "English"},
            {"Class": "1st", "Subject": "Hindi"}, {"Class": "1st", "Subject": "Maths"},
            {"Class": "1st", "Subject": "SST"}, {"Class": "1st", "Subject": "URDU"}
        ]
        df = pd.DataFrame(defaults)
        df.to_excel(SUBJECTS_FILE, index=False)

def get_subjects_by_class(class_name):
    init_subject_file()
    df = pd.read_excel(SUBJECTS_FILE)
    df['Class'] = df['Class'].astype(str).str.strip()
    df['Subject'] = df['Subject'].astype(str).str.strip()
    
    # Filter for the selected class
    filtered_df = df[df['Class'] == str(class_name).strip()]
    # Create an 'id' column using the index so the delete buttons work
    filtered_df = filtered_df.reset_index().rename(columns={'index': 'id'})
    return filtered_df

def add_subject_to_db(class_name, name):
    init_subject_file()
    df = pd.read_excel(SUBJECTS_FILE)
    
    # Check if it already exists
    exists = df[(df['Class'].astype(str).str.strip() == str(class_name).strip()) & 
                (df['Subject'].astype(str).str.strip() == str(name).strip())]
    
    if not exists.empty:
        st.error(f"Subject '{name.strip()}' already exists for Class '{class_name.strip()}'.")
    else:
        new_row = pd.DataFrame([{"Class": str(class_name).strip(), "Subject": str(name).strip()}])
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_excel(SUBJECTS_FILE, index=False)
        st.success(f"Added '{name.strip()}' to Class '{class_name.strip()}'!")

def remove_subject_from_db(row_index):
    df = pd.read_excel(SUBJECTS_FILE)
    df = df.drop(index=int(row_index))
    df.to_excel(SUBJECTS_FILE, index=False)

init_subject_file()

# --- PREMIUM HEADER-CORRECTED PDF GENERATION ENGINE ---
def generate_pdf(student_data, subjects_list, school_details):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    cred_style = ParagraphStyle('Creds', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#002060'), alignment=1)
    
    # New custom centered bold style for government recognition text block
    rec_center_style = ParagraphStyle('RecCenter', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#000000'), alignment=1, spaceAfter=1)
    
    title_style = ParagraphStyle('SchoolTitle', fontName='Helvetica-Bold', fontSize=32, textColor=colors.HexColor('#002060'), alignment=0, leading=32)
    addr_style = ParagraphStyle('SchoolAddr', fontName='Helvetica', fontSize=10, textColor=colors.HexColor('#000000'), alignment=0, leading=12)
    
    rep_style = ParagraphStyle('RepCard', fontName='Helvetica-Bold', fontSize=13, textColor=colors.white, alignment=1)
    body_bold = ParagraphStyle('BBold', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#002060'))
    body_norm = ParagraphStyle('BNorm',fontName='Helvetica-Bold',fontSize=10,leading=12,wordWrap='LTR',textColor=colors.HexColor('#333333'))
    sub_row_style = ParagraphStyle('SubRow', fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#111111'))
    th_style = ParagraphStyle('THead',fontName='Helvetica-Bold',fontSize=7.5,leading=8,textColor=colors.white,alignment=1)
    
    sr_style = ParagraphStyle(
    'SRValue',
    fontName='Helvetica-Bold',
    fontSize=11,
    textColor=colors.HexColor('#8B0000')   # or #002060
)
    # 1. School Metadata Credentials Top Bar
    cred_data = [
        [Paragraph(f"<b>REG. NO.</b><br/>{school_details['reg_no']}", cred_style),
         Paragraph(f"<b>DISE CODE</b><br/>{school_details['dise_code']}", cred_style),
         Paragraph(f"<b>SCHOOL CODE</b><br/>{school_details['school_code']}", cred_style)]
    ]
    cred_table = Table(cred_data, colWidths=[180, 180, 180])
    cred_table.setStyle(TableStyle([
        ('LINEBELOW', (0,0), (-1,-1), 1, colors.HexColor('#D4AF37')),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6)
    ]))
    story.append(cred_table)
    story.append(Spacer(1, 10))

    # 2. Centered & Bold Recognition Banner Block
    story.append(Paragraph("(RECOGNIZED BY GOVT OF RAJASTHAN)", rec_center_style))

    # 3. Clean Left-Side Monogram & Header Text Layout
    logo_path = "logo.png"
    if not os.path.exists(logo_path):
        logo_path = "logo.jpg"
        
    if os.path.exists(logo_path):
        school_text_block = [
            Spacer(1, 4),
            Paragraph(school_details['name'], title_style),
            Spacer(1, 2),
            Paragraph(school_details['address'], addr_style)
        ]
        school_logo = Image(logo_path, width=65, height=70)
        
        header_table = Table([[school_logo, school_text_block]], colWidths=[70, 370])
        header_table.hAlign = 'CENTER'
        header_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ALIGN', (0,0), (0,0), 'LEFT'),
            ('LEFTPADDING', (1,0), (1,0), 12),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))
        story.append(header_table)
    else:
        title_center = ParagraphStyle('TitleCenter', fontName='Helvetica-Bold', fontSize=26, textColor=colors.HexColor('#002060'), alignment=1, leading=32)
        addr_center = ParagraphStyle('AddrCenter', fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#000000'), alignment=1, leading=13)
        story.append(Paragraph(school_details['name'], title_center))
        story.append(Spacer(1, 4))
        story.append(Paragraph(school_details['address'], addr_center))

    story.append(Spacer(1, 10))

    # 4. Ribbon banner block
    ribbon_table = Table([[Paragraph("REPORT CARD SESSION 2026–2027", rep_style)]], colWidths=[540])
    ribbon_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#002060')),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('ALIGN', (0,0), (-1,-1), 'CENTER')
    ]))
    story.append(ribbon_table)
    story.append(Spacer(1, 5))

    # 5. Student Metadata Information Box
    meta_data = [
        [Paragraph("SR No.", body_bold), Paragraph(f":  {student_data['SR No']}", sr_style),
         Paragraph("Father's Name", body_bold), Paragraph(f":  {student_data['Father Name']}", body_norm)],
        [Paragraph("Student Name", body_bold), Paragraph(f":  {student_data['Student Name']}", body_norm),
         Paragraph("Mother's Name", body_bold), Paragraph(f":  {student_data['Mother Name']}", body_norm)],
        [Paragraph("Class", body_bold), Paragraph(f":  {student_data['Class']}", body_norm),
         Paragraph("DOB", body_bold), Paragraph(f":  {student_data['DOB']}", body_norm)]
    ]
    meta_table = Table(meta_data, colWidths=[85, 185, 85, 185])
    meta_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#D1DDF2')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F9FBFC')),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 15))

    # 6. Core Matrix Structure Grid
    headers = [
    Paragraph("Subject", th_style),
    Paragraph("Unit Test 1<br/>(10)", th_style),
    Paragraph("Unit Test 2<br/>(10)", th_style),
    Paragraph("Unit Test 3<br/>(10)", th_style),
    Paragraph("H.Y.<br/>(70)", th_style),
    Paragraph("Total<br/>(100)", th_style),
    Paragraph("Unit Test 4<br/>(30)", th_style),
    Paragraph("Final<br/>(70)", th_style),
    Paragraph("Total<br/>(100)", th_style),
    Paragraph("Grand<br/>Total", th_style),
    Paragraph("Grade", th_style)
]
    table_data = [headers]
    
    for sub_name in subjects_list:
        table_data.append([
    Paragraph(f"<b>{sub_name}</b>", sub_row_style),"", "", "", "", "", "", "", "", "", ""])
        
    score_table = Table(
    table_data,
    colWidths=[
        118,   # Subject
        41,    # UT1
        41,    # UT2
        41,    # UT3
        42,    # HY
        42,    # Total
        42,    # UT4
        42,    # Final
        42,    # Total
        48,    # Grand Total
        42     # Grade
    ]
)
    score_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#002060')),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#B0C4DE')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),  
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 1), (0, -1), 10),
    ]))
    story.append(score_table)
    story.append(Spacer(1, 20))

    # 7. Bottom Signatures & Result Status Footer Box
    f_title = ParagraphStyle('FTitle', fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#333333'), alignment=1)
    f_lbl = ParagraphStyle('FLbl', fontName='Helvetica-Bold', fontSize=7, textColor=colors.HexColor('#777777'), alignment=1)
    result_lbl_style = ParagraphStyle('ResLbl', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#002060'), alignment=1)
    result_box_style = ParagraphStyle('ResBox', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#555555'), alignment=1)
    
    footer_data = [
        ["", "", "", "", ""],
        [Paragraph("<b>Principal Signature</b>", f_title), "", Paragraph("<b>RESULT DECLARATION DATE</b>", f_title), "", Paragraph("<b>Teacher Signature</b>", f_title)],
        ["_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _", "", "______ / ______ / _________", "", "_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _"],
        [Paragraph("PRINCIPAL SIGNATURE", f_lbl), "", "", "", Paragraph("TEACHER SIGNATURE", f_lbl)],
        [Spacer(1, 15), "", "", "", ""],
        [Paragraph("FINAL RESULT: [  PASS  /  FAIL  ]", result_lbl_style), "", "", "", ""],
        [Paragraph("Status: ____________________", result_box_style), "", "", "", ""]
    ]
    
    footer_table = Table(footer_data, colWidths=[160, 30, 160, 30, 160])
    footer_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('SPAN', (0, 5), (0, 5)),
        ('SPAN', (0, 6), (0, 6))
    ]))
    story.append(footer_table)

    def add_background(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor('#002060'))
        canvas.setLineWidth(4)
        canvas.rect(14, 14, doc.pagesize[0]-28, doc.pagesize[1]-28)
        
        canvas.setStrokeColor(colors.HexColor('#D4AF37'))
        canvas.setLineWidth(1)
        canvas.rect(18, 18, doc.pagesize[0]-36, doc.pagesize[1]-36)
        canvas.restoreState()

    doc.build(story, onFirstPage=add_background)
    buffer.seek(0)
    return buffer

# --- STREAMLIT UI SETUP ---
school_details = {
    "name": "AZIZ PUBLIC SCHOOL",
    "address": "6, Aziz Public School Aziz Colony Chardarwaza Jaipur Rajasthan",
    "reg_no": "41/91-92",
    "dise_code": "08122508601",
    "school_code": "267-2000"
}

tab_marking, tab_config = st.tabs(["📄 Print Premium Report Cards", "⚙️ Manage Subjects per Class"])

# --- TAB 2: SUBJECT CURRICULUM CONFIGURATION ---
with tab_config:
    st.header("🛠️ School Curriculum Management")
    c_left, c_right = st.columns(2)
    
    with c_left:
        st.subheader("➕ Add Subject to a Class")
        tgt_class = st.text_input("Enter Target Class (e.g., 1st, 9th):", key="add_tgt_class")
        new_sub_name = st.text_input("Subject Name:", key="add_sub_name")
        if st.button("Save Subject to Database", type="primary"):
            if tgt_class and new_sub_name:
                add_subject_to_db(tgt_class, new_sub_name)
                st.toast(f"💾 Saved {new_sub_name.strip()} for Class {tgt_class.strip()}!")
                st.rerun()

    with c_right:
        st.subheader("❌ Remove Subject")
        
        c_search_box, c_search_btn = st.columns([3, 1])
        with c_search_box:
            view_class = st.text_input("Enter Class to View:", value="9th", key="view_tgt_class")
        with c_search_btn:
            st.write(" ")
            st.write(" ")
            run_sub_search = st.button("🔍 Search Class")
            
        if view_class and (run_sub_search or 'last_sub_class' in st.session_state):
            if run_sub_search:
                st.session_state['last_sub_class'] = view_class
                
            class_subjects = get_subjects_by_class(st.session_state['last_sub_class'])
            if not class_subjects.empty:
                for idx, r in class_subjects.iterrows():
                    c_label, c_btn = st.columns([3, 1])
                    c_label.write(f"📖 **{r['Subject']}**")
                    if c_btn.button("Delete", key=f"del_{r['id']}"):
                        remove_subject_from_db(r['id'])
                        st.rerun()
            else:
                st.info(f"No custom subjects mapped for Class '{st.session_state['last_sub_class']}' yet.")

# --- TAB 1: CARD EXTRACTION ENGINE ---
with tab_marking:
    uploaded_file = st.file_uploader("Step 1: Upload Student Roster Excel File", type=["xlsx"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        for col in df.columns:
            if col in ['Class', 'Student Name', 'Father Name', 'Mother Name', 'SR No']:
                df[col] = df[col].astype(str).str.strip()
                
        st.success("Roster imported successfully!")
        st.markdown("---")
        
        st.markdown("### Step 2: Search Student by SR No")
        
        c_in, c_btn = st.columns([4, 1])
        with c_in:
            search_sr = st.text_input("Type or Scan Student's SR No:")
        with c_btn:
            st.write(" ")
            st.write(" ")
            run_search = st.button("🔍 Search Student", type="primary")
            
        if search_sr and (run_search or 'last_sr' in st.session_state):
            if run_search:
                st.session_state['last_sr'] = search_sr
                
            matched_records = df[df['SR No'] == str(st.session_state['last_sr']).strip()]
            
            if not matched_records.empty:
                student_row = matched_records.iloc[0]
                selected_class = student_row['Class']
                
                db_subjects = get_subjects_by_class(selected_class)
                subjects_list = db_subjects['Subject'].tolist() if not db_subjects.empty else []
                
                c_title, c_dl = st.columns([3, 1])
                with c_title:
                    st.markdown("### 🔍 Live Verification Preview")
                    st.info(f"✅ **Record Found:** {student_row['Student Name']}")
                with c_dl:
                    if subjects_list:
                        pdf_data = generate_pdf(student_row, subjects_list, school_details)
                        st.write(" ")
                        st.download_button(
                            label=f"📥 Download Premium PDF",
                            data=pdf_data,
                            file_name=f"Report_Card_{student_row['SR No']}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                
                col_left, col_right = st.columns(2)
                with col_left:
                    st.markdown(f"📁 **SR No:** `{student_row['SR No']}`")
                    st.markdown(f"👤 **Student Name:** {student_row['Student Name']}")
                    st.markdown(f"🏫 **Class:** {student_row['Class']}")
                    st.markdown(f"📅 **DOB:** {student_row['DOB']}")
                
                with col_right:
                    st.markdown(f"👨‍👦 **Father's Name:** {student_row['Father Name']}")
                    st.markdown(f"👩‍👦 **Mother's Name:** {student_row['Mother Name']}")
                    if subjects_list:
                        st.markdown(f"📖 **Class Subjects:** {', '.join(subjects_list)}")
                    else:
                        st.markdown("📖 **Class Subjects:** *None assigned yet.*")
                st.markdown("---")
                
                if not subjects_list:
                    st.error(f"⚠️ No subjects mapped for Class '{selected_class}'. Assign them in Tab 2 first.")
            else:
                st.error("❌ No student found matching this SR No. Check your Excel values.")
    else:
        st.info("💡 Please upload your student roster file to begin. Ensure your columns match: 'SR No', 'Student Name', 'Father Name', 'Mother Name', 'Class', 'DOB'.")