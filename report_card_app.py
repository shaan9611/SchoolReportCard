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
def generate_pdf(student_data, subjects_list, school_details, marks_dict=None):
    buffer = io.BytesIO()
    # Shrunk margins to prevent signature overflow pages
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=20, bottomMargin=20)
    story = []
    
    styles = getSampleStyleSheet()
    cred_style = ParagraphStyle('Creds', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#002060'), alignment=1)
    rec_center_style = ParagraphStyle('RecCenter', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#000000'), alignment=1, spaceAfter=1)
    
    title_style = ParagraphStyle('SchoolTitle', fontName='Helvetica-Bold', fontSize=32, textColor=colors.HexColor('#002060'), alignment=0, leading=32)
    addr_style = ParagraphStyle('SchoolAddr', fontName='Helvetica', fontSize=10, textColor=colors.HexColor('#000000'), alignment=0, leading=12)
    
    rep_style = ParagraphStyle('RepCard', fontName='Helvetica-Bold', fontSize=13, textColor=colors.white, alignment=1)
    body_bold = ParagraphStyle('BBold', fontName='Helvetica-Bold', fontSize=9, textColor=colors.HexColor('#002060'))
    body_norm = ParagraphStyle('BNorm',fontName='Helvetica-Bold',fontSize=10,leading=12,wordWrap='LTR',textColor=colors.HexColor('#333333'))
    sub_row_style = ParagraphStyle('SubRow', fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#111111'))
    th_style = ParagraphStyle('THead',fontName='Helvetica-Bold',fontSize=7.5,leading=8,textColor=colors.white,alignment=1)
    
    td_style = ParagraphStyle('TData', fontName='Helvetica', fontSize=8.5, alignment=1)
    td_bold = ParagraphStyle('TDataBold', fontName='Helvetica-Bold', fontSize=8.5, alignment=1, textColor=colors.HexColor('#002060'))
    
    sr_style = ParagraphStyle('SRValue', fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor('#8B0000'))

    # 1. School Metadata Credentials Top Bar
    cred_data = [
        [Paragraph(f"<b>REG. NO.</b><br/>{school_details['reg_no']}", cred_style),
         Paragraph(f"<b>DISE CODE</b><br/>{school_details['dise_code']}", cred_style),
         Paragraph(f"<b>SCHOOL CODE</b><br/>{school_details['school_code']}", cred_style)]
    ]
    cred_table = Table(cred_data, colWidths=[180, 180, 180])
    cred_table.setStyle(TableStyle([
        ('LINEBELOW', (0,0), (-1,-1), 1, colors.HexColor('#D4AF37')),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4)
    ]))
    story.append(cred_table)
    story.append(Spacer(1, 6))

    # 2. Centered & Bold Recognition Banner Block
    story.append(Paragraph("(RECOGNIZED BY GOVT OF RAJASTHAN)", rec_center_style))

    # 3. Clean Left-Side Monogram & Header Text Layout
    logo_path = "logo.png"
    if not os.path.exists(logo_path):
        logo_path = "logo.jpg"
        
    if os.path.exists(logo_path):
        school_text_block = [
            Spacer(1, 2),
            Paragraph(school_details['name'], title_style),
            Spacer(1, 1),
            Paragraph(school_details['address'], addr_style)
        ]
        school_logo = Image(logo_path, width=60, height=65)
        
        header_table = Table([[school_logo, school_text_block]], colWidths=[65, 375])
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

    story.append(Spacer(1, 6))

    # 4. Ribbon banner block
    ribbon_table = Table([[Paragraph("REPORT CARD SESSION 2026–2027", rep_style)]], colWidths=[540])
    ribbon_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#002060')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('ALIGN', (0,0), (-1,-1), 'CENTER')
    ]))
    story.append(ribbon_table)
    story.append(Spacer(1, 4))

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
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 10))

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
        Paragraph("Grand Total<br/>(200)", th_style),
        Paragraph("Grade", th_style)
    ]
    table_data = [headers]
    
    for sub_name in subjects_list:
        if marks_dict and sub_name in marks_dict:
            m = marks_dict[sub_name]
            if sub_name.strip().lower() == "attendance":
                table_data.append([
                    Paragraph(f"<b>{sub_name}</b>", sub_row_style),
                    Paragraph(str(m['UT1']) if m['UT1'] != "-" else "-", td_style),
                    Paragraph(str(m['UT2']) if m['UT2'] != "-" else "-", td_style),
                    Paragraph(str(m['UT3']) if m['UT3'] != "-" else "-", td_style),
                    Paragraph(str(m['HY']) if m['HY'] != "-" else "-", td_style),
                    Paragraph("-", td_bold),
                    Paragraph(str(m['UT4']) if m['UT4'] != "-" else "-", td_style),
                    Paragraph(str(m['Final']) if m['Final'] != "-" else "-", td_style),
                    Paragraph("-", td_bold),
                    Paragraph("-", td_bold),
                    Paragraph("-", td_bold)
                ])
            else:
                table_data.append([
                    Paragraph(f"<b>{sub_name}</b>", sub_row_style),
                    Paragraph(str(m['UT1']), td_style),
                    Paragraph(str(m['UT2']), td_style),
                    Paragraph(str(m['UT3']), td_style),
                    Paragraph(str(m['HY']), td_style),
                    Paragraph(str(m['HY Total']), td_bold),
                    Paragraph(str(m['UT4']), td_style),
                    Paragraph(str(m['Final']), td_style),
                    Paragraph(str(m['Final Total']), td_bold),
                    Paragraph(str(m['Grand Total']), td_bold),
                    Paragraph(str(m['Grade']), td_bold)
                ])
        else:
            table_data.append([
                Paragraph(f"<b>{sub_name}</b>", sub_row_style), "", "", "", "", "", "", "", "", "", ""
            ])
        
    score_table = Table(
        table_data,
        colWidths=[118, 41, 41, 41, 42, 42, 42, 42, 42, 48, 42]
    )
    score_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#002060')),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#B0C4DE')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 1), (0, -1), 10),
    ]))
    story.append(score_table)
    story.append(Spacer(1, 12))

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
        [Spacer(1, 8), "", "", "", ""],
        [Paragraph("FINAL RESULT: [  PASS  /  FAIL  ]", result_lbl_style), "", "", "", ""],
        [Paragraph("Status: ____________________", result_box_style), "", "", "", ""]
    ]
    
    footer_table = Table(footer_data, colWidths=[160, 30, 160, 30, 160])
    footer_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
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

tab_classic, tab_smart, tab_config = st.tabs([
    "📄 Classic Report Card", 
    "🤖 Smart Report Card (Marks Entry)", 
    "⚙️ Manage Subjects per Class"
])

# --- TAB 3: SUBJECT CURRICULUM CONFIGURATION ---
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

# --- TAB 1: CLASSIC BLANK MATRIX GENERATOR ---
with tab_classic:
    uploaded_file = st.file_uploader("Step 1: Upload Student Roster Excel File", type=["xlsx"], key="upload_classic")

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
            search_sr = st.text_input("Type or Scan Student's SR No:", key="search_classic")
        with c_btn:
            st.write(" ")
            st.write(" ")
            run_search = st.button("🔍 Search Student", type="primary", key="btn_classic")
            
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
                            use_container_width=True,
                            key="btn_dl_classic"
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
            else:
                st.error("❌ No student found matching this SR No. Check your Excel values.")
    else:
        st.info("💡 Please upload your student roster file to begin. Ensure your columns match: 'SR No', 'Student Name', 'Father Name', 'Mother Name', 'Class', 'DOB'.")

# --- TAB 2: SMART MARKS ENTRY ENGINE ---
with tab_smart:
    uploaded_file_smart = st.file_uploader("Step 1: Upload Student Roster Excel File", type=["xlsx"], key="upload_smart")

    if uploaded_file_smart is not None:
        df_smart = pd.read_excel(uploaded_file_smart)
        for col in df_smart.columns:
            if col in ['Class', 'Student Name', 'Father Name', 'Mother Name', 'SR No']:
                df_smart[col] = df_smart[col].astype(str).str.strip()
                
        st.success("Roster imported successfully for Marks Processing!")
        st.markdown("---")
        st.markdown("### Step 2: Search Student to Begin Evaluation")
        
        c_in_s, c_btn_s = st.columns([4, 1])
        with c_in_s:
            search_sr_smart = st.text_input("Type or Scan Student's SR No:", key="search_smart")
        with c_btn_s:
            st.write(" ")
            st.write(" ")
            run_search_smart = st.button("🔍 Load Student Matrix", type="primary", key="btn_smart")
            
        if search_sr_smart and (run_search_smart or 'last_sr_smart' in st.session_state):
            if run_search_smart:
                st.session_state['last_sr_smart'] = search_sr_smart
                
            matched_records = df_smart[df_smart['SR No'] == str(st.session_state['last_sr_smart']).strip()]
            
            if not matched_records.empty:
                student_row = matched_records.iloc[0]
                selected_class = student_row['Class']
                db_subjects = get_subjects_by_class(selected_class)
                subjects_list = db_subjects['Subject'].tolist() if not db_subjects.empty else []
                
                if not subjects_list:
                    st.error(f"⚠️ No subjects mapped for Class '{selected_class}'. Map subjects in the configuration tab first.")
                else:
                    st.info(f"✅ **Evaluating Record:** {student_row['Student Name']} (Class: {student_row['Class']})")
                    
                    # Build dynamic setup container values for editor
                    init_data = []
                    for sub in subjects_list:
                        init_data.append({
                            "Subject": sub, "UT1": 0.0, "UT2": 0.0, 
                            "UT3": 0.0, "HY": 0.0, "UT4": 0.0, 
                            "Final": 0.0
                        })
                    editor_df = pd.DataFrame(init_data)
                    
                    st.markdown("### 📝 Enter Subject Scores / Attendance Values Below")
                    # Removed strict min/max boundaries entirely from the column config. 
                    # Standard validation calculations happen dynamically in the backend loops.
                    edited_df = st.data_editor(
                        editor_df,
                        column_config={
                            "Subject": st.column_config.TextColumn("Subject", disabled=True),
                            "UT1": st.column_config.NumberColumn("UT1 Score / Days", step=0.5, format="%.1f"),
                            "UT2": st.column_config.NumberColumn("UT2 Score / Days", step=0.5, format="%.1f"),
                            "UT3": st.column_config.NumberColumn("UT3 Score / Days", step=0.5, format="%.1f"),
                            "HY": st.column_config.NumberColumn("HY Score / Days", step=0.5, format="%.1f"),
                            "UT4": st.column_config.NumberColumn("UT4 Score / Days", step=0.5, format="%.1f"),
                            "Final": st.column_config.NumberColumn("Final Score / Days", step=0.5, format="%.1f"),
                        },
                        disabled=["Subject"],
                        hide_index=True,
                        key=f"editor_matrix_{student_row['SR No']}"
                    )
                    
                    # Calculation Pipeline
                    calculated_marks = {}
                    has_error = False
                    error_msg = ""
                    
                    for _, row in edited_df.iterrows():
                        sub = row["Subject"]
                        ut1, ut2, ut3, hy = row["UT1"], row["UT2"], row["UT3"], row["HY"]
                        ut4, final = row["UT4"], row["Final"]
                        
                        # Dynamic Backend Boundary Checks (Skipped for Attendance)
                        if sub.strip().lower() == "attendance":
                            calculated_marks[sub] = {
                                'UT1': ut1, 'UT2': ut2, 'UT3': ut3, 'HY': hy, 'HY Total': "-",
                                'UT4': ut4, 'Final': final, 'Final Total': "-", 
                                'Grand Total': "-", 'Grade': "-"
                            }
                        else:
                            if not (0 <= ut1 <= 10) or not (0 <= ut2 <= 10) or not (0 <= ut3 <= 10) or not (0 <= hy <= 70) or not (0 <= ut4 <= 30) or not (0 <= final <= 70):
                                has_error = True
                                error_msg = f"🚨 Validation Error in '{sub}': UT1-3 must be ≤10, HY ≤70, UT4 ≤30, Final ≤70."
                            
                            hy_total = ut1 + ut2 + ut3 + hy
                            final_total = ut4 + final
                            grand_total = hy_total + final_total
                            
                            pct = (grand_total / 200.0) * 100.0
                            if pct >= 91: grade = "A1"
                            elif pct >= 81: grade = "A2"
                            elif pct >= 71: grade = "B1"
                            elif pct >= 61: grade = "B2"
                            elif pct >= 51: grade = "C1"
                            elif pct >= 41: grade = "C2"
                            elif pct >= 33: grade = "D"
                            else: grade = "E"
                            
                            calculated_marks[sub] = {
                                'UT1': ut1, 'UT2': ut2, 'UT3': ut3, 'HY': hy, 'HY Total': hy_total,
                                'UT4': ut4, 'Final': final, 'Final Total': final_total, 
                                'Grand Total': grand_total, 'Grade': grade
                            }
                    
                    if has_error:
                        st.error(error_msg)
                    else:
                        st.markdown("### 📈 Computed Calculations Summary")
                        preview_list = []
                        for s, m in calculated_marks.items():
                            preview_list.append({
                                "Subject": s, "HY Total (100)": m['HY Total'], 
                                "Final Total (100)": m['Final Total'], "Grand Total (200)": m['Grand Total'], 
                                "Grade": m['Grade']
                            })
                        st.dataframe(pd.DataFrame(preview_list), hide_index=True, use_container_width=True)
                        
                        # Smart PDF Export Action Linked Downloader
                        smart_pdf_data = generate_pdf(student_row, subjects_list, school_details, marks_dict=calculated_marks)
                        st.write(" ")
                        st.download_button(
                            label=f"📥 Download Smart PDF for {student_row['Student Name']}",
                            data=smart_pdf_data,
                            file_name=f"Smart_Report_Card_{student_row['SR No']}.pdf",
                            mime="application/pdf",
                            type="primary",
                            key="btn_dl_smart"
                        )
            else:
                st.error("❌ No student found matching this SR No. Check your Excel values.")
    else:
        st.info("💡 Please upload your student roster file inside this tab to generate report cards containing calculated scores.")