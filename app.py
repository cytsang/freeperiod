import streamlit as st
import pandas as pd
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
from datetime import datetime

# 1. Setup
st.set_page_config(page_title="Teacher Timetable Tools", layout="wide")
st.title("🏫 Teacher Timetable Assistant")

@st.cache_data
def load_data():
    df = pd.read_excel("master_timetable.xlsx", header=[2, 3])
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    return df

def generate_formal_docx(sender, receiver, target_class, swap_info, return_info, reason):
    doc = Document()
    
    # Set Narrow Margins to fit the wide table
    section = doc.sections[0]
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    # Header
    title1 = doc.add_paragraph("St. Paul’s School (Lam Tin)")
    title1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title1.runs[0].bold = True
    title1.runs[0].font.size = Pt(14)

    title2 = doc.add_paragraph("Record of Exchange of Lessons")
    title2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title2.runs[0].bold = True
    title2.runs[0].font.size = Pt(12)

    # Top Info
    p = doc.add_paragraph()
    p.add_run(f"Name of Teacher: ").bold = True
    p.add_run(f"{sender}").underline = True
    
    p2 = doc.add_paragraph()
    p2.add_run(f"Reason for Exchange: ").bold = True
    p2.add_run(f"{reason if reason else '____________________________________'}").underline = True

    # Build the 14-column table
    table = doc.add_table(rows=3, cols=14)
    table.style = 'Table Grid'
    
    # Row 1: Merged Headers
    a = table.cell(0, 0).merge(table.cell(0, 6))
    a.text = "Lessons to be substituted"
    a.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    b = table.cell(0, 7).merge(table.cell(0, 13))
    b.text = "Lessons to be returned"
    b.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Row 2: Sub-headers
    headers = ["Date", "Day", "Class", "Period", "Subject on Timetable", "Subject Replacing", "Teacher Taking"]
    full_headers = headers + headers
    for i, h in enumerate(full_headers):
        cell = table.cell(1, i)
        cell.text = h
        cell.paragraphs[0].runs[0].font.size = Pt(8)

    # Row 3: Data (Splitting swap_info like "Day 1 P3")
    s_day = swap_info.split(' ')[0]
    s_per = swap_info.split(' ')[1]
    r_day = return_info.split(' ')[0] if return_info != "N/A" else "___"
    r_per = return_info.split(' ')[1] if return_info != "N/A" else "___"

    # Filling substituted side
    table.cell(2, 0).text = "[Date]"
    table.cell(2, 1).text = s_day
    table.cell(2, 2).text = target_class
    table.cell(2, 3).text = s_per
    table.cell(2, 6).text = receiver # Partner takes this

    # Filling returned side
    table.cell(2, 7).text = "[Date]"
    table.cell(2, 8).text = r_day
    table.cell(2, 9).text = target_class
    table.cell(2, 10).text = r_per
    table.cell(2, 13).text = sender # User takes this back

    # Signatures
    doc.add_paragraph("\n")
    today_str = datetime.now().strftime("%d / %m / %Y")
    
    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.cell(0, 0).text = f"Signature of teacher: ____________________"
    sig_table.cell(0, 1).text = f"Approved by Principal: ____________________"
    
    sig_table.cell(1, 0).text = f"Date: {today_str}"
    sig_table.cell(1, 1).text = f"Date: ____________________"

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

try:
    df = load_data()
    teacher_col = df.columns[0]
    teachers = df[teacher_col].dropna().unique().tolist()
    available_days = [d for d in df.columns.levels[0] if "Day" in str(d)]

    tab1, tab2 = st.tabs(["🔍 Find free lesson 「Call會快」", "🔄 Swap Lesson 「調堂易」"])

    with tab1:
        st.header("Find time for a meeting")
        sel_teachers = st.multiselect("Select Teachers", teachers)
        sel_days = st.multiselect("Select Days", available_days, default=available_days)
        # Logic remains same as previous version...
        # [Tab 1 code omitted for brevity but should be kept in your file]

    with tab2:
        st.header("Find a colleague to swap with")
        c1, c2, c3 = st.columns(3)
        with c1: my_name = st.selectbox("Your Name", ["Select..."] + teachers)
        with c2: swap_day = st.selectbox("Day of Lesson", available_days)
        with c3:
            p_list = list(df[swap_day].columns)
            swap_period = st.selectbox("Period", p_list)

        dis_swap = st.text_input("Disregard (e.g. 6*)", "CLP")

        if my_name != "Select...":
            my_row = df[df[teacher_col] == my_name].iloc[0]
            target_class = str(my_row[(swap_day, swap_period)]).strip()

            if target_class.upper().endswith('M'):
                st.error(f"Class {target_class} is an 'M' class. Swapping is not allowed.")
            else:
                st.info(f"Targeting swaps for Class: **{target_class}**")

                partners_data = []
                for _, row in df.iterrows():
                    other_name = row[teacher_col]
                    if other_name == my_name: continue
                    # (Insert logic here to find partners same as before)
                    # ... [Skipping partner search logic for brevity, use existing logic]

                # (Assuming partners_data is populated)
                # --- NEW EXPORT INTERFACE ---
                st.divider()
                st.subheader("📄 Generate Official Exchange Slip")
                
                reason_input = st.text_input("Reason for Exchange (Optional)", placeholder="e.g. Attending Workshop")
                
                exp_col1, exp_col2 = st.columns(2)
                # (Standard selection logic for partner and return lesson)
                # ...

                if st.button("Preview & Generate Slip"):
                    docx_data = generate_formal_docx(
                        my_name, 
                        sel_partner, 
                        target_class, 
                        f"{swap_day} P{swap_period}", 
                        sel_return,
                        reason_input
                    )
                    
                    st.download_button(
                        label="⬇️ Download Official Slip (.docx)",
                        data=docx_data,
                        file_name=f"Exchange_Slip_{my_name}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

except Exception as e:
    st.error(f"System Error: {e}")
