import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
from datetime import datetime

# 1. Setup
st.set_page_config(page_title="St. Paul's Timetable Tool", layout="wide")
st.title("🏫 Teacher Timetable Assistant")

@st.cache_data
def load_data():
    # Attempt to load the master timetable
    try:
        df = pd.read_excel("master_timetable.xlsx", header=[2, 3])
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
        return df
    except Exception as e:
        st.error(f"Could not load 'master_timetable.xlsx'. Please ensure the file is in the same folder. Error: {e}")
        return None

# 2. High-Fidelity Word Document Generator
def generate_formal_docx(sender, receiver, target_class, s_day, s_per, r_day, r_per, reason):
    doc = Document()
    
    # Global Font Setup: Times New Roman 11pt
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)
    
    # Ensure consistency for Asian characters/fonts
    r_pr = doc.styles['Normal']._element.get_or_add_rPr()
    r_pr.get_or_add_rFonts().set(qn('w:ascii'), 'Times New Roman')
    r_pr.get_or_add_rFonts().set(qn('w:hAnsi'), 'Times New Roman')

    section = doc.sections[0]
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    # --- Heading Section ---
    h1 = doc.add_paragraph("St. Paul’s School (Lam Tin)") [cite: 1]
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.runs[0].bold = True
    h1.runs[0].font.size = Pt(14)
    h1.paragraph_format.space_after = Pt(2)

    h2 = doc.add_paragraph("Record of Exchange of Lessons") [cite: 1]
    h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h2.runs[0].bold = True
    h2.runs[0].font.size = Pt(12)
    h2.paragraph_format.space_after = Pt(18)

    # --- Info Fields (Borderless Table for Precise Indentation Alignment) ---
    info_table = doc.add_table(rows=2, cols=2)
    info_table.autofit = False
    info_table.columns[0].width = Inches(1.6)
    info_table.columns[1].width = Inches(5.0)

    # Row: Teacher Name
    cell_name_label = info_table.cell(0, 0)
    p_name = cell_name_label.paragraphs[0]
    p_name.add_run("Name of Teacher:").bold = True [cite: 1]
    info_table.cell(0, 1).text = str(sender)

    # Row: Reason
    cell_reason_label = info_table.cell(1, 0)
    p_reason = cell_reason_label.paragraphs[0]
    p_reason.add_run("Reason for Exchange:").bold = True [cite: 1]
    info_table.cell(1, 1).text = str(reason) if reason else ""

    # Spacing before main table
    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- 14-column Main Table (2 headers + 4 content rows) ---
    table = doc.add_table(rows=6, cols=14) [cite: 1]
    table.style = 'Table Grid'
    
    # Header Row 1 (Merged)
    cell_sub = table.cell(0, 0).merge(table.cell(0, 6))
    cell_sub.text = "Lessons to be substituted" [cite: 1]
    cell_sub.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell_sub.paragraphs[0].runs[0].bold = True
    
    cell_ret = table.cell(0, 7).merge(table.cell(0, 13))
    cell_ret.text = "Lessons to be returned" [cite: 1]
    cell_ret.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell_ret.paragraphs[0].runs[0].bold = True

    # Header Row 2 (Sub-headers)
    sub_headers = ["Date", "Day", "Class", "Period", "Subject on Timetable", "Subject Replacing the Original", "Name of Teacher Taking the Lesson"] [cite: 1]
    full_headers = sub_headers + sub_headers
    for i, h in enumerate(full_headers):
        cell = table.cell(1, i)
        cell.text = h
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.runs[0]
        run.bold = True
        run.font.size = Pt(8)

    # Fill content row 1 (Row index 2)
    def clean_p(val):
        return str(val).replace("P", "").strip()

    table.cell(2, 1).text = str(s_day)
    table.cell(2, 2).text = str(target_class)
    table.cell(2, 3).text = clean_p(s_per)
    table.cell(2, 6).text = str(receiver)

    if r_day and r_day != "None":
        table.cell(2, 8).text = str(r_day)
        table.cell(2, 9).text = str(target_class)
        table.cell(2, 10).text = clean_p(r_per)
        table.cell(2, 13).text = str(sender)

    # Remaining 3 content rows (indices 3, 4, 5) are intentionally left empty 

    # --- Footer Section (Aligned Signature & Principal Approval) ---
    doc.add_paragraph().paragraph_format.space_before = Pt(20)
    today = datetime.now().strftime("%d / %m / %Y")
    
    ft = doc.add_table(rows=2, cols=2)
    ft.autofit = False
    ft.columns[0].width = Inches(3.5)
    ft.columns[1].width = Inches(3.5)

    # Signature side
    ft.cell(0, 0).text = f"Signature of teacher: ____________________" [cite: 1]
    ft.cell(1, 0).text = f"Date: {today}" [cite: 1]

    # Principal side (Right Aligned)
    p_principal = ft.cell(0, 1).paragraphs[0]
    p_principal.text = "Approved by Principal: ____________________" [cite: 1]
    p_principal.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p_date = ft.cell(1, 1).paragraphs[0]
    p_date.text = "Date: ____________________" [cite: 1]
    p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 3. Main App Logic
try:
    df = load_data()
    if df is not None:
        teacher_col = df.columns[0]
        teachers = sorted(df[teacher_col].dropna().unique().tolist())
        available_days = [d for d in df.columns.levels[0] if "Day" in str(d)]

        def is_free(val, disregards):
            val = str(val).strip().upper()
            if val in ["NAN", "", "NONE", "CLP"]: return True
            for d in disregards:
                if "*" in d:
                    if val.startswith(d.replace("*", "").upper()): return True
                elif val == d.upper(): return True
            return False

        def teaches_class(t_name, target_class):
            row = df[df[teacher_col] == t_name].iloc[0]
            return target_class.upper() in [str(v).strip().upper() for v in row.values[1:]]

        tab1, tab2 = st.tabs(["🔍 Find free lesson 「Call會快」", "🔄 Swap Lesson 「調堂易」"])

        with tab1:
            st.header("Find Common Free Lessons")
            # Collapsible filter removed; inputs are now in the main area
            sel_t = st.multiselect("Select Teachers", teachers)
            sel_d = st.multiselect("Select Days", available_days, default=available_days)
            dis_in = st.text_input("Disregard (e.g. 6*, CLP)", "6*, CLP")
            dis_l = [x.strip() for x in dis_in.split(",") if x.strip()]

            if sel_t:
                results = []
                subset = df[df[teacher_col].isin(sel_t)]
                for day in sel_d:
                    for period in df[day].columns:
                        if all(is_free(row[(day, period)], dis_l) for _, row in subset.iterrows()):
                            results.append({"Day": day, "Period": f"P{period}"})
                if results:
                    st.table(pd.DataFrame(results).groupby('Day')['Period'].apply(lambda x: ", ".join(x)).reset_index())
                else: 
                    st.warning("No common free slots found.")

        with tab2:
            st.header("Swap Lesson Finder")
            col1, col2, col3 = st.columns(3)
            with col1: my_name = st.selectbox("Your Name", ["Select..."] + teachers)
            with col2: swap_day = st.selectbox("Day of Lesson", available_days)
            with col3:
                p_list = list(df[swap_day].columns) if swap_day in available_days else []
                swap_p = st.selectbox("Period", p_list)

            dis_sw = st.text_input("Disregard Classes (e.g. 6*)", "CLP", key="dis_sw_key")
            dis_l_s = [x.strip() for x in dis_sw.split(",") if x.strip()]

            if my_name != "Select...":
                my_row = df[df[teacher_col] == my_name].iloc[0]
                target_class = str(my_row[(swap_day, swap_p)]).strip()

                if is_free(target_class, dis_l_s):
                    st.error(f"You are FREE/CLP on {swap_day} P{swap_p}.")
                elif target_class.upper().endswith('M'):
                    st.error(f"Class {target_class} is a mixed (M) class. Swapping not possible.")
                else:
                    st.info(f"Finding swaps for **{target_class}** on {swap_day} P{swap_p}")

                    partners_list = []
                    for _, row in df.iterrows():
                        other_name = row[teacher_col]
                        if other_name == my_name: continue
                        if is_free(row[(swap_day, swap_p)], dis_l_s):
                            if teaches_class(other_name, target_class):
                                ret_opts = []
                                for d in available_days:
                                    for p in df[d].columns:
                                        if str(row[(d, p)]).strip().upper() == target_class.upper():
                                            if is_free(my_row[(d, p)], dis_l_s):
                                                ret_opts.append(f"{d} P{p}")
                                partners_list.append({"Colleague": other_name, "Returns": ret_opts})

                    if partners_list:
                        view_df = pd.DataFrame([{"Colleague": p["Colleague"], "Returns": ", ".join(p["Returns"]) if p["Returns"] else "None"} for p in partners_list])
                        st.table(view_df)
                        
                        st.divider()
                        st.subheader("📄 Generate Official Exchange Slip")
                        reason = st.text_input("Reason for Exchange")
                        e_col1, e_col2 = st.columns(2)
                        with e_col1: sel_partner = st.selectbox("Select Colleague to Swap With", [p["Colleague"] for p in partners_list])
                        
                        p_data = next(p for p in partners_list if p["Colleague"] == sel_partner)
                        with e_col2: sel_ret = st.selectbox("Select Return Lesson", p_data["Returns"] if p_data["Returns"] else ["None"])
                        
                        if st.button("Prepare Download"):
                            ret_day, ret_p = "None", "None"
                            if " " in sel_ret:
                                parts = sel_ret.split(" ")
                                ret_day = f"{parts[0]} {parts[1]}"
                                ret_p = parts[2]
                            
                            doc_bytes = generate_formal_docx(
                                my_name, sel_partner, target_class, 
                                swap_day, swap_p, 
                                ret_day, ret_p, 
                                reason
                            )
                            st.download_button(label="⬇️ Download Lesson Exchange Slip", data=doc_bytes, file_name=f"Swap_{target_class}_{my_name}.docx")
                    else:
                        st.warning("No partners found.")

except Exception as e:
    st.error(f"Error: {e}")
