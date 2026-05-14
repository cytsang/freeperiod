import streamlit as st
import pandas as pd
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
from datetime import datetime

# 1. Setup
st.set_page_config(page_title="St. Paul's Timetable Tool", layout="wide")
st.title("🏫 Teacher Timetable Assistant")

@st.cache_data
def load_data():
    # Load Excel - Row 3 is Days, Row 4 is Periods
    df = pd.read_excel("master_timetable.xlsx", header=[2, 3])
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    return df

# 2. Formal Word Document Generator
def generate_formal_docx(sender, receiver, target_class, swap_info, return_info, reason):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    # Header Section
    h1 = doc.add_paragraph("St. Paul’s School (Lam Tin)")
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.runs[0].bold = True
    h1.runs[0].font.size = Pt(14)

    h2 = doc.add_paragraph("Record of Exchange of Lessons")
    h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h2.runs[0].bold = True
    h2.runs[0].font.size = Pt(12)

    doc.add_paragraph(f"Name of Teacher: {sender}").runs[0].bold = True
    doc.add_paragraph(f"Reason for Exchange: {reason if reason else '________________________'}").runs[0].bold = True

    # Build the 14-column table
    table = doc.add_table(rows=3, cols=14)
    table.style = 'Table Grid'
    
    cell_sub = table.cell(0, 0).merge(table.cell(0, 6))
    cell_sub.text = "Lessons to be substituted"
    cell_sub.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    cell_ret = table.cell(0, 7).merge(table.cell(0, 13))
    cell_ret.text = "Lessons to be returned"
    cell_ret.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    headers = ["Date", "Day", "Class", "Period", "Subject", "Replacing", "Taking"]
    for i, h in enumerate(headers + headers):
        cell = table.cell(1, i)
        cell.text = h
        cell.paragraphs[0].runs[0].font.size = Pt(8)

    # Data Splitting
    s_day = swap_info.split(' ')[0]
    s_per = swap_info.split(' ')[1]
    r_day = return_info.split(' ')[0] if "P" in return_info else "___"
    r_per = return_info.split(' ')[1] if "P" in return_info else "___"

    # Fill Table
    table.cell(2, 1).text = s_day
    table.cell(2, 2).text = target_class
    table.cell(2, 3).text = s_per
    table.cell(2, 6).text = receiver
    table.cell(2, 8).text = r_day
    table.cell(2, 9).text = target_class
    table.cell(2, 10).text = r_per
    table.cell(2, 13).text = sender

    # Footer
    doc.add_paragraph("\n")
    today = datetime.now().strftime("%d / %m / %Y")
    footer = doc.add_table(rows=2, cols=2)
    footer.cell(0, 0).text = f"Signature of teacher: ____________________"
    footer.cell(0, 1).text = f"Approved by Principal: ____________________"
    footer.cell(1, 0).text = f"Date: {today}"
    footer.cell(1, 1).text = f"Date: ____________________"

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 3. Application logic
try:
    df = load_data()
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

    # --- TAB 1: CALL會快 ---
    with tab1:
        st.header("Find Common Free Lessons")
        sel_t = st.multiselect("Select Teachers", teachers)
        sel_d = st.multiselect("Select Days", available_days, default=available_days)
        dis_in = st.text_input("Disregard (e.g. 6*, CLP)", "")
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
            else: st.warning("No common free slots found.")

    # --- TAB 2: 調堂易 ---
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
                st.error(f"You are FREE/CLP on {swap_day} P{swap_p}. Nothing to swap!")
            elif target_class.upper().endswith('M'):
                st.error(f"Class {target_class} is a mixed (M) class. Swapping is not allowed.")
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
                    
                    partner_data = next(p for p in partners_list if p["Colleague"] == sel_partner)
                    with e_col2: sel_ret = st.selectbox("Select Return Lesson", partner_data["Returns"] if partner_data["Returns"] else ["N/A"])
                    
                    if st.button("Prepare Download"):
                        doc_bytes = generate_formal_docx(my_name, sel_partner, target_class, f"{swap_day} P{swap_p}", sel_ret, reason)
                        st.download_button(label="⬇️ Download Lesson Exchange Slip", data=doc_bytes, file_name=f"Swap_{target_class}_{my_name}.docx")
                else:
                    st.warning("No partners found teaching this class who are free.")

except Exception as e:
    st.error(f"Error: {e}")
