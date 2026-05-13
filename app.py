import streamlit as st
import pandas as pd
import re

# 1. Setup
st.set_page_config(page_title="Teacher Timetable Tools", layout="wide")
st.title("🏫 Teacher Timetable Assistant")

@st.cache_data
def load_data():
    df = pd.read_excel("master_timetable.xlsx", header=[2, 3])
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    return df

try:
    df = load_data()
    teacher_col = df.columns[0]
    teachers = df[teacher_col].dropna().unique().tolist()
    available_days = [d for d in df.columns.levels[0] if "Day" in str(d)]

    # Helper: Determine if a cell is "Free"
    def is_free(val, disregards):
        val = str(val).strip().upper()
        if val in ["NAN", "", "NONE", "CLP"]: return True
        for d in disregards:
            if "*" in d:
                if val.startswith(d.replace("*", "").upper()): return True
            elif val == d.upper(): return True
        return False

    # Helper: Check if a teacher teaches a specific class anywhere
    def teaches_class(teacher_name, target_class):
        row = df[df[teacher_col] == teacher_name].iloc[0]
        # Ignore the first column (Teacher Name) and check all others
        return target_class.upper() in [str(v).strip().upper() for v in row.values[1:]]

    # Create Tabs for the two different tools
    tab1, tab2 = st.tabs(["🔍 Common Free-Slot Finder", "🔄 Swap Lesson Finder"])

    # --- TAB 1: COMMON FREE SLOTS ---
    with tab1:
        st.header("Find time for a meeting")
        sel_teachers = st.multiselect("Select Teachers", teachers, key="free_t")
        sel_days = st.multiselect("Select Days", available_days, default=available_days, key="free_d")
        dis_input = st.text_input("Disregard (e.g. 6*, CLP)", "", key="free_dis")
        dis_list = [x.strip() for x in dis_input.split(",") if x.strip()]

        if sel_teachers:
            results = []
            subset = df[df[teacher_col].isin(sel_teachers)]
            for day in sel_days:
                for period in df[day].columns:
                    if all(is_free(row[(day, period)], dis_list) for _, row in subset.iterrows()):
                        results.append({"Day": day, "Period": f"P{period}"})
            
            if results:
                st.table(pd.DataFrame(results).groupby('Day')['Period'].apply(lambda x: " | ".join(x)).reset_index())
            else: st.warning("No common free slots.")

    # --- TAB 2: SWAP FINDER ---
    with tab2:
        st.header("Find a colleague to swap with")
        col1, col2, col3 = st.columns(3)
        with col1:
            my_name = st.selectbox("Your Name", ["Select..."] + teachers)
        with col2:
            swap_day = st.selectbox("Day of Lesson", available_days)
        with col3:
            # Dynamically get periods for that day
            p_list = list(df[swap_day].columns)
            swap_period = st.selectbox("Period", p_list)

        dis_swap = st.text_input("Disregard (e.g. 6*)", "CLP", key="swap_dis")
        dis_list_s = [x.strip() for x in dis_swap.split(",") if x.strip()]

        if my_name != "Select...":
            # 1. Identify the class the user wants to give up
            my_row = df[df[teacher_col] == my_name].iloc[0]
            target_class = str(my_row[(swap_day, swap_period)]).strip()

            if is_free(target_class, dis_list_s):
                st.error(f"You appear to be FREE (or {target_class}) on {swap_day} P{swap_period}. Nothing to swap!")
            elif target_class.upper().endswith('M'):
                st.error(f"Class {target_class} is an 'M' class. Swapping is not allowed for mixed classes.")
            else:
                st.info(f"Targeting swaps for Class: **{target_class}** on {swap_day} P{swap_period}")

                # 2. Search for colleagues
                partners = []
                for _, row in df.iterrows():
                    other_name = row[teacher_col]
                    if other_name == my_name: continue
                    
                    # Criteria A: Are they free right now?
                    if is_free(row[(swap_day, swap_period)], dis_list_s):
                        # Criteria B: Do they teach this class elsewhere?
                        if teaches_class(other_name, target_class):
                            # Criteria C: Where do they have this class when YOU are free?
                            return_options = []
                            for d in available_days:
                                for p in df[d].columns:
                                    # If they have the class AND you are free
                                    if str(row[(d, p)]).strip().upper() == target_class.upper():
                                        if is_free(my_row[(d, p)], dis_list_s):
                                            return_options.append(f"{d} P{p}")
                            
                            partners.append({
                                "Colleague": other_name,
                                "Potential Return Lessons": " OR ".join(return_options) if return_options else "None found"
                            })

                if partners:
                    st.success(f"Found {len(partners)} potential swap partners:")
                    st.table(pd.DataFrame(partners))
                else:
                    st.warning("No colleagues found who teach this class and are free at this time.")

except Exception as e:
    st.error(f"Error: {e}")
