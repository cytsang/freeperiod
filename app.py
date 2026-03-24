import streamlit as st
import pandas as pd
import re

# 1. Setup
st.set_page_config(page_title="Teacher Free-Slot Finder", layout="wide")
st.title("📅 Teacher Common Free-Slot Finder")

@st.cache_data
def load_data():
    # We read rows 3 and 4 (index 2 and 3) as headers
    df = pd.read_excel("master_timetable.xlsx", header=[2, 3])
    # Remove any completely empty rows or columns
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
    return df

try:
    df = load_data()
    
    # The teacher names are in the first column
    teacher_col_name = df.columns[0]
    teachers = df[teacher_col_name].dropna().unique().tolist()

    # --- SIDEBAR ---
    st.sidebar.header("Filter Settings")
    selected_teachers = st.sidebar.multiselect("1. Select Teachers", teachers)
    
    # Get the unique Days from the top header level
    available_days = [d for d in df.columns.levels[0] if "Day" in str(d)]
    selected_days = st.sidebar.multiselect("2. Select Days", available_days, default=available_days)
    
    disregard_input = st.sidebar.text_input("3. Disregard Lessons (e.g. CLP, 6*, 1B)", "")
    disregard_list = [x.strip() for x in disregard_input.split(",") if x.strip()]

    def is_free(cell_value):
        val = str(cell_value).strip()
        if val.lower() in ["nan", "", "none", "nan nan"]:
            return True
        for pattern in disregard_list:
            if "*" in pattern:
                clean_pat = "^" + pattern.replace("*", ".*") + "$"
                if re.match(clean_pat, val, re.IGNORECASE): return True
            elif val.lower() == pattern.lower():
                return True
        return False

    # --- SEARCH LOGIC ---
    if selected_teachers:
        results = []
        
        # Filter the dataframe for only selected teachers
        subset = df[df[teacher_col_name].isin(selected_teachers)]
        
        for day in selected_days:
            # Look at all periods available for that day
            day_columns = df[day].columns
            for period in day_columns:
                # Check every selected teacher for this specific Day and Period
                all_free = True
                for _, row in subset.iterrows():
                    cell_content = row[(day, period)]
                    if not is_free(cell_content):
                        all_free = False
                        break
                
                if all_free:
                    results.append({"Day": day, "Period": f"P{period}"})

        # --- DISPLAY ---
        st.subheader(f"Common Free Slots for: {', '.join(selected_teachers)}")
        
        if results:
            res_df = pd.DataFrame(results)
            # Create a nice horizontal view
            final_view = res_df.groupby('Day')['Period'].apply(lambda x: " | ".join(x)).reset_index()
            st.table(final_view)
        else:
            st.warning("No common free slots found.")
    else:
        st.info("Select teachers in the sidebar to begin.")

except Exception as e:
    st.error(f"Error loading data: {e}")
    st.info("Check your Excel: Headers should be in Row 3 (Days) and Row 4 (Periods).")
