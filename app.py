import streamlit as st
import pandas as pd
import re

# 1. Setup the Page
st.set_page_config(page_title="Teacher Free-Slot Finder", layout="wide")
st.title("📅 Teacher Common Free-Slot Finder")

# 2. Load the Data
@st.cache_data
def load_data():
    # Replace 'your_file.xlsx' with your actual filename later
    df = pd.read_excel("master_timetable.xlsx", header=[2, 3]) 
    return df

try:
    df = load_data()
    
    # Clean up the teacher names (Column A)
    teachers = df.iloc[:, 0].dropna().unique().tolist()

    # --- SIDEBAR INPUTS ---
    st.sidebar.header("Filter Settings")
    
    selected_teachers = st.sidebar.multiselect("1. Select Teachers", teachers)
    
    days = ["Day 1", "Day 2", "Day 3", "Day 4", "Day 5", "Day 6"]
    selected_days = st.sidebar.multiselect("2. Select Days", days, default=days)
    
    disregard_input = st.sidebar.text_input("3. Disregard Lessons (e.g. CLP, 6*, 1B)", "")
    disregard_list = [x.strip() for x in disregard_input.split(",") if x.strip()]

    # --- LOGIC ---
    if selected_teachers:
        # Create a function to check if a slot is "Free"
        def is_free(cell_value):
            val = str(cell_value).strip()
            if val == "nan" or val == "" or val == "None":
                return True
            for pattern in disregard_list:
                if "*" in pattern:
                    clean_pat = pattern.replace("*", ".*")
                    if re.match(clean_pat, val): return True
                elif val == pattern:
                    return True
            return False

        # Find the common slots
        results = []
        
        for day in selected_days:
            for period in range(1, 10): # Periods 1-9
                period_str = str(period)
                all_free = True
                
                for teacher in selected_teachers:
                    # Look up the specific teacher's row and the day/period column
                    teacher_row = df[df.iloc[:, 0] == teacher]
                    try:
                        cell_content = teacher_row[(day, period_str)].values[0]
                        if not is_free(cell_content):
                            all_free = False
                            break
                    except:
                        continue
                
                if all_free:
                    results.append({"Day": day, "Period": f"Period {period}"})

        # --- DISPLAY RESULTS ---
        st.subheader(f"Common Free Slots for: {', '.join(selected_teachers)}")
        if results:
            res_df = pd.DataFrame(results)
            # Pivot for better view
            st.table(res_df.groupby('Day')['Period'].apply(list))
        else:
            st.warning("No common free slots found for these teachers.")
    else:
        st.info("Please select at least one teacher from the sidebar to begin.")

except Exception as e:
    st.error("Please ensure 'master_timetable.xlsx' is uploaded to the repository.")
