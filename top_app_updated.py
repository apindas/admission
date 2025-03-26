import streamlit as st
import pandas as pd
import datetime
import matplotlib.pyplot as plt
import os

# Set page configuration
st.set_page_config(page_title="GHSS Cherpu Admission Portal", page_icon="🏫", layout="wide")

# Attempt to import openpyxl
try:
    import openpyxl
    type_openpyxl = True
except ModuleNotFoundError:
    type_openpyxl = False

# Constants
SCHOOL_NAME = "GHSS CHERPU"
ADMISSION_YEAR = 2025
STREAMS = ["BIO", "CS", "HUM", "COM"]
SECOND_LANGUAGES = ["MAL", "HIN", "SKT"]
CASTES = ["GEN", "ETB", "MUSLIM", "SC", "LSA", "OBH", "DV", "VK", "KN", "KU", "ST", "OBCHRISTIAN"]
STATUS_OPTIONS = ["PERMANENT", "TEMPORARY"]

# Initialize session state
def initialize_session():
    if "students_df" not in st.session_state:
        st.session_state.students_df = pd.DataFrame(columns=['Name', 'Rank', 'Stream', 'Second_Language', 'Caste', 'Admission_Status', 'Date_of_Admission'])
    if "tc_issued_df" not in st.session_state:
        st.session_state.tc_issued_df = pd.DataFrame(columns=['Name', 'Rank', 'Stream', 'Date_of_TC'])

initialize_session()

def save_admitted_students():
    if not type_openpyxl:
        return None
    with pd.ExcelWriter("admitted_students.xlsx", engine="openpyxl") as writer:
        for stream in STREAMS:
            stream_df = st.session_state.students_df[st.session_state.students_df['Stream'] == stream]
            stream_df.to_excel(writer, sheet_name=stream, index=False)
    with open("admitted_students.xlsx", "rb") as f:
        return f.read()

def save_tc_students():
    return st.session_state.tc_issued_df.to_csv(index=False)

def admission_form():
    st.header("New Admission Form")
    with st.form("admission_form"):
        name = st.text_input("Student Name")
        rank = st.number_input("Rank", min_value=1)
        stream = st.selectbox("Stream", STREAMS)
        language = st.selectbox("Second Language", SECOND_LANGUAGES)
        caste = st.selectbox("Caste", CASTES)
        status = st.selectbox("Admission Status", STATUS_OPTIONS)
        submit = st.form_submit_button("Submit")

        if submit:
            new_student = {
                "Name": name,
                "Rank": rank,
                "Stream": stream,
                "Second_Language": language,
                "Caste": caste,
                "Admission_Status": status,
                "Date_of_Admission": datetime.date.today()
            }
            st.session_state.students_df = pd.concat([st.session_state.students_df, pd.DataFrame([new_student])], ignore_index=True)
            st.success(f"Admission for {name} recorded successfully!")

def tc_generation():
    st.header("TC Issuance")
    if st.session_state.students_df.empty:
        st.info("No students available for TC issuance.")
        return
    name = st.selectbox("Select Student for TC Issuance", st.session_state.students_df["Name"].unique())
    if st.button("Issue TC"):
        student_data = st.session_state.students_df[st.session_state.students_df["Name"] == name].iloc[0].to_dict()
        student_data["Date_of_TC"] = datetime.date.today()
        st.session_state.tc_issued_df = pd.concat([st.session_state.tc_issued_df, pd.DataFrame([student_data])], ignore_index=True)
        st.session_state.students_df = st.session_state.students_df[st.session_state.students_df["Name"] != name]
        st.success(f"TC Issued for {name}")

def data_analysis():
    st.header("Admission Data Analysis")
    if st.session_state.students_df.empty:
        st.info("No admission data available.")
        return
    analysis_type = st.selectbox("Select Analysis Type", ["Stream-wise Distribution", "Caste-wise Distribution", "Admission Status Analysis", "Second Language Distribution", "Date-wise Admission Analysis"])
    counts = st.session_state.students_df[analysis_type.split("-")[0].strip()].value_counts()
    fig, ax = plt.subplots()
    ax.bar(counts.index, counts.values)
    ax.set_title(f"{analysis_type}")
    plt.xticks(rotation=45)
    st.pyplot(fig)

def main():
    st.title(f"{SCHOOL_NAME} Admission Portal")
    st.subheader(f"Academic Year: {ADMISSION_YEAR}")
    page = st.sidebar.radio("Navigation", ["New Admission", "TC Issuance", "Data Analysis"])
    if page == "New Admission":
        admission_form()
    elif page == "TC Issuance":
        tc_generation()
    elif page == "Data Analysis":
        data_analysis()
    
    if not st.session_state.students_df.empty and type_openpyxl:
        admitted_data = save_admitted_students()
        if admitted_data:
            st.download_button(label="Download Admitted Students Data", data=admitted_data, file_name="admitted_students.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    if not st.session_state.tc_issued_df.empty:
        tc_data = save_tc_students()
        st.download_button(label="Download TC Issued Data", data=tc_data, file_name="tc_issued_students.csv", mime="text/csv")

if __name__ == "__main__":
    main()
