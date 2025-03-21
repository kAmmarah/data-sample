import streamlit as st # type: ignore
import pandas as pd
import os
from io import BytesIO
import openpyxl      
# Check if openpyxl is available
# try:
#     import openpyxl
#     OPENPYXL_AVAILABLE = True
# except ImportError:
#     OPENPYXL_AVAILABLE = False
    
OPENPYXL_ERROR_MSG = "openpyxl is not installed. Please install it to enable Excel export."

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    st.error(OPENPYXL_ERROR_MSG)


# Set page config
st.set_page_config(page_title="Data Sweeper", layout='wide')

st.title("Data Sample ✨📚")
st.write("Transform your files between CSV and Excel formats with built-in data cleaning and visualization! 📊")

# Upload file
st.subheader("Upload your file 📁")
uploaded_files = st.file_uploader("Upload your file (CSV or Excel):", type=['csv', 'xlsx'], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        file_ext = os.path.splitext(file.name)[-1].lower()

        if file_ext == ".csv":
            df = pd.read_csv(file)
        elif file_ext == ".xlsx":
            if OPENPYXL_AVAILABLE:
                df = pd.read_excel(file, engine='openpyxl')
            else:
                st.error(OPENPYXL_ERROR_MSG)
                continue
        else:
            st.error(f"Unsupported file type: {file_ext}")
            continue

        # Display file details
        st.write(f"File Name: {file.name}")
        st.write(f"File Size: {file.size / 1024:.2f} KB")

        # Show preview of the DataFrame with row selection
        st.write(f"Preview of {file.name}:")
        rows_to_display = st.slider(f"Select number of rows to display for {file.name}", 5, len(df), 10)
        st.dataframe(df.head(rows_to_display))

        # Data Cleaning Options
        st.subheader("Data Cleaning Options 🧼")
        if st.checkbox(f"Clean Data for {file.name}"):
            col1, col2 = st.columns(2)

            with col1:
                if st.button(f"Remove Duplicates from {file.name}"):
                    df.drop_duplicates(inplace=True)
                    st.write("Duplicates Removed! ✅")

            with col2:
                if st.button(f"Fill Missing Values for {file.name}"):
                    numeric_cols = df.select_dtypes(include=['number']).columns
                    df[numeric_cols] = df[numeric_cols].fillna(df[numeric_cols].mean())
                    st.write("Missing Values have been filled! ✅")

        # Select Specific Columns
        st.subheader("Select Columns to Keep 📋")
        columns = st.multiselect(f"Choose Columns for {file.name}", df.columns, default=df.columns)

        if columns:
            df = df[columns]
            st.write("Filtered Data:")
            st.dataframe(df)
        else:
            st.warning("Please select at least one column.")

        # Data Summary
        st.subheader("Data Summary 📊")
        if st.checkbox(f"Show Data Types for {file.name}"):
            st.write(df.dtypes)
            st.write(df.describe())

        # Data Visualization
        st.subheader("Data Visualization 📈")
        if st.checkbox(f"Show Visualizations for {file.name}"):
            st.bar_chart(df.select_dtypes(include='number').iloc[:, :2])

        # File Conversion
        st.subheader("Conversion Options 🔄")
        conversion_type = st.radio(f"Convert {file.name} to:", ["CSV", "Excel"], key=file.name)

        if st.button(f"Convert {file.name}"):
            buffer = BytesIO()

            if conversion_type == "CSV":
                df.to_csv(buffer, index=False)
                file_name = file.name.replace(file_ext, ".csv")
                mime_type = "text/csv"
            else:
                if OPENPYXL_AVAILABLE:
                    df.to_excel(buffer, index=False, engine='openpyxl')
                    file_name = file.name.replace(file_ext, ".xlsx")
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                else:
                    st.error("Cannot convert to Excel as openpyxl is not installed.")
                    continue

            buffer.seek(0)
            st.download_button(
                label=f"Download {file.name} as {conversion_type} 📥",
                data=buffer,
                file_name=file_name,
                mime=mime_type
            )

            st.success(f"{file.name} processed successfully! 🎉")

# Add a footer indicating the app creator with a colorful design
st.markdown(
    """
    <div style="background-color: #4CAF50; padding: 10px; border-radius: 5px; text-align: center;">
        <h3 style="color: white;">This app was created by Ammara 🌟</h3>
    </div>
    """,
    unsafe_allow_html=True
)