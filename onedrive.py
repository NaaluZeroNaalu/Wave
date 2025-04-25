import pandas as pd
import io
import requests
import streamlit as st
import base64

# OneDrive shared link for the Excel file
shared_link = "https://1drv.ms/x/s!Axxxxxx"  # Replace with your actual shared link
encoded_link = base64.urlsafe_b64encode(shared_link.encode()).decode()
download_url = f"https://api.onedrive.com/v1.0/shares/u!{encoded_link}/driveItem/content"

# Function to download and read Excel file from OneDrive
def download_excel_from_onedrive(download_url):
    try:
        response = requests.get(download_url, stream=True)
        if response.status_code != 200:
            raise Exception(f"Failed to download file: {response.status_code}")
        bytes_file_obj = io.BytesIO(response.content)
        bytes_file_obj.seek(0)
        df = pd.read_excel(bytes_file_obj)
        return df
    except Exception as e:
        st.error(f"Error downloading file: {e}")
        return None

# Streamlit app
st.title("Display OneDrive Excel File")

# Download and display the Excel file
schedule_df = download_excel_from_onedrive(download_url)
if schedule_df is not None:
    st.write("Excel File Contents:")
    st.dataframe(schedule_df)  # Display the DataFrame
    st.write("Columns:", schedule_df.columns.tolist())  # Show column names
else:
    st.warning("Failed to load Excel file from OneDrive")