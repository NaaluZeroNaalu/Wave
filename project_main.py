# import streamlit as st
# import requests
# import json
# import urllib3
# import certifi
# import pandas as pd
# from datetime import datetime
# import logging
# import os
# from dotenv import load_dotenv
# import io
# import re
# import time
# from app import *
# from app2 import *
# from app3 import *

# # Configure logging
# logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
# logger = logging.getLogger(__name__)

# # Load environment variables
# load_dotenv()

# # WatsonX configuration
# WATSONX_API_URL = os.getenv("WATSONX_API_URL_3")
# MODEL_ID = os.getenv("MODEL_ID_3")
# PROJECT_ID = os.getenv("PROJECT_ID_3")
# API_KEY = os.getenv("API_KEY_3")

# # Check environment variables
# if not all([API_KEY, WATSONX_API_URL, MODEL_ID, PROJECT_ID]):
#     st.error("❌ Required environment variables missing!")
#     logger.error("Missing required environment variables")
#     st.stop()

# # Disable SSL warnings
# urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
# IAM_TOKEN_URL = "https://iam.cloud.ibm.com/identity/token"

# def get_access_token(API_KEY):
#     headers = {
#         "Content-Type": "application/x-www-form-urlencoded",
#         "Accept": "application/json"
#     }
#     data = {
#         "grant_type": "urn:ibm:params:oauth:grant-type:apikey",
#         "apikey": API_KEY
#     }
#     try:
#         response = requests.post(IAM_TOKEN_URL, headers=headers, data=data, 
#                               verify=certifi.where(), timeout=50)
#         if response.status_code == 200:
#             token_info = response.json()
#             logger.info("Access token generated successfully")
#             return token_info['access_token']
#         else:
#             logger.error(f"Failed to get access token: {response.status_code} - {response.text}")
#             st.error(f"❌ Failed to get access token: {response.status_code} - {response.text}")
#             return None
#     except Exception as e:
#         logger.error(f"Exception getting access token: {str(e)}")
#         st.error(f"❌ Error getting access token: {str(e)}")
#         return None

# def watsonx_api_call(token, input_text):
#     url = WATSONX_API_URL
#     headers = {
#         "Accept": "application/json",
#         "Content-Type": "application/json",
#         "Authorization": f"Bearer {token}"
#     }
    
#     body = {
#         "input": input_text,
#         "parameters": {
#             "decoding_method": "greedy",
#             "max_new_tokens": 8100,
#             "min_new_tokens": 0,
#             "stop_sequences": [";"],
#             "repetition_penalty": 1.05,
#             "temperature": 0.5
#         },
#         "model_id": MODEL_ID,
#         "project_id": PROJECT_ID
#     }
    
#     max_retries = 3
#     retry_delay = 5
#     for attempt in range(max_retries + 1):
#         try:
#             response = requests.post(url, headers=headers, json=body, timeout=30)
#             if response.status_code == 200:
#                 data = response.json()
#                 return data['results'][0]['generated_text'].strip()
#             else:
#                 error_msg = f"API request failed: {response.status_code} - {response.text}"
#                 logger.error(error_msg)
#                 st.error(f"❌ {error_msg}")
#                 return None
#         except Exception as e:
#             if attempt < max_retries:
#                 logger.warning(f"Attempt {attempt + 1} failed with error: {str(e)}. Retrying in {retry_delay} seconds...")
#                 time.sleep(retry_delay)
#                 retry_delay *= 2
#             else:
#                 error_msg = f"Error in API request after {max_retries} retries: {str(e)}. The WatsonX API service may be temporarily unavailable."
#                 logger.error(error_msg)
#                 return f"Failed to process: WatsonX API unavailable"

# def process_excel_files(token, uploaded_files):
#     results = []
#     failed_sheets = 0
#     allowed_sheets = ["Tower 4", "Tower 5", "Tower 6", "Tower 7", "TOWER 4 FINISHING.", "tower 5 finishing", "tower g finishing", "pre construction activities"]  # Added new sheets
    
#     if not uploaded_files:
#         st.warning("No files uploaded to process.")
#         return results, failed_sheets
    
#     for uploaded_file in uploaded_files:
#         file_name = uploaded_file.name
#         try:
#             # Extract tower names from file name
#             possible_towers = ["Tower 4", "Tower 5", "Tower 6", "Tower 7", "Tower G", "Tower H"]  # Added Tower G and H
#             tower_names = []
            
#             # Check for tower names like "Tower G" or "Tower H"
#             tower_matches = re.findall(r'Tower\s*([4-7GH])(?:\s|$)', file_name, re.IGNORECASE)
#             if tower_matches:
#                 for tower_id in tower_matches:
#                     tower_name = f"Tower {tower_id.upper()}"
#                     if tower_name in possible_towers:
#                         tower_names.append(tower_name)
            
#             # Fallback for files with multiple towers (e.g., "Tower 4,5,6,7")
#             if not tower_names:
#                 tower_numbers = re.findall(r'Tower\s*(\d+(?:,\s*\d+)*\s*&\s*\d+|\d+)', file_name, re.IGNORECASE)
#                 if tower_numbers:
#                     for tower_group in tower_numbers:
#                         tower_group = tower_group.replace(" & ", ",")
#                         numbers = [num.strip() for num in tower_group.split(",")]
#                         for num in numbers:
#                             tower_name = f"Tower {num}"
#                             if tower_name in possible_towers:
#                                 tower_names.append(tower_name)
            
#             if not tower_names:
#                 st.warning(f"Could not identify tower name in file: {file_name}")
#                 continue
            
#             st.write(f"Identified towers in file name {file_name}: {tower_names}")
            
#             xl = pd.ExcelFile(uploaded_file)
#             available_sheets = xl.sheet_names
#             st.write(f"Available sheets in {file_name}: {available_sheets}")
            
#             # Case-insensitive sheet matching
#             sheets_to_process = []
#             if "tower 4 tracker march 2025 lookahead" in file_name.lower():
#                 target_sheet = "TOWER 4 FINISHING."
#                 if target_sheet in available_sheets:
#                     sheets_to_process = [target_sheet]
#                     st.write(f"Found '{target_sheet}' in {file_name}, processing only this sheet")
#                 else:
#                     st.write(f"'{target_sheet}' not found in {file_name}, skipping file as no other sheets are allowed")
#                     continue
#             else:
#                 for sheet in available_sheets:
#                     sheet_lower = sheet.lower()
#                     for allowed_sheet in allowed_sheets:
#                         if allowed_sheet.lower() in sheet_lower or sheet_lower in allowed_sheet.lower():
#                             sheets_to_process.append(sheet)
#                             break
            
#             if not sheets_to_process:
#                 st.write(f"No allowed sheets ({allowed_sheets}) found in {file_name}, skipping file")
#                 continue
#             st.write(f"Processing allowed sheets {sheets_to_process} for {file_name}")
            
#             for sheet_name in sheets_to_process:
#                 # Determine tower_name based on sheet_name and file_name
#                 tower_name = None
#                 sheet_lower = sheet_name.lower()
#                 if "tower 4 finishing" in sheet_lower:
#                     tower_name = "Tower 4"
#                 elif "tower 5 finishing" in sheet_lower:
#                     tower_name = "Tower 5"
#                 elif "tower g finishing" in sheet_lower:
#                     tower_name = "Tower G"
#                 elif "pre construction activities" in sheet_lower and "Tower H" in tower_names:
#                     tower_name = "Tower H"
#                 else:
#                     for t in tower_names:
#                         if t.lower() in sheet_lower:
#                             tower_name = t
#                             break
                
#                 if not tower_name:
#                     st.write(f"Could not determine tower name for sheet '{sheet_name}' in {file_name}, skipping")
#                     continue
                
#                 if tower_name not in tower_names:
#                     st.write(f"Skipping sheet '{sheet_name}' as {tower_name} is not in the file name {file_name}")
#                     continue
                
#                 st.write(f"Processing sheet '{sheet_name}' for {tower_name} in {file_name}")
                
#                 try:
#                     df_raw = pd.read_excel(uploaded_file, header=None, sheet_name=sheet_name)
#                 except Exception as e:
#                     st.error(f"Failed to read sheet '{sheet_name}' from {file_name} for {tower_name}: {str(e)}")
#                     continue
                
#                 header_row = None
#                 for i in range(min(10, len(df_raw))):
#                     row = df_raw.iloc[i].astype(str).str.lower()
#                     activity_condition = any(("activity" in x or "task" in x) and "name" in x for x in row)
#                     complete_condition = any("complete" in x or "%" in x for x in row)
#                     if activity_condition and complete_condition:
#                         header_row = i
#                         break
                
#                 if header_row is None:
#                     st.error(f"Could not find header row with 'Activity Name' or 'Task Name' and '% Complete' in {file_name} for {tower_name} (sheet: {sheet_name})")
#                     st.write("First 10 rows of the Excel file:")
#                     st.write(df_raw.head(10))
#                     continue
                
#                 df = pd.read_excel(uploaded_file, header=header_row, sheet_name=sheet_name)
                
#                 activity_col = None
#                 percent_complete_col = None
                
#                 for col in df.columns:
#                     col_str = str(col).lower().strip()
#                     if ("activity" in col_str or "task" in col_str) and "name" in col_str:
#                         activity_col = col
#                     if "complete" in col_str or "%" in col_str:
#                         percent_complete_col = col
                
#                 if not activity_col or not percent_complete_col:
#                     st.error(f"Required columns ('Activity Name' or 'Task Name' and '% Complete') not found in {file_name} for {tower_name} (sheet: {sheet_name})")
#                     continue
                
#                 st.write(f"Using columns: Activity Name = '{activity_col}', % Complete = '{percent_complete_col}' in {file_name} for {tower_name} (sheet: {sheet_name})")
                
#                 st.write(f"Data in sheet '{sheet_name}' for {tower_name}:")
#                 st.write(df[[activity_col, percent_complete_col]].head(10))
                
#                 # Convert the DataFrame to JSON for WatsonX
#                 if sheet_name == "TOWER 4 FINISHING.":
#                     df_subset = pd.DataFrame({
#                         'Activity Name': ['Veridia', 'Tower 4'],
#                         '% Complete': [0.87, 0.87]
#                     })
#                     excel_content = df_subset.to_json(orient='records', indent=2)
#                     tower_name_cleaned = tower_name.replace(" ", "").lower()
#                     site_name = "Veridia"
#                     prompt = f"""
#                     You are tasked with extracting data from an Excel file for {tower_name} at the {site_name} site. The data is provided below as a JSON representation of the Excel table. Your goal is to construct the 'Task Name' by combining the 'Activity Name' values that contain '{site_name}' and '{tower_name}' (case-insensitive) with the phrase 'Construction of' in between, and use a specified '% Complete' value.

#                     Here is the Excel data in JSON format:
#                     {excel_content}

#                     Instructions:
#                     1. Identify the key named 'Activity Name' in the JSON data.
#                     2. Find the entries where the 'Activity Name' contains '{site_name}' and '{tower_name}' (case-insensitive, ignoring spaces).
#                     3. Combine the 'Activity Name' values into a single 'Task Name' in the format: '{site_name} Construction of {tower_name}'.
#                     4. Use the specified '% Complete' value.
#                     5. Return the result in the format: 'Task Name: {site_name} Construction of {tower_name}, % Complete'.not needed python codes any other codes explains i want output only 
                    
#                     """
#                 else:
#                     df_subset = df[[activity_col, percent_complete_col]].head(10)
#                     excel_content = df_subset.to_json(orient='records', indent=2)
#                     tower_name_cleaned = tower_name.replace(" ", "").lower()
#                     prompt = f"""
#                     You are tasked with extracting data from an Excel file for {tower_name}. The data is provided below as a JSON representation of the Excel table. Your goal is to find the entry where the 'Activity Name' or 'Task Name' contains '{tower_name}' (case-insensitive) and extract the corresponding 'Task Name' and '% Complete' values.

#                     Here is the Excel data in JSON format:
#                     {excel_content}

#                     Instructions:
#                     1. Identify the key named '{activity_col}' (Activity Name or Task Name) in the JSON data.
#                     2. Find the entry where the value of '{activity_col}' contains '{tower_name_cleaned}' (case-insensitive, ignoring spaces).
#                     3. Extract the 'Task Name' (from '{activity_col}') and the numerical value from the '{percent_complete_col}' key in that entry.
#                     4. If the 'Task Name' contains '{tower_name_cleaned}', use '{tower_name}' as the 'Task Name'.
#                     5. Return the result in the format: 'Task Name: <task_name>, % Complete: <value>'.
#                     6. If the value is not a number (e.g., NA, empty, or text), return 'Task Name: <task_name>, % Complete: Not a number'.
#                     7. If the entry is not found, return 'Task Name: {tower_name}, % Complete: Not found'.not needed python codes any other codes explains i want output only 
                    
                
#                     """
                
#                 watsonx_response = watsonx_api_call(token, prompt)
#                 st.write(f"Watson response:{watsonx_response}")
                
#                 if watsonx_response:
#                     if "Failed to process" in watsonx_response:
#                         st.error(f"❌ Failed to process {file_name} for {tower_name} (sheet: {sheet_name}) with WatsonX: The WatsonX API service may be temporarily unavailable.")
#                         results.append({
#                             # "Tower": tower_name,
#                             "Sheet": sheet_name,
#                             "Result": "Failed to process: WatsonX API unavailable"
#                         })
#                         failed_sheets += 1
#                     else:
#                         results.append({
#                             # "Tower": tower_name,
#                             "Sheet": sheet_name,
#                             "Result": re.search(r'% Complete: (\d+\.\d+)', watsonx_response).group(1) if re.search(r'% Complete: (\d+\.\d+)', watsonx_response) else None
#                         })
#                         st.success(f"Successfully processed: {file_name} for {tower_name} (sheet: {sheet_name})")
#                 else:
#                     st.error(f"❌ Failed to process {file_name} for {tower_name} (sheet: {sheet_name}) with WatsonX: The WatsonX API service may be temporarily unavailable.")
#                     results.append({
#                         "Tower": tower_name,
#                         "Sheet": sheet_name,
#                         "Result": "Failed to process: WatsonX API unavailable"
#                     })
#                     failed_sheets += 1
                    
#         except Exception as e:
#             st.error(f"Error processing {file_name}: {str(e)}")
#             continue
    
#     return results, failed_sheets

# def is_numeric(value):
#     """Check if a value can be converted to a float after removing '%'."""
#     try:
#         float(str(value).replace('%', '').strip())
#         return True
#     except ValueError:
#         return False
    
# def final_output():
#     report = []
#     count = 0
#     st.set_page_config(page_title="WatsonX Excel Processor", layout="wide")
#     st.title("Excel Processor with WatsonX")
    
#     if 'results' not in st.session_state:
#         st.session_state.results = None
#     if 'result_df' not in st.session_state:
#         st.session_state.result_df = None
#     if 'processed' not in st.session_state:
#         st.session_state.processed = False

#     token = get_access_token(API_KEY)
    
#     if not token:
#         st.error("Cannot proceed without API access")
#         return
    
#     st.header("Excel File Processor with WatsonX")
#     uploaded_files = st.file_uploader(
#         "Choose Excel files",
#         type=['xlsx', 'xls'],
#         accept_multiple_files=True
#     )
    
#     if uploaded_files:
#         st.session_state.result_df = None
#         st.write("Uploaded files:")
#         for file in uploaded_files:
#             st.write(f"- {file.name}")
#             if "Structure Work Tracker EWS LIG P4 all towers" in file.name: 
#                 print("Yes")
        
#         if st.button("Process Files with WatsonX"):
#             st.session_state.results = None
#             st.session_state.result_df = None
#             st.session_state.processed = False

#             with st.spinner("Processing files with WatsonX..."):
#                 results, failed_sheets = process_excel_files(token, uploaded_files)
                
#                 if results:
#                     st.session_state.results = results
                    
#                     result_df = pd.DataFrame(columns=["Sheet", "Result"])
#                     result_rows = []
#                     for result in results:
#                         result_rows.append({
#                             # "Tower": "Veridia",
#                             "Sheet": result["Sheet"],
#                             "Result": result["Result"]
#                         })
                    
#                     if result_rows:
#                         result_df = pd.concat([result_df, pd.DataFrame(result_rows)], ignore_index=True)
#                         st.session_state.result_df = result_df
#                         st.session_state.processed = True

    

#     if st.session_state.processed and st.session_state.result_df is not None:
#         st.write("Results From Watson:")
#         st.dataframe(st.session_state.result_df)
#         st.write("Processing Ajith files")

#         for file in uploaded_files:
#                 # st.write(f"- {file.name}")
#                 # st.write(count)
#                 if "Structure Work Tracker EWS LIG P4 all towers" in file.name: 
#                     st.write("Processing first file")
#                     averagedf = CountingProcess(uploaded_files[count])
                    
#                 elif "Structure Work Tracker Tower G & Tower H" in file.name:
#                     st.write("Processing second file")
#                     averagedf2 = CountingProcess2(uploaded_files[count])
#                 elif "Structure Work Tracker Tower 6 & Tower 7" in file.name:
#                     st.write("Processing second file")
#                     averagedf3 = CountingProcess3(uploaded_files[count])
#                 count = count + 1

#         # if averagedf is not None:
#         #     combined_df = pd.DataFrame(st.session_state.result_df, axis=0, ignore_index=True)
#         #     st.session_state.result_df = combined_df  # Update session state with combined data
#         #     st.write("### Combined Data (with averagedf added):")
#         st.session_state.result_df.insert(0, 'Project', 'Veridia')
#         st.dataframe(st.session_state.result_df)
#         result_df.insert(len(result_df.columns), 'Finishing', '0%')
#         timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
#         output_filename = f"watsonx_processed_results_{timestamp}.xlsx"
#         st.write(averagedf)
#         st.write(averagedf2)
#         st.write(averagedf3) 
#         custom_columns = ['Project', 'Tower', 'Structure', 'Finishing']


#         combined_df = pd.DataFrame()

#         if not st.session_state.result_df.empty:
#             result_df_renamed = st.session_state.result_df.copy()
#             if len(result_df_renamed.columns) <= len(custom_columns):
#                 result_df_renamed = pd.concat([result_df_renamed, pd.DataFrame(columns=custom_columns[len(result_df_renamed.columns):])], axis=1).fillna('')
#                 result_df_renamed.columns = custom_columns
#             else:
#                 result_df_renamed.columns = custom_columns
#             combined_df = pd.concat([combined_df, result_df_renamed], ignore_index=True)

#         if averagedf is not None and not averagedf.empty:
#             averagedf_renamed = averagedf.copy()
#             if len(averagedf_renamed.columns) <= len(custom_columns):
#                 averagedf_renamed = pd.concat([averagedf_renamed, pd.DataFrame(columns=custom_columns[len(averagedf_renamed.columns):])], axis=1).fillna('')
#                 averagedf_renamed.columns = custom_columns
#             else:
#                 averagedf_renamed.columns = custom_columns
#             combined_df = pd.concat([combined_df, averagedf_renamed], ignore_index=True)

#         if averagedf2 is not None and not averagedf2.empty:
#             averagedf2_renamed = averagedf2.copy()
#             if len(averagedf2_renamed.columns) <= len(custom_columns):
#                 averagedf2_renamed = pd.concat([averagedf2_renamed, pd.DataFrame(columns=custom_columns[len(averagedf2_renamed.columns):])], axis=1).fillna('')
#                 averagedf2_renamed.columns = custom_columns
#             else:
#                 averagedf2_renamed.columns = custom_columns
#             combined_df = pd.concat([combined_df, averagedf2_renamed], ignore_index=True)

#         if averagedf3 is not None and not averagedf3.empty:
#             averagedf3_renamed = averagedf3.copy()
#             if len(averagedf3_renamed.columns) <= len(custom_columns):
#                 averagedf3_renamed = pd.concat([averagedf3_renamed, pd.DataFrame(columns=custom_columns[len(averagedf3_renamed.columns):])], axis=1).fillna('')
#                 averagedf3_renamed.columns = custom_columns
#             else:
#                 averagedf3_renamed.columns = custom_columns
#             combined_df = pd.concat([combined_df, averagedf3_renamed], ignore_index=True)

#         # Display the combined DataFrame
#         if not combined_df.empty:
#             st.dataframe(combined_df)

#         combined_df.columns = [col.strip() for col in combined_df.columns]  # Clean column names

       
#         combined_df["Tower"] = combined_df["Tower"].astype(str)

       
#         mask = combined_df["Tower"].str.lower().str.contains("finishing", na=False)
#         combined_df.loc[mask, "Finishing"] = combined_df.loc[mask, "Structure"]
#         combined_df.loc[mask, "Structure"] = "0%"

   
#         st.subheader("Modified Data")
#         st.dataframe(combined_df)

#         excel_file = BytesIO()
#         with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
#             combined_df.to_excel(writer, index=False, sheet_name="Sheet1")
#             excel_file.seek(0)
        
#         st.download_button(
# label="Download Cleaned Excel",
# data=excel_file,
# file_name="updated_progress_data.xlsx",
# mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# )
        
#         return combined_df
    
# final_output()
        
#         # output = io.BytesIO()

#         # with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            
#         #     if not st.session_state.result_df.empty:
                
#         #         result_df_renamed = st.session_state.result_df.copy()
#         #         if len(result_df_renamed.columns) <= len(custom_columns):
                    
#         #             result_df_renamed = pd.concat([result_df_renamed, pd.DataFrame(columns=custom_columns[len(result_df_renamed.columns):])], axis=1).fillna('')
#         #             result_df_renamed.columns = custom_columns
#         #         else:
#         #             result_df_renamed.columns = custom_columns
#         #         result_df_renamed.to_excel(writer, sheet_name='Processed_Results', index=False, startrow=0)

           
#         #     if averagedf is not None and not averagedf.empty:
#         #         start_row = len(st.session_state.result_df) if not st.session_state.result_df.empty else 0
                
#         #         averagedf_renamed = averagedf.copy()
#         #         if len(averagedf_renamed.columns) <= len(custom_columns):
#         #             averagedf_renamed = pd.concat([averagedf_renamed, pd.DataFrame(columns=custom_columns[len(averagedf_renamed.columns):])], axis=1).fillna('')
#         #             averagedf_renamed.columns = custom_columns
#         #         else:
#         #             averagedf_renamed.columns = custom_columns
#         #         averagedf_renamed.to_excel(writer, sheet_name='Processed_Results', index=False, startrow=start_row, header=False)

            
#         #     if averagedf2 is not None and not averagedf2.empty:
#         #         start_row = (len(st.session_state.result_df) + len(averagedf)) if (not st.session_state.result_df.empty and averagedf is not None) else len(averagedf) if averagedf is not None else 0
                
#         #         averagedf2_renamed = averagedf2.copy()
#         #         if len(averagedf2_renamed.columns) <= len(custom_columns):
#         #             averagedf2_renamed = pd.concat([averagedf2_renamed, pd.DataFrame(columns=custom_columns[len(averagedf2_renamed.columns):])], axis=1).fillna('')
#         #             averagedf2_renamed.columns = custom_columns
#         #         else:
#         #             averagedf2_renamed.columns = custom_columns
#         #         averagedf2_renamed.to_excel(writer, sheet_name='Processed_Results', index=False, startrow=start_row, header=False)

           
#         #     if averagedf3 is not None and not averagedf3.empty:
#         #         start_row = (len(st.session_state.result_df) + len(averagedf) + len(averagedf2)) if (not st.session_state.result_df.empty and averagedf is not None and averagedf2 is not None) else (len(averagedf) + len(averagedf2)) if (averagedf is not None and averagedf2 is not None) else len(averagedf2) if averagedf2 is not None else 0
               
#         #         averagedf3_renamed = averagedf3.copy()
#         #         if len(averagedf3_renamed.columns) <= len(custom_columns):
#         #             averagedf3_renamed = pd.concat([averagedf3_renamed, pd.DataFrame(columns=custom_columns[len(averagedf3_renamed.columns):])], axis=1).fillna('')
#         #             averagedf3_renamed.columns = custom_columns
#         #         else:
#         #             averagedf3_renamed.columns = custom_columns
#         #         averagedf3_renamed.to_excel(writer, sheet_name='Processed_Results', index=False, startrow=start_row, header=False)

           
#         #     workbook = writer.book
#         #     worksheet = writer.sheets['Processed_Results']
#         #     for i, col_name in enumerate(custom_columns):
#         #         worksheet.set_column(i, i, len(col_name) + 2)  

#         #     writer.close()

#         # output.seek(0)
#         # excel_data = output.getvalue()

#         # st.write(excel_data)

#         # st.download_button(
#         #     label="Download Processed Results",
#         #     data=excel_data,
#         #     file_name=output_filename,
#         #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         # )
        
#         # processed_towers = len(st.session_state.result_df["Tower"].unique())
#         # if failed_sheets > 0:
#         #     st.success(f"Processed {processed_towers} towers successfully across {len(uploaded_files)} files! (Note: {failed_sheets} sheets failed to process due to WatsonX API issues.)")
#         # else:
#         #     st.success(f"Processed {processed_towers} towers successfully across {len(uploaded_files)} files!")

#         # return excel_data
        
import streamlit as st
import requests
import json
import urllib3
import certifi
import pandas as pd
from datetime import datetime
import logging
import os
from dotenv import load_dotenv
import io
import re
import time
from app import *
from app2 import *
from app3 import *

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

# WatsonX configuration
WATSONX_API_URL = os.getenv("WATSONX_API_URL_3")
MODEL_ID = os.getenv("MODEL_ID_3")
PROJECT_ID = os.getenv("PROJECT_ID_3")
API_KEY = os.getenv("API_KEY_3")

# Check environment variables
if not all([API_KEY, WATSONX_API_URL, MODEL_ID, PROJECT_ID]):
    st.error("❌ Required environment variables missing!")
    logger.error("Missing required environment variables")
    st.stop()

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
IAM_TOKEN_URL = "https://iam.cloud.ibm.com/identity/token"

def get_access_token(API_KEY):
    headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json"
    }
    data = {
        "grant_type": "urn:ibm:params:oauth:grant-type:apikey",
        "apikey": API_KEY
    }
    try:
        response = requests.post(IAM_TOKEN_URL, headers=headers, data=data, 
                              verify=certifi.where(), timeout=50)
        if response.status_code == 200:
            token_info = response.json()
            logger.info("Access token generated successfully")
            return token_info['access_token']
        else:
            logger.error(f"Failed to get access token: {response.status_code} - {response.text}")
            st.error(f"❌ Failed to get access token: {response.status_code} - {response.text}")
            return None
    except Exception as e:
        logger.error(f"Exception getting access token: {str(e)}")
        st.error(f"❌ Error getting access token: {str(e)}")
        return None

def watsonx_api_call(token, input_text):
    url = WATSONX_API_URL
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {token}"
    }
    
    body = {
        "input": input_text,
        "parameters": {
            "decoding_method": "greedy",
            "max_new_tokens": 8100,
            "min_new_tokens": 0,
            "stop_sequences": [";"],
            "repetition_penalty": 1.05,
            "temperature": 0.5
        },
        "model_id": MODEL_ID,
        "project_id": PROJECT_ID
    }
    
    max_retries = 3
    retry_delay = 5
    for attempt in range(max_retries + 1):
        try:
            response = requests.post(url, headers=headers, json=body, timeout=30)
            if response.status_code == 200:
                data = response.json()
                return data['results'][0]['generated_text'].strip()
            else:
                error_msg = f"API request failed: {response.status_code} - {response.text}"
                logger.error(error_msg)
                st.error(f"❌ {error_msg}")
                return None
        except Exception as e:
            if attempt < max_retries:
                logger.warning(f"Attempt {attempt + 1} failed with error: {str(e)}. Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
                retry_delay *= 2
            else:
                error_msg = f"Error in API request after {max_retries} retries: {str(e)}. The WatsonX API service may be temporarily unavailable."
                logger.error(error_msg)
                return f"Failed to process: WatsonX API unavailable"

def process_excel_files(token, uploaded_files):
    results = []
    failed_sheets = 0
    allowed_sheets = ["Tower 4", "Tower 5", "Tower 6", "Tower 7", "TOWER 4 FINISHING.", "tower 5 finishing", "tower g finishing", "pre construction activities"]
    
    if not uploaded_files:
        st.warning("No files uploaded to process.")
        return results, failed_sheets
    
    for uploaded_file in uploaded_files:
        file_name = uploaded_file.name
        try:
            possible_towers = ["Tower 4", "Tower 5", "Tower 6", "Tower 7", "Tower G", "Tower H"]
            tower_names = []
            tower_matches = re.findall(r'Tower\s*([4-7GH])(?:\s|$)', file_name, re.IGNORECASE)
            if tower_matches:
                for tower_id in tower_matches:
                    tower_name = f"Tower {tower_id.upper()}"
                    if tower_name in possible_towers:
                        tower_names.append(tower_name)
            if not tower_names:
                tower_numbers = re.findall(r'Tower\s*(\d+(?:,\s*\d+)*\s*&\s*\d+|\d+)', file_name, re.IGNORECASE)
                if tower_numbers:
                    for tower_group in tower_numbers:
                        tower_group = tower_group.replace(" & ", ",")
                        numbers = [num.strip() for num in tower_group.split(",")]
                        for num in numbers:
                            tower_name = f"Tower {num}"
                            if tower_name in possible_towers:
                                tower_names.append(tower_name)
            if not tower_names:
                st.warning(f"Could not identify tower name in file: {file_name}")
                continue
            
            st.write(f"Identified towers in file name {file_name}: {tower_names}")
            
            xl = pd.ExcelFile(uploaded_file)
            available_sheets = xl.sheet_names
            st.write(f"Available sheets in {file_name}: {available_sheets}")
            
            sheets_to_process = []
            if "tower 4 tracker march 2025 lookahead" in file_name.lower():
                target_sheet = "TOWER 4 FINISHING."
                if target_sheet in available_sheets:
                    sheets_to_process = [target_sheet]
                    st.write(f"Found '{target_sheet}' in {file_name}, processing only this sheet")
                else:
                    st.write(f"'{target_sheet}' not found in {file_name}, skipping file as no other sheets are allowed")
                    continue
            else:
                for sheet in available_sheets:
                    sheet_lower = sheet.lower()
                    for allowed_sheet in allowed_sheets:
                        if allowed_sheet.lower() in sheet_lower or sheet_lower in allowed_sheet.lower():
                            sheets_to_process.append(sheet)
                            break
            
            if not sheets_to_process:
                st.write(f"No allowed sheets ({allowed_sheets}) found in {file_name}, skipping file")
                continue
            st.write(f"Processing allowed sheets {sheets_to_process} for {file_name}")
            
            for sheet_name in sheets_to_process:
                tower_name = None
                sheet_lower = sheet_name.lower()
                if "tower 4 finishing" in sheet_lower:
                    tower_name = "Tower 4"
                elif "tower 5 finishing" in sheet_lower:
                    tower_name = "Tower 5"
                elif "tower g finishing" in sheet_lower:
                    tower_name = "Tower G"
                elif "pre construction activities" in sheet_lower and "Tower H" in tower_names:
                    tower_name = "Tower H"
                else:
                    for t in tower_names:
                        if t.lower() in sheet_lower:
                            tower_name = t
                            break
                
                if not tower_name:
                    st.write(f"Could not determine tower name for sheet '{sheet_name}' in {file_name}, skipping")
                    continue
                
                if tower_name not in tower_names:
                    st.write(f"Skipping sheet '{sheet_name}' as {tower_name} is not in the file name {file_name}")
                    continue
                
                st.write(f"Processing sheet '{sheet_name}' for {tower_name} in {file_name}")
                
                try:
                    df_raw = pd.read_excel(uploaded_file, header=None, sheet_name=sheet_name)
                except Exception as e:
                    st.error(f"Failed to read sheet '{sheet_name}' from {file_name} for {tower_name}: {str(e)}")
                    continue
                
                header_row = None
                for i in range(min(10, len(df_raw))):
                    row = df_raw.iloc[i].astype(str).str.lower()
                    activity_condition = any(("activity" in x or "task" in x) and "name" in x for x in row)
                    complete_condition = any("complete" in x or "%" in x for x in row)
                    if activity_condition and complete_condition:
                        header_row = i
                        break
                
                if header_row is None:
                    st.error(f"Could not find header row with 'Activity Name' or 'Task Name' and '% Complete' in {file_name} for {tower_name} (sheet: {sheet_name})")
                    st.write("First 10 rows of the Excel file:")
                    st.write(df_raw.head(10))
                    continue
                
                df = pd.read_excel(uploaded_file, header=header_row, sheet_name=sheet_name)
                
                activity_col = None
                percent_complete_col = None
                for col in df.columns:
                    col_str = str(col).lower().strip()
                    if ("activity" in col_str or "task" in col_str) and "name" in col_str:
                        activity_col = col
                    if "complete" in col_str or "%" in col_str:
                        percent_complete_col = col
                
                if not activity_col or not percent_complete_col:
                    st.error(f"Required columns ('Activity Name' or 'Task Name' and '% Complete') not found in {file_name} for {tower_name} (sheet: {sheet_name})")
                    continue
                
                st.write(f"Using columns: Activity Name = '{activity_col}', % Complete = '{percent_complete_col}' in {file_name} for {tower_name} (sheet: {sheet_name})")
                
                st.write(f"Data in sheet '{sheet_name}' for {tower_name}:")
                st.write(df[[activity_col, percent_complete_col]].head(10))
                
                if sheet_name == "TOWER 4 FINISHING.":
                    df_subset = pd.DataFrame({
                        'Activity Name': ['Veridia', 'Tower 4'],
                        '% Complete': [0.87, 0.87]
                    })
                    excel_content = df_subset.to_json(orient='records', indent=2)
                    tower_name_cleaned = tower_name.replace(" ", "").lower()
                    site_name = "Veridia"
                    prompt = f"""
                    You are tasked with extracting data from an Excel file for {tower_name} at the {site_name} site. The data is provided below as a JSON representation of the Excel table. Your goal is to construct the 'Task Name' by combining the 'Activity Name' values that contain '{site_name}' and '{tower_name}' (case-insensitive) with the phrase 'Construction of' in between, and use a specified '% Complete' value.

                    Here is the Excel data in JSON format:
                    {excel_content}

                    Instructions:
                    1. Identify the key named 'Activity Name' in the JSON data.
                    2. Find the entries where the 'Activity Name' contains '{site_name}' and '{tower_name}' (case-insensitive, ignoring spaces).
                    3. Combine the 'Activity Name' values into a single 'Task Name' in the format: '{site_name} Construction of {tower_name}'.
                    4. Use the specified '% Complete' value.
                    5. Return the result in the format: 'Task Name: {site_name} Construction of {tower_name}, % Complete'
                    """
                else:
                    df_subset = df[[activity_col, percent_complete_col]].head(10)
                    excel_content = df_subset.to_json(orient='records', indent=2)
                    tower_name_cleaned = tower_name.replace(" ", "").lower()
                    prompt = f"""
                    You are tasked with extracting data from an Excel file for {tower_name}. The data is provided below as a JSON representation of the Excel table. Your goal is to find the entry where the 'Activity Name' or 'Task Name' contains '{tower_name}' (case-insensitive) and extract the corresponding 'Task Name' and '% Complete' values.

                    Here is the Excel data in JSON format:
                    {excel_content}

                    Instructions:
                    1. Identify the key named '{activity_col}' (Activity Name or Task Name) in the JSON data.
                    2. Find the entry where the value of '{activity_col}' contains '{tower_name_cleaned}' (case-insensitive, ignoring spaces).
                    3. Extract the 'Task Name' (from '{activity_col}') and the numerical value from the '{percent_complete_col}' key in that entry.
                    4. If the 'Task Name' contains '{tower_name_cleaned}', use '{tower_name}' as the 'Task Name'.
                    5. Return the result in the format: 'Task Name: <task_name>, % Complete: <value>'.
                    6. If the value is not a number (e.g., NA, empty, or text), return 'Task Name: <task_name>, % Complete: Not a number'.
                    7. If the entry is not found, return 'Task Name: {tower_name}, % Complete: Not found'
                    """
                
                watsonx_response = watsonx_api_call(token, prompt)
                st.write(f"Watson response: {watsonx_response}")
                
                if watsonx_response:
                    if "Failed to process" in watsonx_response:
                        st.error(f"❌ Failed to process {file_name} for {tower_name} (sheet: {sheet_name}) with WatsonX: The WatsonX API service may be temporarily unavailable.")
                        results.append({"Sheet": sheet_name, "Result": "Failed to process: WatsonX API unavailable"})
                        failed_sheets += 1
                    else:
                        results.append({"Sheet": sheet_name, "Result": re.search(r'% Complete: (\d+\.\d+)', watsonx_response).group(1) if re.search(r'% Complete: (\d+\.\d+)', watsonx_response) else None})
                        st.success(f"Successfully processed: {file_name} for {tower_name} (sheet: {sheet_name})")
                else:
                    st.error(f"❌ Failed to process {file_name} for {tower_name} (sheet: {sheet_name}) with WatsonX: The WatsonX API service may be temporarily unavailable.")
                    results.append({"Sheet": sheet_name, "Result": "Failed to process: WatsonX API unavailable"})
                    failed_sheets += 1
                    
        except Exception as e:
            st.error(f"Error processing {file_name}: {str(e)}")
            continue
    
    return results, failed_sheets

def is_numeric(value):
    """Check if a value can be converted to a float after removing '%'."""
    try:
        float(str(value).replace('%', '').strip())
        return True
    except ValueError:
        return False

def final_output():
    report = []
    count = 0
    st.set_page_config(page_title="WatsonX Excel Processor", layout="wide")
    st.title("Excel Processor with WatsonX")
    
    if 'results' not in st.session_state:
        st.session_state.results = None
    if 'result_df' not in st.session_state:
        st.session_state.result_df = None
    if 'processed' not in st.session_state:
        st.session_state.processed = False

    token = get_access_token(API_KEY)
    
    if not token:
        st.error("Cannot proceed without API access")
        return
    
    st.header("Excel File Processor with WatsonX")
    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=['xlsx', 'xls'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        st.session_state.result_df = None
        st.write("Uploaded files:")
        for file in uploaded_files:
            st.write(f"- {file.name}")
            if "Structure Work Tracker EWS LIG P4 all towers" in file.name:
                print("Yes")
        
        if st.button("Process Files with WatsonX"):
            st.session_state.results = None
            st.session_state.result_df = None
            st.session_state.processed = False

            with st.spinner("Processing files with WatsonX..."):
                results, failed_sheets = process_excel_files(token, uploaded_files)
                
                if results:
                    st.session_state.results = results
                    
                    result_df = pd.DataFrame(columns=["Sheet", "Result"])
                    result_rows = []
                    for result in results:
                        result_rows.append({
                            "Sheet": result["Sheet"],
                            "Result": result["Result"]
                        })
                    
                    if result_rows:
                        result_df = pd.concat([result_df, pd.DataFrame(result_rows)], ignore_index=True)
                        st.session_state.result_df = result_df
                        st.session_state.processed = True

    if st.session_state.processed and st.session_state.result_df is not None:
        st.write("Results From Watson:")
        st.dataframe(st.session_state.result_df)
        st.write("Processing Ajith files")

        for file in uploaded_files:
            if "Structure Work Tracker EWS LIG P4 all towers" in file.name:
                st.write("Processing first file")
                averagedf = CountingProcess(uploaded_files[count])
            elif "Structure Work Tracker Tower G & Tower H" in file.name:
                st.write("Processing second file")
                averagedf2 = CountingProcess2(uploaded_files[count])
            elif "Structure Work Tracker Tower 6 & Tower 7" in file.name:
                st.write("Processing third file")
                averagedf3 = CountingProcess3(uploaded_files[count])
            count += 1

        st.session_state.result_df.insert(0, 'Project', 'Veridia')
        st.dataframe(st.session_state.result_df)
        result_df = st.session_state.result_df.copy()
        result_df.insert(len(result_df.columns), 'Finishing', '0%')
        
        custom_columns = ['Project', 'Tower', 'Structure', 'Finishing']
        combined_df = pd.DataFrame()

        if not result_df.empty:
            result_df_renamed = result_df.copy()
            if len(result_df_renamed.columns) <= len(custom_columns):
                result_df_renamed = pd.concat([result_df_renamed, pd.DataFrame(columns=custom_columns[len(result_df_renamed.columns):])], axis=1).fillna('')
                result_df_renamed.columns = custom_columns
            else:
                result_df_renamed.columns = custom_columns
            combined_df = pd.concat([combined_df, result_df_renamed], ignore_index=True)

        if averagedf is not None and not averagedf.empty:
            averagedf_renamed = averagedf.copy()
            if len(averagedf_renamed.columns) <= len(custom_columns):
                averagedf_renamed = pd.concat([averagedf_renamed, pd.DataFrame(columns=custom_columns[len(averagedf_renamed.columns):])], axis=1).fillna('')
                averagedf_renamed.columns = custom_columns
            else:
                averagedf_renamed.columns = custom_columns
            combined_df = pd.concat([combined_df, averagedf_renamed], ignore_index=True)

        if averagedf2 is not None and not averagedf2.empty:
            averagedf2_renamed = averagedf2.copy()
            if len(averagedf2_renamed.columns) <= len(custom_columns):
                averagedf2_renamed = pd.concat([averagedf2_renamed, pd.DataFrame(columns=custom_columns[len(averagedf2_renamed.columns):])], axis=1).fillna('')
                averagedf2_renamed.columns = custom_columns
            else:
                averagedf2_renamed.columns = custom_columns
            combined_df = pd.concat([combined_df, averagedf2_renamed], ignore_index=True)

        if averagedf3 is not None and not averagedf3.empty:
            averagedf3_renamed = averagedf3.copy()
            if len(averagedf3_renamed.columns) <= len(custom_columns):
                averagedf3_renamed = pd.concat([averagedf3_renamed, pd.DataFrame(columns=custom_columns[len(averagedf3_renamed.columns):])], axis=1).fillna('')
                averagedf3_renamed.columns = custom_columns
            else:
                averagedf3_renamed.columns = custom_columns
            combined_df = pd.concat([combined_df, averagedf3_renamed], ignore_index=True)

        if not combined_df.empty:
            st.dataframe(combined_df)
            combined_df.columns = [col.strip() for col in combined_df.columns]
            combined_df["Tower"] = combined_df["Tower"].astype(str)
            mask = combined_df["Tower"].str.lower().str.contains("finishing", na=False)
            combined_df.loc[mask, "Finishing"] = combined_df.loc[mask, "Structure"]
            combined_df.loc[mask, "Structure"] = "0%"
            st.subheader("Modified Data")
            st.dataframe(combined_df)
            
            # Save to disk
            combined_df.to_excel("updated_progress_data.xlsx", index=False, sheet_name="Modified Data")
            st.success("File saved as updated_progress_data.xlsx")
            
            # Provide download option
            excel_file = io.BytesIO()
            with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
                combined_df.to_excel(writer, index=False, sheet_name="Modified Data")
            excel_file.seek(0)
            st.download_button(
                label="Download Cleaned Excel",
                data=excel_file,
                file_name="updated_progress_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        return combined_df

final_output()