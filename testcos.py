import streamlit as st
import requests
import json
import ibm_boto3
from ibm_botocore.client import Config
import io
import urllib.parse
import pandas as pd


COS_API_KEY = "VVATXn7oVk51VbC-cCE1DJJEA_UCSAMu1r6ushqG2GQ9"
COS_SERVICE_INSTANCE_ID = "crn:v1:bluemix:public:cloud-object-storage:global:a/fddc2a92db904306b413ed706665c2ff:e99c3906-0103-4257-bcba-e455e7ced9b7:bucket:projectreport"
COS_ENDPOINT = "https://s3.us-south.cloud-object-storage.appdomain.cloud"
COS_BUCKET = "projectreport"


st.session_state.cos_client = ibm_boto3.client(
    's3',
    ibm_api_key_id=COS_API_KEY,
    ibm_service_instance_id=COS_SERVICE_INSTANCE_ID,
    config=Config(signature_version='oauth'),
    endpoint_url=COS_ENDPOINT
)

def get_cos_files():
    try:
        response = st.session_state.cos_client.list_objects_v2(Bucket=COS_BUCKET)
        files = [obj['Key'] for obj in response.get('Contents', []) if obj['Key'].endswith('.xlsx')]
        if not files:
            print("No .json files found in the bucket 'ozonetell'. Please ensure JSON files are uploaded.")
        return files
    except Exception as e:
        print(f"Error fetching COS files: {e}")
        return []
    
# for i in get_cos_files():
#     st.write(i)

response = st.session_state.cos_client  .get_object(Bucket=COS_BUCKET, Key="Tower 4 Tracker March 2025 Lookahead.xlsx")
datas = pd.read_excel(io.BytesIO(response['Body'].read()), sheet_name="TOWER 4 FINISHING.", skiprows=15)
st.write(datas.head())
