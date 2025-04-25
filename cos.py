import streamlit as st
import requests
import json
import ibm_boto3
from ibm_botocore.client import Config
import io
import urllib.parse

# IBM COS Credentials
COS_API_KEY = "2PjLRmZ3Ay-WQpuE33qGaQzDohwVJIzocHlABKayUsNV"
COS_SERVICE_INSTANCE_ID = "crn:v1:bluemix:public:cloud-object-storage:global:a/fddc2a92db904306b413ed706665c2ff:e99c3906-0103-4257-bcba-e455e7ced9b7::"
COS_ENDPOINT = "https://s3.us-south.cloud-object-storage.appdomain.cloud"
COS_BUCKET = "ozonetell"

# Initialize COS client
cos_client = ibm_boto3.client(
    's3',
    ibm_api_key_id=COS_API_KEY,
    ibm_service_instance_id=COS_SERVICE_INSTANCE_ID,
    config=Config(signature_version='oauth'),
    endpoint_url=COS_ENDPOINT
)

# Existing API credentials for Watsonx
api_key = "3Vj-0udUsnRjiJwBKNGAEcHNpiS-xi6VX-5tU2VZWHij"
url = "https://us-south.ml.cloud.ibm.com/ml/v1/text/generation?version=2023-05-29"
project_id = "4152f31e-6a49-40aa-9b62-0ecf629aae42"
model_id = "meta-llama/llama-3-2-90b-vision-instruct"
auth_url = "https://iam.cloud.ibm.com/identity/token"

# Session state initialization
st.session_state.transcript = "Select a JSON file to get transcript"
st.session_state.insights = "Select a JSON file to get insights"
st.session_state.callquality = "Select a JSON file to get call quality"
st.session_state.separate = "Select a JSON file to get call transcript"
st.session_state.raw_transcript = "Raw transcript will appear here"
st.session_state.failed_files = []  # To track files with invalid AudioFile
st.session_state.json_output = None  # To store JSON output for download

# Page config
st.set_page_config(layout="wide", page_title="Wave Infra Call Insights")
st.markdown("<h1 style='text-align:center'>Wave Infra Call Insights</h1>", unsafe_allow_html=True)
st.markdown("<link href='https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css' rel='stylesheet' integrity='sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH' crossorigin='anonymous'>", unsafe_allow_html=True)
st.markdown("""
    <style>
    .stApp {
        background-color: #0e1117;
    }
    .css-1d8v1u7, .css-1a32f9v {
        color: #ffffff !important; /* Bright white text to match other boxes */
        font-size: 16px !important; /* Consistent text size */
    }
    </style>
""", unsafe_allow_html=True)

top1, top2 = st.columns(2)
bottom1, bottom2 = st.columns(2)

def access_token():
    headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}
    data = {"grant_type": "urn:ibm:params:oauth:grant-type:apikey", "apikey": api_key}
    response = requests.post(auth_url, headers=headers, data=data)
    return response.json()['access_token']

def Getcallquality(trans):
    body = {
        "input": f"""
        Below is a transcription of a conversation between two people. Need call quality analysis for the given transcription.
        The output should look like this:
        Add-on Request by Customer: customer request...........
        Action Taken for the request: action taken by ........
        call rating: rating out of 10
        Reason: reason for the rating
        Transcription:
        {trans}
        """,
        "parameters": {
            "decoding_method": "greedy",
            "max_new_tokens": 300,
            "min_new_tokens": 30,
            "stop_sequences": [";"],
            "repetition_penalty": 1.05,
            "temperature": 0.5
        },
        "model_id": model_id,
        "project_id": project_id
    }
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {access_token()}"
    }
    response = requests.post(url, headers=headers, json=body)
    st.session_state.callquality = response.json()['results'][0]['generated_text']

def Getinsights(trans):
    body = {
        "input": f"""
        Below is a transcription of a conversation between two people. Need insight summary from the given transcription.
        The output should look like this:
        Insights Summary: insight summary of the conversation.....
        Sentiment: sentiment of the conversation.....
        I need Insights Summary and Sentiment only.
        Transcription:
        {trans}
        """,
        "parameters": {
            "decoding_method": "greedy",
            "max_new_tokens": 300,
            "min_new_tokens": 30,
            "stop_sequences": [";"],
            "repetition_penalty": 1.05,
            "temperature": 0.5
        },
        "model_id": model_id,
        "project_id": project_id
    }
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {access_token()}"
    }
    response = requests.post(url, headers=headers, json=body)
    st.session_state.insights = response.json()['results'][0]['generated_text']

def Separatespeakers(trans):
    prompt = f"""
    You are a conversation formatter.
    
    I'm going to give you a raw call transcript. Convert it into a clean conversation format where each turn is clearly labeled as either "Speaker1:" or "Speaker2:" (no bold, no asterisks).
    
    Each speaker should be on a new line. Make sure to preserve the natural flow of conversation, with questions followed by answers and logical turn-taking.
    
    Here is the transcript:
    {trans}
    
    Format your response exactly like this example:
    Speaker1: [First speaker's text]
    Speaker2: [Second speaker's text]
    Speaker1: [First speaker's next turn]
    Speaker2: [Second speaker's response]

    Important: ONLY output the formatted conversation. Do not include any explanations, headers, or notes.
    """
    body = {
        "input": prompt,
        "parameters": {
            "decoding_method": "greedy",
            "max_new_tokens": 1000,
            "min_new_tokens": 30,
            "repetition_penalty": 1.05,
            "temperature": 0.1
        },
        "model_id": model_id,
        "project_id": project_id
    }
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": f"Bearer {access_token()}"
    }
    response = requests.post(url, headers=headers, json=body)
    response_text = response.json()['results'][0]['generated_text']
    lines = response_text.split('\n')
    clean_lines = []
    for line in lines:
        line = line.strip()
        if not line or "transcript" in line.lower() or "conversation" in line.lower():
            continue
        if line.startswith('Speaker1:') or line.startswith('Speaker2:'):
            clean_lines.append(line)
        elif line.startswith('Speaker 1:'):
            clean_lines.append(line.replace('Speaker 1:', 'Speaker1:'))
        elif line.startswith('Speaker 2:'):
            clean_lines.append(line.replace('Speaker 2:', 'Speaker2:'))
        elif line.startswith('**Speaker1**:'):
            clean_lines.append(line.replace('**Speaker1**:', 'Speaker1:'))
        elif line.startswith('**Speaker2**:'):
            clean_lines.append(line.replace('**Speaker2**:', 'Speaker2:'))
        elif line.startswith('**Speaker 1**:'):
            clean_lines.append(line.replace('**Speaker 1**:', 'Speaker1:'))
        elif line.startswith('**Speaker 2**:'):
            clean_lines.append(line.replace('**Speaker 2**:', 'Speaker2:'))
    if clean_lines:
        st.session_state.separate = '\n'.join(clean_lines)
    else:
        sentences = trans.split('.')
        alternate_lines = []
        is_speaker1 = True
        for sentence in sentences:
            sentence = sentence.strip()
            if sentence:
                speaker = "Speaker1" if is_speaker1 else "Speaker2"
                alternate_lines.append(f"{speaker}: {sentence}.")
                is_speaker1 = not is_speaker1
        st.session_state.separate = '\n'.join(alternate_lines)

def create_json_output(file_key):
    output_data = {
        "transcript": st.session_state.transcript,
        "insights": st.session_state.insights,
        "call_quality": st.session_state.callquality,
        "separated_transcript": st.session_state.separate
    }
    json_str = json.dumps(output_data, indent=2)
    st.session_state.json_output = json_str
    return json_str, f"output_{file_key.replace('.json', '')}.json"

def get_cos_files():
    try:
        response = cos_client.list_objects_v2(Bucket=COS_BUCKET)
        files = [obj['Key'] for obj in response.get('Contents', []) if obj['Key'].endswith('.json')]
        if not files:
            st.warning("No .json files found in the bucket 'ozonetell'. Please ensure JSON files are uploaded.")
        return files
    except Exception as e:
        st.error(f"Error fetching COS files: {e}")
        return []

def process_audio_from_cos(file_key):
    try:
        file_obj = cos_client.get_object(Bucket=COS_BUCKET, Key=file_key)
        json_data = json.loads(file_obj['Body'].read().decode('utf-8'))
        if 'AudioFile' not in json_data or not json_data['AudioFile']:
            st.error(f"No valid 'AudioFile' URL found in {file_key}. Adding to failed files list for review.")
            st.session_state.failed_files.append(file_key)
            return
        audio_url = json_data['AudioFile']
        result = urllib.parse.urlparse(audio_url)
        if not result.scheme:
            st.error(f"Invalid URL '{audio_url}' in {file_key}: No scheme supplied. Perhaps you meant https://{audio_url}? Adding to failed files list for review.")
            st.session_state.failed_files.append(file_key)
            return
        audio_response = requests.get(audio_url, stream=True)
        if audio_response.status_code == 200:
            audio_data = audio_response.content
        else:
            st.error(f"Failed to download audio from {audio_url} in {file_key}. Status code: {audio_response.status_code}")
            st.session_state.failed_files.append(file_key)
            return
        url = "https://dev.assisto.tech/workflow_apis/process_file"
        payload = {}
        files = [('file', (file_key, io.BytesIO(audio_data), 'audio/mpeg'))]
        headers = {}
        response = requests.post(url, headers=headers, data=payload, files=files)
        if response.status_code == 200:
            transcript = response.json()['result'][0]['message']
            st.session_state.raw_transcript = transcript
            st.session_state.transcript = transcript
            Getinsights(st.session_state.transcript)
            Getcallquality(st.session_state.transcript)
            Separatespeakers(st.session_state.transcript)
            st.session_state.json_output = None  # Reset before creating new JSON
            create_json_output(file_key)  # Create JSON for download
        else:
            st.error(f"Error from external API for {file_key}: {response.status_code} - {response.text}")
            st.session_state.transcript = response.text
    except Exception as e:
        st.error(f"An error occurred while processing the file {file_key}: {e}")
        st.session_state.transcript = str(e)

# Layout
with top1:
    st.subheader("Select Audio from IBM COS")
    json_files = get_cos_files()
    selected_file = st.selectbox("Choose an audio file", json_files, placeholder="No options to select.")
    if st.button("Process Audio", type="primary", use_container_width=True) and selected_file:
        process_audio_from_cos(selected_file)
    if st.session_state.failed_files:
        st.warning(f"The following files failed due to invalid 'AudioFile' URLs: {st.session_state.failed_files}")
        st.write("Please update these JSON files to include a valid 'AudioFile' key with a URL (e.g., https://example.com/audio.mp3).")
    # Download button for JSON output
    if st.session_state.json_output:
        json_str, filename = create_json_output(selected_file)
        st.download_button(
            label="Download JSON Output",
            data=json_str,
            file_name=filename,
            mime="application/json",
            type="primary",
            use_container_width=True
        )

with top2:
    st.subheader("Call Transcript using Watsonx Speech-To-text")
    with st.container(height=300):
        if st.session_state.separate == "Select a JSON file to get call transcript":
            st.write("Transcript will appear here")
        else:
            st.write(st.session_state.separate)

with bottom1:
    st.subheader("Call Insights by watsonx.ai")
    with st.container(height=300, key=2):
        st.write(st.session_state.insights)

with bottom2:
    st.subheader("Call quality analysis by watsonx.ai")
    with st.container(height=300, key=3):
        st.write(st.session_state.callquality)