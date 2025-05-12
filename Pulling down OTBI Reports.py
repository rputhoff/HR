from dotenv import load_dotenv
import os
import requests
from requests.auth import HTTPBasicAuth

# Load environment variables from .env file
load_dotenv()

# Define the API endpoint and credentials
api_endpoint = "https://hcor.fa.us2.oraclecloud.com/analytics/saw.dll?catalog"
username = os.getenv("OTBI_USERNAME")
password = os.getenv("OTBI_PASSWORD")
params = {
    "path": "/shared/Custom/your_report.xdo"
}

# Make the API call to download the report
response = requests.get(api_endpoint, params=params, auth=HTTPBasicAuth(username, password))
with open("report.xdo", "wb") as file:
    file.write(response.content)
