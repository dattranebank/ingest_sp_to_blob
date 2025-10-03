import azure.functions as func
import logging
import os
import requests
from msal import ConfidentialClientApplication
from azure.storage.blob import BlobServiceClient
from azure.core.exceptions import ResourceExistsError

app = func.FunctionApp()


# Timer Trigger cập nhật
@app.schedule(schedule="0 0 * * * *", arg_name="myTimer", run_on_startup=True)
def copy_sp_to_blob(myTimer: func.TimerRequest) -> None:

    # Auth setup
    client_id = os.environ["CLIENT_ID"]
    tenant_id = os.environ["TENANT_ID"]
    client_secret = os.environ["CLIENT_SECRET"]
    app_auth = ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}"
    )
    token = app_auth.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in token:
        logging.error("Authentication failed")
        return
    access_token = token["access_token"]
    headers = {"Authorization": f"Bearer {access_token}"}
    
    # Fetch files từ thư mục cụ thể trong SharePoint
    site_id = os.environ["SHAREPOINT_SITE_ID"]
    folder_path = os.environ.get("SHAREPOINT_FOLDER_PATH", "")  # Ví dụ: "01. Landing_Raw/DMS/DMS_BC_1.7"
    if folder_path:
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_path}:/children"
    else:
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
    
    response = requests.get(graph_url, headers=headers)
    if response.status_code != 200:
        logging.error(f"Failed to fetch files: {response.text}")
        return
    files = response.json().get("value", [])
    
    # Lọc chỉ Excel files
    excel_files = [f for f in files if f.get("name", "").endswith(('.xlsx', '.xls')) and "file" in f]
    
    # Upload từng file Excel lên Blob
    connection_string = os.environ["AzureWebJobsStorage"]
    container_name = os.environ["BLOB_CONTAINER_NAME"]  # Ví dụ: "sharepoint-files"
    subfolder = os.environ.get("BLOB_SUBFOLDER", "")  # Ví dụ: dms/dms17
    blob_service_client = BlobServiceClient.from_connection_string(connection_string)
    container_client = blob_service_client.get_container_client(container_name)
    
    for file_item in excel_files:
        file_name = file_item["name"]
        download_url = file_item["@microsoft.graph.downloadUrl"]
        file_content = requests.get(download_url, headers=headers).content

        # Nối subfolder + filename
        if subfolder:
            blob_path = f"{subfolder}/{file_name}"
        else:
            blob_path = file_name

        blob_client = container_client.get_blob_client(blob_path)

        try:
            blob_client.upload_blob(file_content, overwrite=True)
            logging.info(f"Uploaded Excel file to {blob_path}")
        except ResourceExistsError:
            logging.info(f"File {blob_path} already exists, skipping")
        except Exception as e:
            logging.error(f"Upload failed for {blob_path}: {str(e)}")
    
    logging.info(f"Processed {len(excel_files)} Excel files from SharePoint folder: {folder_path}")