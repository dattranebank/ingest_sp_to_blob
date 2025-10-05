import azure.functions as func
import logging
import os
import requests
from msal import ConfidentialClientApplication
from azure.storage.blob import BlobServiceClient
from azure.core.exceptions import ResourceExistsError

app = func.FunctionApp()
@app.function_name(name="copy_sp_to_blob")
@app.route(
    route="copy_sp_to_blob",
    methods=["POST"],
    auth_level=func.AuthLevel.FUNCTION  # giữ nguyên FUNCTION
)
def copy_sp_to_blob(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("Function started")

    try:
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
            logging.error("Authentication failed: no token returned")
            return func.HttpResponse("Authentication failed", status_code=401)
        access_token = token["access_token"]

        headers = {"Authorization": f"Bearer {access_token}"}
        site_id = os.environ["SHAREPOINT_SITE_ID"]
        folder_path = os.environ.get("SHAREPOINT_FOLDER_PATH", "")
        graph_url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_path}:/children"
            if folder_path
            else f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
        )

        response = requests.get(graph_url, headers=headers)
        if response.status_code != 200:
            logging.error(f"Failed to fetch files: {response.text}")
            return func.HttpResponse(f"Graph API error: {response.text}", status_code=response.status_code)

        files = response.json().get("value", [])
        excel_files = [f for f in files if f.get("name", "").endswith(('.xlsx', '.xls')) and "file" in f]

        connection_string = os.environ["AzureWebJobsStorage"]
        container_name = os.environ["BLOB_CONTAINER_NAME"]
        subfolder = os.environ.get("BLOB_SUBFOLDER", "")

        blob_service_client = BlobServiceClient.from_connection_string(connection_string)
        container_client = blob_service_client.get_container_client(container_name)

        uploaded_count = 0
        for file_item in excel_files:
            file_name = file_item["name"]
            download_url = file_item["@microsoft.graph.downloadUrl"]
            file_content = requests.get(download_url, headers=headers).content
            blob_path = f"{subfolder}/{file_name}" if subfolder else file_name

            blob_client = container_client.get_blob_client(blob_path)
            blob_client.upload_blob(file_content, overwrite=True)
            uploaded_count += 1
            logging.info(f"Uploaded {blob_path}")

        message = f"✅ Uploaded {uploaded_count}/{len(excel_files)} files from SharePoint folder: {folder_path}"
        logging.info(message)
        return func.HttpResponse(message, status_code=200)

    except Exception as e:
        logging.error(f"Unhandled exception: {str(e)}")
        return func.HttpResponse(f"Error: {str(e)}", status_code=500)
