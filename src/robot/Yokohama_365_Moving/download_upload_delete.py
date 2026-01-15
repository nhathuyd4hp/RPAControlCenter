import logging
import os
import re
from pathlib import Path

import requests
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

# 📚 Microsoft Graph API Credentials
current_dir = Path(__file__).resolve().parent
env_path = current_dir.parent.parent.parent / ".env"
load_dotenv(dotenv_path=env_path)
CLIENT_ID = os.getenv("API_SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("API_SHAREPOINT_CLIENT_SECRET")
TENANT_ID = os.getenv("API_SHAREPOINT_TENANT_ID")

# 🌎 Microsoft Graph API URLs
BASE_URL = "https://graph.microsoft.com/v1.0"
SEARCH_URL = f"{BASE_URL}/search/query"

# 📂 Drives to Search (Multiple Sites)
DRIVE_IDS = [
    # Site: 2021
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMgqFmxUFSTDS5xlMmIATcY_",  # か行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMhUTAdjY1jnSK7YfX-IfcQs",  # さ行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMgxXBpRecuxQp40qCQ96qCw",  # た行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMhn9jPDMigKTKcQq4biVQTp",  # ドキュメント・2
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMgLltFBcoLeSJyVrkeRVc-u",  # な行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMiGv-OjvRfuSIZjru4KPfrt",  # ドキュメント
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMje-cYmil_oQ6oMx_OlS8au",  # ま行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMi2WzsMrIhERLjKzloPS0YK",  # や・ら・わ行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMihKuZWYmqkTqqy3R9t3aff",  # は行
    "b!XKuZyeFTlkSp4cWfSxd10AfxU7PGA3xBi27uhOfFFMhYgcNNyw5IQKep4L6_VFIk",  # あ行
    # Site: Kantou
    "b!CGMwpFZqO0aR13-uULpoA739OTZDETFKpDsa-PGqFCBe0TiC03OyTLyZUjcaE8e9",  # ドキュメント
    "b!CGMwpFZqO0aR13-uULpoA739OTZDETFKpDsa-PGqFCBMszIEG92nQ76ejmAOfnzy",  # 新ドキュメント(関東)
    "b!CGMwpFZqO0aR13-uULpoA739OTZDETFKpDsa-PGqFCCdvvkNEUAUTb8Gxjm9oin3",  # 検索構成リスト
    # Site: 2019
    "b!sCgCnWR2UkGKdRInfBWzdlcnAGNMtfdEjamzCOTJHvCO1eFDmXWzRpY7g3QpUVA-",  # Documents
    "b!sCgCnWR2UkGKdRInfBWzdlcnAGNMtfdEjamzCOTJHvANeCwNSd0wTZ7-9-ersYK5",  # タマホーム
    # Site: Shuuko
    "b!vArDktlKE0uGKwPHe6i71cHlFfas-b9DhL0W0_9h3SLpFob0RyrQRrPmZvYxcvot",  # Documents
    "b!vArDktlKE0uGKwPHe6i71cHlFfas-b9DhL0W0_9h3SLNUAN_SP5FRZzzVzIygXm8",  # Search Config List
]

# 📁 Base Folder for Downloads
DOWNLOAD_DIR = os.path.join(os.getcwd(), "Ankens")

# 📄 Excel Path for Ankens
EXCEL_PATH = os.path.join(os.getcwd(), "Data.xlsx")

# ✏️ Column Names
ANKEN_COLUMN = "案件番号"
STATUS_COLUMN = "Download Status"

# ✨ Folder Paths
BASE_DIR = os.getcwd()
CSV_INPUT_FOLDER = os.path.join(BASE_DIR, "CSV")  # Folder where CSVs are stored
EXCEL_OUTPUT_FOLDER = os.path.join(BASE_DIR, "Excels")  # Folder where Excel files are saved
DOWNLOAD_DIR = os.path.join(BASE_DIR, "Ankens")  # Folder where PDFs are downloaded

# ✨ Other configs
REGION = "JPN"  # For Microsoft Graph Search API (Japan site)

# Download Config
BATCH_SIZE = 5  # Only 5 ankens in parallel to avoid 429
MAX_RETRIES = 3  # Retry 3 times on 429
RETRY_SLEEP = 5  # 5 seconds sleep on 429
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


class SharePointHandler:
    def __init__(self):
        self.get_access_token()
        self.headers = {"Authorization": f"Bearer {self.get_access_token()}"}

    def get_access_token(self):
        app = ConfidentialClientApplication(
            CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}", client_credential=CLIENT_SECRET
        )
        result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
        if "access_token" in result:
            return result["access_token"]
        raise Exception("Unable to acquire token")

    # def download_folder(self,drive_id, folder_path, local_path):
    #     access_token =  self.access_token
    #     headers = {"Authorization": f"Bearer {access_token}"}

    #     url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}:/children"
    #     response = requests.get(url, headers=headers)
    #     response.raise_for_status()
    #     items = response.json().get("value", [])

    #     for item in items:
    #         item_name = item["name"]
    #         item_drive_path = f"{folder_path}/{item_name}"
    #         local_item_path = os.path.join(local_path, item_name)

    #         if "folder" in item:
    #             os.makedirs(local_item_path, exist_ok=True)
    #             self.download_folder(drive_id, item_drive_path, local_item_path)
    #         elif "file" in item:
    #             download_url = item["@microsoft.graph.downloadUrl"]
    #             file_content = requests.get(download_url).content
    #             with open(local_item_path, "wb") as f:
    #                 f.write(file_content)
    #             print(f"✅ Downloaded: {local_item_path}")

    def get_drive_id_by_name(self, drive_name):
        access_token = self.get_access_token()
        headers = {"Authorization": f"Bearer {access_token}"}
        site_url = "https://graph.microsoft.com/v1.0/sites/nskkogyo.sharepoint.com:/sites/nouhinsumi-kantou:/drives"
        resp = requests.get(site_url, headers=headers)
        if resp.ok:
            for drive in resp.json().get("value", []):
                if drive["name"] == drive_name:
                    return drive["id"]
        else:
            print(f"❌ Error: {resp.status_code} — {resp.text}")
            return None

    # def upload_folder(self,local_folder_path, sharepoint_folder_path):
    #         access_token = self.access_token
    #         headers = {"Authorization": f"Bearer {access_token}"}

    #         path_parts = sharepoint_folder_path.strip("/").split("/")
    #         drive_name = path_parts[0]
    #         base_sharepoint_folder = "/".join(path_parts[1:])  # Path inside drive
    #         drive_id = self.get_drive_id_by_name(drive_name)

    #         if not drive_id:
    #             print(f"❌ Drive '{drive_name}' not found.")
    #             return

    #         folder_name = os.path.basename(local_folder_path.rstrip("/\\"))  # Folder name only
    #         sharepoint_base = f"{base_sharepoint_folder}/{folder_name}".strip("/")

    #         print(f"📁 Uploading folder: {local_folder_path} ➜ SharePoint: {sharepoint_base}")

    #         for root, _, files in os.walk(local_folder_path):
    #             for file_name in files:
    #                 local_file_path = os.path.join(root, file_name)
    #                 relative_path = os.path.relpath(local_file_path, local_folder_path).replace("\\", "/")
    #                 sharepoint_file_path = f"{sharepoint_base}/{relative_path}"

    #                 upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sharepoint_file_path}:/content"

    #                 with open(local_file_path, "rb") as file_data:
    #                     resp = requests.put(upload_url, headers=headers, data=file_data)
    #                     if resp.ok:
    #                         print(f"✅ Uploaded: {sharepoint_file_path}")
    #                     else:
    #                         print(f"❌ Failed: {sharepoint_file_path} — {resp.status_code} | {resp.text}")
    #         return True

    def get_item_metadata(self, drive_id, item_id):
        access_token = self.get_access_token()
        headers = {"Authorization": f"Bearer {access_token}"}
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()

    def get_drive_id_by_name(self, drive_name):  # noqa
        access_token = self.get_access_token()
        headers = {"Authorization": f"Bearer {access_token}"}
        site_url = "https://graph.microsoft.com/v1.0/sites/nskkogyo.sharepoint.com:/sites/nouhinsumi-kantou:/drives"
        resp = requests.get(site_url, headers=headers)
        if resp.ok:
            for drive in resp.json().get("value", []):
                if drive["name"] == drive_name:
                    return drive["id"]
        else:
            print(f"❌ Error: {resp.status_code} — {resp.text}")
            return None

    def upload_folder(self, local_folder_path, sharepoint_folder_path):
        access_token = self.get_access_token()
        headers = {"Authorization": f"Bearer {access_token}"}

        path_parts = sharepoint_folder_path.strip("/").split("/")
        drive_name = path_parts[0]
        base_sharepoint_folder = "/".join(path_parts[1:])  # Path inside drive
        drive_id = self.get_drive_id_by_name(drive_name)
        logging.info(f"driverid:{drive_id}")

        if not drive_id:
            logging.info(f"Drive '{drive_name}' not found.")
            return False

        folder_name = os.path.basename(local_folder_path.rstrip("/\\"))  # Folder name only
        sharepoint_base = f"{base_sharepoint_folder}/{folder_name}".strip("/")

        logging.info(f"Uploading folder: {local_folder_path} SharePoint: {sharepoint_base}")

        for root, _, files in os.walk(local_folder_path):
            for file_name in files:
                local_file_path = os.path.join(root, file_name)
                relative_path = os.path.relpath(local_file_path, local_folder_path).replace("\\", "/")
                sharepoint_file_path = f"{sharepoint_base}/{relative_path}"

                upload_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sharepoint_file_path}:/content"
                logging.info(f"upload_url:{upload_url}")

                with open(local_file_path, "rb") as file_data:
                    resp = requests.put(upload_url, headers=headers, data=file_data)
                    if resp.ok:
                        logging.info(f"Uploaded: {sharepoint_file_path}")
                    else:
                        logging.info(f"Failed: {sharepoint_file_path} — {resp.status_code} | {resp.text}")

        folder_metadata_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sharepoint_base}"
        folder_resp = requests.get(folder_metadata_url, headers=headers)

        # if folder_resp.ok:
        #     folder_info = folder_resp.json()
        #     web_url = folder_info.get("webUrl")
        #     logging.info(f"SharePoint Folder Link: {web_url}")
        # else:
        #     logging.warning(f"Failed to fetch folder link: {folder_resp.status_code} | {folder_resp.text}")
        #     web_url = "none"  # In case it's needed for later

        if folder_resp.ok:
            folder_info = folder_resp.json()
            folder_id = folder_info.get("id")
            web_url = folder_info.get("webUrl")
            logging.info(f"Long SharePoint Folder Link: {web_url}")

            # 📌 Generate short sharing link using folder ID
            sharing_api = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/createLink"
            payload = {"type": "view", "scope": "organization"}  # or "edit"  # or "organization"
            resp = requests.post(sharing_api, headers=headers, json=payload)

            if resp.ok:
                sharing_data = resp.json()
                web_url = sharing_data.get("link", {}).get("webUrl")
                logging.info(f"Short SharePoint Link: {web_url}")
            else:
                logging.warning(f"Failed to generate sharing link: {resp.status_code} | {resp.text}")
                web_url = "none"
        else:
            logging.warning(f"Failed to fetch folder metadata: {folder_resp.status_code} | {folder_resp.text}")
            web_url = "none"

        return web_url

    def search_anken_folder(self, anken_number):
        search_url = "https://graph.microsoft.com/v1.0/search/query"
        payload = {
            "requests": [{"entityTypes": ["driveItem"], "query": {"queryString": anken_number}, "region": "JPN"}]
        }
        headers = {"Authorization": f"Bearer {self.get_access_token()}", "Content-Type": "application/json"}

        resp = requests.post(search_url, headers=headers, json=payload)

        if resp.status_code != 200:
            raise Exception(f"Search failed: {resp.text}")

        results = resp.json()

        items = results.get("value", [])[0].get("hitsContainers", [])[0].get("hits", [])

        if not items:
            logging.info(f"No folder found for:{anken_number}")
            return None

        first = items[0]
        item = first.get("resource", {})
        folder_id = item.get("id")
        folder_name = item.get("name")
        drive_id = item.get("parentReference", {}).get("driveId", "")
        site_id = item.get("parentReference", {}).get("siteId", "")
        folder_url = item.get("webUrl", "")

        if not folder_url:
            logging.info(f"No folder URL found for:{anken_number}")
            return None
        if "nouhinsumi-kantou" in folder_url:
            logging.info(f"Folder URL is nouhinsumi-kantou: {folder_url}")
            logging.info(f"Skipping folder: {folder_name}")
            return None
        else:
            logging.info(f"Folder URL is not nouhinsumi-kantou: {folder_url}")
            logging.info(f"Continuing with folder: {folder_name}")

        # print(f"folder_id:{folder_id}", f"folder_name:{folder_name}", f"drive_id:{drive_id}", f"site_id:{site_id}")
        # input("a")

        logging.info(f"Found folder: {folder_name} (ID: {folder_id})")

        # Get full folder path hierarchy
        folder_chain = self.get_folder_path_chain(drive_id, folder_id)

        if not folder_chain:
            logging.info("Could not retrieve full folder path. Returning minimal data.")
            return {
                "name": folder_name,
                "id": folder_id,
                "driveId": drive_id,
                "siteId": site_id,
                "path": f"{folder_name}",
            }

        logging.info(" Folder Path Hierarchy:")
        for folder in folder_chain:
            logging.info(f"  {folder['name']}")

        # Include all folders in the path, including "root"
        full_path = "/".join(f["name"] for f in folder_chain)
        logging.info(f"\nFull Path (with root): {full_path}\n")

        return {"name": folder_name, "id": folder_id, "driveId": drive_id, "siteId": site_id, "path": full_path}

    def get_folder_path_chain(self, drive_id, folder_id):
        chain = []
        current_id = folder_id

        while current_id:
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{current_id}"
            headers = {"Authorization": f"Bearer {self.get_access_token()}"}

            resp = requests.get(url, headers=headers)
            if resp.status_code != 200:
                logging.info(f"Failed to retrieve folder with ID: {current_id}")
                break

            item = resp.json()
            chain.insert(0, {"name": item.get("name", ""), "id": item.get("id")})

            parent_ref = item.get("parentReference", {})
            parent_id = parent_ref.get("id")

            if not parent_id or parent_id == item.get("id"):
                break  # Reached root

            current_id = parent_id

        return chain

    def get_drive_name_by_id(self, drive_id):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
        headers = {"Authorization": f"Bearer {self.get_access_token()}"}
        resp = requests.get(url, headers=headers)
        if resp.ok:
            return resp.json().get("name")
        return None

    def list_children(self, drive_id, item_id):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
        headers = {"Authorization": f"Bearer {self.get_access_token()}"}

        resp = requests.get(url, headers=headers)
        if resp.status_code != 200:
            # raise Exception(f"Failed to list children: {resp.text}")
            logging.info(f"Failed to list children: {resp.text}")
            return None
        return resp.json().get("value", [])

    def download_recursive(self, drive_id, folder_id, local_path):
        os.makedirs(local_path, exist_ok=True)
        children = self.list_children(drive_id, folder_id)
        if not children:
            logging.info(f"No files found in folder {folder_id} : Download Nashi")
            return False
        for item in children:
            item_name = item["name"]

            if item.get("folder"):
                subfolder_path = os.path.join(local_path, item_name)
                self.download_recursive(drive_id, item["id"], subfolder_path)
            else:
                download_url = item.get("@microsoft.graph.downloadUrl")
                if download_url:
                    file_path = os.path.join(local_path, item_name)
                    file_resp = requests.get(download_url)
                    with open(file_path, "wb") as f:
                        f.write(file_resp.content)
                        logging.info(f"Downloaded: {file_path}")

    def delete_folder(self, drive_id, item_id, access_token):
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        headers = {"Authorization": f"Bearer {access_token}"}
        response = requests.delete(url, headers=headers)

        if response.status_code == 204:
            logging.info(f"Successfully deleted folder with ID: {item_id}")
            return True
        else:
            logging.error(f"Failed to delete folder {item_id}: {response.status_code} {response.text}")
            return False

    def download_entire_folder(self, anken_number, download_dir):
        try:
            # Step 1: Search for the folder
            folder_info = self.search_anken_folder(anken_number)
            if not folder_info:
                logging.info(f"Folder not found for: {anken_number}")
                return False, "案件見つけない"

            folder_id = folder_info["id"]
            drive_id = folder_info["driveId"]
            folder_name = folder_info["name"]
            full_path = folder_info["path"]
            drive_name = self.get_drive_name_by_id(drive_id)

            if not drive_name:
                logging.info(f"Could not determine drive name for ID: {drive_id}")
                return False, "drive_name_by_id無"

            parts = full_path.replace("root/", "").split("/")

            # Find index of the first part that starts with a 6-digit number (案件番号)
            ank_index = None
            for i, part in enumerate(parts):
                if re.match(r"^\d{6}", part):  # starts with 6-digit number
                    ank_index = i
                    break

            # Extract path before 案件番号
            if ank_index and ank_index > 0:
                desired_path = "/".join(parts[:ank_index])
                relative_sharepoint_path = desired_path
            else:
                desired_path = "/Unknown"
                return False, "Upload_path無"

            # print(desired_path)

            logging.info(f"Relative SharePoint Upload Path: {relative_sharepoint_path}")

            # Step 2: Create local save path
            save_path = os.path.join(download_dir, folder_name)
            os.makedirs(save_path, exist_ok=True)

            # Step 3: Check if folder already downloaded
            full_downloaded_path = None
            for fname in os.listdir(download_dir):
                fpath = os.path.join(download_dir, fname)
                if os.path.isdir(fpath) and anken_number in fname:
                    full_downloaded_path = os.path.abspath(fpath)
                    logging.info(f"Folder containing anken '{anken_number}' found at: {full_downloaded_path}")

            # Step 4: Download recursively
            # Step 5: Perform download
            donwload = self.download_recursive(drive_id, folder_id, save_path)
            if donwload is False:
                logging.info(f"Download failed for folder: {folder_name}")
                return False, "Download失敗"
            logging.info(f"Completed download for folder: {folder_name} → {save_path}")

            # Step 6: Upload to SharePoint
            if full_downloaded_path:
                logging.info(f"Uploading from existing folder: {full_downloaded_path}")
                upload = self.upload_folder(full_downloaded_path, relative_sharepoint_path)

            if upload is not False:
                logging.info(f"Upload successful for {folder_name} and sharepoint link is : {upload}")
                # return True, upload
                #  Step 6: Delete folder
                delete_folder = self.delete_folder(drive_id, folder_id, self.get_access_token())
                if delete_folder is True:
                    logging.info(f"Deleted folder: {folder_name}")
                    return True, upload
                else:
                    logging.info(f"Failed to delete folder: {folder_name}")
                    return False, "削除無"

            else:
                logging.info(f"Upload failed for {folder_name}")
                return False, "UP無"

        except Exception as e:
            logging.info(f"Error during download: {e}")
            return False, "Error"
