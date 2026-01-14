import logging
import os

import requests
from config import BASE_URL
from token_manager import get_access_token


def upload_pdfs_with_structure(local_folder_path, Ankensfolder_name):
    try:
        site_name = "Shuuko"
        site_id = get_site_id(site_name)
        drive_id = get_drive_id(site_id)

        # Create案件フォルダ
        anken_folder_id = create_folder_if_not_exists(drive_id, None, Ankensfolder_name)

        # 🚀 Walk through local folder recursively
        for root, _, files in os.walk(local_folder_path):
            for file_name in files:
                if file_name.endswith(".pdf"):
                    file_path = os.path.join(root, file_name)

                    # Determine relative path from the案件 folder
                    relative_path = os.path.relpath(file_path, start=local_folder_path)
                    relative_dirs = os.path.dirname(relative_path).split(os.sep)

                    # Start from 案件 folder
                    current_folder_id = anken_folder_id

                    # Create subfolders step-by-step if needed
                    for subfolder in relative_dirs:
                        if subfolder:  # avoid empty
                            current_folder_id = create_folder_if_not_exists(drive_id, current_folder_id, subfolder)

                    # Upload the PDF into the correct SharePoint folder
                    upload_file(drive_id, current_folder_id, file_path)

        logging.info(f"Upload completed with structure for: {Ankensfolder_name}")
    except Exception as e:
        logging.exception(f"Upload with structure failed for {Ankensfolder_name}: {e}")


def get_site_id(site_name):
    """
    Retrieves the site ID for a given SharePoint site name.
    """
    url = f"{BASE_URL}/sites/nskkogyo.sharepoint.com:/sites/{site_name}"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        return resp.json()["id"]
    else:
        logging.error(f"Failed to get site ID for {site_name}: {resp.text}")
        raise Exception("Failed to retrieve site ID")


def get_drive_id(site_id):
    """
    Retrieves the drive ID for a given site ID.
    """
    url = f"{BASE_URL}/sites/{site_id}/drive"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        return resp.json()["id"]
    else:
        logging.error(f"Failed to get drive ID: {resp.text}")
        raise Exception("Failed to retrieve drive ID")


def get_folder_id_recursive(drive_id, parent_id, folder_list):
    """Recursively navigate folders to get final folder id"""
    current_id = parent_id
    for folder_name in folder_list:
        current_id = create_folder_if_not_exists(drive_id, current_id, folder_name)
    return current_id


def create_folder_if_not_exists(drive_id, parent_id, folder_name):
    """
    Checks if a folder exists under the specified parent and creates it if not present.
    """
    folder_url = (
        f"{BASE_URL}/drives/{drive_id}/items/{parent_id}/children"
        if parent_id
        else f"{BASE_URL}/drives/{drive_id}/root/children"
    )
    headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}
    resp = requests.get(folder_url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Failed to list folders: {resp.text}")

    for item in resp.json().get("value", []):
        if item["name"] == folder_name:
            return item["id"]

    payload = {"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"}
    resp = requests.post(folder_url, headers=headers, json=payload)
    if resp.status_code in [200, 201]:
        return resp.json()["id"]
    else:
        raise Exception(f"Failed to create folder '{folder_name}': {resp.text}")


def upload_file(drive_id, folder_id, file_path):
    """
    Uploads a file to a specified folder in SharePoint.
    """
    file_name = os.path.basename(file_path)
    url = f"{BASE_URL}/drives/{drive_id}/items/{folder_id}:/{file_name}:/content"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    with open(file_path, "rb") as file_data:
        resp = requests.put(url, headers=headers, data=file_data)
    if resp.status_code not in [200, 201]:
        raise Exception(f"Upload failed for '{file_name}': {resp.text}")
