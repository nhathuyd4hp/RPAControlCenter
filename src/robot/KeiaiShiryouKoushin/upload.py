from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
from graph_uploader import get_site_id, get_drive_id, create_folder_if_not_exists, upload_file
from token_manager import get_access_token
from config import BASE_URL
import requests, os, logging, shutil
from tqdm import tqdm
from colorama import Fore
from graph_searcher import search_anken_folder
import time


def fileUpload_graph_api(upload_path, 案件番号):
    local_folder = Path(upload_path)
    old_files_dir = local_folder / "old files"
    old_files_dir.mkdir(exist_ok=True)

    try:
        site_id = get_site_id("Kantou")
        drive_id = get_drive_id(site_id)
        資料_folder_id = resolve_anken_folder_path(drive_id, 案件番号)

        # Step 1: Download all existing files
        download_all_files_from_sharepoint(drive_id, 資料_folder_id, old_files_dir)

        wait_for_download(old_files_dir, timeout=60)

        # Step 2: Delete them from SharePoint
        delete_all_files_in_sharepoint_folder(drive_id, 資料_folder_id)

        # Step 3: Upload old files into 資料/old files/
        old_files_folder_id = create_folder_if_not_exists(drive_id, 資料_folder_id, "old files")
        upload_folder_to_sharepoint(drive_id, old_files_folder_id, old_files_dir)

        # Step 4: Upload new files into 資料/
        upload_folder_to_sharepoint(drive_id, 資料_folder_id, local_folder)

        logging.info(Fore.CYAN + f"✅ All operations completed for: {案件番号}")

    except Exception as e:
        logging.exception(Fore.RED + f"🔥 Upload process failed: {e}")


def resolve_anken_folder_path(drive_id, 案件番号):
    result = search_anken_folder(案件番号)
    if not result:
        raise Exception(f"❌ Folder containing 案件番号 '{案件番号}' not found via search.")
    return create_folder_if_not_exists(drive_id, result["id"], "資料")


def download_all_files_from_sharepoint(drive_id, 資料_folder_id, local_folder):
    logging.info(Fore.CYAN + "📥 Downloading all existing files from SharePoint 資料 folder...")
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    items = get_all_children(drive_id, 資料_folder_id)

    for item in items:
        if "file" not in item:
            continue
        file_name = item["name"]
        item_id = item["id"]
        download_url = f"{BASE_URL}/drives/{drive_id}/items/{item_id}/content"
        local_path = local_folder / file_name
        try:
            resp = requests.get(download_url, headers=headers)
            resp.raise_for_status()
            with open(local_path, "wb") as f:
                f.write(resp.content)
            logging.info(Fore.GREEN + f"⬇️ Downloaded: {file_name}")
        except Exception as e:
            logging.error(Fore.RED + f"❌ Failed to download {file_name}: {e}")


def delete_all_files_in_sharepoint_folder(drive_id, 資料_folder_id):
    logging.info(Fore.CYAN + "🗑️ Deleting all files in SharePoint 資料 folder...")
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    items = get_all_children(drive_id, 資料_folder_id)

    for item in items:
        if "file" not in item:
            continue
        file_name = item["name"]
        item_id = item["id"]
        delete_url = f"{BASE_URL}/drives/{drive_id}/items/{item_id}"
        try:
            resp = requests.delete(delete_url, headers=headers)
            if resp.status_code in [204, 202]:
                logging.info(Fore.MAGENTA + f"🗑️ Deleted: {file_name}")
            else:
                logging.warning(Fore.YELLOW + f"⚠️ Failed to delete {file_name}: {resp.text}")
        except Exception as e:
            logging.error(Fore.RED + f"❌ Exception deleting {file_name}: {e}")


def upload_folder_to_sharepoint(drive_id, target_folder_id, local_folder):
    files = list(local_folder.iterdir())
    logging.info(Fore.CYAN + f"🚀 Uploading {len(files)} files from {local_folder.name}...")

    def upload(file_path):
        try:
            upload_file(drive_id, target_folder_id, str(file_path))
            logging.info(Fore.GREEN + f"✅ Uploaded: {file_path.name}")
        except Exception as e:
            logging.error(Fore.RED + f"❌ Upload failed: {file_path.name} | {e}")

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(upload, file) for file in files if file.is_file()]
        for _ in tqdm(as_completed(futures), total=len(futures), desc="Uploading", dynamic_ncols=True):
            pass


def get_all_children(drive_id, folder_id):
    url = f"{BASE_URL}/drives/{drive_id}/items/{folder_id}/children"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    items = []
    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def wait_for_download(folder_path: str, timeout: int = 60):
    start_time = time.time()
    while True:
        in_progress = False
        for file in os.listdir(folder_path):
            if file.endswith('.download') or file.endswith('.tmp'):
                in_progress = True
                break
            full_path = os.path.join(folder_path, file)
            if os.path.isfile(full_path):
                try:
                    with open(full_path, 'rb') as f:
                        f.read(1)
                except (PermissionError, OSError):
                    in_progress = True
                    break

        if not in_progress:
            return

        if time.time() - start_time > timeout:
            raise TimeoutError(f"Download not completed within {timeout} seconds")

        time.sleep(1)
