import requests
from config import BASE_URL
from token_manager import get_access_token


def search_anken_folder(anken_number):
    search_url = f"{BASE_URL}/search/query"
    payload = {"requests": [{"entityTypes": ["driveItem"], "query": {"queryString": anken_number}, "region": "JPN"}]}
    headers = {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}
    resp = requests.post(search_url, headers=headers, json=payload)
    if resp.status_code != 200:
        raise Exception(f"Search failed: {resp.text}")
    results = resp.json()
    items = results.get("value", [])[0].get("hitsContainers", [])[0].get("hits", [])
    if not items:
        return None
    first = items[0]
    item = first.get("resource", {})
    return {"name": item.get("name"), "id": item.get("id"), "parentReference": item.get("parentReference", {})}


def list_children(drive_id, folder_id):
    url = f"{BASE_URL}/drives/{drive_id}/items/{folder_id}/children"
    headers = {"Authorization": f"Bearer {get_access_token()}"}
    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"List children failed: {resp.text}")
    return resp.json().get("value", [])
