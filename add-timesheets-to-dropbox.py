import dropbox
import requests

import os
from dotenv import load_dotenv
load_dotenv()

# load environment variables
ACCESS_TOKEN = os.getenv('TEAM_ACCESS_TOKEN')
FOLDER_PATH = os.getenv('TEAM_FOLDER_PATH')  # Specify the folder path you want to search
DROPBOX_ROOT_ID = os.getenv('DROPBOX_ROOT_ID') # This can be found at the endpoint /users/get_current_account

# Initialize Dropbox client
dbx = dropbox.Dropbox(ACCESS_TOKEN)
KEYWORDS = ['"biomass"']

def search_files_and_folders(keyword, path):
    """Search for files and folders containing a specific keyword in a specified folder."""
    url = "https://api.dropboxapi.com/2/files/search_v2"
    headers = {
        "Authorization": f"Bearer {ACCESS_TOKEN}",
        "Content-Type": "application/json",
        "Dropbox-Api-Path-Root": f"{{\".tag\": \"root\", \"root\": \"3211714451\"}}"
    }
    data = {
        "query": keyword,
        "options": {
            "path": path,
            "file_status": "active",
            "file_categories":[{".tag":"folder"},{".tag":"document"},{".tag":"pdf"},{".tag":"spreadsheet"},{".tag":"image"}]
        }
    }
    try:
        response = requests.post(url, headers=headers, json=data)
        response.raise_for_status()
        results = response.json()
        print(results)
        return results.get('matches', [])
    except requests.exceptions.RequestException as err:
        print(f"Failed to search files and folders with keyword '{keyword}' in folder '{path}': {err}")
        print(err)
        return []

def main():
    for keyword in KEYWORDS:
        print(f"Searching for files and folders with keyword: {keyword} in folder: {FOLDER_PATH}")
        # # Process the root folder and its subfolders
        # process_folder(keyword, FOLDER_PATH)
    # print(f"Tagged {count} files and folders.")

if __name__ == '__main__':
    main()