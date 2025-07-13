import requests
import logging
from typing import Tuple, Union, Dict, Any
from base import get_onedrive_client
import base64
import os

# Configure logging
logger = logging.getLogger(__name__)

def outlookMail_list_folders(include_hidden: bool = True) -> dict:
    """
    List mail folders in the signed-in user's mailbox.

    Args:
        include_hidden (bool, optional): Whether to include hidden folders. Defaults to True.

    Returns:
        dict: JSON response with list of folders or an error.
    """
    client = get_onedrive_client()
    if not client:
        logging.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/mailFolders"
    params = {}
    if include_hidden:
        params["includeHiddenFolders"] = "true"

    try:
        response = requests.get(url, headers=client['headers'], params=params)
        response.raise_for_status()
        logging.info("Fetched mail folders")
        return response.json()
    except Exception as e:
        logging.error(f"Could not get mail folders from {url}: {e}")
        return {"error": f"Could not get mail folders from {url}: {e}"}

def outlookMail_get_mail_folder(folder_id: str) -> dict:
    """
    Get details of a specific mail folder by its ID.

    Args:
        folder_id (str): The unique ID of the mail folder.

    Returns:
        dict: JSON response from Microsoft Graph with folder details,
              or error info if request fails.
    """
    client = get_onedrive_client()
    if not client:
        logging.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/mailFolders/{folder_id}"

    try:
        response = requests.get(url, headers=client['headers'])
        response.raise_for_status()
        logging.info(f"Retrieved mail folder {folder_id}")
        return response.json()
    except Exception as e:
        logging.error(f"Could not get mail folder at {url}: {e}")
        return {"error": f"Could not get mail folder at {url}"}