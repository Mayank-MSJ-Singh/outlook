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

def outlookMail_create_mail_folder(display_name: str, is_hidden: bool = False) -> dict:
    """
    Create a new mail folder in the signed-in user's mailbox.

    Args:
        display_name (str): The name of the new folder.
        is_hidden (bool, optional): Whether the folder is hidden. Defaults to False.

    Returns:
        dict: JSON response from Microsoft Graph with the created folder info,
              or error info if request fails.
    """
    client = get_onedrive_client()
    if not client:
        logging.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/mailFolders"
    payload = {
        "displayName": display_name,
        "isHidden": is_hidden
    }

    try:
        response = requests.post(url, headers=client['headers'], json=payload)
        response.raise_for_status()
        logging.info(f"Created mail folder: {display_name}")
        return response.json()
    except Exception as e:
        logging.error(f"Could not create mail folder at {url}: {e}")
        return {"error": f"Could not create mail folder at {url}"}

def outlookMail_list_child_folders(folder_id: str, include_hidden: bool = False) -> dict:
    """
    List child folders of a specific Outlook mail folder.

    Args:
        folder_id (str): ID of the parent folder.
        include_hidden (bool, optional): Whether to include hidden folders. Defaults to False.

    Returns:
        dict: JSON response from Microsoft Graph with the list of child folders,
              or error info if request fails.
    """
    client = get_onedrive_client()
    if not client:
        logging.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/mailFolders/{folder_id}/childFolders"
    params = {}
    if include_hidden:
        params['includeHiddenFolders'] = 'true'

    try:
        response = requests.get(url, headers=client['headers'], params=params)
        response.raise_for_status()
        logging.info(f"Retrieved child folders of folder: {folder_id}")
        return response.json()
    except Exception as e:
        logging.error(f"Could not get child folders from {url}: {e}")
        return {"error": f"Could not get child folders from {url}"}

def outlookMail_create_child_folder(parent_folder_id: str, display_name: str, is_hidden: bool = False) -> dict:
    """
    Create a child mail folder inside a specified Outlook parent folder.

    Args:
        parent_folder_id (str): ID of the parent mail folder.
        display_name (str): Display name for the new folder.
        is_hidden (bool, optional): Whether the new folder should be hidden. Defaults to False.

    Returns:
        dict: Created folder details on success, or error message.
    """
    client = get_onedrive_client()
    if not client:
        logging.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/mailFolders/{parent_folder_id}/childFolders"

    payload = {
        "displayName": display_name,
        "isHidden": is_hidden
    }

    try:
        response = requests.post(url, headers=client['headers'], json=payload)
        response.raise_for_status()
        logging.info(f"Created child folder '{display_name}' under folder: {parent_folder_id}")
        return response.json()
    except Exception as e:
        logging.error(f"Could not create child folder at {url}: {e}")
        return {"error": f"Could not create child folder at {url}"}

def outlookMail_list_messages_from_folder(folder_id: str, top: int = 10) -> dict:
    """
    Retrieve messages from a specific Outlook mail folder.

    Args:
        folder_id (str): The ID of the folder.
        top (int, optional): Number of messages to return. Defaults to 10.

    Returns:
        dict: JSON response with list of messages, or error details.
    """
    client = get_onedrive_client()
    if not client:
        logging.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/mailFolders/{folder_id}/messages"
    params = {'$top': top}

    try:
        response = requests.get(url, headers=client['headers'], params=params)
        response.raise_for_status()
        logging.info(f"Fetched messages from folder: {folder_id}")
        return response.json()
    except Exception as e:
        logging.error(f"Failed to fetch messages from {url}: {e}")
        return {"error": f"Failed to fetch messages from {url}"}

def outlookMail_update_folder_display_name(folder_id: str, display_name: str) -> dict:
    """
    Update the display name of an Outlook mail folder.

    Args:
        folder_id (str): ID of the mail folder to update.
        display_name (str): New display name.

    Returns:
        dict: JSON response on success, or error details.
    """
    client = get_onedrive_client()
    if not client:
        logging.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/mailFolders/{folder_id}"
    payload = {"displayName": display_name}

    try:
        response = requests.patch(url, headers=client['headers'], json=payload)
        response.raise_for_status()
        logging.info(f"Updated folder {folder_id} display name to '{display_name}'")
        return response.json()
    except Exception as e:
        logging.error(f"Failed to update folder at {url}: {e}")
        return {"error": f"Failed to update folder at {url}"}