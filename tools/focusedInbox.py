import requests
import logging
from typing import Tuple, Union, Dict, Any
from base import get_onedrive_client
import base64
import os

def outlookMail_update_inference_override(override_id: str, classify_as: str = "focused") -> dict:
    """
    Update an existing inference classification override.

    Args:
        override_id (str): The ID of the override to update.
        classify_as (str): "focused" or "other".

    Returns:
        dict: JSON response from Microsoft Graph API, or an error message.
    """
    client = get_onedrive_client()  # your function to get the authenticated client
    if not client:
        logging.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/inferenceClassification/overrides/{override_id}"
    payload = {
        "classifyAs": classify_as
    }

    try:
        response = requests.patch(url, headers=client['headers'], json=payload)
        response.raise_for_status()
        logging.info("Updated inference classification override")
        return response.json()
    except Exception as e:
        logging.error(f"Could not update inference classification override at {url}: {e}")
        return {"error": f"Could not update inference classification override at {url}"}

def outlookMail_delete_inference_override(override_id: str) -> str:
    """
    Delete an inference classification override by ID.

    Args:
        override_id (str): The ID of the override to delete.

    Returns:
        str: "Deleted" on success, or an error message.
    """
    client = get_onedrive_client()  # your helper to get the authenticated client
    if not client:
        logging.error("Could not get Outlook client")
        return "Could not get Outlook client"

    url = f"{client['base_url']}/me/inferenceClassification/overrides/{override_id}"

    try:
        response = requests.delete(url, headers=client['headers'])
        if response.status_code == 204:
            logging.info("Deleted inference classification override")
            return "Deleted"
        else:
            try:
                return response.json()
            except Exception:
                return f"Unexpected response: {response.status_code}"
    except Exception as e:
        logging.error(f"Could not delete inference classification override at {url}: {e}")
        return f"Error: {e}"