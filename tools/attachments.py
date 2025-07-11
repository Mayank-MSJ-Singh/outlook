import requests
import logging
from typing import Tuple, Union, Dict, Any
from base import get_onedrive_client
import base64

# Configure logging
logger = logging.getLogger(__name__)

def outlookMail_get_attachment(message_id: str, attachment_id: str, expand: str = None):
    """
    Get a specific attachment from an Outlook mail message.

    Args:
        message_id (str): ID of the message that has the attachment.
        attachment_id (str): ID of the attachment to retrieve.
        expand (str, optional):
            An OData $expand expression to include related entities.
            Example: "microsoft.graph.itemattachment/item" to expand item attachments.

    Returns:
        dict: JSON response from Microsoft Graph API with attachment details,
              or an error message if the request fails.

    Notes:
        - Requires an authenticated Outlook client.
        - This function sends a GET request to:
          https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments/{attachment_id}
          and adds $expand if provided.
    """
    client = get_onedrive_client()  # same client you’re using for other outlook calls
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}/attachments/{attachment_id}"

    params = {}
    if expand:
        params['$expand'] = expand

    try:
        response = requests.get(url, headers=client['headers'], params=params)
        logger.info("Fetched attachment from Outlook mail")
        return response.json()
    except Exception as e:
        logger.error(f"Could not get Outlook attachment at {url}: {e}")
        return {"error": f"Could not get Outlook attachment at {url}"}

def outlookMail_download_attachment(message_id: str, attachment_id: str, save_path: str):
    """
    Download an attachment from an Outlook mail message and save it locally.

    Args:
        message_id (str): The ID of the message containing the attachment.
        attachment_id (str): The ID of the attachment to download.
        save_path (str): Local file path to save the downloaded attachment.

    Returns:
        str: Path to saved file if successful, or error message.
    """
    client = get_onedrive_client()  # same function you're using for other Outlook calls
    if not client:
        logging.error("Could not get Outlook client")
        return "Could not get Outlook client"

    url = f"{client['base_url']}/me/messages/{message_id}/attachments/{attachment_id}"

    try:
        response = requests.get(url, headers=client['headers'])
        response.raise_for_status()
        data = response.json()

        # 'contentBytes' is base64 encoded
        content_bytes = data.get("contentBytes")
        if content_bytes:
            file_data = base64.b64decode(content_bytes)
            with open(save_path, "wb") as f:
                f.write(file_data)
            logging.info(f"Attachment saved to {save_path}")
            return save_path
        else:
            logging.error(f"No contentBytes found in attachment: {attachment_id}")
            return "Attachment does not have downloadable content"

    except Exception as e:
        logging.error(f"Failed to download attachment at {url}: {e}")
        return f"Error: {e}"
