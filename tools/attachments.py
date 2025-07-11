import requests
import logging
from typing import Tuple, Union, Dict, Any
from base import get_onedrive_client

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
