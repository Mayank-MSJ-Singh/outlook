import requests
import logging
from base import get_onedrive_client

# Configure logging
logger = logging.getLogger(__name__)

def outlookMail_list_inbox_rules() -> dict:
    """
    List all message rules (inbox rules) from the user's Inbox folder.

    Returns:
        dict: JSON response from Microsoft Graph API with the list of rules,
              or an error message if the request fails.
    """
    client = get_onedrive_client()  # same method you already use to get auth + base_url
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/mailFolders/inbox/messageRules"

    try:
        response = requests.get(url, headers=client['headers'])
        response.raise_for_status()
        logger.info("Retrieved inbox message rules")
        return response.json()
    except Exception as e:
        logger.error(f"Could not get inbox message rules from {url}: {e}")
        return {"error": f"Could not get inbox message rules from {url}: {e}"}
