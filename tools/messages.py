import requests
import logging
import os
from typing import Tuple, Union, Dict, Any
from base import get_onedrive_client
import json

# Configure logging
logger = logging.getLogger(__name__)

def outlookMail_list_messages(top: int = 10, filter_query: str = None, orderby: str = None):
    """
        Retrieve a list of Outlook mail messages from the signed-in user's mailbox.

        Args:
            top (int, optional):
                The maximum number of messages to return. Defaults to 10.
            filter_query (str, optional):
                An OData $filter expression to filter messages by specific criteria.
                Example filters you can use:
                    - "isRead eq false"                          → Only unread emails
                    - "importance eq 'high'"                     → Emails marked as high importance
                    - "from/emailAddress/address eq 'example@example.com'" → Emails sent by a specific address
                    - "subject eq 'Welcome'"                      → Emails with a specific subject
                    - "receivedDateTime ge 2025-07-01T00:00:00Z" → Emails received after a date
                    - "hasAttachments eq true"                    → Emails that include attachments
                    - Combine filters: "isRead eq false and importance eq 'high'"
            orderby (str, optional):
                An OData $orderby expression to sort results.
                Example: "receivedDateTime desc" (newest first)

        Returns:
            dict: JSON response from the Microsoft Graph API containing the list of messages
                  or an error message if the request fails.

        Notes:
            - Requires an authenticated Outlook client.
            - This function internally builds the API request to:
              GET https://graph.microsoft.com/v1.0/me/messages
              with the provided query parameters.
        """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}
    logging.info("Retrieving Outlook mail messages")

    url = f"{client['base_url']}/me/messages"
    params = {'$top': top}

    if filter_query:
        params['$filter'] = filter_query
    if orderby:
        params['$orderby'] = orderby
    try:
        response = requests.get(url, headers=client['headers'], params=params)
        logger.info("Retrieved Outlook mail messages")
        return response.json()
    except Exception as e:
        logger.error(f"Could not get Outlook messages from {url}: {e}")
        return {"error": f"Could not get Outlook messages from {url}"}

def outlookMail_list_messages_from_folder(folder_id : str, top: int = 10, filter_query: str = None, orderby: str = None):
    """
    Retrieve a list of Outlook mail messages from a specific folder in the signed-in user's mailbox.

    Args:
        folder_id (str):
            The unique ID of the Outlook mail folder to retrieve messages from.
            Example: 'AQMkADAwATNiZmYAZS05YmUxLTk3NDYtMDACLTAwCgAuAAAD...'
        top (int, optional):
            The maximum number of messages to return. Defaults to 10.
        filter_query (str, optional):
            An OData $filter expression to filter messages by specific criteria.
            Example filters you can use:
                - "isRead eq false"                          → Only unread emails
                - "importance eq 'high'"                     → Emails marked as high importance
                - "from/emailAddress/address eq 'example@example.com'" → Emails sent by a specific address
                - "subject eq 'Welcome'"                      → Emails with a specific subject
                - "receivedDateTime ge 2025-07-01T00:00:00Z" → Emails received after a date
                - "hasAttachments eq true"                    → Emails that include attachments
                - Combine filters: "isRead eq false and importance eq 'high'"
        orderby (str, optional):
            An OData $orderby expression to sort results.
            Example: "receivedDateTime desc" (newest first)

    Returns:
        dict: JSON response from the Microsoft Graph API containing the list of messages,
              or an error message if the request fails.

    Notes:
        - Requires an authenticated Outlook client.
        - This function sends a GET request to:
          https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages
          with the provided query parameters.
    """

    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}
    url = f"{client['base_url']}/me/mailFolders/{folder_id}/messages"

    params = {'$top': top}

    if filter_query:
        params['$filter'] = filter_query
    if orderby:
        params['$orderby'] = orderby
    try:
        response = requests.get(url, headers=client['headers'], params=params)
        logger.info("Retrieved Outlook mail messages")
        return response.json()
    except Exception as e:
        logger.error(f"Could not get Outlook messages from {url}: {e}")
        return {"error": f"Could not get Outlook messages from {url}"}


def outlookMail_create_draft(subject: str, body_content: str, to_recipients: list, importance: str = "Normal"):
    """
    Create a draft Outlook mail message for the signed-in user.

    Args:
        subject (str): The subject of the draft email.
        body_content (str): The HTML content of the email body.
        to_recipients (list): List of recipient email addresses as strings.
                              Example: ["someone@example.com", "another@example.com"]
        importance (str, optional): Importance level ("Low", "Normal", "High"). Defaults to "Normal".

    Returns:
        dict: JSON response from Microsoft Graph API with the created draft's details,
              or an error message if the request fails.

    Notes:
        - Requires an authenticated Outlook client.
        - The draft won't be sent automatically; it's saved in the Drafts folder.
    """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages"

    # Build recipient list in required format
    recipients = [
        {"emailAddress": {"address": email}}
        for email in json.loads(to_recipients)
    ]

    payload = {
        "subject": subject,
        "importance": importance,
        "body": {
            "contentType": "HTML",
            "content": body_content
        },
        "toRecipients": recipients
    }

    try:
        response = requests.post(url, headers=client['headers'], json=payload)
        logger.info("Created draft Outlook mail message")
        return response.json()
    except Exception as e:
        logger.error(f"Could not create Outlook draft message at {url}: {e}")
        return {"error": f"Could not create Outlook draft message at {url}"}


if __name__ == "__main__":
    #print(outlookMail_list_messages(top = 1))
    #print(outlookMail_list_messages_from_folder(folder_id='AQMkADAwATNiZmYAZS05YmUxLTk3NDYtMDACLTAwCgAuAAADb25xEWuFWEWCX6SpYNrvPwEAp93M14k-O06xyivtWYvXZgAAAgEMAAAA'))
    draft = outlookMail_create_draft(
        subject="Test draft from API",
        body_content="<p>Hello, this is a draft.</p>",
        to_recipients='["mayank.msj.singh@gmail.com"]'
    )
    print(draft)

    pass