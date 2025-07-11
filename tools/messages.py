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
        for email in to_recipients
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

def outlookMail_create_draft_in_folder(folder_id : str, subject: str, body_content: str, to_recipients: list, importance: str = "Normal"):
    """
    Create a draft Outlook mail message in specific folder in the signed-in user's mailbox.

    Args:
        folder_id (str):
            The unique ID of the Outlook mail folder to retrieve messages from.
            Example: 'AQMkADAwATNiZmYAZS05YmUxLTk3NDYtMDACLTAwCgAuAAAD...'
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

    url = f"{client['base_url']}/me/mailFolders/{folder_id}/messages"

    # Build recipient list in required format
    recipients = [
        {"emailAddress": {"address": email}}
        for email in to_recipients
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

def outlookMail_update_draft(message_id: str, subject: str = None, body_content: str = None, to_recipients: list = None, importance: str = "Normal"):
    """
    Update an existing Outlook draft message by message ID.

    Args:
        message_id (str): The ID of the draft message to update.
        subject (str, optional): New subject.
        body_content (str, optional): New HTML content.
        to_recipients (list, optional): New list of email addresses.

    Returns:
        dict: JSON response from Microsoft Graph API with updated draft details,
              or an error message if the request fails.
    """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}"

    payload = {}

    if subject:
        payload["subject"] = subject

    if body_content:
        payload["body"] = {
            "contentType": "HTML",
            "content": body_content
        }

    if to_recipients:
        # make sure it's in the right format
        recipients = [{"emailAddress": {"address": email}} for email in to_recipients]
        payload["toRecipients"] = recipients

    if importance:
        payload["importance"] = importance

    try:
        response = requests.patch(url, headers=client['headers'], json=payload)
        logger.info("Updated draft Outlook mail message")
        return response.json()
    except Exception as e:
        logger.error(f"Could not update Outlook draft message at {url}: {e}")
        return {"error": f"Could not update Outlook draft message at {url}"}

def outlookMail_delete_draft(message_id: str):
    """
    Delete an existing Outlook draft message by message ID.

    Args:
        message_id (str): The ID of the draft message to Delete.

    Returns:
        dict: JSON response from Microsoft Graph API with updated draft details,
              or an error message if the request fails.
    """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}"

    try:
        logger.info(f"Deleting draft Outlook mail message at {url}")
        response = requests.delete(url, headers=client['headers'])
        if response.status_code == 204:
            logger.info("Deleted draft Outlook mail message successfully")
            return "Deleted"
        else:
            logger.warning(f"Unexpected status code: {response.status_code}")
            # try to parse error if there is one
            try:
                error_response = response.json()
                logger.error(f"Delete failed with response: {error_response}")
                return error_response
            except Exception as parse_error:
                logger.error(f"Could not parse error response: {parse_error}")
                return {"error": f"Unexpected response: {response.status_code}"}
    except Exception as e:
        logger.error(f"Could not delete Outlook draft message at {url}: {e}")
        return {"error": f"Could not delete Outlook draft message at {url}"}

def outlookMail_copy_message(message_id: str, destination_folder_id: str):
    """
    Copy an existing Outlook draft message by message ID.

    Args:
        message_id (str): The ID of the draft message to Delete.
        folder_id (str): The ID of the destination folder.

    Returns:
        dict: JSON response from Microsoft Graph API with updated draft details,
              or an error message if the request fails.
    """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}"
    payload = {
  "destinationId": destination_folder_id,
}

    try:
        logger.info(f"Coping draft Outlook mail message at {url}")
        response = requests.delete(url, headers=client['headers'], json=payload)
        if response.status_code == 204:
            logger.info("Copied draft Outlook mail message successfully")
            return "Copied"
        else:
            logger.warning(f"Unexpected status code: {response.status_code}")
            # try to parse error if there is one
            try:
                error_response = response.json()
                logger.error(f"Copy failed with response: {error_response}")
                return error_response
            except Exception as parse_error:
                logger.error(f"Could not parse error response: {parse_error}")
                return {"error": f"Unexpected response: {response.status_code}"}
    except Exception as e:
        logger.error(f"Could not copy Outlook draft message at {url}: {e}")
        return {"error": f"Could not copy Outlook draft message at {url}"}

def outlookMail_create_forward_draft(message_id: str, comment: str, to_recipients: list):
    """
    Create a draft forward message for an existing Outlook message.

    Args:
        message_id (str): ID of the original message to forward.
        comment (str): Comment to include in the forwarded message.
        to_recipients (list): List of recipient email addresses as strings.

    Returns:
        dict: JSON response from Microsoft Graph API with the created draft forward's details,
              or an error message if the request fails.
    """
    client = get_onedrive_client()  # same method you used to get your client and headers
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}/createForward"

    # Build recipient list in required format
    recipients = [{"emailAddress": {"address": email}} for email in to_recipients]

    payload = {
        "comment": comment,
        "toRecipients": recipients
    }

    try:
        response = requests.post(url, headers=client['headers'], json=payload)
        logger.info("Created draft forward Outlook mail message")
        return response.json()
    except Exception as e:
        logger.error(f"Could not create Outlook forward draft message at {url}: {e}")
        return {"error": f"Could not create Outlook forward draft message at {url}"}

def outlookMail_create_reply_draft(message_id: str, comment: str):
    """
    Create a draft reply message to an existing Outlook message.

    Args:
        message_id (str): ID of the original message to reply to.
        comment (str): Comment to include in the reply.


    Returns:
        dict: JSON response from Microsoft Graph API with the created draft reply's details,
              or an error message if the request fails.
    """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}/createReply"

    payload = {
        "comment": comment
    }


    try:
        response = requests.post(url, headers=client['headers'], json=payload)
        logger.info("Created draft reply Outlook mail message")
        return response.json()
    except Exception as e:
        logger.error(f"Could not create Outlook reply draft message at {url}: {e}")
        return {"error": f"Could not create Outlook reply draft message at {url}"}


def outlookMail_create_reply_all_draft(message_id: str, comment: str = ""):
    """
    Create a reply-all draft to an existing Outlook message.

    Args:
        message_id (str): The ID of the message you want to reply to.
        comment (str, optional): Text to include in the reply body.

    Returns:
        dict: JSON response from Microsoft Graph API with the draft details,
              or an error message if the request fails.
    """
    client = get_onedrive_client()  # reuse your existing token logic
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}/createReplyAll"

    payload = {
        "comment": comment
    }

    try:
        response = requests.post(url, headers=client['headers'], json=payload)
        logger.info("Created reply-all draft Outlook mail message")
        return response.json()
    except Exception as e:
        logger.error(f"Could not create reply-all draft at {url}: {e}")
        return {"error": f"Could not create reply-all draft at {url}"}

def outlookMail_forward_message(message_id: str, to_recipients: list, comment: str = ""):
    """
    Forward an Outlook message by message ID.

    Args:
        message_id (str): ID of the message to forward.
        to_recipients (list): List of email addresses to forward to.
        comment (str, optional): Comment to include above the forwarded message.

    Returns:
        dict: JSON response or error.
    """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}/forward"

    recipients = [{"emailAddress": {"address": email}} for email in to_recipients]

    payload = {
        "comment": comment,
        "toRecipients": recipients
    }

    try:
        response = requests.post(url, headers=client['headers'], json=payload)
        if response.status_code in (202, 200):
            logger.info("Forwarded Outlook mail message")
            return {"success": True}
        else:
            logger.error(f"Forward failed: {response.status_code} {response.text}")
            return response.json()
    except Exception as e:
        logger.error(f"Could not forward Outlook message at {url}: {e}")
        return {"error": f"Could not forward Outlook message at {url}"}

def outlookMail_move_message(message_id: str, destination_folder_id: str):
    """
    Move an Outlook mail message to another folder.

    Args:
        message_id (str): ID of the message to move.
        destination_folder_id (str): ID of the target folder.
                                     Example: 'deleteditems' or actual folder ID.

    Returns:
        dict: JSON response from Microsoft Graph API with moved message details,
              or an error message if it fails.
    """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}/move"

    payload = {
        "destinationId": destination_folder_id
    }

    try:
        response = requests.post(url, headers=client['headers'], json=payload)
        logger.info(f"Moved Outlook mail message to folder {destination_folder_id}")
        return response.json()
    except Exception as e:
        logger.error(f"Could not move Outlook mail message at {url}: {e}")
        return {"error": f"Could not move Outlook mail message at {url}"}

def outlookMail_send_reply_custom(message_id: str, comment: str, to_recipients: list):
    """
    Send a reply to an Outlook mail message, with custom recipients and names.

    Args:
        message_id (str): ID of the message to reply to.
        comment (str): Text to include in the reply.
        to_recipients (list): List of dicts like:
            [{"name": "Samantha Booth", "address": "samanthab@contoso.com"}, ...]

    Returns:
        str: "Sent" if successful, or error message.
    """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}/reply"

    recipients = [
        {"emailAddress": {"address": r["address"], "name": r["name"]}}
        for r in to_recipients
    ]

    payload = {
        "message": {
            "toRecipients": recipients
        },
        "comment": comment
    }

    try:
        response = requests.post(url, headers=client['headers'], json=payload)
        if response.status_code in [200, 202]:
            logger.info(f"Replied (custom) to message {message_id}")
            return "Sent"
        else:
            try:
                return response.json()
            except:
                return {"error": f"Unexpected response: {response.status_code}"}
    except Exception as e:
        logger.error(f"Could not reply (custom) to Outlook mail message at {url}: {e}")
        return {"error": f"Could not reply (custom) to Outlook mail message at {url}"}


def outlookMail_reply_all(message_id: str, comment: str):
    """
    Reply all to a message by ID with a comment.

    Args:
        message_id (str): The ID of the message.
        comment (str): Your reply comment.

    Returns:
        dict: JSON response from Microsoft Graph API.
    """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}/replyAll"
    payload = {
        "comment": comment
    }

    try:
        response = requests.post(url, headers=client['headers'], json=payload)
        logger.info("Replied all to Outlook message")
        return response.json()
    except Exception as e:
        logger.error(f"Could not reply all to Outlook message at {url}: {e}")
        return {"error": f"Could not reply all to Outlook message at {url}"}

def outlookMail_send_draft(message_id: str):
    """
    Send an existing draft Outlook mail message by message ID.

    Args:
        message_id (str): The ID of the draft message to send.

    Returns:
        dict: Empty response if successful, or error details.
    """
    client = get_onedrive_client()
    if not client:
        logger.error("Could not get Outlook client")
        return {"error": "Could not get Outlook client"}

    url = f"{client['base_url']}/me/messages/{message_id}/send"

    try:
        response = requests.post(url, headers=client['headers'])
        if response.status_code == 202 or response.status_code == 200 or response.status_code == 204:
            logger.info("Draft sent successfully")
            return {"success": "Draft sent successfully"}
        else:
            try:
                return response.json()
            except Exception:
                return {"error": f"Unexpected response: {response.status_code}"}
    except Exception as e:
        logger.error(f"Could not send Outlook draft message at {url}: {e}")
        return {"error": f"Could not send Outlook draft message at {url}"}


if __name__ == "__main__":
    #print(outlookMail_list_messages(top = 1))
    print(outlookMail_list_messages_from_folder(folder_id='AQMkADAwATNiZmYAZS05YmUxLTk3NDYtMDACLTAwCgAuAAADb25xEWuFWEWCX6SpYNrvPwEAp93M14k-O06xyivtWYvXZgAAAgEMAAAA'))
    '''
    draft = outlookMail_create_draft(
        subject="Test draft",
        body_content="<p>Hello, this is a draft.</p>",
        to_recipients=["mayank.msj.singh@gmail.com"]
    )
    print(draft)
    '''


    '''
    draft = outlookMail_create_draft_in_folder(
        folder_id='AQMkADAwATNiZmYAZS05YmUxLTk3NDYtMDACLTAwCgAuAAADb25xEWuFWEWCX6SpYNrvPwEAp93M14k-O06xyivtWYvXZgAAAgEMAAAA',
        subject="Test draft",
        body_content="<p>Hello, this is a draft.</p>",
        to_recipients=["mayank.msj.singh@gmail.com"]
    )
    print(draft)
    '''
    '''
    draft_id = "AQMkADAwATNiZmYAZS05YmUxLTk3NDYtMDACLTAwCgBGAAADb25xEWuFWEWCX6SpYNrvPwcAp93M14k-O06xyivtWYvXZgAAAgEPAAAAp93M14k-O06xyivtWYvXZgAAAn1oAAAA"

    result = outlookMail_update_draft(
        message_id=draft_id,
        subject="Updated subject",
        body_content="<p>Updated body content from API</p>",
        importance="High"
    )
    print(result)
    '''
    '''
    draft_id = "AQMkADAwATNiZmYAZS05YmUxLTk3NDYtMDACLTAwCgBGAAADb25xEWuFWEWCX6SpYNrvPwcAp93M14k-O06xyivtWYvXZgAAAgEPAAAAp93M14k-O06xyivtWYvXZgAAAn1oAAAA"

    result = outlookMail_update_draft(
            message_id=draft_id,
            subject="Updated subject",
            body_content="<p>Updated body content from API</p>",
            importance="High"
        )
    print(result)
    '''
    #print(outlookMail_delete_draft("AQMkADAwATNiZmYAZS05YmUxLTk3NDYtMDACLTAwCgBGAAADb25xEWuFWEWCX6SpYNrvPwcAp93M14k-O06xyivtWYvXZgAAAgEPAAAAp93M14k-O06xyivtWYvXZgAAAn1rAAAA"))
    #print(outlookMail_copy_message('AQMkADAwATNiZmYAZS05YmUxLTk3NDYtMDACLTAwCgBGAAADb25xEWuFWEWCX6SpYNrvPwcAp93M14k-O06xyivtWYvXZgAAAgEPAAAAp93M14k-O06xyivtWYvXZgAAAn1sAAAA','AQMkADAwATNiZmYAZS05YmUxLTk3NDYtMDACLTAwCgAuAAADb25xEWuFWEWCX6SpYNrvPwEAp93M14k-O06xyivtWYvXZgAAAgEMAAAA'))
    pass