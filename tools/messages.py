import requests
import logging
import os
from typing import Tuple, Union, Dict, Any
from base import get_onedrive_client

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


if __name__ == "__main__":
    print(outlookMail_list_messages(top = 1))
    pass