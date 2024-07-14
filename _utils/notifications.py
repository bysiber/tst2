import requests
import json
from config import Credentials


def send_email(
        emailBody: str,
        emailSubject: str,
        recipientsTo: list,
        emailStyle: str = None,
        recipientsCC: list = None,
        recipientsBCC: list = None,
        attachments: list = None):
    """
    This function triggers a logic app to send an email as specified in the parameters.

    :params
        emailBody: Body of the email as plain text or as HTML code.
        emailSubject: Subject of the email as plain text.
        emailStyle: Styling for the email HTML header as HTML code.
        recipientsTo: Main recipient of the email.
        recipientsCC: CC recipient of the email.
        recipientsBCC: BCC recipient of the email.
        attachments: sttachments as a list with each item as a dict containing
                     ContentBytes as a base64encoded string
                     Name of the file as a string
                     Example:
                         [
                             {
                                 "ContentBytes": "UEsDBBQAAAAIAHdc+1QHQU1igQAAAL...redacted",
                                 "Name": "testfile1.txt"
                             },
                             {
                                 "ContentBytes": "UEsDBBQAAAAIAHdc+nhtbE2OPQsCMR...redacted",
                                 "Name": "testfile2.txt"
                             }
                         ]
    """

    if not isinstance(recipientsTo, list):
        raise (TypeError("recipientsTo must be a list"))
    if not isinstance(recipientsCC, list) and recipientsCC is not None:
        raise (TypeError("recipientsCC must be a list"))
    if not isinstance(recipientsBCC, list) and recipientsBCC is not None:
        raise (TypeError("recipientsBCC must be a list"))

    call_url = Credentials.send_email_call_url
    headers = {"Content-Type": "application/json"}
    payload = json.dumps(
        {
            "emailBody": emailBody,
            "emailSubject": emailSubject,
            "emailStyle": emailStyle,
            "recipientsTo": ";".join(recipientsTo),
            "recipientsCC": ";".join(recipientsCC) if recipientsCC else None,
            "recipientsBCC": ";".join(recipientsBCC) if recipientsBCC else None,
            "attachments": attachments,
        }
    )

    return requests.post(url=call_url, headers=headers, data=payload)
