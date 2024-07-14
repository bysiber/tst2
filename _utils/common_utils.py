import json
import os
# from datetime import date, datetime, timedelta
from pathlib import Path
# from typing import Optional
# import psutil
import requests
# from twilio.rest import Client
# from _utils.config import Credentials
from _utils.logger import logger
import zipfile


def zip_files(zip_filepath, files_to_zip):
    with zipfile.ZipFile(zip_filepath, 'w') as zipf:
        for file in files_to_zip:
            zipf.write(file, arcname=os.path.basename(file))
            logger.info(f"Added {file} to the zip.")
    logger.info(f'ZIP file was created from {len(files_to_zip)} files.')
    return zip_filepath


def send_email(
    emailBody: str,
    emailSubject: str,
    recipientsTo: list = [],
    emailStyle: str = None,
    recipientsCC: list = None,
    recipientsBCC: list = None,
    attachments: list = None,
    special_sender: str = None
) -> None:
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

    call_url = ""
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
            "specialSender": special_sender
        }
    )

    return requests.post(url=call_url, headers=headers, data=payload)


# def get_last_sms_body(phone_number: str, date_sent_after: Optional[datetime] = None) -> str:
#     """
#     Return the last SMS body sent to the given phone number

#     If date_sent_after parameter is not given, it is set to look for in the last 5 seconds.

#     params:
#         phone_number: Twilio phone number to get SMS body from
#         date_sent_after: Initial time to search for messages
#     return:
#         SMS body if there is a message after the given time, otherwise None
#     """

#     client = Client(Credentials.Twilio.account_sid, Credentials.Twilio.auth_token)

#     if not date_sent_after:
#         date_sent_after = datetime.utcnow() - timedelta(seconds=5)

#     message = client.messages.list(
#         from_=phone_number,
#         limit=1,
#         date_sent_after=date_sent_after,
#     )

#     if not message:
#         return None
#     else:
#         return message[0].body
    

def get_filename(path : Path):
        """"
            This function returns the last modified file in a directory.
            :param path: path of the directory
        """
        # List all files in the directory
        files = os.listdir(path)

        # Sort files by modification time
        files.sort(key= lambda x: os.path.getmtime(os.path.join(path, x)))
        last_file = files[-1]

        return last_file
    
