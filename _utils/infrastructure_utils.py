import argparse
import base64
import json
import traceback
from datetime import datetime
from os.path import exists
from time import sleep
from _utils.azureutils import AdlsLogHandler
from pathlib import Path

# from _utils.common_utils import send_email #--> if common utils throwing error need to setup requirements file

import requests
from _utils.logger import LOG_FILE, logger

import argparse

SLEEP = 2
PAYLOAD = {}

def send_email(
        emailBody: str,
        emailSubject: str,
        recipientsTo: list,
        emailStyle: str = None,
        recipientsCC: list = None,
        recipientsBCC: list = None,
        attachments: list = None,
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
        }
    )

    return requests.post(url=call_url, headers=headers, data=payload)


####################### CLI arguments ################################
def parse_arguments_and_return_as_dict() -> dict:
    parser = argparse.ArgumentParser()
    parser.add_argument("--module", help="Module name")
    parser.add_argument("--process", help="Process name")
    parser.add_argument("--email_list", help="Email list")
    args = parser.parse_args()

    global PAYLOAD
    # open the JSON file
    payload_path = Path.cwd() / 'payload.json'
    if not exists(payload_path):
        raise FileNotFoundError(f'File not found: {payload_path}')
    else:
        with open(payload_path, encoding='utf-16') as file: # while using ps script to create payload use utf-8-sig encoding 
            # read the contents of the file as a string
            file_contents = file.read()
            # load the data from the file string
            PAYLOAD = json.loads(file_contents)
            file.close()

    return {
        "payload": PAYLOAD,
        "module": args.module,
        "process": args.process,
        "email_list": args.email_list,
    }


####################### CLI arguments ################################
def parse_cli_args() -> dict:
    try:
        print(1)
        # sleep(SLEEP)
        CLI = argparse.ArgumentParser()
        print(2)
        # sleep(SLEEP)
        argument_flags = ["--module", "--process", "--payload", "--email_list"]
        print(3)
        # sleep(SLEEP)
        for argument in argument_flags:
            CLI.add_argument(
                argument,  # name on the CLI - drop the `--` for positional/required parameters
                nargs=1 if not argument == "--email_list" else "*",
                type=json.loads if argument == "--payload" else str,
                default={"payload": None}
                if argument == "--payload"
                else None,  # default if nothing is provided
            )
        print(4)
        # sleep(SLEEP)
        print(CLI)
        # sleep(SLEEP)
        try:
            args = CLI.parse_args()
        except:
            print(traceback.format_exc(chain=True))
            sleep(5)
            raise Exception("Process failed parsing arguments.")
        print(5)
        sleep(SLEEP)
        return {
            "module": args.module[0],
            "process": args.process[0],
            "payload": args.payload[0],
            "email_list": args.email_list,
        }
    except:
        print(traceback.format_exc(chain=True))
        sleep(SLEEP)


####################### Master Wrapper ################################
def master_wrapper(func):
    def wrap(*args, **kwargs):
        arguments = parse_arguments_and_return_as_dict()
        start_time = datetime.now()
        log_file_name = f'Log_{arguments["module"]}_{arguments["process"]}_{start_time.strftime("%m-%d-%YT%H-%M-%S")}.txt'
        is_failed, error_msg = False, ""
        try:
            func()
            logger.info("FINAL STATUS: Successfully ran the process.")
        except Exception as e:
            exception_log = traceback.format_exc(chain=True)
            is_failed, error_msg = True, e
            logger.error(exception_log)
            logger.error(e)
            logger.error("FINAL STATUS: Failed to run process.")
            failed_time = datetime.now()
            emailBody = f"""<table style="border-collapse: collapse; width: 600px; margin: 0 auto;">
   <tr>
    <td colspan="2" style="background-color: #f44336; color: #fff; font-size: 20px; font-weight: bold; text-align: center; padding: 8px;">PROCESS FAILURE</td>
  </tr>
  <tr>
    <td colspan="2" style="border: 1px solid black; text-align: center;">The process below has failed to run. Look at the end of the message for the exception log.</td>
  </tr>
  <tr>
    <td style="border: 1px solid black; text-align: right; font-weight: bold; padding-right: 10px;">Started:</td>
    <td style="border: 1px solid black; padding-left: 10px;">{start_time}</td>
  </tr>
  <tr>
    <td style="border: 1px solid black; text-align: right; font-weight: bold; padding-right: 10px;">Failed:</td>
    <td style="border: 1px solid black; padding-left: 10px;">{failed_time}</td>
  </tr>
  <tr>
    <td style="border: 1px solid black; text-align: right; font-weight: bold; padding-right: 10px;">Process Name:</td>
    <td style="border: 1px solid black; padding-left: 10px;">{arguments["module"]}--{arguments["process"]}</td>
  </tr>
  <tr>
    <td style="border: 1px solid black; text-align: right; font-weight: bold; padding-right: 10px;">Payload:</td>
    <td style="border: 1px solid black; padding-left: 10px;">{arguments["payload"] if len(str(arguments["payload"])) < 1000 else "too long to display, refer logs"}</td>
  </tr>
  <tr>
    <td style="border: 1px solid black; text-align: right; font-weight: bold; padding-right: 10px;">Error message:</td>
    <td style="border: 1px solid black; padding-left: 10px;">{exception_log}</td>
  </tr>
  <tr>
    <td colspan="2" style="border: 1px solid black; text-align: center; font-weight: bold;">End of message</td>
  </tr>
</table>
"""
            emailSubject = (
                f'PROCESS FAILURE: HIGHGATE-{str(arguments["module"]).upper()}--{str(arguments["process"]).upper()}'
            )
            recipientsTo = arguments["email_list"].split(",")
            log_data = open(LOG_FILE, "rb").read()
            log_file_encoded = base64.b64encode(log_data).decode("UTF-8")
            send_email(
                emailBody=emailBody,
                emailSubject=emailSubject,
                recipientsTo=recipientsTo,
                attachments=[{"ContentBytes": log_file_encoded, "Name": log_file_name}],
            )
        finally:
            # Upload logfile to ADLS
            is_prod = True
            if is_prod:
                end_time = datetime.now()
                time_taken = end_time - start_time
                log_folder = f'Highgate/PowerAutomateLogs/{start_time.year}-{start_time.strftime("%B")}/{start_time.day:02}/{arguments["module"]}/{arguments["module"]}_{arguments["process"]}_T{start_time.strftime("%H-%M")}'
                adls_log = AdlsLogHandler()
                entity = {
                    "PartitionKey": start_time.strftime("%m-%d-%Y"),
                    "RowKey": start_time.strftime("%m%d%YT%H%M%S"),
                    "client": "Highgate",
                    "module": str(arguments["module"]).capitalize(),
                    "process": str(arguments["process"]).capitalize(),
                    "start_time": start_time.strftime("%d-%m-%YT%H-%M-%S"),
                    "end_time": end_time.strftime("%d-%m-%YT%H-%M-%S"),
                    "time_take_seconds": time_taken.seconds,
                    "status": ("Completed" if not is_failed else "Failed"),
                    "error_msg": str(error_msg),
                    "log_folder_location": log_folder,
                    "log_link" : "https://blueoceansp.blob.core.windows.net/root/" + log_folder + "/automation.log",
                    "email_list" : str(arguments['email_list'])
                }
                adls_log.save_run_logs(entity=entity)
                # upload artificats directory files
                for file in Path(Path.cwd() / "output").iterdir():
                    if file.is_file():
                        adls_log.upload(
                            local_path=Path.cwd() / "output",
                            fname=str(file.name),
                            adls_path=log_folder,
                        )

    return wrap
