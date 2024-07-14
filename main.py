from datetime import datetime
from utils.logger import logger
from utils.eom_pcard_report_generator import ReportGenerator as EOMPcardReportGenerator
from utils.web_automation_steps import WebAutomation
from utils.biweekly_report_generator import ReportGenerator as BIWeeklyReportGenerator
import config as config
import os
import json
import requests
import calendar


def get_last_business_day_in_month(year: int, month: int) -> int:
    return max(calendar.monthcalendar(year, month)[-1][:5])

def create_folders():
    # This methods create all the folder structure needed to run the process.
    logger.info(f'Creating folder structure for process execution..')

    for path in [config.Paths.resources_filepath, 
                 config.Paths.rules_filepath,
                 config.Paths.downloads_filepath,
                 config.Paths.processed_filepath]:
        if not os.path.exists(path):
            os.makedirs(path)
            logger.info(f'Directory {path} created locally')
            
def delete_existing_files():
    # This methods delete all the generated reports to clean the environment.
    logger.info(f'Deleting all created files for the process execution..')

    for path in [config.Paths.resources_filepath, 
                 config.Paths.rules_filepath,
                 config.Paths.downloads_filepath,
                 config.Paths.processed_filepath]:
        if os.path.exists(path):
            for filename in os.listdir(path):
                file_path = os.path.join(path, filename)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                except Exception as e:
                    logger.error(f'Error while deleting {file_path}: {e}')


def main(payload: dict = None):
    #Payload example. Please uncomment to test the process locally. 
    payload = {
         "report_type": "EOMPCARD",
         "send_to_vpos": False,
         "user_email": "testtest@ttt.ai"
    }

    # Payload validations.          
    try:
        if payload and len(payload) > 0:
            logger.info(f"Work Item Payload: {payload}")
            auxiliary_email = None
            send_to_vpos = None
            
            if "report_type" in payload:
                report_type = payload.pop("report_type")
            if "run_date" in payload:
                run_date = datetime.strptime(payload.pop("run_date"), "%m-%d-%Y").date()
                logger.info(f"Running reports for an specific date: {run_date}")
            if "send_to_vpos" in payload: 
                send_to_vpos = payload.pop("send_to_vpos")
                if send_to_vpos == False:
                    auxiliary_email = payload.pop("user_email")
        else:
            logger.error("Payload not provided.")
            raise Exception()
                
    except Exception:
        logger.error("Report generator module was not executed.")
        raise Exception(f"Provided Payload is invalid.")
    
    # Getting current date and last business day to check if the process needs to be run for the execution date or not.
    current_datetime = datetime.today()
    last_business_day = get_last_business_day_in_month(current_datetime.year, current_datetime.month)

    # Validation day of run based on the type of report. (We are also able to trigger this validation by specifying Test payload (send_to_vpos == False))
    if (current_datetime.day in [14, last_business_day] and report_type == "BIWEEKLY") or (current_datetime.day == 1 and report_type == "EOMPCARD") or (send_to_vpos == False and auxiliary_email):
        logger.info(f"""Dates were validated. Running process for the {report_type} report for day {current_datetime.day} and month {current_datetime.month}.""")

        logger.info("Creating folder structure and removing existing files if exists")
        create_folders()
        delete_existing_files()
        
        logger.info("Downloading Input files through web automation")
        web_automation = WebAutomation()
        
        # Downloading both Monthly and Biweekly reports from BOA (Bank of America).
        web_automation.download_reports(running_for_previous_month=False)

        # Connecting to BlueOcean Sharepoint to upload and synchronize files.
        #logger.info("Successfully connected to SharePoint.")
        #logger.info(f"{report_type} report selected. Starting the process.")

        #download_path, processed_path = None, None

        if report_type.upper() == "BIWEEKLY":
            logger.info("Processing BIWEEKLY report generation")
            #biweekly_report_generator = BIWeeklyReportGenerator(sharepoint_client)
            #logger.info("Creating reports..")
            #sharepoint_client = None
            biweekly_report_generator = BIWeeklyReportGenerator() # No need to pass sharepoint
            biweekly_report_generator.create_reports(web_automation.processed_directory)

            #logger.info("Uploading reports..")
            #download_path, processed_path = biweekly_report_generator.upload_reports(web_automation.download_directory, web_automation.processed_directory)

            #logger.info("Sending email reports..")
            #biweekly_report_generator.send_report_emails(auxiliary_email, web_automation.processed_directory)

        elif report_type.upper() == "EOMPCARD":
            logger.info("Processing EOM PCARD report generation")
            eom_report_generator = EOMPcardReportGenerator() #eom_report_generator = EOMPcardReportGenerator(sharepoint_client)
            logger.info("Creating reports..")
            eom_report_generator.create_reports()
            
            

            #logger.info("Uploading reports..")
            #download_path, processed_path = eom_report_generator.upload_reports(web_automation.download_directory, web_automation.processed_directory)

            #logger.info("Sending email reports..")
            #eom_report_generator.send_report_emails(auxiliary_email, web_automation.processed_directory)

        else:
            logger.error("Report generator module was not executed.")
            raise Exception(f"Provided ReportType is invalid: {report_type}")

        
        #reportpaths = [download_path.__str__(), processed_path.__str__()]
        #reportpaths = [path.replace('\\', '/') for path in reportpaths]

        # Trigger Logic App to transfer output to Highgate's SharePoint site.
        #payload = {"reportpaths": reportpaths}
        #logger.info(f"Transfer Logic App - Payload: {payload}")
        #data = json.dumps(payload)
        #headers = {
        #    "Content-Type": "application/json"
        #}
        #response = requests.post(url=config.TRANSFER_LOGIC_URL, headers=headers, data=data)
        #logger.info(f"Transfer Logic App - Response: {response}")

        delete_existing_files()

        logger.info("Automation process was successfully.")
    else:
        logger.info("The process was not run. The days are not matched as expected. Or the payload contins invalid flags.")
        #logger.info("Biweekly report will go out on the 14th and last business day of each month at 8:00 am EST.")
        #logger.info("Monthly report will run at 1:30 pm EST on the 1st of each month.")

if __name__ == '__main__':
    main()
