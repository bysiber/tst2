import os
from pathlib import Path


TRANSFER_LOGIC_URL = ""

class LoginConfiguration:
    LOGIN_URL: str = ""
    LOGIN_EMAIL: str = ""
    LOGIN_NAME: str = ""
    LOGIN_PASSWORD: str = ""


class ReportConfiguration:  
    # REPORTS_TO_RUN: list = ["01 - EOM Pcard Report"]
    REPORTS_TO_RUN: list = ["01 - EOM Pcard Report", "02 - BI-WeeklyPcard Outstanding - VPO_VPF"]
    OUTPUT_REPORTS: list = ["EOM Pcard Report - FS & SS", "Full Service EOM Report", "Select Service EOM Report", "Suspended Hotels Report"]


class Emails:
    eom_pcard_default_email = ''


class FileNames:
    rules_master_filename = Path('Masterfile-Rules.xlsx')


class Paths:
    resources_filepath = Path(os.getcwd(), 'boa_pcard_reporting/resources')
    rules_filepath = Path(resources_filepath, 'rules')
    downloads_filepath = Path(resources_filepath, 'downloads')
    processed_filepath = Path(resources_filepath, 'processed')
    
    sp_root = Path("02 - BoA PCard Reporting")
    sp_rules_master_filepath = Path(sp_root, "01 - Rules", FileNames.rules_master_filename)
    sp_processed_filepath = Path(sp_root, "02 - Reports")


class Credentials:
    sp_email: str = ''
    sp_password: str = ''
    sp_site_name: str = ''
    send_email_call_url = ""
