import os
import time
from selenium.webdriver import Chrome, ChromeOptions
from selenium.webdriver.chrome.service import Service
import config as config
from po.create_report_page import *
from po.home_page import HomePage
from po.login_page import LoginPage
from po.reports_page import ReportsPage
from po.completed_reports import CompletedReportsPage
from _utils import helper
from _utils.logger import logger
from _utils.waiter import Waiter
from pathlib import Path


class WebAutomation:
    driver: Chrome
    waiter: Waiter
    login_page: LoginPage
    home_page: HomePage
    reports_page: ReportsPage
    create_report_page: CreateReportPage
    completed_reports_page: CompletedReportsPage
    download_directory: str = config.Paths.downloads_filepath.__str__()
    processed_directory: str = config.Paths.processed_filepath.__str__()
    
    def initialize_driver(self):
        chrome_options = ChromeOptions()
        logger.info(f"Browser download directory set to: {self.download_directory}")
            
        chrome_options.add_argument(f'--download.default_directory={self.download_directory}')
        chrome_options.add_argument('--disable-gpu')
        
        prefs = {
            'download.default_directory': self.download_directory,
            'download.prompt_for_download': False,
            'download.directory_upgrade': True,
            'safebrowsing.enabled': False
        }
        chrome_options.add_experimental_option('prefs', prefs)
        service = Service()
        self.driver = Chrome(service=service, options=chrome_options)
        
    def open_driver_and_initialize_page_objects(self):
        self.driver.get(config.LoginConfiguration.LOGIN_URL)
        self.driver.maximize_window()
        self.waiter = Waiter(self.driver)

        self.login_page = LoginPage(self.driver, self.waiter)
        self.home_page = HomePage(self.driver, self.waiter)
        self.reports_page = ReportsPage(self.driver, self.waiter)
        self.create_report_page = CreateReportPage(self.driver, self.waiter)
        self.completed_reports_page = CompletedReportsPage(self.driver, self.waiter)
        
    def download_reports(self, running_for_previous_month: bool):
        logger.info("Initializing Driver in order to star the automated process.")
        self.initialize_driver()

        logger.info("Opening Driver to download all the corresponding reports.")
        self.open_driver_and_initialize_page_objects()

        logger.info("Executing automated flow to download desired reports.")
        self.download_reports_from_page(running_for_previous_month)

        logger.info("Closing Driver after downloading all the corresponding reports.")
        self.close_driver()

    def download_reports_from_page(self, running_for_previous_month: bool):
        self.login_page.login(email=config.LoginConfiguration.LOGIN_EMAIL,
                              login_name=config.LoginConfiguration.LOGIN_NAME,
                              password=config.LoginConfiguration.LOGIN_PASSWORD)

        for report in config.ReportConfiguration.REPORTS_TO_RUN:
            logger.info(F"Downloading report: {report}")
            
            self.home_page.click_section("Reports", "Template Library")
            self.reports_page.click_on_template(report, "Modify / Run")
            self.create_report_page.submit_report(helper.get_report_index(report), running_for_previous_month)
            self.completed_reports_page.wait_and_download_last_report()
            self.wait_file_downloaded(report=report, timeout=30)
            logger.info("Report downloaded successfully.")

    def wait_file_downloaded(self, report: str, timeout: int) -> None:
        report_filepath = Path(self.download_directory, report + '.xls')
        if not report_filepath.exists():
            for i in range(timeout):
                time.sleep(1)
                if report_filepath.name in os.listdir(report_filepath.parent):
                    return
            raise FileNotFoundError(f'File {report} was not downloaded for some reason.')

    def close_driver(self):
        self.driver.close()
        self.driver.quit()
