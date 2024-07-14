from selenium.webdriver.common.by import By
from po.basic_page import BasicPage


class CompletedReportsPage(BasicPage):
    
    COMPLETED_REPORTS_LIST = (By.ID, 'completedReportsTable')
    ROWS_FROM_LIST = (By.XPATH, "//tr[@role='row']")
    COMPLETED_ICON = (By.XPATH, "//img[@src='/works/resources/dyn/93ef6b6d/themes/default/images/icons/small_success_icon.png']")
    XLS_DOWNLOAD = (By.XPATH, "//td[@class=' data_left']//a[@href and text()='XLS']")
    
    
    def __init__(self, driver, waiter):
        super().__init__(driver, waiter)


    def wait_and_download_last_report(self):
        self.waiter.wait_until_text_disappears(text="Awaiting Processing", timeout=45)
        # TODO - we can add here an assert to validate that the first row of this list has the same name of the submited report and the same day of the server.
        completed_reports_table = self.waiter.wait_to_be_present(self.COMPLETED_REPORTS_LIST)
        latest_completed_report = completed_reports_table.find_elements(*self.ROWS_FROM_LIST)[1]
        
        # TODO - Review this line.
        # self.waiter.wait_to_be_present_within_element(latest_completed_report, self.COMPLETED_ICON, timeout=10)
        
        xls_download_element = self.waiter.wait_to_be_present_within_element(latest_completed_report, self.XLS_DOWNLOAD, timeout=40)
        xls_download_element.click()
