from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from po.basic_page import BasicPage

class CreateReportPage(BasicPage):
    
    LOADER = (By.ID, 'busyIndicator')
    SUBMIT_BUTTON = (By.XPATH, "//input[@name='submit' and @value='Submit Report']")
    CALENDAR_ICON = (By.ID, "cfpDATE_RANGEdtPostDate")
    CALENDAR_DIALOG = (By.XPATH, "//div[@class='datePickerBox ui-dialog-content ui-widget-content']")
    CYCLE_TO_DATE_RADIO_BUTTON = (By.ID, "ctd")
    PREVIOUS_MONTH_RADIO_BUTTON = (By.ID, "pcm")
    OK_BUTTON = (By.XPATH, "//button[@class='ui-button ui-corner-all ui-widget' and text()='OK' and following-sibling::button]")


    
    def __init__(self, driver, waiter):
        super().__init__(driver, waiter)
        
    
    def submit_report(self, report_index: str, running_for_previous_month: bool):
        try:
            self.waiter.wait_until_not_present(self.driver.find_element(*self.LOADER), timeout=15)
        except NoSuchElementException:
            pass
        
        if report_index == "01":
            pass
        elif report_index == "02":
            calendar_element = self.waiter.wait_to_be_clickable(self.CALENDAR_ICON)
            calendar_element.click()
            
            self.waiter.wait_to_be_visible(self.CALENDAR_DIALOG)

            if running_for_previous_month:
                previous_month_element = self.waiter.wait_to_be_clickable(self.PREVIOUS_MONTH_RADIO_BUTTON)
                previous_month_element.click()
            else:
                cycle_to_date_element = self.waiter.wait_to_be_clickable(self.CYCLE_TO_DATE_RADIO_BUTTON)
                cycle_to_date_element.click()
            
            ok_button_element = self.waiter.wait_to_be_clickable(self.OK_BUTTON)
            ok_button_element.click()
            
        else:
            return None
        
        submit_button_element = self.waiter.wait_to_be_clickable(self.SUBMIT_BUTTON)
        submit_button_element.click()