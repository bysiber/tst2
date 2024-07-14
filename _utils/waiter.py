from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webelement import WebElement


class Waiter:
        
    def __init__(self, driver):
        self.driver = driver


    def wait_to_be_clickable(self, element, timeout: int = 7):
        if isinstance(element, tuple):
            return WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable(element)
            )
        elif isinstance(element, WebElement):
            return WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable(element)
            )
        
        
    def wait_to_be_present(self, by: tuple, timeout: int = 7):
        return WebDriverWait(self.driver, timeout).until(
        EC.presence_of_element_located(by)
    )
        
        
    def wait_to_be_visible(self, by: tuple, timeout: int = 7):
        return WebDriverWait(self.driver, timeout).until(
        EC.visibility_of_element_located(by)
    )
        
    
    def wait_to_be_present_within_element(self, parent_element: WebElement, by: tuple, timeout: int = 7):
        return WebDriverWait(parent_element, timeout).until(
        EC.presence_of_element_located(by)
    )
        
        
    def wait_until_not_visible(self, by: tuple, timeout: int = 7):
        return WebDriverWait(self.driver, timeout).until(
            EC.invisibility_of_element(by)
        )
    
    
    def wait_until_not_present(self, by: tuple, timeout: int = 7):
        return WebDriverWait(self.driver, timeout).until(
            EC.staleness_of(by)
        )
        

    def wait_until_text_disappears(self, text, timeout=30):
            WebDriverWait(self.driver, timeout).until(
            EC.invisibility_of_element_located((By.XPATH, f"//*[contains(text(), '{text}')]")))