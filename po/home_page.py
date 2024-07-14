from selenium.webdriver.common.by import By
from po.basic_page import BasicPage

class HomePage(BasicPage):
    
    def __init__(self, driver, waiter):
        super().__init__(driver, waiter)
