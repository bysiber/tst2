from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains

class BasicPage:
    
    TAB_LIST = "//ul[@class='sf-menu sf-js-enabled']//span[text()='{tab_name}']"
    LIST_OPTION = "//a[text()='{section_name}']"
    

    def __init__(self, driver, waiter):
        self.driver = driver
        self.waiter = waiter

    def click_tab(self, tab_name):
        self.TAB_LIST.format(tab_name=tab_name)
        tab_element = self.driver.find_element(By.XPATH, tab_name)
        
        tab_element.click()
        
    def hover_tab(self, tab_name):
        tab_element = self.driver.find_element(By.XPATH, self.TAB_LIST.format(tab_name=tab_name))
        
        actions = ActionChains(self.driver)
        actions.move_to_element(tab_element).perform()

    def click_section(self, tab_name, section_name):
        self.hover_tab(tab_name=tab_name)
        final_xpath = (f"{self.TAB_LIST}/ancestor::li{self.LIST_OPTION}").format(tab_name=tab_name, section_name=section_name)
        option_element = self.driver.find_element(By.XPATH, final_xpath)
    
        option_element.click()