from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from po.basic_page import BasicPage

class ReportsPage(BasicPage):
    
    TEMPLATE_LIST = "//div[@id='sharedTemplateTable_wrapper']//tbody//tr[@role='row']//a[@class='link_menu_link preventClick' and text()='{template_name}']"
    TEMPLATE_ACTION = "//div[contains(@class, 'linkActionItem') and contains(text(), '{template_action}')]"

    
    def __init__(self, driver, waiter):
        super().__init__(driver, waiter)

    def click_on_template(self, template_name, template_action):
        template_element = self.driver.find_element(By.XPATH, self.TEMPLATE_LIST.format(template_name=template_name))
        template_element.click()
        
        templates_actions = template_element.find_elements(By.XPATH, self.TEMPLATE_ACTION.format(template_action=template_action))
        
        # TODO - Correct this cycle, this should not be necessary. We have to filter the element from XPATH.
        for template_actions in templates_actions:
            try:
                run_menu_element = self.waiter.wait_to_be_clickable(template_actions, timeout=0.1)
                run_menu_element.click()
                break
            except TimeoutException:
                pass