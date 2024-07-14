from selenium.webdriver.common.by import By

class LoginPage:

    EMAIL_INPUT = (By.ID, 'email')
    LOGIN_NAME_INPUT = (By.ID, 'username')
    PASSWORD_INPUT = (By.ID, 'password')
    LOGIN_BUTTON = (By.NAME, 'submit')

    def __init__(self, driver, waiter):
        self.driver = driver
        self.waiter = waiter

    def set_email(self, email):
        email_element = self.driver.find_element(*self.EMAIL_INPUT)
        email_element.clear()
        email_element.send_keys(email)

    def set_login_name(self, login_name):
        login_name_element = self.driver.find_element(*self.LOGIN_NAME_INPUT)
        login_name_element.clear()
        login_name_element.send_keys(login_name)

    def set_password(self, password):
        password_element = self.driver.find_element(*self.PASSWORD_INPUT)
        password_element.clear()
        password_element.send_keys(password)

    def click_submit(self):
        login_button_element = self.driver.find_element(*self.LOGIN_BUTTON)
        login_button_element.click()

    def login(self, email, login_name, password):
        self.set_email(email)
        self.set_login_name(login_name)
        self.set_password(password)
        self.click_submit()
