from selenium import webdriver
from selenium.webdriver.common.keys import Keys

driver = webdriver.Firefox()

# Navegar a la página de inicio de sesión
driver.get("https://sistemaspeiqaa.fincomun.com.mx/spei/login.jsp")


#driver.quit()