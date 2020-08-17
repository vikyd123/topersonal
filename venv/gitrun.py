from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import os

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
driver = webdriver.Chrome(executable_path = ChromeDriverManager().install(), options=chrome_options)
driver.implicitly_wait(2)

driver.get('https://www.jiffyshirts.com')
driver.implicitly_wait(1)
print(driver.find_element_by_xpath('//*[@id="content"]/section/aside/aside/div[1]/div').text)
print(driver.find_element_by_xpath('//*[@id="content"]/section/aside/aside/div[2]/div').text)
print(driver.find_element_by_xpath('//*[@id="content"]/section/aside/aside/div[3]/div').text)
print(driver.find_element_by_xpath('//*[@id="content"]/section/aside/aside/div[4]/div').text)
driver.close()
os.system('"taskkill /im chromedriver.exe /f"')