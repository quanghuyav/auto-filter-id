from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import time
import openpyxl

def get_value_excel(filename, cellname):
    wb = openpyxl.load_workbook(filename)
    Sheet1 = wb['Sheet1']
    wb.close()
    return Sheet1[cellname].value
	

def update_value_excel(filename, cellname, value):
    wb = openpyxl.load_workbook(filename)
    Sheet1 = wb['Sheet1']
    Sheet1[cellname].value = value
    wb.close()
    wb.save(filename)
filename='data.xlsx'
chrome_options = webdriver.ChromeOptions()
prefs = {
	"profile.managed_default_content_settings.images": 2
}
chrome_options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome('./chromedriver', options = chrome_options)
driver.get('https://id.atpsoftware.vn/')
for i_row in range(1,696):
	link_facebook="%s%s"%("A",i_row)
	id_facebook="%s%s"%("B",i_row)
	account=get_value_excel(filename, link_facebook)
	

	
	driver.find_element_by_css_selector('input[name="linkCheckUid"').clear()
	input_search=driver.find_element_by_css_selector('input[name="linkCheckUid"')
	input_search.send_keys(account)
	driver.find_element_by_css_selector('#menu1 > form > div > div > button').click()
	try:
		a = driver.find_element_by_css_selector('#menu1 > textarea').text
	except:
		a = ("x") 
	print(a)
	update_value_excel(filename, id_facebook, a)
driver.close()
