
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import DesiredCapabilities
from time import sleep
import urllib.parse
import pandas as pd 
import pymsgbox
driver = None
Link = "https://web.whatsapp.com/"
wait = None
missed=[]
def whatsapp_login():
    global wait, driver, Link
    chrome_options = Options()
    chrome_options.add_argument("user-data-dir=TempWhatsapp") 
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)
    driver.maximize_window()
    driver.get(Link)
    




def send_message_to_unsavaed_contact(number,msg):
    # Reference : https://faq.whatsapp.com/en/android/26000030/
    params = {'phone': str(number), 'text': str(msg)}
    end = urllib.parse.urlencode(params)
    final_url = Link + 'send?' + end
    return final_url

def GenerateMessage(msg,orderMsg,noOfvars,row):
	try:
		orderMsg=orderMsg.split(',')
	except:
		orderMsg=list(orderMsg)
	final_msg=msg
	for i in range(1,noOfvars+1):		
		final_msg=final_msg.replace('"VAR{}"'.format(i),str(row[orderMsg[i-1]]))
	return final_msg


#df is temp dataframe,current is the department,colList is the list of departments
def Send(df,msg,orderMsg,noOfvars,ContactNumber):
	global driver,missed
	for index, row in df.iterrows():
		try:
			no=str(row[ContactNumber])
			temp_msg=''
			temp_msg=GenerateMessage(msg,orderMsg,int(noOfvars),row)
			if(temp_msg==''):
				missed.append(row['Name'])
				continue
			print(temp_msg)
			driver.get(send_message_to_unsavaed_contact('91'+no,temp_msg))

			wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[4]/div/footer/div[1]/div[3]/button/span'))).click()
			driver.execute_script("window.onbeforeunload = function() {};")
		except:
			missed.append(row['Name'])
	if(len(missed)>0):
		print('missedlist',missed)
			



def Load_excel(df,msg,orderMsg,noOfvars,ContactNumber):
	global wait, driver, Link
	check=pymsgbox.confirm(text='Click Ok to Send', title='Confirm', buttons=['OK', 'Cancel'])
	if(check=='OK'):
		options = Options()
		options.add_argument("user-data-dir=TempWhatsapp")
		# chrome_options.add_argument('--headless')
		options.add_argument("--window-size=1920,1080");
		options.add_argument("--disable-gpu");
		options.add_argument("--disable-extensions");
		# options.setExperimentalOption("useAutomationExtension", false);
		options.add_argument("'--proxy-server=direct://'");
		options.add_argument("--proxy-bypass-list=*");
		options.add_argument("--start-maximized");
		options.add_argument("--headless");
		driver = webdriver.Chrome(options=options)
		wait = WebDriverWait(driver, 20)
		driver.maximize_window()
		driver.get(Link)
		Send(df,msg,orderMsg,noOfvars,ContactNumber)
	else:
		pass

	



# Load_excel('Hockey','hey "VAR1","VAR2" "VAR3"','Name,Department,Present Year',3)