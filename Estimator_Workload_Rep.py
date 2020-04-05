import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import os
import win32com.client as win32


options = webdriver.ChromeOptions()
prefs= {'safebrowsing.enabled': 'false', 'download.default_directory': 'P:\\Service Department\\Programs\\Estimator Workload Report'}
options.add_experimental_option("prefs", prefs)
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument('download.prompt_for_download = False')
options.add_argument("safebrowsing.enabled = False")

driver= webdriver.Chrome(options=options)


driver.get('https://dfeast2prweb2.dataforma.com/dflowslope/pages/security/LoginForm.action')




id_box = driver.find_element_by_name('servicecode')
id_box.send_keys("Service code here")
company_code = driver.find_element_by_name('B4')
company_code.click()

username_box= driver.find_element_by_name('j_username')
username_box.send_keys('Username Here')
pass_box=driver.find_element_by_name('j_password')
pass_box.send_keys('Password here')
login = driver.find_element_by_name('B3')
login.click()
time.sleep(3)
tab=driver.find_element_by_xpath('/html/body/div/div[3]/div[3]/div[1]/div[2]/div/div[2]/ul')
tab.find_element_by_xpath('//*[@id="mboard_workorder"]').click()
time.sleep(10)
cur_es_req = driver.find_element_by_xpath('/html/body/div/div[3]/div[3]/div[1]/div[2]/div/div[2]/div/div[2]/div/div/div/div[12]/div/div[1]/a')
cur_es_req.click()
driver.switch_to.window(driver.window_handles[-1])
time.sleep(5)

frame=driver.switch_to.frame('frame2')
#Testing used for elements in website
#testing_elements=driver.find_elements_by_xpath('//*[@name]')
#for ii in testing_elements:
    #print(ii.get_attribute('name'))
select_op= Select(driver.find_element_by_name("queryfunction1"))
select_op.select_by_visible_text('export')
apply_button=driver.find_element_by_name("B5")
apply_button.click()
driver.switch_to.window(driver.window_handles[-1])
export_select=driver.find_element_by_xpath('//*[@id="export-form"]/div[3]/div[2]/div[1]')
export_select.click()

export_button=driver.find_element_by_class_name('export-btn')
export_button.click()

xmlfile= "P:\\Service Department\\Programs\\Estimator Workload Report\\KPC_Estimator_Work_Load_Report.xml"
def converter():
    xlapp = win32.Dispatch('Excel.Application')
    xlapp.DisplayAlerts = False
    xlapp.Visible = False

    xlmac = xlapp.Workbooks.Open('excelsheetstopdf.xlsm')
    xlbook = xlapp.Workbooks.Open(xmlfile)
    xlbook.Worksheets("Estimator Work Load Report").Select()
    xlapp.Application.Run("excelsheetstopdf.xlsm!PDFActiveSheetNoPrompt")

    xlbook.Save()
    xlbook.Close()
    xlapp.Quit()

    del xlbook
    del xlapp
time.sleep(10)
driver.quit()
converter()


attachment = 'Path\\Estimator Work Load Report.pdf'

contacts = """
Emails for whoever you want to send it to."""

def Emailer(text1, subject, recipient, cc):
    import win32com.client as win32

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Cc = cc
    mail.Subject = subject
    mail.HtmlBody = text1
    mail.Attachments.Add(attachment)
    mail.Send()
time.sleep(5)

Emailer("""<p>Good Afternoon!
        </p><p>Attached you will find the Estimator Workload Report.
        """,
        "Daily Estimate Requested Report for {0}".format(datetime.today().strftime('%m/%d/%Y')),
        contacts, "")

os.remove(attachment)
os.remove(xmlfile)




