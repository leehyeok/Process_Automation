import email
from tabula import read_pdf
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
from os.path import isfile, join
import os
from os import listdir
import datetime
import smtplib
import ssl
from email.message import EmailMessage

def typing(path, string):
    global driver
    typ = driver.find_element(by=By.XPATH, value=path).send_keys(string)

def wclick(path, method):
    global driver
    Button = None
    i=0
    while Button == None and i <=3:
        if method == 1:
            wait = WebDriverWait(driver, 30)
            Button = wait.until(EC.visibility_of_element_located((By.XPATH,path)))
        
        else:
            wait = WebDriverWait(driver, 30)
            Button = wait.until(EC.element_to_be_clickable((By.XPATH,path)))
        
        i+=1
  
    try: 
        Button.click()

    except: 
        print('Timeouterror {}'.format(path))
        time.sleep(5)

def left(s, amount): 
    return s[:amount]

def right(s, amount): 
    return s[len(s)-amount:]

startime = datetime.datetime.now()
EMAIL_ADDRESS = os.environ.get('EMAIL_ADDRESS') # my gmail, I'm using environment variable
password = os.environ.get('COMPLEX_PASS') # password for hydroone
EMAIL_PASSWORD = os.environ.get('APP_PASSWORD') # password for gmail
path_location = 'C:/Users/leehy/Downloads/' # input your path here, it will be different from mine

#Downloading sample invoice from web
print("Downloading bill from hydro-one")
chrome_options = Options()
chrome_options.add_argument('--log-level=3')
driver = webdriver.Chrome('C:/Users/leehy/code/chromedriver.exe', options=chrome_options)
driver.maximize_window()
driver.get('https://www.hydroone.com/login?EC=')
wclick('//*[@id="username"]',1)
typing('//*[@id="username"]',EMAIL_ADDRESS)
wclick('//*[@id="password"]',1)
typing('//*[@id="password"]',password)
wclick('//*[@id="btnSubmit"]',1)
wclick('//*[@id="ctl00_ctl50_g_105f7d2a_7b1f_4ebf_90e2_76ccf434e9e6"]/div/enhanced-account-profile/div[1]/div/div[2]/div[1]/div/div[3]/a',1)
time.sleep(3)
driver.quit()

# retrieving the latest PDF file from download
download_dir = glob.glob('{}*.pdf'.format(path_location)) 
pdf_download = max(download_dir, key=os.path.getctime)

# parsing first page
print('Parsing PDF document')
Uppertable = read_pdf(pdf_download,pages=1,pandas_options={'header':None},area=[45,45,470,590])
Owed_amount = Uppertable[0][0][6]
Usage = Uppertable[0][1][7]
Due_date = datetime.datetime.strptime(Uppertable[0][1][6]+right(Uppertable[0][1][8],4), '%b %d,%Y').date()
Bill_date = datetime.datetime.strptime(right(Uppertable[0][1][1],len(Uppertable[0][1][1])-Uppertable[0][1][1].find(':')-2), '%B %d, %Y').date()
Vs_Prev = right(Uppertable[0][0][15],len(Uppertable[0][0][15])-(Uppertable[0][0][15].find('has')+4))+' '+left(Uppertable[0][0][16],Uppertable[0][0][16].find('%')+1)

# parsing second page
Lowertable = read_pdf(pdf_download,pages=2,pandas_options={'header':None},area=[40,38,449,595])
Prev_balance = Lowertable[0][1][0]
Curr_balance = Lowertable[0][1][3]
Electricity_amount = Lowertable[0][1][8]
Delivery_amount = Lowertable[0][1][16]
Reg_charge = Lowertable[0][1][23]
OnESP = Lowertable[0][1][28]
Tax_amount = Lowertable[0][1][29]
Tax_type = left(Lowertable[0][0][29],3)
On_peak_usg = round(float(Lowertable[0][3][9].split()[1]),2)
On_peak_rate = Lowertable[0][3][9].split()[2]
On_peak_cost = Lowertable[0][3][9].split()[3]
Mid_peak_usg = round(float(Lowertable[0][3][11].split()[1]),2)
Mid_peak_rate = Lowertable[0][3][11].split()[2]
Mid_peak_cost = Lowertable[0][3][11].split()[3]
Off_peak_usg = round(float(Lowertable[0][3][12].split()[1]),2)
Off_peak_rate = Lowertable[0][3][12].split()[2]
Off_peak_cost = Lowertable[0][3][12].split()[3]
Rebate = Lowertable[0][1][30]
sum_of_usage = On_peak_usg+Mid_peak_usg+Off_peak_usg

# converting parsed information from PDF into pandas dataframe and saving it as excel
print('Converting parsed information into Pandas dataframe')
df = pd.DataFrame([Bill_date, Due_date, Owed_amount, Prev_balance, Curr_balance, Electricity_amount, On_peak_cost, Mid_peak_cost,Off_peak_cost,Delivery_amount,Reg_charge,OnESP,Rebate,Tax_amount,Tax_type,Usage,On_peak_usg,Mid_peak_usg,Off_peak_usg,Vs_Prev]).T
df.columns = ['Bill Date','Due Date','Total Owed','Previous Balance','Current Balance',	'Electricity Amount','On-peak amount','Mid-peak amount','Off-peak amount','Delivery amount','Regulatory charges','OESP','Ontario Electricity Rebate','Tax amount','Tax Type','Total usage','On-peak usage','Mid-peak usage','Off-peak usage','VS last month']
print('Saving dataframe as Excel')
df.to_excel('{}bill.xlsx'.format(path_location),engine='xlsxwriter',index=False)

#Email send portion
print('Sending email to {}'.format(EMAIL_ADDRESS))
msg = EmailMessage()
msg['Subject'] = 'Bill overview'
msg['From'] = EMAIL_ADDRESS
msg['To'] = EMAIL_ADDRESS
msg.set_content('Bill due date: {}'.format(Due_date))

#Basic HTML for the template of the email
msg.add_alternative("""\
<!DOCTYPE html>
<html> 
    <body>
        <h2 style="color:Grey;">Bill Summary {} compared to last month</h2>
        <h3 style="color:Grey;">Invoice is due on : {}</h3>
        <h3 style="color:Grey;">Owed amount of : {}</h3>
                <table>
            <tr>
                <td style="border-bottom: 1px dashed grey; text-align: center; font-size: 14px; width:14em;"></td>
                <td style="border-bottom: 1px dashed grey; text-align: center; font-size: 14px; width:14em;">Amount</td>
            </tr>
            <tr>
                <td style="text-align: left; font-size: 14px; width:14em;">Electricity</td>
                <td style="text-align: center; font-size: 14px; width:14em;">{}</td>
                </tr>
            <tr>
                <td style="text-align: left; font-size: 14px; width:14em;">Delivery Charge</td>
                <td style="text-align: center; font-size: 14px; width:14em;">{}</td>
            </tr>
            <tr>
                <td style="text-align: left; font-size: 14px; width:14em;">Regulatory Charges</td>
                <td style="text-align: center; font-size: 14px; width:14em;">{}</td>
            </tr>
            <tr>
                <td style="text-align: left; font-size: 14px; width:14em;">OESP</td>
                <td style="text-align: center; font-size: 14px; width:14em;">{}</td>
            </tr>
            <tr>
                <td style="text-align: left; font-size: 14px; width:14em;">{}</td>
                <td style="text-align: center; font-size: 14px; width:14em;">{}</td>
            </tr>            
            <tr>
                <td style="text-align: left; font-size: 14px; width:14em;">Ontario Electricity Rebate</td>
                <td style="text-align: center; font-size: 14px; width:14em;">{}</td>
            </tr>       
            <tr>
                <td style="border-top: 1px dashed grey; text-align: left; font-size: 14px; width:14em;">Total</td>
                <td style="border-top: 1px dashed grey; text-align: center; font-size: 14px; width:14em;">{}</td>
            </tr>
        </table>
        <h3 style="color:Grey;">Total Usage was : {}</h3>
            <table>
            <tr>
                <td style="border-bottom: 1px dashed grey; text-align: center; font-size: 14px; width:7em;"></td>
                <td style="border-bottom: 1px dashed grey; text-align: center; font-size: 14px; width:7em;">Usage</td>
                <td style="border-bottom: 1px dashed grey; text-align: center; font-size: 14px; width:7em;">Cost</td>
                <td style="border-bottom: 1px dashed grey; text-align: center; font-size: 14px; width:7em;">% Usage</td>
                <td style="border-bottom: 1px dashed grey; text-align: center; font-size: 14px; width:7em;">% Cost</td>
            </tr>
            <tr>
                <td style="text-align: left; font-size: 14px; width:7em;">On-Peak</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
                </tr>
            <tr>
                <td style="text-align: left; font-size: 14px; width:7em;">Mid-Peak</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
            </tr>
            <tr>
                <td style="text-align: left; font-size: 14px; width:7em;">Off-Peak</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="text-align: center; font-size: 14px; width:7em;">{}</td>
            </tr>
            <tr>
                <td style="border-top: 1px dashed grey; text-align: left; font-size: 14px; width:7em;">Total</td>
                <td style="border-top: 1px dashed grey; text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="border-top: 1px dashed grey; text-align: center; font-size: 14px; width:7em;">{}</td>
                <td style="border-top: 1px dashed grey; text-align: center; font-size: 14px; width:7em;"></td>
                <td style="border-top: 1px dashed grey; text-align: center; font-size: 14px; width:7em;"></td>
                            </tr>
        </table>
    </body>

        

</html>
""".format(Vs_Prev,Due_date,Owed_amount,Electricity_amount,Delivery_amount,Reg_charge,OnESP,Tax_type,Tax_amount,Rebate,Curr_balance,Usage,On_peak_usg,On_peak_cost,str(round(On_peak_usg/sum_of_usage*100,2))+' %',str(round(float(On_peak_cost.strip('$'))/float(Electricity_amount.strip('$'))*100,2))+' %',Mid_peak_usg,Mid_peak_cost,str(round(Mid_peak_usg/sum_of_usage*100,2))+' %',str(round(float(Mid_peak_cost.strip('$'))/float(Electricity_amount.strip('$'))*100,2))+' %',Off_peak_usg,Off_peak_cost,str(round(Off_peak_usg/sum_of_usage*100,2))+' %',str(round(float(Off_peak_cost.strip('$'))/float(Electricity_amount.strip('$'))*100,2))+' %',sum_of_usage,Electricity_amount), subtype='html')

with open('{}bill.xlsx'.format(path_location),'rb') as f:
    file_data = f.read()
    ctype = 'application/octet-stream' 
    maintype, subtype = ctype.split('/',1)

msg.add_attachment(file_data,maintype=maintype,subtype=subtype,filename='bill.xlsx')

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    smtp.send_message(msg)

print(f'Execution Time: {datetime.datetime.now() - startime}')