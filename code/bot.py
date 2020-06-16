#Auto MAKAUT report fill bot
#@Soumojit V1.0

from selenium import webdriver 
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait 
import time
import xlrd
import pyfiglet
from colorama import init, Fore, Back, Style

init(convert=True)

def autofillup(week, sub,topic, platform, takenby ,link, duration, note, a_rev, a_sub, test):
    report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div/div[1]/div/ul/li[12]/a')
    report.click()
    time.sleep(2)
    week_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[2]/div[1]/div/select')
    semester_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[2]/div[2]/div/select')
    course_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[2]/div[3]/div/div/button')
    subject_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[3]/div[1]/div/select')
    topic_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[3]/div[2]/div/textarea')
    platform_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[3]/div[3]/div/textarea')
    takenby_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[4]/div[1]/div/div/button')
    date_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[4]/div[2]/div/input')
    link_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[4]/div[3]/div/input')
    duration_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[4]/div[4]/div/input')
    note_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[5]/div[1]/div/input')
    a_rev_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[5]/div[2]/div/textarea')
    a_sub_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[5]/div[3]/div/input')
    test_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[6]/div[1]/div/input')
    daily_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[6]/div[2]/div/textarea')
    remark_report=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[6]/div[3]/div/textarea')
    week_report.send_keys(week)
    semester_report.send_keys(sem)
    course_report.send_keys(dep+'\n')
    time.sleep(2)
    subject_report.send_keys(sub)
    topic_report.send_keys(topic)
    platform_report.send_keys(platform)
    takenby_report.send_keys(takenby +'\n')
    link_report.send_keys(link)
    duration_report.send_keys(duration)
    note_report.send_keys(note)
    a_rev_report.send_keys(a_rev)
    a_sub_report.send_keys(a_sub)
    test_report.send_keys(test)
    daily_report.send_keys(daily_act)
    remark_report.send_keys(remark)
    
print(Fore.CYAN+Style.BRIGHT+pyfiglet.figlet_format("MAKAUT",font="alphabet"),end="")
print(pyfiglet.figlet_format("Auto Report Bot",font="contessa"))
time.sleep(2)
print(Fore.YELLOW+pyfiglet.figlet_format("@Soumojit",font="big"))

print(Fore.GREEN+"----V 1.0 Changelog---- \n ")
print("* Made single exe file")
print("* Added support for auto user login")
print("* Added support for all semester and departments \n")
print("*** Make sure chrome driver is in same directory *** \n")
print("*** Make sure excel spread sheet is in same directory *** \n\n")


user_id=input("Enter Roll number: ")
passwd=input("Enter password: ")
sem=input("Semester: ")
dep=input("Department: ")
sub=input("\n\n Subject Name: ")
takenby=input("Taken by:")
daily_act=input("Daily Activity: ")
remark=input("Remarks: ")
loc=input("Name of excel spreadsheet: ")
n=input("No of entry: ")
n=int(n)
#sem=input("Semester: ");
#course=input("Course:");
print("Running Bot")

  
browser = webdriver.Chrome("C:\\chromedriver_win32\\chromedriver.exe")
browser.maximize_window()

wb = xlrd.open_workbook(loc)   
sheet = wb.sheet_by_name('Sheet1')
print("Logging in.........")
browser.get("https://makaut1.ucanapply.com/smartexam/public/")
student_login = browser.find_element_by_xpath('/html/body/div[2]/div/div[2]/div[1]/div/div/div[1]/a/div/div[2]/div')
student_login.click()
time.sleep(3)
roll_form = browser.find_element_by_xpath('/html/body/div[4]/div/div/div[2]/div/form/div[1]/div/div/input')
pass_form = browser.find_element_by_xpath('/html/body/div[4]/div/div/div[2]/div/form/div[2]/div/div/input')
roll_form.send_keys(user_id)
pass_form.send_keys(passwd)
submit=browser.find_element_by_xpath('/html/body/div[4]/div/div/div[2]/div/form/div[4]/div/a')
submit.click()
print("starting autofill")
time.sleep(2)
i=0
while(i<n):
    week=sheet.cell_value(i, 1)
    topic=sheet.cell_value(i, 3)
    platform=sheet.cell_value(i, 4)
    link=sheet.cell_value(i, 6)
    duration=str(sheet.cell_value(i, 7))
    note=sheet.cell_value(i, 8)
    a_rev=sheet.cell_value(i, 9)
    a_sub=sheet.cell_value(i, 10)
    test=sheet.cell_value(i, 11)
    autofillup(week, sub,topic, platform, takenby ,link, duration, note, a_rev, a_sub, test)
    input("Enter date and time in form and press any key........")
    time.sleep(1)
    submit_form=browser.find_element_by_xpath('/html/body/div[4]/div/div/div/div[2]/div/div/form/div/div[7]/div/input')
    submit_form.click()
    time.sleep(3)
    browser.get("https://makaut1.ucanapply.com/smartexam/public/student/dashboard")
    i=i+1
    print(i)
    print("complete\n\n")
    
    


