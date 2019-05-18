#coding=UTF-8

import os
import csv
import time
import xlrd
import xlwt
import requests
import linecache
from xlutils.copy import copy
from selenium import webdriver
from requests_toolbelt import MultipartEncoder
from selenium.webdriver.support.select import Select
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, date, timedelta

#定义各种变量（防止忘记）
#Absence_Number = "NULL"
#Absence_Name = "NULL"
#Absence_Gender = "NULL"
#Absence_Age = "NULL"
#Absence_Dormitory = "NULL"
#Absence_Bed = "NULL"
#Absence_PhoneNumber = "NULL"
#Absence_Reason = "NULL"
#The_Date_of_symptoms = "NULL"
#The_Date_of_Leaving = "NULL"
#The_Date_of_Return = "NULL"

#定义写入Excel表的格式
ExcelStyle = xlwt.XFStyle()
alignment = xlwt.Alignment()
alignment.horz = xlwt.Alignment.HORZ_CENTER
alignment.vert = xlwt.Alignment.VERT_CENTER
ExcelStyle.alignment = alignment

#上送文件
def UpLoad():
    print('请选择上送方式：')
    print('[A]直接发送Post请求登录后Send_Key(默认)')
    print('[B]使用无头浏览器模拟点击访问(需要安装Fiefox以及对应WebDriver)')
    Flag = str(input())
    Grade = linecache.getline('Config.txt', 1)
    Class = linecache.getline('Config.txt', 2)
    FilePath = linecache.getline('Config.txt', 3)
    Grade = Grade[Grade.find(':')+1:-1]
    Class = Class[Class.find(':')+1:-1]
    FilePath = FilePath[FilePath.find(':')+1:-1] 
    if Flag == 'B':
        #无图形化界面
        firefox_options = webdriver.FirefoxOptions()
        firefox_options.set_headless()
        driver = webdriver.Firefox(firefox_options=firefox_options)
        #图形化界面
        #driver = webdriver.Firefox()
        driver.get("http://10.10.88.8/")
        driver.implicitly_wait(20)
        #选择下拉框中的年级，班级，填充密码
        s = Select(driver.find_element_by_name("select1"))
        s.select_by_value(Grade)
        t = Select(driver.find_element_by_name("select2"))
        t.select_by_value(Class)
        driver.find_element_by_name("UserPwd1").send_keys("56789")
        driver.find_element_by_name("submit").click()
        #选择文件并上传
        driver.find_element_by_name('fAtt').send_keys(FilePath)
        driver.find_element_by_name("bSubmit").click()
        os.system('taskkill /f /IM Firefox.exe')
        os.system('taskkill /f /IM geckodriver.exe')
    else: 
        ssion = requests.session()
        Headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.99 Safari/537.36"}
        Login_Data = {"hPageName":"LoginApply","select1":Grade,"select2":Class,"UserPwd1":"56789","submit":"%B5%C7%C2%BC"}
        ssion.post("http://10.10.88.8/", data = Login_Data, headers = Headers)
        Response = ssion.get("http://10.10.88.8/main.asp", headers = Headers)
        if 'fAtt' in Response.text:
            print('登录成功')
            #模拟multipart/form-data格式
            UpLoadFiles = {'hPageName':(None, 'AttachmentAdd'),
                           'bSubmit':(None, '%EF%BF%BD%EF%BF%BD+++%EF%BF%BD%EF%BF%BD'),
                           'fAtt':('UpLoad.xls', open(FilePath, 'rb'), 'application/vnd.ms-excel')}
            UpLoad_Resopnse = ssion.post('http://10.10.88.8/weblog_attachmentadd2.asp', files = UpLoadFiles)
        else:
            print('出现未知错误，即将返回主菜单')
            time.sleep(0.5)
            return
        
    print('上传完成，写入日志后将回到主菜单')
    log = open("log.txt", "a+", encoding = "UTF-8")
    log.write(time.strftime('[%Y-%m-%d %H:%M:%S]',time.localtime(time.time()))+' 上送完成\n')
    log.close()
    time.sleep(0.5)

#填充数据
def WriteExcel():
    #读取配置文件信息
    Grade = linecache.getline('Config.txt', 1)
    Class = linecache.getline('Config.txt', 2)
    FilePath = linecache.getline('Config.txt', 3)
    Grade = Grade[Grade.find(':')+1:-1]
    Class = Class[Class.find(':')+1:-1]
    FilePath = FilePath[FilePath.find(':')+1:-1]
    #载入表格提取信息
    WorkBook = xlrd.open_workbook(FilePath, formatting_info = True)
    Sheets = WorkBook.sheet_names()
    WorkSheet = WorkBook.sheet_by_name(Sheets[0])
    Rows_Old = WorkSheet.nrows
    #从原表复制一个工作簿准备写入
    New_WorkBook = copy(WorkBook)
    New_WorkSheet = New_WorkBook.get_sheet(0)
    #数据写入操作
    print('请输入今天缺勤人数：')
    Absence_Number = str(input())
    New_WorkSheet.write_merge(Rows_Old, Rows_Old+1, 0, 0, time.strftime('%m-%d',time.localtime(time.time())), ExcelStyle)
    New_WorkSheet.write_merge(Rows_Old, Rows_Old+1, 1, 1, Absence_Number, ExcelStyle)
    New_WorkSheet.write_merge(Rows_Old, Rows_Old+1, 2, 2, Absence_Number, ExcelStyle)
    if Absence_Number != '0':
        K = 1
        for i in range(1, int(Absence_Number)+1):
            print('请输入第'+str(i)+'位缺勤同学的学号(例：0101)：')
            StudentNumber = str(input())
            with open('data.csv', 'r', encoding = "UTF-8") as csvfile:
                Reader = csv.DictReader(csvfile)
                for Row in Reader:
                    if Row['学号'] == StudentNumber:
                        print('目前正在输入的是'+Row.get('姓名')+'的相关信息')
                        #程序自动填充部分
                        New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 3, 3, Row.get('姓名'), ExcelStyle)
                        New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 4, 4, Row.get('性别'), ExcelStyle)
                        New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 5, 5, '16', ExcelStyle)
                        New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 6, 6, Row.get('宿舍号')+'-'+Row.get('床号'), ExcelStyle)
                        if(Row.get('父亲手机') == ''):
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 7, 7, Row.get('母亲手机'), ExcelStyle)
                        else:
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 7, 7, Row.get('父亲手机'), ExcelStyle)
                        New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 11, 11, '是', ExcelStyle)

                        #考勤员手动选择部分（也可以自己输入）
                        print('请选择一个缺勤原因或输入一个字符串作为缺勤原因')
                        print('可选项：')
                        print('[A]发热')
                        print('[B]头疼')
                        print('[C]咳嗽')
                        print('[D]呕吐')
                        print('[E]腹泻')
                        Absence_Reason = str(input())
                        if Absence_Reason == 'A':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 8, 8, "发热", ExcelStyle)
                        elif Absence_Reason == 'B':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 8, 8, "头疼", ExcelStyle)
                        elif Absence_Reason == 'C':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 8, 8, "咳嗽", ExcelStyle)
                        elif Absence_Reason == 'D':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 8, 8, "呕吐", ExcelStyle)
                        elif Absence_Reason == 'E':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 8, 8, "腹泻", ExcelStyle)
                        else:
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 8, 8, Absence_Reason, ExcelStyle)
                        
                        print('请选择一个症状出现日期或输入一个字符串作为症状出现日期')
                        print('可选项：')
                        print('[A]'+(time.strftime('%m-%d',time.localtime(time.time()))+'早'))
                        print('[B]'+(date.today() + timedelta(days = -1)).strftime("%m-%d")+'早')
                        print('[C]'+(date.today() + timedelta(days = -1)).strftime("%m-%d")+'午')
                        print('[D]'+(date.today() + timedelta(days = -1)).strftime("%m-%d")+'晚')
                        The_Date_of_symptoms = str(input())
                        if The_Date_of_symptoms == 'A':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 9, 9, (time.strftime('%m-%d',time.localtime(time.time()))+'早'), ExcelStyle)
                        elif The_Date_of_symptoms == 'B':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 9, 9, (date.today() + timedelta(days = -1)).strftime("%m-%d")+'早', ExcelStyle)
                        elif The_Date_of_symptoms == 'C':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 9, 9, (date.today() + timedelta(days = -1)).strftime("%m-%d")+'午', ExcelStyle)
                        elif The_Date_of_symptoms == 'D':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 9, 9, (date.today() + timedelta(days = -1)).strftime("%m-%d")+'晚', ExcelStyle)
                        else:
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 9, 9, The_Date_of_symptoms, ExcelStyle)

                        print('请选择一个离校时间或输入一个字符串作为症状出现日期')
                        print('可选项：')
                        print('[A]'+(time.strftime('%m-%d',time.localtime(time.time()))+'早'))
                        print('[B]'+(date.today() + timedelta(days = -1)).strftime("%m-%d")+'早')
                        print('[C]'+(date.today() + timedelta(days = -1)).strftime("%m-%d")+'午')
                        print('[D]'+(date.today() + timedelta(days = -1)).strftime("%m-%d")+'晚')
                        The_Date_of_Leaving = str(input())
                        if The_Date_of_Leaving == 'A':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 10, 10, (time.strftime('%m-%d',time.localtime(time.time()))+'早'), ExcelStyle)
                        elif The_Date_of_Leaving == 'B':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 10, 10, (date.today() + timedelta(days = -1)).strftime("%m-%d")+'早', ExcelStyle)
                        elif The_Date_of_Leaving == 'C':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 10, 10, (date.today() + timedelta(days = -1)).strftime("%m-%d")+'午', ExcelStyle)
                        elif The_Date_of_Leaving == 'D':
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 10, 10, (date.today() + timedelta(days = -1)).strftime("%m-%d")+'晚', ExcelStyle)
                        else:
                            New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 10, 10, The_Date_of_Leaving, ExcelStyle)

                        print('该同学是否已经回校？（Y/N）')
                        Flag = str(input())
                        if Flag == 'Y':
                            print('请选择一个回校时间或输入一个字符串作为症状出现日期')
                            print('可选项：')
                            print('[A]'+(time.strftime('%m-%d',time.localtime(time.time()))+'早'))
                            print('[B]'+(date.today() + timedelta(days = -1)).strftime("%m-%d")+'早')
                            print('[C]'+(date.today() + timedelta(days = -1)).strftime("%m-%d")+'午')
                            print('[D]'+(date.today() + timedelta(days = -1)).strftime("%m-%d")+'晚')
                            The_Date_of_Return = str(input())
                            if The_Date_of_Return == 'A':
                                New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 13, 13, (time.strftime('%m-%d',time.localtime(time.time()))+'早'), ExcelStyle)
                            elif The_Date_of_Return == 'B':
                                New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 13, 13, (date.today() + timedelta(days = -1)).strftime("%m-%d")+'早', ExcelStyle)
                            elif The_Date_of_Return == 'C':
                                New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 13, 13, (date.today() + timedelta(days = -1)).strftime("%m-%d")+'午', ExcelStyle)
                            elif The_Date_of_Return == 'D':
                                New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 13, 13, (date.today() + timedelta(days = -1)).strftime("%m-%d")+'晚', ExcelStyle)
                            else:
                                New_WorkSheet.write_merge(Rows_Old+K-1, Rows_Old+K, 13, 13, The_Date_of_Return, ExcelStyle)
                        print('请核对该同学的信息：')
                        print('缺勤人姓名：'+Row.get('姓名'))
                        print('缺勤人性别：'+Row.get('性别'))
                        print('缺勤人宿舍：'+Row.get('宿舍号')+'-'+Row.get('床号'))
                        print('缺勤原因：'+Absence_Reason)
                        print('症状出现日期：'+The_Date_of_symptoms)
                        print('离校时间：'+The_Date_of_Leaving)
                        if Flag == 'Y':
                            print('回校时间：'+The_Date_of_Return)
                        os.system('pause')
            K += 2
    New_WorkBook.save(FilePath)
    print('写入成功，准备回到主菜单')

#修改配置文件
def UpDateConfig():
    Document = open("Config.txt", "a+", encoding = "UTF-8")
    Document.seek(os.SEEK_SET)
    Text = Document.read()
    Document.close()
    print('-----当前文件内容-----')
    print(Text)
    print('----------------------')
    print('是否要继续修改配置文件？(Y/N) ')
    Flag = str(input())
    if Flag == 'Y' :
        print('请输入当前所在年级(1,2,3)： ')
        NowGrade = str(input())
        print('请输入当前所在班级(1,2,...,N)： ')
        NowClass = str(input())
        print('请输入考勤表所在路径(绝对路径)： ')
        NowPath = str(input())
        os.remove("Config.txt")
        NewDocument = open("Config.txt", "a+", encoding = "UTF-8")
        NewDocument.write('Grade:'+NowGrade+'\n')
        NewDocument.write('Class:'+NowClass+'\n')
        NewDocument.write('Path:'+NowPath)
        NewDocument.close()
        print('修改完成，即将回到主菜单')
        time.sleep(0.5)  
    
#打印菜单栏
def Print_Menu():
    print('1:上送考勤表')
    print('2:填充考勤表')
    print('3:更新配置文件')
    print('4:退出程序')
    print('输入一个参数以进一步操作： ')
    Parameter = int(input())
    if Parameter == 1 :
        UpLoad()
    if Parameter == 2 :
        WriteExcel()
    if Parameter == 3 :
        UpDateConfig()
    if Parameter == 4 :
        exit()
    
#开始处理各种情况
print('::::::::::::::::::::::::::::::::::::::::::::::::::::::::')
print('::::              JMYZ 自动化考勤系统               ::::')
print('::::::::::::::::::::::::::::::::::::::::::::::::::::::::')

while True:
    Print_Menu()
    os.system('cls')
