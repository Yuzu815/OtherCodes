import os
import ctypes
import shutil
import platform
import hashlib
import time, datetime
from getpass import getpass

dicts = dict()

def WriteFile(filename,content):
	with open(filename,'a', encoding = "UTF-8") as fw:
		fw.write(str(content)+'\n')

def BUG():
    os.system("cls")
    print("请按照提示进行操作：")
    problem = str(input("问题（必填）："))
    username = str(input("反馈人（选填）："))
    WriteFile("BUG.txt", "问题："+problem)
    WriteFile("BUG.txt", "反馈人："+username)
    WriteFile("BUG.txt", "反馈时间："+str(datetime.date.today()))
    WriteFile("BUG.txt", "处理情况：未解决")
    WriteFile("BUG.txt", "-------------------------------")
    print("已经成功提交...，即将返回主菜单...")
    time.sleep(0.5)

def Suggest():
    os.system("cls")
    print("请按照提示进行操作，非常感谢：")
    problem = str(input("建议："))
    username = str(input("反馈人（选填）："))
    WriteFile("SUG.txt", "建议："+problem)
    WriteFile("SUG.txt", "反馈人："+username)
    WriteFile("SUG.txt", "反馈时间："+str(datetime.date.today()))
    WriteFile("SUG.txt", "处理情况：未解决")
    WriteFile("SUG.txt", "-------------------------------")
    print("已经成功提交...，即将返回主菜单...")
    time.sleep(0.5)

def nline(filename, count):
    cnt = 0
    f = open(filename, 'r', encoding='UTF-8')
    line = f.readline()
    while line:
        cnt += 1
        if cnt == count:
            return line
        line = f.readline()

def solve1():
    solve_cnt1 = 1
    print("需要解决的问题有：")
    count = len(open("BUG.txt", 'r', encoding='UTF-8').readlines())
    print("-------------------------------")
    for i in range(1, count):
        #print("DEBUG:"+nline("BUG.txt", i))
        if "未解决" in nline("BUG.txt", i):
            print("编号：P"+str(solve_cnt1)+'\n')
            print(nline("BUG.txt", i-3))
            print(nline("BUG.txt", i-2))
            print(nline("BUG.txt", i-1))
            print(nline("BUG.txt", i), end='')
            print("-------------------------------")
            dicts['P'+str(solve_cnt1)] = i
            solve_cnt1 += 1

def solve2():
    solve_cnt2 = 1
    print("需要查看的建议有：")
    count = len(open("SUG.txt", 'r', encoding='UTF-8').readlines())
    print("-------------------------------")
    for i in range(1, count):
        if "未解决" in nline("SUG.txt", i):
            print("编号：S"+str(solve_cnt2)+'\n')
            print(nline("SUG.txt", i-3))
            print(nline("SUG.txt", i-2))
            print(nline("SUG.txt", i-1))
            print(nline("SUG.txt", i), end='')
            print("-------------------------------")
            dicts['S'+str(solve_cnt2)] = i
            solve_cnt2 += 1

def rewrite(num):
        if 'P' in num:
                f1 = open("BUG.txt", 'r', encoding='UTF-8')
                line = f1.readlines()
                f1.close()
                f1 = open("BUG.txt", 'w', encoding='UTF-8')
                line[int(dicts[num])-1] = "处理情况：已解决\n"
                f1.writelines(line)
                f1.close()
        if 'S' in num:
                f2 = open("SUG.txt", 'r', encoding='UTF-8')
                line = f2.readlines()
                f2.close()
                f2 = open("SUG.txt", 'w', encoding='UTF-8')
                line[int(dicts[num])-1] = "处理情况：已解决\n"
                f2.writelines(line)
                f2.close()

def alreadysolve():
    while(1):
            print("是否有已经解决的问题或建议？")
            print("如果是，请输入对应的编号.如果没有，请输入0")
            num = str(input())
            if num == "0":
                    return
            else:
                    print("即将把编号为"+num+"的对象属性设为[已解决]，确定？(Y/N)")
            flag = str(input())
            if flag == 'Y':
                    rewrite(num)
            time.sleep(0.5)

def count():
    count1 = 0
    count2 = 0
    f1 = open("BUG.txt", 'r', encoding='UTF-8')
    line1 = f1.readline()
    f2 = open("SUG.txt", 'r', encoding='UTF-8')
    line2 = f2.readline()
    while line1:
        if "未解决" in line1:
            count1 += 1
        line1 = f1.readline()
    while line2:
        if "未解决" in line2:
            count2 += 1
        line2 = f2.readline()
    print("当前有"+str(count1)+"个问题待解决")
    print("当前有"+str(count2)+"个建议待处理")
    if count1 > 0:
        solve1()
    if count2 > 0:
        solve2()
    if count1 > 0 or count2 > 0:
        alreadysolve()

def man():
    os.system("cls")
    print("===计算机维护管理报修管理系统后台===")
    count()
    os.system("pause")

def DJ():
    os.system("cls")
    print("请输入对应的密码：")
    number = getpass()
    md5 = hashlib.md5()
    md5.update(number.encode())
    #请自行替换此处密码，使用md5加密
    if(md5.hexdigest() == "4cd00976f9ef1128caca7e0e2a15a80a"):
        print("验证成功，正在跳转管理页面...")
        man()
    else:
        print("验证失败，即将退出...")
        time.sleep(0.5)
        exit()

def History():
    os.system("cls")
    print("以下为提出的问题：")
    print("-------------------------------")
    f1 = open("BUG.txt", 'r', encoding='UTF-8')
    print(f1.read())
    print("以下为提出的建议：")
    print("-------------------------------")
    f2 = open("SUG.txt", 'r', encoding='UTF-8')
    print(f2.read())
    os.system("pause")

def menu():
    print("===计算机维护管理报修===")
    print("输入下列选项对应的数字以开始程序：")
    print("1.我要反馈现在电脑存在的BUG.")
    print("2.我有改进现在电脑的好的建议.")
    print("3.我要查看已经提出的问题和建议.")
    print("4.退出程序")
    number = int(input())
    if number == 0:
        DJ()
    if number == 1:
        BUG()
    if number == 2:
        Suggest()
    if number == 3:
        History()
    if number == 4:
        exit()

while(1):
    menu()
    os.system("cls")
