import os
import ctypes
import shutil
import time, datetime
        
def WriteFile(filename,content):
	with open(filename,'a', encoding = "UTF-8") as fw:
		fw.write(str(content)+'\n')
    
def countfiles():
        path = r"C:\Users\Students\Desktop\批量转换"
        ls = os.listdir(path)
        count = 0
        WriteFile("logs.txt", "文件列表如下：")
        for i in ls:
                if os.path.isfile(os.path.join(path,i)):
                        if not ".pdf" in i:
                                if ".ppt" in i or ".doc" in i:
                                        WriteFile("logs.txt", "        "+str(i))
                                        count += 1
        WriteFile("logs.txt", "[合计转换"+str(count)+"个文件]")
        
WriteFile("logs.txt", (datetime.datetime.now()).strftime("[%Y-%m-%d %H:%M:%S]")+" 准备开始转换任务")
countfiles()
