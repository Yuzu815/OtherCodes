import os
import shutil
import time, datetime

SETDAYS = 7
FilesDir = "NULL"
NewPos = "NULL"
NowTime = "NULL"

def WriteFile(filename,content):
	with open(filename,'a') as fw:
		fw.write(str(content)+'\n')
	
def Move():
	for DirPath, DirName, FileName in os.walk(FilesDir):
		for NowName in FileName:
                        #若要遍历包括子目录下的文件则注释下面两行
			if DirPath != FilesDir:
				return;
			FullName = DirPath + "\\" + NowName
			CreateTime = os.stat(FullName).st_mtime
			EndFileTime = time.time()-3600*24*int(SETDAYS)
			if EndFileTime > CreateTime:
				shutil.move(FullName, NewPos+'\\'+NowTime+'\\'+NowName)
				WriteFile("Logs.txt", "[正在移动文件]"+NowName)
				WriteFile("Logs.txt", "原目录:"+FullName)
				WriteFile("Logs.txt", "新目录:"+NewPos+NowTime+'\\'+NowName)
def Init():
	k = 1
	global FilesDir, NewPos, SETDAYS, NowTime
	NowTime = str(datetime.date.today())
	fl = open("RegularCleanConfig.txt", "r")
	lines = fl.readlines()
	for line in lines:
		if k == 1:
			FilesDir = line[line.find(':')+1:-1]
			#print(FilesDir)
		elif k == 2:
			NewPos = line[line.find(':')+1:-1]
			#print(NewPos)
		elif k == 3:
			SETDAYS = line[line.find(':')+1:-1]
			#print(SETDAYS)
		k = k + 1
	#print(NewPos+'\\'+NowTime)   
	if os.path.exists(NewPos+'\\'+NowTime) == 0:
		os.makedirs(NewPos+'\\'+NowTime)
		WriteFile("Logs.txt", "[创建目录]"+NewPos+'\\'+NowTime)
		
def Check():
	if os.path.exists("RegularCleanConfig.txt"):
		print("存在配置文件，正在按照预定配置清理中...")
		WriteFile("Logs.txt",str(datetime.date.today()))
		Init()
		Move()
		WriteFile("Logs.txt","---操作完成---")
	else:
		print("当前未程序尚未配置，请根据指引配置后再使用")
		print("请输入需要定期归档的文件夹路径（必须存在）：")
		MakeFolderPath = str(input())
		print("请输入归档的文件夹路径(必须存在)：")
		BackFolderPath = str(input())
		print("请输入归档几天内的文件（默认为7天）")
		Cycle = str(input())
		if os.path.exists(MakeFolderPath) and os.path.exists(BackFolderPath):
			WriteFile("RegularCleanConfig.txt","MakeFolderPath:"+MakeFolderPath)
			WriteFile("RegularCleanConfig.txt","BackFolderPath:"+BackFolderPath)
			WriteFile("RegularCleanConfig.txt","Cycle:"+Cycle)
			print("配置完成，请重新打开该程序...")
		else:
			print("配置文件有误，请检查是否输入了空路径")
Check()
	
		
		
