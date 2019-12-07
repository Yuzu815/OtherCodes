import os

def WriteFile(filename,content):
	with open(filename,'a', encoding = "UTF-8") as fw:
		fw.write(str(content)+'\n')

os.system(r"cd /d %cd% && del pre.txt")
WriteFile("pre.txt", "[批量翻译内容填写]")
