import requests
import json
import time, datetime
import os

word = ""
nowres = "NULL"

def WriteFile(filename,content):
	with open(filename,'a', encoding = "UTF-8") as fw:
		fw.write(str(content)+'\n')
		
class King(object):
    def __init__(self, word):
        self.url = 'http://fy.iciba.com/ajax.php?a=fy'
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/76.0.3809.100 Safari/537.36",
        }
        self.data = {
            'f': 'auto',
            't': 'auto',
            'w': word
        }

    def get_data(self):
        response = requests.post(url=self.url, headers=self.headers, data=self.data)
        return response.content

    def parse_data(self, data):
        dict_data = json.loads(data)
        try:
            wordmean = dict_data['content']['out']
            WriteFile(nowres, word[:-1]+"  [NULL]")
            WriteFile(nowres, wordmean)
            WriteFile("logs.txt", word[:-1])
        except:
            soundmark = "["+dict_data['content']['ph_en']+"]"
            wordmean = str(dict_data['content']['word_mean'])
            WriteFile(nowres, word[:-1]+"  "+soundmark)
            WriteFile(nowres, wordmean)
            WriteFile("logs.txt", word[:-1]) 
            

    def run(self):
        # 发送请求，获取响应
        response = self.get_data()
        # 数据解析
        self.parse_data(response)


if __name__ == '__main__':
    WriteFile("logs.txt", (datetime.datetime.now()).strftime("[%Y-%m-%d %H:%M:%S]")+" 准备开始翻译")
    file_p = open("pre.txt", "r", encoding='UTF-8')
    line_p = file_p.readline()
    while(line_p):
        word = line_p
        if '[' in word:
            line_p = file_p.readline()
            continue
        if '@' in word:
            if nowres != "NULL" and nowres != word[2:-1]+".txt": 
                WriteFile(nowres, (datetime.datetime.now()).strftime("---%Y-%m-%d %H:%M:%S---"))
            nowres = word[2:-1]+".txt"
            line_p = file_p.readline()
            continue
        king = King(word)
        king.run()
        line_p = file_p.readline()
    WriteFile("logs.txt", '---------------------\n')
    WriteFile(nowres, (datetime.datetime.now()).strftime("---%Y-%m-%d %H:%M:%S---"))
