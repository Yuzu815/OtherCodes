@echo off
color F0
echo 按指示进行操作即可...
cd /d %~dp0
python init.py
start pre.txt
msg %username% 在打开的文件中输入单词，对于每一组来源的数据，第一行输入来源，格式为"@ XXXXXX"，接下来的每一行一个单词，接着到下一组数据，完成所有数据输入后保存并继续操作。
pause
python start.py
msg %username% 完成翻译，结果存放于对应文件内。
exit