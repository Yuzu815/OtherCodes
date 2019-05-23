这是一个放各种小型代码的仓库，主要是自己在日常生活中需要实现需求时写的轮子。

#### 简要介绍

- Clean（定期归档文件）

  ​	用于对教室电脑桌面的清理。

  ​	由于较多的同学都喜欢随手把课件扔在桌面上影响美观，因此写了一个这样子的程序来定期整理桌面上的文件。支持简单的日志。

- Automatic attendance system（简化晨报系统）

  ​	这个程序只适用于校内的考勤系统，用于简化晨报人员的工作。

  ​	一般的晨报流程：打开Excel->填写相关信息->打开Firefox->登录网站并上传。

  ​	目前的晨报流程：简单检查人数->输入今日情况->一键上传（全程不需要打开Excel和Firefox）

  ​	特别是班内电脑配置"较高"，打开Excel和Firefox时都会比较卡。因此简化整个考勤步骤就变得很有必要了。

  ![](https://github.com/YZ-HL/OtherCodes/blob/master/images/Automatic%20attendance%20system.png?raw=true)



- Office To PDF（Office批量转换为PDF）

  ​	班上用Kindle，IReader的同学越来越多，因此一个个转换课件就很成问题。

  ​	因此写了两个简单的vbs，支持转换与vbs同目录下的Office文件（目前只包括doc，docx，ppt，pptx）。原理是调用Office的导出功能来实现批量转换，因此使用这个程序的电脑上一定要预装Office套件（Microsoft或者WPS都行）（UPD：目前在Windows-LTSC，WPS2019专业版下可用，其他情况下暂未发现可用。）

  ​	![](https://github.com/YZ-HL/OtherCodes/blob/master/images/Office%20To%20PDF.png?raw=true)



- ZHIXUE Query（智学网成绩查询）

  ​	这个程序基于[anwenhu](https://github.com/anwenhu/zhixuewang)提供的智学网API及相关包，然后[Origami404](https://github.com/Origami404/CommandZhixue)在原有的基础上进行了封装。

  ​	现在将其打包放在这里，方便提供智学网的查询服务。

  ![](https://github.com/YZ-HL/OtherCodes/blob/master/images/ZHIXUE%20Query.png?raw=true)
