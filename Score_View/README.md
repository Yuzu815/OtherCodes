### 关于这个脚本
这是作者希望可视化平时各种考试的成绩而写的，目前只针对以下两种格式的表格：
（好像有一个是智学网用的格式，另一个是学校自己搞的格式）

---

![](https://github.com/YZ-HL/OtherCodes/blob/master/images/table1.png?raw=true)
![](https://github.com/YZ-HL/OtherCodes/blob/master/images/table2.png?raw=true)

---

### 脚本功能

1.  能够分析历次考试总体的班排和级排
2.  能够比较历次考试单科成绩与年级平均分的差距
3.  能够比较历次考试单科成绩与指定对象的差距
4.  能够比较历次考试单科的班排和级排（新增）
5.  能够自动跳过缺考的考试（感谢@TTTTT缺考英语让我发现了这个BUG）

### 脚本效果

![](https://github.com/YZ-HL/OtherCodes/blob/master/images/score_view_1.png?raw=true)

![](https://github.com/YZ-HL/OtherCodes/blob/master/images/score_view_2.png?raw=true)

![](https://github.com/YZ-HL/OtherCodes/blob/master/images/score_view_3.png?raw=true)

### 使用指南

1.  你或许需要先准备一套完整的考试成绩数据，数据格式如上面截图所示

2.  在``start.py``下创建一个文件``config.txt``，以如下格式描述你的数据位置。

    ![](https://github.com/YZ-HL/OtherCodes/blob/master/images/score_view_config.png?raw=true)

    其中，config文件内每一行的格式为：

    [考试名称]#[绝对路径]#[考试成绩文件格式]
    
    若文件格式为上面第一幅图的格式，则填写1，第二幅图的格式填写2。其他格式暂不支持。
    
3.  尽情享用吧！

### 需要完善

1.  目前只支持语数英物化生的查询（以全局变量写死在了程序里...以后可能会重写）
2.  堆积情况比较严重，考虑优化观感和使用体验
3.  未完成js本地化，必须联网才能使用。



**感谢pyecharts！**