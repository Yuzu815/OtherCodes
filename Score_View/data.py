import xlrd
import xlwt
import xlutils
import os
import linecache
from pyecharts.charts import Bar
from pyecharts.charts import Tab
from pyecharts.charts import Line
from pyecharts.charts import Radar
from pyecharts.charts import Grid
from pyecharts.charts import Page
from pyecharts import options as opts

#Author:YZ_HL
#Date:2019.10.13

##########################
##程序所需常量及全局变量##
##########################

#关键字字典列表
#文科班的需要修改很多常量
Subject = ["语文","数学","英语","物理","化学","生物"]
ScoreType = ["成绩","班排","级排"]
ExcelPosition_1 = [2, 3, 4, 7, 8, 9]
ExcelPosition_2 = [5, 8, 11, 14, 17, 20]

Class_Rank = []
Grade_Rank = []
All_Exam_Name = []

############################
##各类图生成函数及辅助函数##
############################

#统计文件行数
def Line_num(filename):
    cnt = 0
    f = open(filename, 'r', encoding='UTF-8')
    line = f.readline()
    while line:
        cnt += 1
        line = f.readline()
    return cnt

#处理[num1][num2]型数据
def Analysis_Type_1(raw_data):
    new_data = raw_data.replace('][', '#')
    new_data = new_data.replace('[', '')
    new_data = new_data.replace(']', '')
    return new_data

#处理num1[num2][num3]型数据
def Analysis_Type_2(raw_data):
    new_data = raw_data.replace('][', '#')
    new_data = new_data.replace('[', '#')
    new_data = new_data.replace(']', '')
    return new_data

#返回对应学号的学生数据
def Find_In_Excel(WorkSheet, Exam_Number, Exam_Data_Format):
    StudentScore = []
    StudentScore.append(Exam_Name)
    if Exam_Data_Format == "1\n":
        for j in range(0, WorkSheet.nrows):
            Temp = WorkSheet.cell_value(j, 0)
            if Temp == Exam_Number:
                StudentScore.append(WorkSheet.cell_value(j, 1))
                for k in range(2, WorkSheet.ncols):
                    StudentScore.append(WorkSheet.cell_value(j, k))

    if Exam_Data_Format == "2\n":
        for j in range(0, WorkSheet.nrows):
            Temp = WorkSheet.cell_value(j, 1)
            #该处是考号匹配的处理，迁移使用时要注意修改
            if Temp == "2020"+Exam_Number:
                StudentScore.append(WorkSheet.cell_value(j, 1))
                for k in range(2, WorkSheet.ncols):
                    StudentScore.append(WorkSheet.cell_value(j, k))

    return StudentScore

#返回已处理的列表结果（过滤出数据里的分数）
def Solve_Data_Score(raw_List, Exam_Data_Format):
    Subject_Score = []
    if Exam_Data_Format == "1\n":
        for j in range(0, 6):
            Subject_Score.append(Analysis_Type_2(raw_List[ExcelPosition_1[j]]).split('#', 3)[0])
    if Exam_Data_Format == "2\n":
        for j in range(0, 6):
            Subject_Score.append(raw_List[ExcelPosition_2[j]])
    return Subject_Score

#返回已处理的列表结果（过滤出数据里的排名）

#平均分计算
def Aveage_Score(WorkSheet, Exam_Data_Format):
    Aveage = []
    if Exam_Data_Format == "1\n":
        for i in range(0, 6):
            truth_number = WorkSheet.nrows
            Sum = 0
            #一段针对表格1格式的代码
            for j in range(1, WorkSheet.nrows):
                if WorkSheet.cell_value(j, ExcelPosition_1[i]) == "缺考":
                    truth_number -= 1
                    continue
                now_score = Analysis_Type_2(WorkSheet.cell_value(j, ExcelPosition_1[i])).split('#', 3)[0]
                Sum += float(now_score)
            Aveage.append(int(Sum/truth_number))
    if Exam_Data_Format == "2\n":
        for i in range(0, 6):
            truth_number = WorkSheet.nrows
            Sum = 0
            #一段针对表格2格式的代码
            for j in range(4, WorkSheet.nrows):
                if WorkSheet.cell_value(j, ExcelPosition_2[i]) == "-" or WorkSheet.cell_value(j, ExcelPosition_2[i]) == "未扫描":
                    truth_number -= 1
                    continue
                Sum += int(WorkSheet.cell_value(j, ExcelPosition_2[i]))
            Aveage.append(int(Sum/truth_number))
    return Aveage

#生成对应的雷达图
def Create_Radar(raw_data, ave_data, exam_name):
    Score_Radar = (
        Radar(init_opts=opts.InitOpts(bg_color="#FFFFFF"))
        .add_schema(
            schema=[
                opts.RadarIndicatorItem(name="语文", max_=150),
                opts.RadarIndicatorItem(name="数学", max_=150),
                opts.RadarIndicatorItem(name="英语", max_=150),
                opts.RadarIndicatorItem(name="物理", max_=110),
                opts.RadarIndicatorItem(name="化学", max_=100),
                opts.RadarIndicatorItem(name="生物", max_=90),
            ],
            splitarea_opt=opts.SplitAreaOpts(
                is_show=True, areastyle_opts=opts.AreaStyleOpts(opacity=1)
            ),
        )
        .add(
            series_name="成绩",
            data=[raw_data],
            linestyle_opts=opts.LineStyleOpts(color="#f9713c",width=3),
            areastyle_opts=opts.AreaStyleOpts(color="#ea3a2e", opacity=0.3),
        )
        .add(
            series_name="平均分",
            data=[ave_data],
            linestyle_opts=opts.LineStyleOpts(color="#77ccff",width=3),
            areastyle_opts=opts.AreaStyleOpts(color="#66ccff", opacity=0.3),
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title=exam_name), legend_opts=opts.LegendOpts()
        )
    )
    filename = "[雷达图]"+exam_name[1:-1]+".html"
    Score_Radar.render(filename)

#生成对应的柱状图
def Create_Bar(my_data, other_data, exam_name):
    Score_Bar = (
        Bar()
        .add_xaxis(Subject)
        .add_yaxis("你的成绩", my_data)
        .add_yaxis("ta的成绩", other_data)
        .set_global_opts(title_opts=opts.TitleOpts(title=exam_name))
    )
    filename = "[柱状图]"+exam_name[1:-1]+".html"
    Score_Bar.render(filename)

#生成对应的折线图
def Create_Line(class_data, grade_data):
    Class_Line =(
        Line()
        .add_xaxis(xaxis_data=All_Exam_Name)
        .add_yaxis(
            series_name="班排",
            y_axis=class_data,
            symbol="emptyCircle",
            linestyle_opts=opts.LineStyleOpts(color="#C67570", width=2),
            areastyle_opts=opts.AreaStyleOpts(opacity=0.5, color="#C67570"),
        )
    )
    Grade_Line = (
        Line()
        .add_xaxis(xaxis_data=All_Exam_Name)
        .add_yaxis(
            series_name="级排",
            y_axis=grade_data,
            symbol="emptyCircle",
            linestyle_opts=opts.LineStyleOpts(color="#C67570", width=2),
            areastyle_opts=opts.AreaStyleOpts(opacity=0.5, color="#C67570"),
        )
    )
    tab = Tab()
    tab.add(Grade_Line, "级排")
    tab.add(Class_Line, "班排")

    filename = "[折线图]"+"近期考试波动分析"+".html"
    tab.render(filename)


############################
##用户操作及程序判断的部分##
############################
    
#开始输入用户信息
print("请输入学号:", end='')
Exam_Number = str(input())

#Exam_Number = "0922"
print("请输入操作类型：")
print("1. 与他人对比成绩（柱状图呈现）")
print("2. 自己历次成绩的对比（折线图呈现）")
print("3. 某一次考试的具体分析（雷达图呈现）")
Operate_Type = int(input())

Other_Exam_Number = "NULL"
if Operate_Type == 1:
    print("请输入另一人的学号：", end = '')
    Other_Exam_Number = str(input())

#读取配置文件
for i in range(1, Line_num("Config.txt")+1):
    Config_Data = linecache.getline('Config.txt', i)

    #配置文件注释
    if Config_Data[0] == '#':
        continue;

    #分割配置文件内容
    Exam_Name = Config_Data.split('#', 2)[0]
    Exam_Data_Path = Config_Data.split('#', 2)[1]
    Exam_Data_Format = Config_Data.split('#', 2)[2]

    #打开对应的Excel工作表
    WorkBook = xlrd.open_workbook(Exam_Data_Path+r"\data.xls", formatting_info = True)

    #该处是针对不同成绩类型所对应的数据处理
    #如果增加了新的数据类型需要修改
    WorkSheet = ''
    if Exam_Data_Format == "1\n":
        WorkSheet = WorkBook.sheet_by_name("理班")
    if Exam_Data_Format == "2\n":
        WorkSheet = WorkBook.sheet_by_name("全部考生成绩汇总")

    #分数提取调试
    #print(Find_In_Excel(WorkSheet, Exam_Number, Exam_Data_Format))
    #print(Solve_Data_Score(Find_In_Excel(WorkSheet, Exam_Number, Exam_Data_Format), Exam_Data_Format))

    #提取自己的信息
    Result_List_Me = Find_In_Excel(WorkSheet, Exam_Number, Exam_Data_Format)
    
    if Operate_Type == 1:       
        Result_List_Other = Find_In_Excel(WorkSheet, Other_Exam_Number, Exam_Data_Format)
        Create_Bar(Solve_Data_Score(Result_List_Me, Exam_Data_Format), Solve_Data_Score(Result_List_Other, Exam_Data_Format), Exam_Name)
    elif Operate_Type == 2:
        All_Exam_Name.append(Exam_Name)
        if Exam_Data_Format == "1\n":
            if Result_List_Me[WorkSheet.ncols-1] == "缺考":
                continue
            Class_Rank.append(Analysis_Type_1(Result_List_Me[WorkSheet.ncols-1]).split('#', 2)[0])
            Grade_Rank.append(Analysis_Type_1(Result_List_Me[WorkSheet.ncols-1]).split('#', 2)[1])
        if Exam_Data_Format == "2\n":
            if Result_List_Me[WorkSheet.ncols-1] == "-":
                continue
            Class_Rank.append(Result_List_Me[WorkSheet.ncols-1])
            Grade_Rank.append(Result_List_Me[WorkSheet.ncols-2])
    elif Operate_Type == 3:
        Create_Radar(Solve_Data_Score(Result_List_Me, Exam_Data_Format), Aveage_Score(WorkSheet, Exam_Data_Format), Exam_Name)

if Operate_Type == 2:
    Create_Line(Class_Rank,Grade_Rank)

    
