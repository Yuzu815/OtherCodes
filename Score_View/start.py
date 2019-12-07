import xlrd
import xlwt
import xlutils
import os
import time
import linecache
from pyecharts.charts import Bar
from pyecharts.charts import Tab
from pyecharts.charts import Line
from pyecharts.charts import Radar
from pyecharts.charts import Grid
from pyecharts.charts import Page
from pyecharts.charts import Timeline
from pyecharts import options as opts
from pyecharts.commons.utils import JsCode
from pyecharts.components import Table
from pyecharts.options import ComponentTitleOpts

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

cnt = 0
Class_Rank = []
Grade_Rank = []
All_Exam_Name = []
All_Exam_Ave = []
All_Exam_Score = [[] for j in range(10)]
All_Exam_RANK = []
ALL_Exam_Tab = []

############################
##各类图生成函数及辅助函数##
############################

#统计文件行数
#参数说明：配置文件的文件名（要求在同一目录下）
def Line_num(filename):
    cnt = 0
    f = open(filename, 'r', encoding='UTF-8')
    line = f.readline()
    while line:
        cnt += 1
        line = f.readline()
    return cnt

#处理[num1][num2]型数据
#参数说明：形如[num1][num2]的字符串
def Analysis_Type_1(raw_data):
    new_data = raw_data.replace('][', '#')
    new_data = new_data.replace('[', '')
    new_data = new_data.replace(']', '')
    return new_data

#处理num1[num2][num3]型数据
#参数说明：形如num1[num2][num3]的字符串
def Analysis_Type_2(raw_data):
    new_data = raw_data.replace('][', '#')
    new_data = new_data.replace('[', '#')
    new_data = new_data.replace(']', '')
    return new_data

#返回对应学号的学生数据
#参数说明：数据集（工作簿），学号（格式类似"0722"），数据格式
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
#参数说明：关于我的数据（一行，未处理分隔符），数据格式
def Solve_Data_RANK(raw_List, Exam_Data_Format):
    Subject_Score = []
    if Exam_Data_Format == "1\n":
        for j in range(0, 6):
            Subject_Score.append(Analysis_Type_2(raw_List[ExcelPosition_1[j]]).split('#', 3)[2])
    if Exam_Data_Format == "2\n":
        for j in range(0, 6):
            Subject_Score.append(raw_List[ExcelPosition_2[j]+1])
    return Subject_Score

#返回已处理的列表结果（过滤出数据里的年级排名）
#参数说明：关于我的数据（一行，未处理分隔符），数据格式
def Solve_Data_Score(raw_List, Exam_Data_Format):
    Subject_Score = []
    if Exam_Data_Format == "1\n":
        for j in range(0, 6):
            Subject_Score.append(Analysis_Type_2(raw_List[ExcelPosition_1[j]]).split('#', 3)[0])
    if Exam_Data_Format == "2\n":
        for j in range(0, 6):
            Subject_Score.append(raw_List[ExcelPosition_2[j]])
    return Subject_Score

#处理排名并放置于全局变量
#参数说明：关于我的数据（一行，未处理分隔符），考试名称
def Solve_Data_Rank(Result_List_Me, Exam_Name):
    All_Exam_Name.append(Exam_Name)
    if Exam_Data_Format == "1\n":
        if Result_List_Me[WorkSheet.ncols-1] != "缺考":
            Class_Rank.append(Analysis_Type_1(Result_List_Me[WorkSheet.ncols-1]).split('#', 2)[0])
            Grade_Rank.append(Analysis_Type_1(Result_List_Me[WorkSheet.ncols-1]).split('#', 2)[1])
    if Exam_Data_Format == "2\n":
        if Result_List_Me[WorkSheet.ncols-1] != "-":
            Class_Rank.append(Result_List_Me[WorkSheet.ncols-1])
            Grade_Rank.append(Result_List_Me[WorkSheet.ncols-2])
            
#平均分计算
#参数说明：传入的原数据（工作簿），数据格式类型
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

#生成对应的柱状图（对比类）
#参数说明：某次考试我的分数，某次考试它的分数，考试名称，x轴名称，y轴名称，输出html或者返回图表
def Create_Bar(my_data, other_data, exam_name, x_name, y_name, flag):
    Score_Bar =(
        Bar()
        .add_xaxis(Subject)
        .add_yaxis(x_name, my_data)
        .add_yaxis(y_name, other_data)
        .set_global_opts(title_opts=opts.TitleOpts(title=exam_name))
    )
    if flag == True:
        filename = "[柱状图]"+exam_name[1:-1]+".html"
        Score_Bar.render(filename)
    else:
        return Score_Bar
#生成表格
def table_base() -> Table:
    table = Table()

    headers = ["年份", "华南理工大学", "深圳大学", "广州工业大学", "五邑大学"]
    rows = [
        ["2019", 55, 169, 377, 745],
		["2018", 58, 167, 452, 792],
		["2017", 87, 432, 406, 768],
		["2016", 58, 470, 350, 800],
    ]
    table.add(headers, rows).set_global_opts(
        title_opts=ComponentTitleOpts(title="江门一中最低录取排名（不包含自主招生）")
    )
    return table
	
	
#生成对应的折线图
#参数说明：每次考试的班排，每次考试的级排，输出html或者返回图表
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
        .set_global_opts(
            title_opts=opts.TitleOpts(title="班排", pos_top="0%"),
            legend_opts=opts.LegendOpts(pos_top="0%"),
        )
    )
    #用于处理合并图表和单独输出的参数

    pos = "50%"
        
    Grade_Line =(
        Line()
        .add_xaxis(xaxis_data=All_Exam_Name)
        .add_yaxis(
            series_name="级排",
            y_axis=grade_data,
            symbol="emptyCircle",
            linestyle_opts=opts.LineStyleOpts(width=2),
            areastyle_opts=opts.AreaStyleOpts(opacity=0.5),
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title="级排", pos_top=pos),
            legend_opts=opts.LegendOpts(pos_top=pos),
        )
    )
    grid = (
	    Grid()
	    .add(Class_Line, grid_opts=opts.GridOpts(pos_bottom="60%"))
	    .add(Grade_Line, grid_opts=opts.GridOpts(pos_top="60%"))
    )
    return grid

#生成单科排名折线图
def Create_Line_RANK(class_data, grade_data):
    Class_Line =(
        Line()
        .add_xaxis(xaxis_data=All_Exam_Name)
        .add_yaxis(
            series_name="分数",
            y_axis=class_data,
            symbol="emptyCircle",
            linestyle_opts=opts.LineStyleOpts(color="#C67570", width=2),
            areastyle_opts=opts.AreaStyleOpts(opacity=0.5, color="#C67570"),
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title="分数", pos_top="0%"),
            legend_opts=opts.LegendOpts(pos_top="0%"),
        )
    )
    
    pos = "50%"
        
    Grade_Line =(
        Line()
        .add_xaxis(xaxis_data=All_Exam_Name)
        .add_yaxis(
            series_name="级排",
            y_axis=grade_data,
            symbol="emptyCircle",
            linestyle_opts=opts.LineStyleOpts(width=2),
            areastyle_opts=opts.AreaStyleOpts(opacity=0.5),
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title="级排", pos_top=pos),
            legend_opts=opts.LegendOpts(pos_top=pos),
        )
    )
    grid = (
        Grid()
        .add(Class_Line, grid_opts=opts.GridOpts(pos_bottom="60%"))
        .add(Grade_Line, grid_opts=opts.GridOpts(pos_top="60%"))
    )
    return grid


#生成单科成绩对比的折线图
#参数说明：每次考试的某科分数，某科对应的位置（见常量设定），输出html或者返回图表
def Create_Line_Score(raw_data, subject_pos, flag):
    Score_Line =(
        Line()
        .add_xaxis(xaxis_data=All_Exam_Name)
        .add_yaxis(
            series_name=Subject[subject_pos],
            y_axis=raw_data,
            symbol="emptyCircle",
            linestyle_opts=opts.LineStyleOpts(color="#C67570", width=2),
            areastyle_opts=opts.AreaStyleOpts(opacity=0.5, color="#C67570"),
        )
        .set_global_opts(
            title_opts=opts.TitleOpts(title=Subject[subject_pos], pos_top="0%"),
        )
    )
    if flag == True:
        Score_Line.repend()
    else:
        return Score_Line

#生成组合图表（使用标签页）
#参数说明：所有的分数数据，所有的平均分数据，所有的考试名称，所有的班排，所有的级排
def Create_Combine(my_data, ave_data, exam_name, class_data, grade_data, Exam_Number):
    #历次排名对比
    tab = Tab()
    The_Line = Create_Line(class_data, grade_data)
    tab.add(The_Line, "历次排名")
	
    tab.add(table_base(), "参考上线排名")
	
    #不同考试成绩概况
    t1 = Timeline()
    for i in range(len(ave_data)):
        Now_Bar = Create_Bar(my_data[i], ave_data[i], exam_name[i], "你的成绩", "平均分", False)
        t1.add(Now_Bar, exam_name[i])
    tab.add(t1, "近段考试概况")
    
    #历次单科成绩对比
    for i in range(0, 6):
        clone_score = []
        for j in range(len(exam_name)):
            clone_score.append(my_data[j][i])
        clone_rank = []
        for j in range(len(exam_name)):
            clone_rank.append(All_Exam_RANK[j][i])
            
        tab.add(Create_Line_RANK(clone_score, clone_rank), Subject[i]+"走势")
    tab.render(str(Exam_Number)+"的综合分析.html")

#生成组合图表-对比用-多标签页
def Create_Combine_Compare(exam_name):
    tab = Tab()
    for i in range(len(exam_name)):
        tab.add(ALL_Exam_Tab[i], exam_name[i])
    tab.render("成绩对比结果.html")
    
############################
##用户操作及程序判断的部分##
############################
    
#开始输入用户信息
os.system("color F0")
print("请输入学号:", end='')
Exam_Number = str(input())

print("请输入操作类型：")
print("1. 与他人对比成绩")
print("2. 生成历次考试的走势分析报告")
print("3. 退出程序")
Operate_Type = int(input())

if Operate_Type == 3:
    exit()

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

    Result_List_Me = Find_In_Excel(WorkSheet, Exam_Number, Exam_Data_Format)

	#处理缺考情况
    flag_Q = True
    if Exam_Data_Format == "1\n":
        for j in ExcelPosition_1:
            if Result_List_Me[j] == "缺考":
                flag_Q = False
    if Exam_Data_Format == "2\n":
        for j in ExcelPosition_2:
            if Result_List_Me[j] == "-":
                flag_Q = False

    if flag_Q == False:
        continue
	#已解决BUG
		
    All_Exam_RANK.append(Solve_Data_RANK(Result_List_Me, Exam_Data_Format))

    if(Operate_Type == 1):
        Result_List_Other = Find_In_Excel(WorkSheet, Other_Exam_Number, Exam_Data_Format)
        ALL_Exam_Tab.append(Create_Bar(Solve_Data_Score(Result_List_Me, Exam_Data_Format), Solve_Data_Score(Result_List_Other, Exam_Data_Format), Exam_Name, "你的成绩", "ta的成绩", False))

    All_Exam_Score[cnt] =  Solve_Data_Score(Result_List_Me, Exam_Data_Format)
    All_Exam_Ave.append(Aveage_Score(WorkSheet, Exam_Data_Format))
    Solve_Data_Rank(Result_List_Me, Exam_Name)
    cnt += 1
        
if Operate_Type == 2:
    Create_Combine(All_Exam_Score, All_Exam_Ave, All_Exam_Name, Class_Rank, Grade_Rank, Exam_Number)
elif Operate_Type == 1:
    Create_Combine_Compare(All_Exam_Name)

print("已完成，结果放于当前目录，正在退出程序...")
time.sleep(1)
