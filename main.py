#!/usr/bin/python
# -*- coding: UTF-8 -*-

from time import time
import xlrd, xlwt, os, xlutils,datetime,time,configparser;
from xlutils.copy import copy

formatTime = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())
def timeDiffer(time1, time2):
    time1 = datetime.datetime.strptime(time1,"%H:%M")
    procTime1 = time1.hour*60 + time1.minute
    time2 = datetime.datetime.strptime(time2,"%H:%M")
    procTime2 = time2.hour*60 + time2.minute
    # print(procTime1,procTime2)
    return procTime1-procTime2

#文件选项
inputFileName = "打卡日报.xlsx" #输入文件名
inputSheetName = "上下班打卡_日报" #输入文件中目标工作表名
outputFolder = "output" #输出文件夹
templateName = "template.xls" #输出模板文件夹
tempFileName = "temp.xls" #临时文件名
outputFileName = "考勤系统 " + formatTime + ".xls" #输出文件名，formatTime为时间戳
outputSheetWidth = 12 #输出工作表需要进行格式化的列数

# inputFileName = ""
# inputSheetName = ""
# outputFolder = ""
# templateName = ""
# tempFileName = ""
# outputFileName = "考勤系统 " + formatTime + ".xls"
# outputSheetWidth = 0

#格式选项
formatLines = 5 # 表中第一个员工前的行数
outFormatLines = 2 #输出表中第一个员工前的行数
NameColumn = 0
MorningCheckColumn = 0
EveningCheckColumn = 0
# formatLines = 0
# outFormatLines = 0

sheetStyle = xlwt.XFStyle()
#边框设置
sheetBorders = xlwt.Borders()
sheetBorders.left = xlwt.Borders.THIN
sheetBorders.right = xlwt.Borders.THIN
sheetBorders.top = xlwt.Borders.THIN
sheetBorders.bottom = xlwt.Borders.THIN
#对齐设置
sheetAlignment = xlwt.Alignment()
sheetAlignment.horz = xlwt.Alignment.HORZ_CENTER
sheetAlignment.vert = xlwt.Alignment.VERT_CENTER
#字体设置
sheetFont = xlwt.Font()
sheetFont.name = "宋体"
sheetFont.height = 20 * 12 #格式为 20*磅数

sheetStyle.borders = sheetBorders
sheetStyle.alignment = sheetAlignment
sheetStyle.font = sheetFont
#行高
rowHeight = 20 * 18.95 #格式为 20*磅数

#设置
MorningCheckTime = "08:00"
EveningCheckTime = "17:00"
MorningTimeLevel = [10, 20, 30]
MorningTimePoint = [5, 10, 15]

# MorningCheckTime = ""
# EveningCheckTime = ""
# MorningTimeLevel = []
# MorningTimePoint = []

cfgFile = configparser.ConfigParser();
cfgFile.read("config.ini",encoding="utf-8-sig");
#文件设置
currentLine = cfgFile["File Settings"]["inputFileName"]
inputFileName = currentLine
currentLine = cfgFile["File Settings"]["inputSheetName"]
inputSheetName = currentLine
currentLine = cfgFile["File Settings"]["outputFolder"]
outputFolder = currentLine
currentLine = cfgFile["File Settings"]["templateName"]
templateName = currentLine
currentLine = cfgFile["File Settings"]["tempFileName"]
tempFileName = currentLine
currentLine = cfgFile["File Settings"]["outputSheetWidth"]
outputSheetWidth = int(currentLine)

#格式设置
currentLine = cfgFile["Format Settings"]["formatLines"]
formatLines = int(currentLine)
currentLine = cfgFile["Format Settings"]["outFormatLines"]
outFormatLines = int(currentLine)
currentLine = cfgFile["Format Settings"]["NameColumn"]
NameColumn = int(currentLine) - 1
currentLine = cfgFile["Format Settings"]["MorningCheckColumn"]
MorningCheckColumn = int(currentLine) - 1
currentLine = cfgFile["Format Settings"]["EveningCheckColumn"]
EveningCheckColumn = int(currentLine) - 1
currentLine = cfgFile["Format Settings"]["outputNameColumn"]
outputNameColumn = int(currentLine) - 1
currentLine = cfgFile["Format Settings"]["outputBValueMorningColumn"]
outputBValueMorningColumn = int(currentLine) - 1
currentLine = cfgFile["Format Settings"]["outputBValueEveningColumn"]
outputBValueEveningColumn = int(currentLine) - 1

#时间设置
currentLine = cfgFile["Time Set"]["MorningCheckTime"]
MorningCheckTime = currentLine
currentLine = cfgFile["Time Set"]["EveningCheckTime"]
EveningCheckTime = currentLine

#早班打卡规则
currentLine = cfgFile["Morning Check Settings"]["MorningTimeLevel"];
MorningTimeLevel = currentLine.split(',')
currentLine = cfgFile["Morning Check Settings"]["MorningTimePoint"];
MorningTimePoint = currentLine.split(',')

currentLine = cfgFile["Evening Check Settings"]["EveningBasePoint"]
EveningBasePoint = int(currentLine)

#目录操作
os.system("mkdir " + outputFolder)

#打开工作表
INPworkbook = xlrd.open_workbook(inputFileName)
TEMPLATEworkbook = xlrd.open_workbook(templateName,formatting_info=True)
INPsheet = INPworkbook.sheet_by_name(inputSheetName)
TARGETworkbook = copy(TEMPLATEworkbook)
TARGETsheet = TARGETworkbook.get_sheet(0)
lineCount = INPsheet.nrows - formatLines;

employeeDatabaseMorningCheck = {}
employeeDatabaseEveningCheck = {}


MorningRuleLen = len(MorningTimeLevel)
for i in range(lineCount):
    employeeName = INPsheet.cell(formatLines+i,NameColumn).value
    employeeMorningCheckTime = INPsheet.cell(formatLines+i,MorningCheckColumn).value
    employeeEveningCheckTime = INPsheet.cell(formatLines+i,EveningCheckColumn).value
    bValMorning = 0
    bValEvening = 0
    
    if((employeeMorningCheckTime != "未打卡")and(employeeMorningCheckTime != "--")) :
        CheckTimeDiffer = timeDiffer(MorningCheckTime,employeeMorningCheckTime)
        for j in range(MorningRuleLen):
            if(j==MorningRuleLen-1):
                if(CheckTimeDiffer>=int(MorningTimeLevel[j])):
                    bValMorning += int(MorningTimePoint[j])
                    break
            else:
                if((CheckTimeDiffer>=int(MorningTimeLevel[j]))and(CheckTimeDiffer<int(MorningTimeLevel[j+1]))):
                    bValMorning += int(MorningTimePoint[j])
                    break
                
        # if((CheckTimeDiffer>=10)and(CheckTimeDiffer<20)):
        #     bVal += 5
        # if((CheckTimeDiffer>=20)and(CheckTimeDiffer<30)):
        #     bVal += 10
        # if(CheckTimeDiffer>=30):
        #     bVal += 15
    
    if((employeeEveningCheckTime != "未打卡")and(employeeEveningCheckTime != "--")) :
        CheckTimeDiffer = -timeDiffer(EveningCheckTime,employeeEveningCheckTime)
        if(CheckTimeDiffer>0):
            bValEvening += EveningBasePoint*(CheckTimeDiffer/60)
            
    #Existence Check
    if(employeeDatabaseMorningCheck.get(employeeName,-1)==-1):
        employeeDatabaseMorningCheck[employeeName] = 0
        employeeDatabaseEveningCheck[employeeName] = 0
    employeeDatabaseMorningCheck[employeeName] += bValMorning
    employeeDatabaseEveningCheck[employeeName] += bValEvening
    
# print(employeeDatabase)

employeeCount = 0
for employee in employeeDatabaseMorningCheck:
    TARGETsheet.row(outFormatLines+employeeCount)
    for i in range(outputSheetWidth):
        TARGETsheet.write(outFormatLines+employeeCount,i,"",sheetStyle)
    TARGETsheet.write(outFormatLines+employeeCount,outputNameColumn,employee,sheetStyle)
    TARGETsheet.write(outFormatLines+employeeCount,outputBValueMorningColumn,round(employeeDatabaseMorningCheck[employee]),sheetStyle)
    TARGETsheet.write(outFormatLines+employeeCount,outputBValueEveningColumn,round(employeeDatabaseEveningCheck[employee]),sheetStyle)
    employeeCount += 1
    # print(employee,employeeDatabase[employee]);
TARGETworkbook.save(tempFileName);


os.system("copy /Y " + tempFileName + " " + outputFolder)
os.system("rename " + outputFolder + "\\" + tempFileName + " " + "\"" + outputFileName + "\"")
os.system("del /F /Q " + tempFileName)

print("成功设置早班打卡" + str(MorningRuleLen) + "条规则")
print("成功读入" + str(INPsheet.nrows) + "条考勤信息")
print("共" + str(employeeCount) + "名员工考勤信息被录入")

print("文件将储存在" + os.path.abspath('.') + "\\" + outputFolder + "\\" + outputFileName)