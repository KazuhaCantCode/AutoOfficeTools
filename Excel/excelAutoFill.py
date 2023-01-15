from openpyxl import Workbook
import easygui

#GPLv3.0，you can use it anywhere

#chinese and english gui and var ready
easygui.msgbox("欢迎使用eAF程序，该程序属于AutoOfficeTools的一部分，该程序应用于填充大量数据，没有的功能自己把代码Clone下来改就行，代码开放的，生成的xlsx文件默认名eAF。Welcome to the eAF program, which is part of AutoOfficeTools, which should be used to fill a large amount of data, without the function to change the code Clone down on the line, the code is open, and the default name of the generated xlsx file is eAF。")
fillData=easygui.enterbox("请输入单元格内填充的数据,指向变量fillData，Please enter the data filled in the cell, pointing to the variable fillData")
fillColumn=easygui.enterbox("请输入填充的列数，Value>=1，指向变量fillColumn，Please enter the number of columns to fill, Value>=1, pointing to the variable fillColumn")
fillRow=easygui.enterbox("请输入填充的行数，Value>=1，指向变量fillRow，Please enter the number of rows to fill, Value>=1, pointing to the variable fillRow")


#var ready
rIndex=1;
cIndex=1;
eAF_wb = Workbook()
eAF_ws = eAF_wb.active


#fill action
for i in range(int(fillRow)*int(fillColumn)):
    eAF_ws.cell(row=rIndex, column=cIndex).value = fillData;
    cIndex+=1;
    if(cIndex>int(fillColumn)):
        cIndex=1;
        rIndex+=1;
    if(rIndex>int(fillRow)):
        break


#save file
eAF_wb.save("eAF.xlsx")
