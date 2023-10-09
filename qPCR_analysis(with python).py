import errno
from openpyxl.utils import get_column_letter
import openpyxl as excel
from openpyxl.styles import Alignment
import math
import statistics
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from openpyxl.chart.marker import DataPoint
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.styles.borders import Border, Side

#シート読み込み
sheet = excel.load_workbook("qpcr-data.xlsx", data_only=True).active
#ブック作成
book = excel.Workbook()
resultSheet = book.active

data = [] #元データ
data1 = [] #アクチン
data2 = [] #遺伝子1
data3 = [] #遺伝子2
data4 = [] #遺伝子3
data5 = [] #遺伝子4
data6 = [] #遺伝子5

for i in range(1,7):
    if i == 1:
        for j in range(3,27):
            value = sheet.cell(row=j,column=i).value
            if value != None:
                data1.append(value)
    if i == 2:
        for j in range(3,27):
            value = sheet.cell(row=j,column=i).value
            if value != None:
                data2.append(value)
    if i == 3:
        for j in range(3,27):
            value = sheet.cell(row=j,column=i).value
            if value != None:
                data3.append(value)
    if i == 4:
        for j in range(3,27):
            value = sheet.cell(row=j,column=i).value
            if value != None:
                data4.append(value)
    if i == 5:
        for j in range(3,27):
            value = sheet.cell(row=j,column=i).value
            if value != None:
                data5.append(value)
    if i == 6:
        for j in range(3,27):
            value = sheet.cell(row=j,column=i).value
            if value != None:
                data6.append(value)

#各サンプルの合計数をカウント
length1 = len(data1)
length2 = len(data2)
length3 = len(data3)
length4 = len(data4)
length5 = len(data5)
length6 = len(data6)

#入力してもらったサンプル数（N）を取得
n = sheet.cell(row=2,column=9).value
number_of_data = len(data1)/n
#遺伝子名を配列で取得
genes = []
for i in range(1,7):
    gene = sheet.cell(row=2,column=i).value
    if gene == None:
        gene = "未入力"
    genes.append(gene)

resultSheet.append (["計算("+genes[1]+")",genes[0],genes[1],"power","avepower","相対値PR1","相対値平均PR1","S.E.（誤差）"])
topLabels2 = ["計算("+genes[2]+")",genes[0],genes[2],"power","avepower","相対値PR1","相対値平均PR1","S.E.（誤差）"]
topLabels3 = ["計算("+genes[3]+")",genes[0],genes[3],"power","avepower","相対値PR1","相対値平均PR1","S.E.（誤差）"]
topLabels4 = ["計算("+genes[4]+")",genes[0],genes[4],"power","avepower","相対値PR1","相対値平均PR1","S.E.（誤差）"]
topLabels5 = ["計算("+genes[5]+")",genes[0],genes[5],"power","avepower","相対値PR1","相対値平均PR1","S.E.（誤差）"]

if data3[0] != None:
    for i, label in enumerate(topLabels2):
        resultSheet.cell(row=29,column=i+1,value=label)
if data4[0] != None:
    for i, label in enumerate(topLabels3):
        resultSheet.cell(row=58,column=i+1,value=label)
if data5[0] != None:
    for i, label in enumerate(topLabels4):
        resultSheet.cell(row=87,column=i+1,value=label)
if data6[0] != None:
    for i, label in enumerate(topLabels5):
        resultSheet.cell(row=116,column=i+1,value=label)

#各データの書き込み
# data1(act)
a = 2
for i ,data in enumerate(data1):
    resultSheet["B" + str(a)] = data
    a += 1
a = 30 
for i ,data in enumerate(data1):
    resultSheet["B" + str(a)] = data
    a += 1
a = 59 
for i ,data in enumerate(data1):
    resultSheet["B" + str(a)] = data
    a += 1
a = 88
for i ,data in enumerate(data1):
    resultSheet["B" + str(a)] = data
    a += 1
a = 117
for i ,data in enumerate(data1):
    resultSheet["B" + str(a)] = data
    a += 1
# data2(gene1)
a = 2
for i ,data in enumerate(data2):
    resultSheet["C" + str(a)] = data
    a+=1
# data3(gene2)
a = 30 
for i ,data in enumerate(data3):
    resultSheet["C" + str(a)] = data
    a += 1
# data4(gene3)
a = 59 
for i ,data in enumerate(data4):
    resultSheet["C" + str(a)] = data
    a += 1
# data5(gene4)
a = 88
for i ,data in enumerate(data5):
    resultSheet["C" + str(a)] = data
    a += 1
# data6(gene5)
a = 117
for i ,data in enumerate(data6):
    resultSheet["C" + str(a)] = data
    a += 1


#サンプルネームを取得し設定
sampleNames = []
for i in range(9,14):
    sampleName = sheet.cell(row=3,column=i).value
    if sampleName == None:
        sampleName = "未入力"
    sampleNames.append(sampleName)

a=2
for number , sampleName in enumerate(sampleNames): 
    if a > length1:
        break
    if a % 1 == 0 or a % n == 0:
        resultSheet.cell(row=a,column=1,value = sampleName)
        resultSheet.merge_cells(range_string="A"+str(a)+":"+"A"+str(n+a-1))
        resultSheet["A"+str(a)].alignment = Alignment(horizontal="center", vertical="center")
    a += n

a=2
for number , sampleName in enumerate(sampleNames): 
    if a > length1:
        break
    if a % 1 == 0 or a % n == 0:
        resultSheet.cell(row=a+28,column=1,value = sampleName)
        resultSheet.merge_cells(range_string="A"+str(a+28)+":"+"A"+str(n+a-1+28))
        resultSheet["A"+str(a+28)].alignment = Alignment(horizontal="center", vertical="center")
    a += n

a=2
for number , sampleName in enumerate(sampleNames): 
    if a > length1:
        break
    if a % 1 == 0 or a % n == 0:
        resultSheet.cell(row=a+57,column=1,value = sampleName)
        resultSheet.merge_cells(range_string="A"+str(a+57)+":"+"A"+str(n+a-1+57))
        resultSheet["A"+str(a+57)].alignment = Alignment(horizontal="center", vertical="center")
    a += n

a=2
for number , sampleName in enumerate(sampleNames): 
    if a > length1:
        break
    if a % 1 == 0 or a % n == 0:
        resultSheet.cell(row=a+86,column=1,value = sampleName)
        resultSheet.merge_cells(range_string="A"+str(a+86)+":"+"A"+str(n+a-1+86))
        resultSheet["A"+str(a+86)].alignment = Alignment(horizontal="center", vertical="center")
    a += n

a=2
for number , sampleName in enumerate(sampleNames): 
    if a > length1:
        break
    if a % 1 == 0 or a % n == 0:
        resultSheet.cell(row=a+115,column=1,value = sampleName)
        resultSheet.merge_cells(range_string="A"+str(a+115)+":"+"A"+str(n+a-1+115))
        resultSheet["A"+str(a+115)].alignment = Alignment(horizontal="center", vertical="center")
    a += n

resultSheet["J2"] = "結果1（相対値平均pr1）"
resultSheet["J3"] = sampleNames[0]
resultSheet["J4"] = sampleNames[1]
resultSheet["J5"] = sampleNames[2]
resultSheet["J6"] = sampleNames[3]
resultSheet["J7"] = sampleNames[4]
resultSheet["K2"] = genes[1]
resultSheet["L2"] = genes[2]
resultSheet["M2"] = genes[3]
resultSheet["N2"] = genes[4]
resultSheet["O2"] = genes[5]
resultSheet["Q2"] = "結果2（誤差）"
resultSheet["Q3"] = sampleNames[0]
resultSheet["Q4"] = sampleNames[1]
resultSheet["Q5"] = sampleNames[2]
resultSheet["Q6"] = sampleNames[3]
resultSheet["Q7"] = sampleNames[4]
resultSheet["R2"] = genes[1]
resultSheet["S2"] = genes[2]
resultSheet["T2"] = genes[3]
resultSheet["U2"] = genes[4]
resultSheet["V2"] = genes[5]

#power値を算出
#gene1
for i in range(length2):
    target = resultSheet.cell(row=i+2,column=3).value
    housekeeping = resultSheet.cell(row=i+2,column=2).value

    if target == "N/A" or housekeeping == "N/A" :
        continue

    power = 2 ** (-(target- housekeeping))
    resultSheet.cell(row= (i+2), column=4,value=power)
#gene2
for i in range(length3):
    target = resultSheet.cell(row=i+30,column=3).value
    housekeeping = resultSheet.cell(row=i+30,column=2).value

    if target == "N/A" or housekeeping == "N/A" :
        continue

    power = 2 ** (-(target- housekeeping))
    resultSheet.cell(row= (i+30), column=4,value=power)
#gene3
for i in range(length4):
    target = resultSheet.cell(row=i+59,column=3).value
    housekeeping = resultSheet.cell(row=i+59,column=2).value

    if target == "N/A" or housekeeping == "N/A" :
        continue

    power = 2 ** (-(target- housekeeping))
    resultSheet.cell(row= (i+59), column=4,value=power)
#gene4
for i in range(length5):
    target = resultSheet.cell(row=i+88,column=3).value
    housekeeping = resultSheet.cell(row=i+88,column=2).value

    if target == "N/A" or housekeeping == "N/A" :
        continue

    power = 2 ** (-(target- housekeeping))
    resultSheet.cell(row= (i+88), column=4,value=power)
#gene5
for i in range(length6):
    target = resultSheet.cell(row=i+117,column=3).value
    housekeeping = resultSheet.cell(row=i+117,column=2).value

    if target == "N/A" or housekeeping == "N/A" :
        continue

    power = 2 ** (-(target- housekeeping))
    resultSheet.cell(row= (i+117), column=4,value=power)


#avepower
# data2(gene1)
a=2
for number , sample in enumerate(data2): 
    if a > length2:
        break
    # 計算処理 aからn個下までの列の数値から平均を算出
    if a % 1 == 0 or a % n == 0:
        sum = 0
        copy_n = n
        for i in range(n):
            power_i = resultSheet.cell(row= a+i,column=4).value
            if power_i == None:
                copy_n -= 1
                continue
            else:
                sum += power_i 
            avepower = sum/copy_n
            resultSheet.cell(row=a,column=5,value = avepower)
    a += n
# data3(gene2)
a=2
for number , sample in enumerate(data3): 
    if a > length3:
        break
    if a % 1 == 0 or a % n == 0:
        sum = 0
        copy_n = n
        for i in range(n):
            power_i = resultSheet.cell(row= a + i + 28,column=4).value
            if power_i == None:
                copy_n -= 1
                continue
            else:
                sum += power_i 
            avepower = sum/copy_n
            resultSheet.cell(row=a+28,column=5,value = avepower)
    a += n
# data4(gene3)
a=2
for number , sample in enumerate(data4): 
    if a > length4:
        break
    if a % 1 == 0 or a % n == 0:
        sum = 0
        copy_n = n
        for i in range(n):
            power_i = resultSheet.cell(row= a + i + 57,column=4).value
            if power_i == None:
                copy_n -= 1
                continue
            else:
                sum += power_i 
            avepower = sum/copy_n
            resultSheet.cell(row=a+57,column=5,value = avepower)
    a += n
# data5(gene4)
a=2
for number , sample in enumerate(data5): 
    if a > length5:
        break
    if a % 1 == 0 or a % n == 0:
        sum = 0
        copy_n = n
        for i in range(n):
            power_i = resultSheet.cell(row= a + i + 86,column=4).value
            if power_i == None:
                copy_n -= 1
                continue
            else:
                sum += power_i 
            avepower = sum/copy_n
            resultSheet.cell(row=a+86,column=5,value = avepower)
    a += n
# data6(gene5)
a=2
for number , sample in enumerate(data6): 
    if a > length6:
        break
    if a % 1 == 0 or a % n == 0:
        sum = 0
        copy_n = n
        for i in range(n):
            power_i = resultSheet.cell(row= a + i + 115,column=4).value
            if power_i == None:
                copy_n -= 1
                continue
            else:
                sum += power_i 
            avepower = sum/copy_n
            resultSheet.cell(row=a+115,column=5,value = avepower)
    a += n

#相対値PR1値を算出
# data2(gene1)
for i in range(length2):
    # power取得
    power = resultSheet.cell(row= (i+2), column=4).value
    if power == None:
        continue
    else:
        # 各pr1と比較するbase_avepower取得
        base_avepower = resultSheet.cell(row= 2, column=5).value
        #計算結果を出力
        pr1 = power/base_avepower
        resultSheet.cell(row=i+2,column=6,value = pr1)
# data3(gene2)
for i in range(length3):
    # power取得
    power = resultSheet.cell(row= (i+30), column=4).value
    if power == None:
        continue
    else:
        # 各pr1と比較するbase_avepower取得
        base_avepower = resultSheet.cell(row= 30, column=5).value
        #計算結果を出力
        pr1 = power/base_avepower
        resultSheet.cell(row=i+30,column=6,value = pr1)
# data4(gene3)
for i in range(length4):
    # power取得
    power = resultSheet.cell(row= (i+59), column=4).value
    if power == None:
        continue
    else:
        # 各pr1と比較するbase_avepower取得
        base_avepower = resultSheet.cell(row= 59, column=5).value
        #計算結果を出力
        pr1 = power/base_avepower
        resultSheet.cell(row=i+59,column=6,value = pr1)
# data5(gene4)
for i in range(length5):
    # power取得
    power = resultSheet.cell(row= (i+88), column=4).value
    if power == None:
        continue
    else:
        # 各pr1と比較するbase_avepower取得
        base_avepower = resultSheet.cell(row= 88, column=5).value
        #計算結果を出力
        pr1 = power/base_avepower
        resultSheet.cell(row=i+88,column=6,value = pr1)
# data6(gene5)
for i in range(length5):
    # power取得
    power = resultSheet.cell(row= (i+117), column=4).value
    if power == None:
        continue
    else:
        # 各pr1と比較するbase_avepower取得
        base_avepower = resultSheet.cell(row= 117, column=5).value
        #計算結果を出力
        pr1 = power/base_avepower
        resultSheet.cell(row=i+117,column=6,value = pr1)

#相対値平均PR1算出
#data2(gene1)
c=2
for number , sample in enumerate(data2): 
    if c > length2:
        break
    # 計算処理 aからn個下までの列の数値から平均を算出
    if c % 1 == 0 or c % n == 0:
        sum = 0
        copy_n = n
        for i in range(n):
            pr1 = resultSheet.cell(row= c+i,column=6).value
            if pr1 == None:
                copy_n -= 1
                continue
            else:
                sum += pr1
            avepr1 = sum/copy_n
            resultSheet.cell(row=c,column=7,value = avepr1)
    c += n
#求めた相対値平均PR1を要約スペースにも書き込み
e = 2
for i in range(n-1):
    avepr1 = resultSheet. cell(row=e,column=7).value
    resultSheet.cell(row=i+3,column=11,value=avepr1)
    e += n

#data3(gene2)
c=2
for number , sample in enumerate(data3): 
    if c > length3:
        break
    # 計算処理 aからn個下までの列の数値から平均を算出
    if c % 1 == 0 or c % n == 0:
        sum = 0
        copy_n = n
        for i in range(n):
            pr1 = resultSheet.cell(row= c+i+28,column=6).value
            if pr1 == None:
                copy_n -= 1
                continue
            else:
                sum += pr1
            avepr1 = sum/copy_n
            resultSheet.cell(row=c+28,column=7,value = avepr1)
    c += n
#求めた相対値平均PR1を要約スペースにも書き込み
e = 30
for i in range(n-1):
    avepr1 = resultSheet. cell(row=e,column=7).value
    resultSheet.cell(row=i+3,column=12,value=avepr1)
    e += n

#data4(gene3)
c=2
for number , sample in enumerate(data4): 
    if c > length4:
        break
    # 計算処理 aからn個下までの列の数値から平均を算出
    if c % 1 == 0 or c % n == 0:
        sum = 0
        copy_n = n
        for i in range(n):
            pr1 = resultSheet.cell(row= c+i+57,column=6).value
            if pr1 == None:
                copy_n -= 1
                continue
            else:
                sum += pr1
            avepr1 = sum/copy_n
            resultSheet.cell(row=c+57,column=7,value = avepr1)
    c += n
#求めた相対値平均PR1を要約スペースにも書き込み
e = 59
for i in range(n-1):
    avepr1 = resultSheet. cell(row=e,column=7).value
    resultSheet.cell(row=i+3,column=13,value=avepr1)
    e += n

#data5(gene4)
c=2
for number , sample in enumerate(data5): 
    if c > length5:
        break
    # 計算処理 aからn個下までの列の数値から平均を算出
    if c % 1 == 0 or c % n == 0:
        sum = 0
        copy_n = n
        for i in range(n):
            pr1 = resultSheet.cell(row= c+i+86,column=6).value
            if pr1 == None:
                copy_n -= 1
                continue
            else:
                sum += pr1
            avepr1 = sum/copy_n
            resultSheet.cell(row=c+86,column=7,value = avepr1)
    c += n
#求めた相対値平均PR1を要約スペースにも書き込み
e = 88
for i in range(n-1):
    avepr1 = resultSheet. cell(row=e,column=7).value
    resultSheet.cell(row=i+3,column=14,value=avepr1)
    e += n

#data6(gene5)
c=2
for number , sample in enumerate(data6): 
    if c > length6:
        break
    # 計算処理 aからn個下までの列の数値から平均を算出
    if c % 1 == 0 or c % n == 0:
        sum = 0
        copy_n = n
        for i in range(n):
            pr1 = resultSheet.cell(row= c+i+115,column=6).value
            if pr1 == None:
                copy_n -= 1
                continue
            else:
                sum += pr1
            avepr1 = sum/copy_n
            resultSheet.cell(row=c+115,column=7,value = avepr1)
    c += n
#求めた相対値平均PR1を要約スペースにも書き込み
e = 117
for i in range(n-1):
    avepr1 = resultSheet. cell(row=e,column=7).value
    resultSheet.cell(row=i+3,column=15,value=avepr1)
    e += n

#誤算算出 data2はまだ値が入っている前提でしか作れていなかった？？？
#data2
a=2
for number , sample in enumerate(data2): 
    if a > length2:
        break
    if a % 1 == 0 or a % n == 0:
        pr1_row = []
        copy_n = n
        for i in range(n):
            pr1 = resultSheet.cell(row= a+i,column=6).value

            if pr1 == None:
                copy_n -= 1
                continue
            else:
                pr1_row.append(pr1)

        pr1_stdev = statistics.stdev(pr1_row)
        gosa = pr1_stdev / math.sqrt(copy_n)
        resultSheet.cell(row=a,column=8,value = gosa)
    a += n
# data3
a=2
for number , sample in enumerate(data3): 
    if a > length3:
        break
    if a % 1 == 0 or a % n == 0:
        pr1_row = []
        copy_n = n
        for i in range(n):
            pr1 = resultSheet.cell(row= a+i+28,column=6).value

            if pr1 == None:
                copy_n -= 1
                continue
            else:
                pr1_row.append(pr1)

        pr1_stdev = statistics.stdev(pr1_row)
        gosa = pr1_stdev / math.sqrt(copy_n)
        resultSheet.cell(row=a+28,column=8,value = gosa)
    a += n
# data4
a=2
for number , sample in enumerate(data4): 
    if a > length4:
        break
    if a % 1 == 0 or a % n == 0:
        pr1_row = []
        copy_n = n
        for i in range(n):
            pr1 = resultSheet.cell(row= a+i+57,column=6).value

            if pr1 == None:
                copy_n -= 1
                continue
            else:
                pr1_row.append(pr1)

        pr1_stdev = statistics.stdev(pr1_row)
        gosa = pr1_stdev / math.sqrt(copy_n)
        resultSheet.cell(row=a+57,column=8,value = gosa)
    a += n
# data5
a=2
for number , sample in enumerate(data5): 
    if a > length5:
        break
    if a % 1 == 0 or a % n == 0:
        pr1_row = []
        copy_n = n
        for i in range(n):
            pr1 = resultSheet.cell(row= a+i+86,column=6).value
            if pr1 == None:
                copy_n -= 1
                continue
            else:
                pr1_row.append(pr1)

        pr1_stdev = statistics.stdev(pr1_row)
        gosa = pr1_stdev / math.sqrt(copy_n)

        resultSheet.cell(row=a+86,column=8,value = gosa)
    a += n
# data6
a=2
for number , sample in enumerate(data6): 
    if a > length6:
        break
    if a % 1 == 0 or a % n == 0:
        pr1_row = []
        copy_n = n
        for i in range(n):
            pr1 = resultSheet.cell(row= a+i+115,column=6).value

            if pr1 == None:
                copy_n -= 1
                continue
            else:
                pr1_row.append(pr1)

        pr1_stdev = statistics.stdev(pr1_row)
        gosa = pr1_stdev / math.sqrt(copy_n)
        resultSheet.cell(row=a+115,column=8,value = gosa)
    a += n


#グラフ表示
chart = BarChart()
chart.width = 20
chart.height = 14
chart.title = "Result"              
chart.y_axis.title = '発現量'
chart.legend.position = 'b' 
chart_data = Reference(resultSheet,min_col=11, max_col= 9+len(genes) , min_row=2, max_row= 2+number_of_data)
chart_category= Reference(resultSheet,min_col=10, min_row=3, max_row= 2+number_of_data)
chart.add_data(chart_data, titles_from_data=True)
chart.set_categories(chart_category) 
chart.style = 7
chart.type = "col"
chart.grouping = "standard"
chart.gapwidth = 15
resultSheet.add_chart(chart,"J10")

# 罫線設定(行目よりn数ごとに分割)
# gene1
border = Border(bottom=Side(style='thin', color='000000'))
for number , sample in enumerate(data1):
    if (number + 1)  %  n  == 0:
        for j in range(2,9):
            resultSheet.cell(row= number + 2 ,column= j).border = border
# gene2
border = Border(bottom=Side(style='thin', color='000000'))
for number , sample in enumerate(data1):
    if (number + 1)  %  n  == 0:
        for j in range(2,9):
            resultSheet.cell(row= number + 30 ,column= j).border = border
# gene3
border = Border(bottom=Side(style='thin', color='000000'))
for number , sample in enumerate(data1):
    if (number + 1)  %  n  == 0:
        for j in range(2,9):
            resultSheet.cell(row= number + 59 ,column= j).border = border
# gene4
border = Border(bottom=Side(style='thin', color='000000'))
for number , sample in enumerate(data1):
    if (number + 1)  %  n  == 0:
        for j in range(2,9):
            resultSheet.cell(row= number + 88 ,column= j).border = border
# gene5
border = Border(bottom=Side(style='thin', color='000000'))
for number , sample in enumerate(data1):
    if (number + 1)  %  n  == 0:
        for j in range(2,9):
            resultSheet.cell(row= number + 117 ,column= j).border = border
            
#セルの幅調節
for col in resultSheet.columns:
    max_length = 0
    column = col[0].column

    for cell in col:
        if len(str(cell.value)) > max_length:
            max_length = len(str(cell.value))

    adjusted_width = (max_length + 2) * 1.2
    resultSheet.column_dimensions[get_column_letter(column)].width = adjusted_width

#ファイルの保存
book.save("result_qpcr.xlsx")
#ファイルオープン

#誤差の要約スペース　グラフへの反映