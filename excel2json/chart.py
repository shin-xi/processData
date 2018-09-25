import openpyxl
import pprint
import json


def toChart(file,sheet,target, legend):
    wb = openpyxl.load_workbook(file)
    sheet = wb[sheet]
    chartData = [['x', 'data', 'legend']]

    for i in range(2, 26):
        item = []
        for j in range(2):
            item.append(sheet.cell(row=i, column=j + 1).value)
        item[0] = item[0].strftime('%H:%M:%S')
        chartData.append(item + [legend])

    pprint.pprint(chartData)

    chart = open(target, 'w')
    chart.write(json.dumps(chartData))


toChart('./chart/chart.xlsx','Sheet1','./chart/chart1.json','l2')