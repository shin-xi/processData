import openpyxl
import pprint
import json


def toChart(sheet, legend):
    chartData = [['x', 'data', 'legend']]

    for i in range(2, 26):
        item = []
        for j in range(2):
            item.append(sheet.cell(row=i, column=j + 1).value)
        item[0] = item[0].strftime('%H:%M:%S')
        chartData.append(item + [legend])
    return chartData


wb = openpyxl.load_workbook('./chart/chart.xlsx')
sheet = wb['Sheet1']
pprint.pprint(toChart(sheet, 'l1'))

chart1 = open('./chart/chart1.json', 'w')
chart1.write(json.dumps(toChart(sheet, 'l1')))
