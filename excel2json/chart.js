const fs = require('fs');
const Excel = require('exceljs');
const wb = new Excel.Workbook();

function toChart(file, sheet, target, lengend = 'l1') {
    wb.xlsx.readFile(file).then(() => {
        const ws = wb.getWorksheet(sheet);
        const listData = ['x', 'data', 'legend'];
        const max_columns = ws.getColumn(1).values.length;

        for (let i = 2; i < max_columns; i++) {
            let item = [];
            ws.getRow(i).values.forEach((v, i) => {
                if (i === 1) {
                    item.push(new Date(v).toISOString().split('T')[1].slice(0, 5));
                } else if (i > 1) {
                    item.push(v);
                }
            });
            listData.push([...item, lengend]);
        }

        console.log(listData);

        fs.writeFile(target, JSON.stringify(listData), (err) => {
            if (err) throw err;
            // console.log('文件已保存！');
        });
    });
}

toChart('./chart/chart.xlsx', 'Sheet1', './chart/chart1.json');