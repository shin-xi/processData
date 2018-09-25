const fs = require('fs');
const Excel = require('exceljs');
const wb = new Excel.Workbook();

function toList(file, sheet, target) {
    wb.xlsx.readFile(file).then(() => {
        const ws = wb.getWorksheet(sheet);
        const header = ws.getRow(1).values;
        const listData = [];
        const max_columns = ws.getColumn(1).values.length;

        for (let i = 2; i < max_columns; i++) {
            let item = {};
            ws.getRow(i).values.forEach((v, i) => {
                item[header[i]] = v;
            });
            listData.push(item);
        }

        console.log(listData);

        fs.writeFile(target, JSON.stringify(listData), (err) => {
            if (err) throw err;
            // console.log('文件已保存！');
        });
    });
}

toList('./list/list.xlsx', 'Sheet1', './list/list1.json');