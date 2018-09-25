const Excel = require('exceljs');
const wb = new Excel.Workbook();

wb.xlsx.readFile('./list/list.xlsx').then((data) => {
    const ws = wb.getWorksheet('Sheet1');
    const header = ws.getRow(1).values;
    const listData = [];

    const max_row = ws.getRow(1).values.length;
    const max_columns = ws.getColumn(1).values.length;

    for (let i = 2; i < max_columns; i++) {
        // console.log(i);
        let item = {};
        ws.getRow(i).values.forEach((v, i) => {
            item[header[i]] = v;
        });
        listData.push(item);
    }

    console.log(listData);

    // console.log(ws.getCell('A1').value);
    // console.log(max_row, max_columns);
    // console.log(ws.getRow(1).values);
    // console.log(ws.getColumn(1).values);
    // console.log(ws.getColumn(1).values.length);
});