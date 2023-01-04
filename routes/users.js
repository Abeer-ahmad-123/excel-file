var express = require('express');
var router = express.Router();
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Worksheet Name');

/* GET users listing. */
router.get('/', function (req, res, next) {
  const data = [
    {
      firstName: 'John',
      lastName: 'Bailey',
      purchasePrice: 1000,
      paymentsMade: 100,
    },
    {
      firstName: 'Leonard',
      lastName: 'Clark',
      purchasePrice: 1000,
      paymentsMade: 150,
    },
    {
      firstName: 'Phil',
      lastName: 'Knox',
      purchasePrice: 1000,
      paymentsMade: 200,
    },
    {
      firstName: 'Sonia',
      lastName: 'Glover',
      purchasePrice: 1000,
      paymentsMade: 250,
    },
    {
      firstName: 'Adam',
      lastName: 'Mackay',
      purchasePrice: 1000,
      paymentsMade: 350,
    },
    {
      firstName: 'Lisa',
      lastName: 'Ogden',
      purchasePrice: 1000,
      paymentsMade: 400,
    },
    {
      firstName: 'Elizabeth',
      lastName: 'Murray',
      purchasePrice: 1000,
      paymentsMade: 500,
    },
    {
      firstName: 'Caroline',
      lastName: 'Jackson',
      purchasePrice: 1000,
      paymentsMade: 350,
    },
    {
      firstName: 'Kylie',
      lastName: 'James',
      purchasePrice: 1000,
      paymentsMade: 900,
    },
    {
      firstName: 'Harry',
      lastName: 'Peake',
      purchasePrice: 1000,
      paymentsMade: 1000,
    },
  ];

  const Excel = require('exceljs');

  // need to create a workbook object. Almost everything in ExcelJS is based off of the workbook object.
  let workbook = new Excel.Workbook();

  let worksheet = workbook.addWorksheet('Debtors');

  worksheet.columns = [
    { header: 'Sku', key: 'seller-sku' },
    { header: 'Item Name', key: 'item-name' },
    { header: 'Open Date', key: 'open-date' },
    { header: 'Asin', key: 'asin1' },
    { header: 'Product Id', key: 'product-id' },
    { header: 'Status', key: 'status' },
    { header: 'Relation', key: 'relation' },
    { header: 'Avg Landed Costs', key: 'avg_landed_costs' },
  ];

  // force the columns to be at least as long as their header row.
  // Have to take this approach because ExcelJS doesn't have an autofit property.
  worksheet.columns.forEach((column) => {
    column.width = column.header.length < 12 ? 12 : column.header.length;
  });

  // Make the header bold.
  // Note: in Excel the rows are 1 based, meaning the first row is 1 instead of 0.
  worksheet.getRow(1).font = { bold: true };

  // Dump all the data into Excel
  data.forEach((e) => {
    // By using destructuring we can easily dump all of the data into the row without doing much
    // We can add formulas pretty easily by providing the formula property.
    worksheet.addRow(e);
  });

  // Set the way columns C - F are formatted
  const figureColumns = [3, 4, 5, 6];
  figureColumns.forEach((i) => {
    worksheet.getColumn(i).numFmt = '$0.00';
    worksheet.getColumn(i).alignment = { horizontal: 'center' };
  });

  // loop through all of the rows and set the outline style.
  worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
    worksheet.getCell(`A${rowNumber}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'none' },
    };

    const insideColumns = ['B', 'C', 'D'];

    insideColumns.forEach((v) => {
      worksheet.getCell(`${v}${rowNumber}`).border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'none' },
        right: { style: 'none' },
      };
    });
  });
  // Create a freeze pane, which means we'll always see the header as we scroll around.
  worksheet.views = [
    { state: 'frozen', xSplit: 0, ySplit: 1, activeCell: 'B2' },
  ];

  // Keep in mind that reading and writing is promise based.
  workbook.xlsx.writeFile('Debtors.xlsx');
  res.status(200).json('File added successfully');
});

module.exports = router;
