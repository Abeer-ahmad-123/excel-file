var express = require('express');
var router = express.Router();
const fs = require('fs');
const path = require('path');
const Excel = require('exceljs');
const xl = require('excel4node');
const wb = new xl.Workbook();
let arr = []; /*read json report file and parse it*/
let data = []; /*read json products file and parse it*/

/* GET home page. */
router.get('/', function (req, res, next) {
  let pathForFile = path.join(__dirname);
  data = JSON.parse(
    fs.readFileSync(`${pathForFile}/products.json`, 'utf8', 'r')
  );
  arr = JSON.parse(fs.readFileSync(`${pathForFile}/report.json`, 'utf8', 'r'));

  console.log('arr, data =========', arr.length, data.length);
  // console.log(data);
  readfile(arr, data);
  res.status(200).send('Excel File Created');
});

module.exports = router;

function readfile(arr, data) {
  let products = [];
  let productAsins = [];

  for (let prod of data) {
    if (!productAsins.includes(prod.asin)) {
      products.push(prod);
      productAsins.push(prod.asin);
    }
  }

  let allAsins = [];
  let activeArray = [];
  let inactiveArray = [];
  let inCompleteaArray = [];
  for (let a of arr) {
    if (a.status === 'Inactive') inactiveArray.push(a.asin1);
    if (a.status === 'Incomplete') inCompleteaArray.push(a.asin1);
    if (a.status === 'Active') activeArray.push(a.asin1);
  }

  let actualInActiveArray = [];
  for (let elem of inactiveArray) {
    if (!activeArray.includes(elem)) actualInActiveArray.push(elem);
  }
  inactiveArray = [...actualInActiveArray];

  let parentsData = { list: [], keyValuePair: {} };
  let childrenData = { list: [], keyValuePair: {} };
  let anonymousData = { list: [], keyValuePair: {} };

  let duplicateData = {
    children: [],
    parent: [],
    anonymous: [],
    allAsins: [],
  };
  let inactiveData = {
    children: [],
    parent: [],
    anonymous: [],
    allAsins: [],
  };
  let activeData = { children: [], parent: [], anonymous: [], allAsins: [] };

  let brandsWithAsins = {};

  let i = 0;
  for (let prod of products) {
    if (prod.variations?.length) {
      let variation = prod.variations[0];
      if (variation.variationType === 'PARENT') {
        if (parentsData.keyValuePair[prod.asin]) {
          duplicateData.parent.push(prod.asin);
          duplicateData.allAsins.push(prod.asin);
        } else {
          if (inactiveArray.includes(prod.asin)) {
            inactiveData.allAsins.push(prod.asin);
            inactiveData.parent.push(prod.asin);
          } else {
            activeData.parent.push(prod.asin);
            activeData.allAsins.push(prod.asin);
          }
        }
        parentsData.list.push(prod.asin);
        parentsData.keyValuePair[prod.asin] = prod;
      } else {
        if (childrenData.keyValuePair[prod.asin]) {
          duplicateData.allAsins.push(prod.asin);
          duplicateData.children.push(prod.asin);
        } else {
          if (inactiveArray.includes(prod.asin)) {
            inactiveData.allAsins.push(prod.asin);
            inactiveData.children.push(prod.asin);
          } else {
            activeData.allAsins.push(prod.asin);
            activeData.children.push(prod.asin);
          }
        }
        childrenData.keyValuePair[prod.asin] = prod;
        childrenData.list.push(prod.asin);
      }
    } else {
      if (anonymousData.keyValuePair[prod.asin]) {
        duplicateData.allAsins.push(prod.asin);
        duplicateData.anonymous.push(prod.asin);
      } else {
        if (inactiveArray.includes(prod.asin)) {
          inactiveData.allAsins.push(prod.asin);
          inactiveData.anonymous.push(prod.asin);
        } else {
          activeData.allAsins.push(prod.asin);
          activeData.anonymous.push(prod.asin);
        }
      }
      anonymousData.list.push(prod.asin);
      anonymousData.keyValuePair[prod.asin] = prod;
    }
    allAsins.push(prod.asin);
  }

  for (let prod of products) {
    if (
      prod.summaries &&
      prod.summaries?.length &&
      prod.summaries[0].brandName
    ) {
      let brandName = prod.summaries[0].brandName;
      if (!brandsWithAsins[brandName]) {
        brandsWithAsins[brandName] = {
          allAsins: [],
          parents: [],
          inactiveParents: [],

          children: [],
          inactiveChildren: [],

          inactive: [],
          active: [],

          anonymous: [],
          inactiveAnonymous: [],
        };
      }

      let brandsData = brandsWithAsins[brandName];
      if (parentsData.list.includes(prod.asin)) {
        brandsData.parents.push(prod.asin);
        if (inactiveData.parent.includes(prod.asin)) {
          brandsData.inactiveParents.push(prod.asin);
        }
      }
      if (childrenData.list.includes(prod.asin)) {
        brandsData.children.push(prod.asin);
        if (inactiveData.children.includes(prod.asin)) {
          brandsData.inactiveChildren.push(prod.asin);
        }
      }
      if (anonymousData.list.includes(prod.asin)) {
        brandsData.anonymous.push(prod.asin);
        if (inactiveData.anonymous.includes(prod.asin)) {
          brandsData.inactiveAnonymous.push(prod.asin);
        }
      }
      if (activeData.allAsins.includes(prod.asin)) {
        brandsData.active.push(prod.asin);
      }
      if (inactiveData.allAsins.includes(prod.asin)) {
        brandsData.inactive.push(prod.asin);
      }
      brandsData.allAsins.push(prod.asin);
    } else {
      i++;
    }
  }

  let brandsDataStats = {};

  const createIndividualObject = (data) => {
    let obj = {};
    for (let key of Object.keys(data)) {
      obj[key] = data[key].length;
    }
    return obj;
  };

  for (var key of Object.keys(brandsWithAsins)) {
    brandsDataStats[key] = createIndividualObject(brandsWithAsins[key]);
  }

  console.log('total elements', data.length);
  console.log('total child products', childrenData.list.length);
  console.log('total parent products', parentsData.list.length);
  console.log('total anonymous products', anonymousData.list.length);

  console.log('children duplicateData', duplicateData.children.length);
  console.log('parent duplicateData', duplicateData.parent.length);
  console.log('anonymous duplicateData', duplicateData.anonymous.length);

  console.log('active parent length', activeData.parent.length);
  console.log('inActive parent length', inactiveData.parent.length);

  console.log('active children length', activeData.children.length);
  console.log('inActive children length', inactiveData.children.length);

  console.log('active anonymous length', activeData.anonymous.length);
  console.log('inActive anonymous length', inactiveData.anonymous.length);

  let active = [];
  let inactive = [];
  for (let d of arr) {
    if (d.status.toLowerCase() !== 'active') inactive.push(d);
    else active.push(d);
  }

  const formatActiveAndInActive = () => {
    let formattedActive = [];
    let formattedInActive = [];

    let keys = [
      'seller-sku',
      'item-name',
      'open-date',
      'asin1',
      'product-id',
      'status',
      'relation',
      'avg_landed_costs',
    ];
    let parentProductList = parentsData.list;
    const addRequiredColumns = (data) => {
      let formattedData = {};
      for (let index in keys) {
        let key = keys[index];
        formattedData[key] = data[key] || '';
      }
      formattedData['relation'] = parentProductList.includes(data.asin1)
        ? 'parent'
        : anonymousData.list.includes(data.asin1)
        ? 'missing'
        : 'child';
      return formattedData;
    };

    for (let p of inactive) {
      formattedInActive.push(addRequiredColumns(p));
    }
    for (let p of active) {
      formattedActive.push(addRequiredColumns(p));
    }
    return {
      formattedActive,
      formattedInActive,
    };
  };

  const { formattedActive, formattedInActive } = formatActiveAndInActive();
  console.log('formattedActive', formattedActive.length);
  console.log('formattedInActive', formattedInActive.length);

  // need to create a workbook object. Almost everything in ExcelJS is based off of the workbook object.
  let workbook = new Excel.Workbook();

  let worksheetActive = workbook.addWorksheet('Active');
  let worksheetInActive = workbook.addWorksheet('InActive');

  worksheetActive.columns = [
    { header: 'Sku', key: 'seller-sku' },
    { header: 'Item Name', key: 'item-name' },
    { header: 'Open Date', key: 'open-date' },
    { header: 'Asin', key: 'asin1' },
    { header: 'Product Id', key: 'product-id' },
    { header: 'Status', key: 'status' },
    { header: 'Relation', key: 'relation' },
    { header: 'Avg Landed Costs', key: 'avg_landed_costs' },
  ];
  worksheetInActive.columns = [
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
  worksheetActive.columns.forEach((column) => {
    if (column.header === 'Item Name') {
      column.width = column.header.length < 12 ? 150 : column.header.length;
    } else if (column.header === 'Sku') {
      column.width = column.header.length < 12 ? 33 : column.header.length;
    } else {
      column.width = column.header.length < 12 ? 12 : column.header.length;
    }
  });
  worksheetInActive.columns.forEach((column) => {
    if (column.header === 'Item Name') {
      column.width = column.header.length < 12 ? 140 : column.header.length;
    } else if (column.header === 'Sku') {
      column.width = column.header.length < 12 ? 30 : column.header.length;
    } else {
      column.width = column.header.length < 12 ? 12 : column.header.length;
    }
  });

  //

  // Make the header bold.
  // Note: in Excel the rows are 1 based, meaning the first row is 1 instead of 0.
  worksheetActive.getRow(1).font = { bold: true };
  worksheetInActive.getRow(1).font = { bold: true };

  // Dump all the data into Excel
  formattedActive.forEach((e) => {
    // By using destructuring we can easily dump all of the data into the row without doing much
    // We can add formulas pretty easily by providing the formula property.
    worksheetActive.addRow(e);
  });
  formattedInActive.forEach((e) => {
    // By using destructuring we can easily dump all of the data into the row without doing much
    // We can add formulas pretty easily by providing the formula property.
    worksheetInActive.addRow(e);
  });
  // Set the way columns C - F are formatted
  const figureColumns = [3, 4, 5, 6];
  figureColumns.forEach((i) => {
    worksheetActive.getColumn(i).numFmt = '$0.00';
    worksheetActive.getColumn(i).alignment = { horizontal: 'center' };
  });
  const figureColumnsInActive = [3, 4, 5, 6];
  figureColumnsInActive.forEach((i) => {
    worksheetInActive.getColumn(i).numFmt = '$0.00';
    worksheetInActive.getColumn(i).alignment = { horizontal: 'center' };
  });

  // loop through all of the rows and set the outline style.
  worksheetActive.eachRow({ includeEmpty: false }, function (row, rowNumber) {
    worksheetActive.getCell(`A${rowNumber}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'none' },
    };
    worksheetInActive.eachRow(
      { includeEmpty: false },
      function (row, rowNumber) {
        worksheetInActive.getCell(`A${rowNumber}`).border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'none' },
        };

        const insideColumns = ['B', 'C', 'D'];

        insideColumns.forEach((v) => {
          worksheetActive.getCell(`${v}${rowNumber}`).border = {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'none' },
            right: { style: 'none' },
          };
        });
      }
    );
    const insideColumns = ['B', 'C', 'D'];
    insideColumns.forEach((v) => {
      worksheetInActive.getCell(`${v}${rowNumber}`).border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'none' },
        right: { style: 'none' },
      };
    });

    const widthColumns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
    widthColumns.forEach((v) => {
      worksheetInActive.getCell(`${v}${rowNumber}`).style = { width: 120 };
    });
  });

  // Create a freeze pane, which means we'll always see the header as we scroll around.
  worksheetActive.views = [
    { state: 'frozen', xSplit: 0, ySplit: 1, activeCell: 'B2' },
  ];
  worksheetInActive.views = [
    { state: 'frozen', xSplit: 0, ySplit: 1, activeCell: 'B2' },
  ];

  // Keep in mind that reading and writing is promise based.
  workbook.xlsx.writeFile('Data.xlsx');
}
