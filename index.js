const express = require("express"); //Import the express dependency
const app = express(); //Instantiate an express app, the main work horse of this server
const port = 9000; //Save the port number where your server will be listening
const fs = require("fs");
const path = require("path");
const Excel = require("exceljs");
let arr = []; /*read json report file and parse it*/
let data = []; /*read json products file and parse it*/
let checkForValidFile;
let mainPath = path.join(__dirname, "data");

app.get("/", (req, res) => {
  fs.readdir(mainPath, function (err, folders) {
    //handling error
    if (err) {
      return console.log("Unable to scan directory: " + err);
    }

    const mainFolders = folders.filter((res) =>
      fs.lstatSync(path.resolve(mainPath, res)).isDirectory()
    );

    mainFolders.forEach(function (folder) {
      fs.readdir(path.join(mainPath, folder), function (err, files) {
        //listing all files using forEach
        console.log("folder", folder);

        files.forEach(function (file) {
          const extension = file.split(".").pop();
          checkForValidFile = extension === "json" ? true : false;
          if (checkForValidFile) {
            if (file === "products.json") {
              data = JSON.parse(
                fs.readFileSync(path.join(mainPath, folder, file), "utf8", "r")
              );
            }
            if (file === "reports.json") {
              arr = JSON.parse(
                fs.readFileSync(path.join(mainPath, folder, file), "utf8", "r")
              );
            }
          }
        });
        if (checkForValidFile) {
          readfile(folder, arr, data);
        }
      });
    });
  });

  res.status(200).send("Excel File Created");
});

function readfile(folder, arr, data) {
  let products = [...data];
  // let productAsins = [];

  // cost per iten against seller sku
  let landedCostData = {};

  // for (let prod of data) {
  //   if (!productAsins.includes(prod.asin)) {
  //     products.push(prod);
  //     productAsins.push(prod.asin);
  //   }
  // }

  let allAsins = [];

  let parentsData = { list: [], keyValuePair: {} };
  let childrenData = { list: [], keyValuePair: {} };
  let anonymousData = { list: [], keyValuePair: {} };

  let duplicateData = {
    children: [],
    parent: [],
    anonymous: [],
    all: [],
    allAsins: [],
  };
  let inactiveData = {
    children: [],
    parent: [],
    anonymous: [],
    all: [],
    allAsins: [],
  };
  let activeData = {
    children: [],
    parent: [],
    anonymous: [],
    all: [],
    allAsins: [],
  };
  let inCompleteData = {
    children: [],
    parent: [],
    anonymous: [],
    all: [],
    allAsins: [],
  };

  let brandsWithAsins = {};

  let i = 0;
  for (let prod of products) {
    landedCostData[prod.sellerSku] = prod.costPerItem;
    if (prod.variations?.length) {
      let variation = prod.variations[0];
      if (variation.variationType === "PARENT") {
        if (parentsData.keyValuePair[prod.asin]) {
          duplicateData.parent.push(prod.asin);
          duplicateData.all.push(prod);
          duplicateData.allAsins.push(prod.asin);
        }
        if (prod.status === "Inactive") {
          inactiveData.allAsins.push(prod.asin);
          inactiveData.parent.push(prod.asin);
          inactiveData.all.push(prod);
        } else if (prod.status === "Incomplete") {
          inCompleteData.allAsins.push(prod.asin);
          inCompleteData.parent.push(prod.asin);
          inCompleteData.all.push(prod);
        } else {
          activeData.parent.push(prod.asin);
          activeData.allAsins.push(prod.asin);
          activeData.all.push(prod);
        }
        parentsData.list.push(prod.asin);
        parentsData.keyValuePair[prod.asin] = prod;
      } else {
        if (childrenData.keyValuePair[prod.asin]) {
          duplicateData.allAsins.push(prod.asin);
          duplicateData.children.push(prod.asin);
          duplicateData.all.push(prod);
        }
        if (prod.status === "Inactive") {
          inactiveData.allAsins.push(prod.asin);
          inactiveData.children.push(prod.asin);
          inactiveData.all.push(prod);
        } else if (prod.status === "Incomplete") {
          inCompleteData.allAsins.push(prod.asin);
          inCompleteData.children.push(prod.asin);
          inCompleteData.all.push(prod);
        } else {
          activeData.allAsins.push(prod.asin);
          activeData.children.push(prod.asin);
          activeData.all.push(prod);
        }
        childrenData.list.push(prod.asin);
        childrenData.keyValuePair[prod.asin] = prod;
      }
    } else {
      if (anonymousData.keyValuePair[prod.asin]) {
        duplicateData.allAsins.push(prod.asin);
        duplicateData.anonymous.push(prod.asin);
        duplicateData.all.push(prod);
      }
      if (prod.status === "Inactive") {
        inactiveData.allAsins.push(prod.asin);
        inactiveData.anonymous.push(prod.asin);
        inactiveData.all.push(prod);
      } else if (prod.status === "Incomplete") {
        inCompleteData.allAsins.push(prod.asin);
        inCompleteData.anonymous.push(prod.asin);
        inCompleteData.all.push(prod);
      } else {
        activeData.allAsins.push(prod.asin);
        activeData.anonymous.push(prod.asin);
        activeData.all.push(prod);
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

  // console.log('total elements', data.length);
  // console.log('total child products', childrenData.list.length);
  // console.log('total parent products', parentsData.list.length);
  // console.log('total anonymous products', anonymousData.list.length);

  // console.log('children duplicateData', duplicateData.children.length);
  // console.log('parent duplicateData', duplicateData.parent.length);
  // console.log('anonymous duplicateData', duplicateData.anonymous.length);

  // console.log('active parent length', activeData.parent.length);
  // console.log('inActive parent length', inactiveData.parent.length);

  // console.log('active children length', activeData.children.length);
  console.log("inActive all length", inactiveData.all.length);
  console.log("active all length", activeData.all.length);
  console.log("in complete all length", inCompleteData.all.length);
  // console.log('inActive anonymous length', inactiveData.anonymous.length);

  // console.log('brandsWithAsins', brandsWithAsins);
  // console.log('brandsDataStats', brandsDataStats);
  let parentProductList = parentsData.list;
  const formatActiveAndInActive = () => {
    let formattedActive = [];
    let formattedInActive = [];
    let formattedInComplete = [];

    let keys = [
      "sellerSku",
      "itemName",
      "asin1",
      "productId",
      "relation",
      "costPerItem",
    ];

    const addRequiredColumns = (pData) => {
      let formattedData = {};
      for (let index in keys) {
        let key = keys[index];
        formattedData[key] = pData[key] || "";
      }
      formattedData["relation"] = parentProductList.includes(pData.asin1)
        ? "parent"
        : anonymousData.list.includes(pData.asin1)
        ? "missing"
        : "child";
      return formattedData;
    };

    for (let p of inactiveData.all) {
      formattedInActive.push(addRequiredColumns(p));
    }
    for (let p of activeData.all) {
      formattedActive.push(addRequiredColumns(p));
    }
    for (let p of inCompleteData.all) {
      formattedInComplete.push(addRequiredColumns(p));
    }
    return {
      formattedActive,
      formattedInActive,
      formattedInComplete,
    };
  };

  const { formattedActive, formattedInActive, formattedInComplete } =
    formatActiveAndInActive();

  console.log("formattedActive", formattedActive.length);
  console.log(
    "formattedInActive",
    formattedInActive.length,
    "inactiveData.all",
    inactiveData.all?.length
  );
  console.log("formattedInComplete", formattedInComplete.length);

  // need to create a workbook object. Almost everything in ExcelJS is based off of the workbook object.
  let workbook = new Excel.Workbook();

  let worksheetActive = workbook.addWorksheet("Active");
  let worksheetInActive = workbook.addWorksheet("InActive");
  let worksheetInComplete = workbook.addWorksheet("InComplete");

  worksheetActive.columns = [
    { header: "Sku", key: "sellerSku" },
    { header: "Item Name", key: "itemName" },
    { header: "Asin", key: "asin1" },
    { header: "Product Id", key: "productId" },
    { header: "Relation", key: "relation" },
    { header: "Avg Landed Costs", key: "costPerItem" },
  ];
  worksheetInActive.columns = [
    { header: "Sku", key: "sellerSku" },
    { header: "Item Name", key: "itemName" },
    { header: "Asin", key: "asin1" },
    { header: "Product Id", key: "productId" },
    { header: "Relation", key: "relation" },
    { header: "Avg Landed Costs", key: "costPerItem" },
  ];
  worksheetInComplete.columns = [
    { header: "Sku", key: "sellerSku" },
    { header: "Item Name", key: "itemName" },
    { header: "Asin", key: "asin1" },
    { header: "Product Id", key: "productId" },
    { header: "Relation", key: "relation" },
    { header: "Avg Landed Costs", key: "costPerItem" },
  ];

  // force the columns to be at least as long as their header row.
  // Have to take this approach because ExcelJS doesn't have an autofit property.
  worksheetActive.columns.forEach((column) => {
    if (column.header === "Item Name") {
      column.width = column.header.length < 12 ? 160 : column.header.length;
    } else if (column.header === "Sku") {
      column.width = column.header.length < 12 ? 33 : column.header.length;
    } else if (column.header === "Open Date") {
      column.width = column.header.length < 12 ? 21 : column.header.length;
    } else {
      column.width = column.header.length < 12 ? 12 : column.header.length;
    }
  });
  worksheetInActive.columns.forEach((column) => {
    if (column.header === "Item Name") {
      column.width = column.header.length < 12 ? 150 : column.header.length;
    } else if (column.header === "Sku") {
      column.width = column.header.length < 12 ? 30 : column.header.length;
    } else if (column.header === "Open Date") {
      column.width = column.header.length < 12 ? 21 : column.header.length;
    } else {
      column.width = column.header.length < 12 ? 12 : column.header.length;
    }
  });
  worksheetInComplete.columns.forEach((column) => {
    if (column.header === "Item Name") {
      column.width = column.header.length < 12 ? 150 : column.header.length;
    } else if (column.header === "Sku") {
      column.width = column.header.length < 12 ? 30 : column.header.length;
    } else if (column.header === "Open Date") {
      column.width = column.header.length < 12 ? 21 : column.header.length;
    } else {
      column.width = column.header.length < 12 ? 12 : column.header.length;
    }
  });
  //

  // Make the header bold.
  // Note: in Excel the rows are 1 based, meaning the first row is 1 instead of 0.
  worksheetActive.getRow(1).font = { bold: true };
  worksheetInActive.getRow(1).font = { bold: true };
  worksheetInComplete.getRow(1).font = { bold: true };

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
  formattedInComplete.forEach((e) => {
    // By using destructuring we can easily dump all of the data into the row without doing much
    // We can add formulas pretty easily by providing the formula property.
    worksheetInComplete.addRow(e);
  });
  // Set the way columns C - F are formatted
  const figureColumns = [3, 4, 5, 6];
  figureColumns.forEach((i) => {
    worksheetActive.getColumn(i).numFmt = "$0.00";
    worksheetActive.getColumn(i).alignment = { horizontal: "center" };
  });
  const figureColumnsInActive = [3, 4, 5, 6];
  figureColumnsInActive.forEach((i) => {
    worksheetInActive.getColumn(i).numFmt = "$0.00";
    worksheetInActive.getColumn(i).alignment = { horizontal: "center" };
  });

  const figureColumnsInComplete = [3, 4, 5, 6];
  figureColumnsInComplete.forEach((i) => {
    worksheetInComplete.getColumn(i).numFmt = "$0.00";
    worksheetInComplete.getColumn(i).alignment = { horizontal: "center" };
  });

  // loop through all of the rows and set the outline style.
  worksheetActive.eachRow({ includeEmpty: false }, function (row, rowNumber) {
    worksheetActive.getCell(`A${rowNumber}`).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "none" },
    };

    const insideColumns = ["B", "C", "D"];

    insideColumns.forEach((v) => {
      worksheetActive.getCell(`${v}${rowNumber}`).border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "none" },
        right: { style: "none" },
      };
    });
    const widthColumns = ["A", "B", "C", "D", "E", "F", "G", "H"];
    widthColumns.forEach((v) => {
      worksheetActive.getCell(`${v}${rowNumber}`).style = { width: 120 };
    });
  });

  worksheetInActive.eachRow({ includeEmpty: false }, function (row, rowNumber) {
    worksheetInActive.getCell(`A${rowNumber}`).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "none" },
    };

    const insideColumns = ["B", "C", "D"];

    insideColumns.forEach((v) => {
      worksheetInActive.getCell(`${v}${rowNumber}`).border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "none" },
        right: { style: "none" },
      };
    });
    const widthColumnsInActive = ["A", "B", "C", "D", "E", "F", "G", "H"];
    widthColumnsInActive.forEach((v) => {
      worksheetInActive.getCell(`${v}${rowNumber}`).style = { width: 120 };
    });
  });

  worksheetInComplete.eachRow(
    { includeEmpty: false },
    function (row, rowNumber) {
      worksheetInComplete.getCell(`A${rowNumber}`).border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "none" },
      };

      const insideColumns = ["B", "C", "D"];

      insideColumns.forEach((v) => {
        worksheetInComplete.getCell(`${v}${rowNumber}`).border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "none" },
          right: { style: "none" },
        };
      });
      const widthColumnsInComplete = ["A", "B", "C", "D", "E", "F", "G", "H"];
      widthColumnsInComplete.forEach((v) => {
        worksheetInComplete.getCell(`${v}${rowNumber}`).style = { width: 120 };
      });
    }
  );

  // Create a freeze pane, which means we'll always see the header as we scroll around.
  worksheetActive.views = [
    { state: "frozen", xSplit: 1, ySplit: 1, activeCell: "B2" },
  ];
  worksheetInActive.views = [
    { state: "frozen", xSplit: 1, ySplit: 1, activeCell: "B2" },
  ];
  worksheetInComplete.views = [
    { state: "frozen", xSplit: 1, ySplit: 1, activeCell: "B2" },
  ];

  // Keep in mind that reading and writing is promise based.
  workbook.xlsx.writeFile(`${mainPath}/${folder}/${folder}.xlsx`);
}

app.listen(port, () => {
  //server starts listening for any attempts from a client to connect at port: {port}
  console.log(`Now listening on port ${port}`);
});
