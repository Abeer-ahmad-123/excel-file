var express = require('express');
var router = express.Router();
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Worksheet Name');

/* GET users listing. */
router.get('/', function (req, res, next) {
  res.send('Users');
});

module.exports = router;
