const helpers = require('./helpers');
const express = require('express');
const multer = require('multer');
const path = require('path');
const excelToJson = require('convert-excel-to-json');
const fs = require('fs');

const app = express();
const port = process.env.PORT || 3000;
app.use(express.static('public'));

app.listen(port, () => console.log(`Listening on port ${port}...`));

app.post("/image", (req, res) => {
  var storage = multer.diskStorage({
    destination: function(req, file, cb) {
      cb(null, '../uploads');
    },
    filename: function (req, file, cb) {
      cb(null , file.originalname);
    }
  });
  const upload = multer({storage: storage, limits : {fileSize : 250000000}}).single("excel_file");

  upload(req, res, (err) => {
    console.log(req.body.code_room)

    if(err) {
      res.status(400).send("Something went wrong!\nmsg : "+ err);
    }
    var XLSX = require('xlsx');
    var workbook = XLSX.readFile(req.file.path);
    var sheet_name_list = workbook.SheetNames;
    var data = [];
    sheet_name_list.forEach(function(y) {
      var worksheet = workbook.Sheets[y];
      var headers = {};
      for(z in worksheet) {
        if(z[0] === '!') continue;
        //parse out the column, row, and value
        var tt = 0;
        for (var i = 0; i < z.length; i++) {
          if (!isNaN(z[i])) {
            tt = i;
            break;
          }
        };
        var col = z.substring(0,tt);
        var row = parseInt(z.substring(tt));
        var value = worksheet[z].v;

        //store header names
        if(row == 1 && value) {
          headers[col] = value;
          continue;
        }

        if(!data[row]) data[row]={};
        data[row][headers[col]] = value;
      }
      //drop those first two rows which are empty
      data.shift();
      data.shift();
      // console.log(data);
    });
    const data2 = { status:false, codeRoom: req.body.code_room }
    data.map(data => Object.assign(data, data2))
    res.send({"file" : req.file, "excel" : data});
  });
});
