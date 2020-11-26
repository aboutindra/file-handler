const express = require('express');
const multer = require('multer');
const excelToJson = require("convert-excel-to-json");
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
    var result = excelToJson({
      sourceFile: req.file.path,
      header: {
        rows: 1
      },
      columnToKey: {
        '*': '{{columnHeader}}'
      }
    });
    var keys = Object.keys(result);
    console.log(keys);
    const hasil = [];
    for(let i = 0; i < keys.length; i++){
      hasil.push(result[keys[i]])
    }
    const data2 = { status:false, codeRoom: req.body.code_room }
    var finalData = [].concat.apply([], hasil);
    finalData.map(data => Object.assign(data, data2))
    res.send({"file" : req.file, "excel" : finalData});
  });
});
