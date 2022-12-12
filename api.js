const express = require('express')
const app = express()
const cors = require('cors');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const xlsxFile = require('read-excel-file/node');

const workbook = XLSX.readFile("students.xlsx");
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

var students = [];

xlsxFile('./students.xlsx').then((rows) => {
  for (var i = 1; i < rows.length; i++) {
    students.push({ name: rows[i][0], age: rows[i][1] });
  }
});

app.use(cors({
  origin: 'https://mohamadalasaed.github.io/web0Project/'
}));
app.use(bodyParser.json());

app.get('/', function (req, res) {
  res.send('Hello World')
})

app.post('/addstudent', (req, res) => {
  console.log(req.body);
  XLSX.utils.sheet_add_aoa(worksheet, [[req.body.name, req.body.age]], { origin: -1 });
  XLSX.writeFile(workbook, "students.xlsx");
  students.push({ name: req.body.name, age: req.body.age });
  res.send(students);
})

app.get('/data', function (req, res) {
  res.send(students)
})

app.listen(3000)
