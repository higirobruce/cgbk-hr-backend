var ADODB = require("node-adodb");
const express = require("express");
const app = express();
var bodyParser = require('body-parser')
const exportToExcel = require('export-to-excel')
var Excel = require('exceljs');
var workbook = new Excel.Workbook();
const path = require('path');
const download = require('download-pdf')
var moment = require('moment')

var fonts = {
  Roboto: {
    normal: 'fonts/Roboto-Regular.ttf',
    bold: 'fonts/Roboto-Medium.ttf',
    italics: 'fonts/Roboto-Italic.ttf',
    bolditalics: 'fonts/Roboto-MediumItalic.ttf'
  }
};

var PdfPrinter = require('pdfmake');
var printer = new PdfPrinter(fonts);
var fs = require('fs');
ADODB.debug = true;

// Connect to the MS Access DB

const port = 3670;
app.use('/pdf', express.static(__dirname + 'Lateness Report.pdf'));
app.use(bodyParser.urlencoded({ extended: false }))

// parse application/json
app.use(bodyParser.json())

app.get("/lateness_rpt", (req, res) => {
  console.info("Running Lateness Rpt...");
  let userid = req.query.userid ? req.query.userid : '*';
  let unitid = req.query.unitid ? req.query.unitid : '*';
  let freq = req.query.freq ? req.query.freq : 'undefined';
  // console.log(userid)
  getReport(userid, unitid, 1, freq, res);

});

app.get("/overtime_rpt", (req, res) => {
  console.info("Running Lateness Rpt...");
  let userid = req.query.userid ? req.query.userid : '*';
  let unitid = req.query.unitid ? req.query.unitid : '*';
  let freq = req.query.freq ? req.query.freq : 'undefined';
  // console.log(userid)
  getReport(userid, unitid, 2, freq, res);
});

app.get("/dailymvt_rpt", (req, res) => {
  console.info("Running Daily Mvt Rpt...");
  let userid = req.query.userid ? req.query.userid : '*';
  let unitid = req.query.unitid ? req.query.unitid : '*';
  let freq = req.query.freq ? req.query.freq : 'undefined';
  // console.log(userid)
  getReport(userid, unitid, 3, freq, res);
});

app.get("/pdf", (req, res) => {
  var file = fs.createReadStream("./Lateness Report.pdf");
  file.pipe(res);
});

app.get("/departments",(req,res)=>{
  getDepartments().then(data => {
    // console.log(query)
    if (data.length === 0) {
      res.send( {
        'data':[],
        'message':'Unsuccessfull'
      })
    } else {
      res.send({
        'data':data,
        'message':'Successfull'
      }) 
    }

  })
  .catch(e => {
    console.log(e);
    
  });
});

app.listen(port, () => {
  console.log(`Listening on port ${port}`);
});

// getDatesFromQuater();

function getDatesFromQuater() {
  let resp = {}
  let currentQuarter = moment().quarter();
  if (currentQuarter === 1) {
    let yr = moment().year() - 1
    let start_month = 10
    let end_month = 12

    resp = {
      'year': yr,
      'start_month': start_month,
      'end_month': end_month
    }
  } else if (currentQuarter === 2) {
    let yr = moment().year()
    let start_month = 1
    let end_month = 3

    resp = {
      'year': yr,
      'start_month': start_month,
      'end_month': end_month
    }
  } else if (currentQuarter === 3) {
    let yr = moment().year()
    let start_month = 4
    let end_month = 6

    resp = {
      'year': yr,
      'start_month': start_month,
      'end_month': end_month
    }
  } else if (currentQuarter === 4) {
    let yr = moment().year()
    let start_month = 7
    let end_month = 9

    resp = {
      'year': yr,
      'start_month': start_month,
      'end_month': end_month
    }
  }
  return resp;

}

function getReport(userid, unitid, type, freq, res) {
  //get current month and formated date
  let d = new Date();
  let da = d.getDate();
  let mon = d.getMonth() + 1;
  let yr = d.getFullYear();
  let d_format = '#' + mon + '/' + da + '/' + yr + '#';


  //build queries for different requests
  let query = ``;
  let subquery_uid = userid == '*' ? `` : `and (a.userid = '${userid}')`;
  let subquery_did = unitid == '*' ? `` : `and (b.user_group = ${unitid})`;

  if (type === 1) {
    query = `select cint(a.userid) as userid,b.name,b.lastname,b.user_group,c.gname,sum(a.shorthour)  from attendance a,[user] b,user_group c
    where a.userid=b.userid and b.user_group=c.id and month(a.date)='${mon}' and year(a.date)='${yr}'
    ${subquery_uid} ${subquery_did}
    group by a.userid,b.name,b.lastname,b.user_group,c.gname
    order by 1`;
  } else if (type === 2) {
    query = `select a.userid,b.name,b.lastname,b.user_group,c.gname,sum(a.workhour -9)  as overtime from attendance a,[user] b,user_group c
    where a.userid=b.userid and b.user_group=c.id and month(a.date)=${mon} and a.daytype='W'  and a.date < ${d_format}
    ${subquery_uid} ${subquery_did}
    group by a.userid,b.name,b.lastname,b.user_group,c.gname
    having sum(a.workhour -9) > 0
    `
  } else if (type === 3) {
    query = `SELECT cint(userid) as userid,checktime from auditdata where 
    attenddate=${d_format} order by 1,2`
  }
  var connection = ADODB.open(
    // "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Program Files (x86)\\FingerTec\\FingerTec TCMS V3\\TCMS V3;Persist Security Info=False;User id=ingress;Password=ingress;"
    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Program Files (x86)\\FingerTec\\FingerTec TCMS V3\\TCMS V3\\ingress.mdb;Jet OLEDB:Database Password=ingress;"
  );
  connection
    .query(
      query
    )
    .then(data => {
      // console.log(query)
      if (data.length === 0) {
        res.send(
          {
            'message': 'No Data found',
            'query': query
          });
        console.log('Length :' + data.length);
      } else {
        save(data, type);
        sendExcel(res, type);
      }

    })
    .catch(e => {
      console.log(e);
      res.send(e);
    });
}

function getDepartments(){
  let query=`SELECT id as [key],gname as [text],id as [value]
  FROM user_group
  WHERE id in (
  select distinct parentId from user_group
  )`;

  var connection = ADODB.open(
    // "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Program Files (x86)\\FingerTec\\FingerTec TCMS V3\\TCMS V3;Persist Security Info=False;User id=ingress;Password=ingress;"
    "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Program Files (x86)\\FingerTec\\FingerTec TCMS V3\\TCMS V3\\ingress.mdb;Jet OLEDB:Database Password=ingress;"
  );
  return connection
    .query(
      query
    )
    
}
function save(t, type) {
  // console.log(t);
  if (type === 1) {
    exportToExcel.exportXLSX({
      filename: "Lateness Report",
      sheetname: "Report",
      title: [
        {
          fieldName: "userid",
          displayName: "User ID",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "name",
          displayName: "Name",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "lastname",
          displayName: "Last_name",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "gname",
          displayName: "Department",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "Expr1005",
          displayName: "Lateness",
          cellWidth: 8
          // type: "datetime" // 1:是  0:否
        }
      ],
      data: t
    });
  } else if (type === 2) {
    exportToExcel.exportXLSX({
      filename: "Overtime Report",
      sheetname: "Report",
      title: [
        {
          fieldName: "userid",
          displayName: "User ID",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "name",
          displayName: "Name",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "lastname",
          displayName: "Last_name",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "gname",
          displayName: "Department",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "overtime",
          displayName: "Overtime",
          cellWidth: 8
          // type: "datetime" // 1:是  0:否
        }
      ],
      data: t
    });
  } else if (type === 3) {
    console.log(t[0]);
    let data = [];

    for (let i = 0; i < t.length; i++) {

      if ((i > 0 && t[i].userid !== t[i - 1].userid)) {
        let userId = t[i].userid;
        let checkTimes = [];

        checkTimes.push(moment(t[i].checktime).format('HH:mm:ss'));

        for (let j = 2; j <= 6; j++) {
          if (t.length >= i + j + 1) {
            if (t[i].userid === t[i + j].userid) {
              checkTimes.push(moment(t[i + j].checktime).format('HH:mm:ss'));
            } else {
              checkTimes.push('-');
            }
          } else {
            checkTimes.push('-');
          }
        }

        let obj =
        {
          'userId': userId,
          'checkTime1': checkTimes[0],
          'checkTime2': checkTimes[1],
          'checkTime3': checkTimes[2],
          'checkTime4': checkTimes[3],
          'checkTime5': checkTimes[4],
          'checkTime6': checkTimes[5]
        };
        data.push(obj);
      }

    }

    exportToExcel.exportXLSX({
      filename: "Daily Movement Report",
      sheetname: "Report",
      title: [
        {
          fieldName: "userId",
          displayName: "User ID",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "checkTime1",
          displayName: "Punch-In 1",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "checkTime2",
          displayName: "Punch-Out 1",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "checkTime3",
          displayName: "Punch-In 2",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "checkTime4",
          displayName: "Punch-Out 2",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "checkTime5",
          displayName: "Punch-In 3",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
        {
          fieldName: "checkTime6",
          displayName: "Punch-Out 3",
          cellWidth: 8
          // type: "bool" // 1:是  0:否
        },
      ],
      data: data
    });

  }
}

function sendExcel(res, type) {
  if (type === 1) {
    res.download('./Lateness Report.xlsx', () => {
      console.info('done')
    })
  } else if (type === 2) {
    res.download('./Overtime Report.xlsx', () => {
      console.info('done')
    })
  } else if (type === 3) {
    res.download('./Daily Movement Report.xlsx', () => {
      console.info('done')
    })
  }
}

async function sendWorkbook(data, response) {
  var fileName = 'Lateness Report.xlsx';
  let data_array = [];
  data_array.push(
    ['UserId', 'First Name', 'Last Name', 'Department', 'Lateness']
  )
  data.forEach(row => {
    let row_array = [];
    row_array.push(row.userid)
    row_array.push(row.name)
    row_array.push(row.lastname)
    row_array.push(row.gname)
    row_array.push(row.Expr1005)

    data_array.push(row_array);
  });
  console.log(data_array[0]);
  var docDefinition = {
    content: [
      { text: 'Lateness Report', style: 'header' },
      {
        style: 'tableExample',
        table: {
          body: data_array
        }
      }
    ]
  };

  var pdfDoc = printer.createPdfKitDocument(docDefinition);
  pdfDoc.pipe(fs.createWriteStream('./Lateness Report.pdf'));
  pdfDoc.end();


  var file = fs.createReadStream("./Lateness Report.pdf");
  file.pipe(response);

}



