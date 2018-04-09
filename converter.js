var argv = require('minimist')(process.argv.slice(2));
// fileToOpen: -i
// fileNameToWrite: -o
// rowToStartIndexing: -r
var XLSX = require('xlsx');
var workbook = XLSX.readFile(argv.i+'.xlsx'); // 'example.xlsx');


/* ---------------------------------------------------------------------------- */

var worksheet = workbook.Sheets['Opened Last Week'];
var obj = XLSX.utils.sheet_to_json(worksheet)
var data = []
var rowStart = 0
// data.push({"Requested Item":""})

// get worksheet row start number
for ( a = 0; a < obj.length; a++ ) {
  // console.log('a -- ' + JSON.stringify(obj[a]))
  if (obj[a].__EMPTY === 'Interaction ID') {
    // console.log('row to begin is ' + a)
    rowStart = (a+1)
  }
}

for( x = rowStart; x < obj.length; x++ ){
  // console.log('============'+(x-argv.r)+'==============')
  // console.log(obj[x])
  var col9Data = obj[x].__EMPTY_9;   // console.log(col9Data)
  var col9SplitArr = obj[x].__EMPTY_9.split('\n')
  // console.log('# items in big field: '+col9SplitArr.length)

  var dataObj = {}
  var complaintString = ""
  var complaintIndex = 999
  for( y = 0; y < col9SplitArr.length; y++ ){
    var key = col9SplitArr[y].split(':')[0]
    var val = col9SplitArr[y].split(':')[1]
    if(key == 'What is the complaint?' || key == 'What is the feedback / enhancement?'){
      //console.log(y + ' - EQUALS key: '+key+', val: '+val)
      complaintIndex = y
      complaintString += val
    }
    else if(y > complaintIndex){
      // console.log(y + ' - CHECKING.. key: '+key+', val: '+val)
      if(key){
        //console.log(y + ' - ..ADDING key: '+key+', val: '+val)
        complaintString += key
      }
    }
    else{
      // console.log(y + ' - key: '+key+', val: '+val)
      if (key == 'Is this a complaint?') {
        key = 'Type of Feedback'
        // console.log('yes'+val)
        var valSplit = val.split('Yes')
        // console.log('1- val.length: '+val.length)
        if (val.length == 4) {
          val = 'Feedback'
        }
        else {
          // console.log('weeeee')
          val = 'Complaint'
        }
        dataObj[key] = val
      }
      else if (key == 'Requested for') {
        key = 'Reporter\'s Name'
        dataObj[key] = val
      }
      else if (key == 'Requested For GID') {
        key = 'Reporter\'s Global ID'
        dataObj[key] = val
      }
      else if (key == 'Requested for Phone Number') {
        key = 'Reporter\'s Phone'
        dataObj[key] = val
      }
      else if (key == 'Reporter\'s Name') {
        key = 'User\'s Name'
        dataObj[key] = val
      }
      else if (key == 'Reporter\'s Contact Number') {
        key = 'User\'s Phone'
        dataObj[key] = val
      }
      else if (key == 'Master Customer Number') {
        key = 'Master Customer #'
        dataObj[key] = val
      }
      else if (key == 'Which site?') {
        key = 'Product'
        dataObj[key] = val
      }
    }
  }
  // console.log('complaintString: "' + complaintString+'"')
  dataObj['ID'] = obj[x].__EMPTY // __EMPTY: 'SD3167542',
  dataObj['Source'] = obj[x].__EMPTY_3 // __EMPTY_3: 'Self Service',
  dataObj['Opened Date'] = obj[x].__EMPTY_4 // __EMPTY_4: '2/16/2018',
  dataObj['Product'] = obj[x].__EMPTY_8 // __EMPTY_8: 'Ferguson.com (WNA)',

  dataObj['Description'] = complaintString

  dataObj['Requested By'] = obj[x].__EMPTY_2 // __EMPTY_2: 'Associate\r\n',


  // __EMPTY_6: 'Closed',
  // __EMPTY_7: 'Feedback / Enhancement / Complaint',


  data.push(dataObj)
}
// console.log(data)

const wb = { SheetNames: [], Sheets: {} };
wb.Props = {
   Title: "FOL / FEI Enhancement Request Review",
   Author: "Matthew Evert"
};

var headers = [
      'ID',
      'Source',
      'Opened Date',
      'Product',
      'Type of Feedback',
      'Area(s) Affected',
      'Activity',
      'Tags',
      'Description',
      'Notes',
      'Feedback Loop Status',
      'Journey',
      'Experience',
      'Magnitude',
      'Requested By',
      'Reporter\'s Name',
      'Reporter\'s Email',
      'Reporter\'s Global ID',
      'Reporter\'s Phone',
      'User\'s Name',
      'User\'s Email',
      'User\'s Phone',
      'Master Customer #',
      'Tier',
      'KOB',
      'Region',
      'Logon/Branch',
      'Device',
      'Browser'
    ]
// create sheet data & add to workbook
var ws = XLSX.utils.json_to_sheet(data, {header:headers});
var wscols = [
    {wch:15}, // 'ID',
    {wch:15}, // 'Source',
    {wch:15}, // 'Opened Date',
    {wch:20}, // 'Product',
    {wch:20}, // 'Type of Feedback',
    {wch:15}, // 'Area(s) Affected',
    {wch:15}, // 'Activity',
    {wch:15}, // 'Tags',
    {wch:150}, // 'Description',
    {wch:75}, // 'Notes',
    {wch:15}, // 'Feedback Loop Status',
    {wch:15}, // 'Journey',
    {wch:15}, // 'Experience',
    {wch:15}, // 'Magnitude',
    {wch:15}, // 'Requested By',
    {wch:25}, // 'Reporter\'s Name',
    {wch:25}, // 'Reporter\'s Email',
    {wch:15}, // 'Reporter\'s Global ID',
    {wch:25}, // 'Reporter\'s Phone',
    {wch:25}, // 'User\'s Name',
    {wch:25}, // 'User\'s Email',
    {wch:25}, // 'User\'s Phone',
    {wch:15}, // 'Master Customer #',
    {wch:15}, // 'Tier',
    {wch:15}, // 'KOB',
    {wch:15}, // 'Region',
    {wch:15}, // 'Logon/Branch',
    {wch:15}, // 'Device',
    {wch:15}  // 'Browser'
];
ws['!cols'] = wscols;
var ws_name = "Opened Last Week";
XLSX.utils.book_append_sheet(wb, ws, ws_name);

/* ---------------------------------------------------------------------------- */

var worksheet2 = workbook.Sheets['Outstanding'];
var obj2 = XLSX.utils.sheet_to_json(worksheet2)
var data2 = []
var rowStart2 = 0
// get worksheet row start number
for ( b = 0; b < obj2.length; b++ ) {
  // console.log('a -- ' + JSON.stringify(obj[a]))
  if (obj2[b].__EMPTY === 'Interaction ID') {
    // console.log('row to begin is ' + a)
    rowStart2 = (a+1)
  }
}

for( y = rowStart2; y < obj2.length; y++ ){
  // console.log('============'+(x-argv.r)+'==============')
  // console.log(obj[x])
  var col9Data2 = obj2[y].__EMPTY_9;   // console.log(col9Data)
  var col9SplitArr2 = obj2[y].__EMPTY_9.split('\n')
  // console.log('# items in big field: '+col9SplitArr.length)

  var dataObj2 = {}
  var complaintString2 = ""
  var complaintIndex2 = 999
  for( z = 0; z < col9SplitArr2.length; z++ ){
    var key2 = col9SplitArr2[z].split(':')[0]
    var val2 = col9SplitArr2[z].split(':')[1]
    if(key2 == 'What is the complaint?' || key2 == 'What is the feedback / enhancement?'){
      //console.log(y + ' - EQUALS key: '+key+', val: '+val)
      complaintIndex2 = z
      complaintString2 += val2
    }
    else if(z > complaintIndex2){
      // console.log(y + ' - CHECKING.. key: '+key+', val: '+val)
      if(key2){
        //console.log(y + ' - ..ADDING key: '+key+', val: '+val)
        complaintString2 += key2
      }
    }
    else{
      // console.log(y + ' - key: '+key+', val: '+val)
      if (key2 == 'Is this a complaint?') {
        key2 = 'Type of Feedback'
        var valSplit2 = val2.split('Yes')
        // console.log('2- val.length: '+val2.length)
        if (val2.length == 4) {
          val2 = 'Feedback'
        }
        else {
          // console.log('weeeee')
          val2 = 'Complaint'
        }
        dataObj2[key2] = val2
      }
      else if (key2 == 'Requested for') {
        key2 = 'Reporter\'s Name'
        dataObj2[key2] = val2
      }
      else if (key2 == 'Requested For GID') {
        key2 = 'Reporter\'s Global ID'
        dataObj2[key2] = val2
      }
      else if (key2 == 'Requested for Phone Number') {
        key2 = 'Reporter\'s Phone'
        dataObj2[key2] = val2
      }
      else if (key2 == 'Reporter\'s Name') {
        key2 = 'User\'s Name'
        dataObj2[key2] = val2
      }
      else if (key2 == 'Reporter\'s Contact Number') {
        key2 = 'User\'s Phone'
        dataObj2[key2] = val2
      }
      else if (key2 == 'Master Customer Number') {
        key2 = 'Master Customer #'
        dataObj2[key2] = val2
      }
      else if (key2 == 'Which site?') {
        key2 = 'Product'
        dataObj2[key2] = val2
      }
    }
  }
  // console.log('complaintString: "' + complaintString+'"')
  dataObj2['ID'] = obj2[y].__EMPTY // __EMPTY: 'SD3167542',
  dataObj2['Source'] = obj2[y].__EMPTY_3 // __EMPTY_3: 'Self Service',
  dataObj2['Opened Date'] = obj2[y].__EMPTY_5 // __EMPTY_4: '2/16/2018',
  dataObj2['Product'] = obj2[y].__EMPTY_8 // __EMPTY_8: 'Ferguson.com (WNA)',

  dataObj2['Description'] = complaintString2

  dataObj2['Requested By'] = obj2[y].__EMPTY_2 // __EMPTY_2: 'Associate\r\n',


  // __EMPTY_6: 'Closed',
  // __EMPTY_7: 'Feedback / Enhancement / Complaint',


  data2.push(dataObj2)
}

// create sheet data & add to workbook
var ws2 = XLSX.utils.json_to_sheet(data2, {header:headers});
var wscols2 = [
    {wch:15}, // 'ID',
    {wch:15}, // 'Source',
    {wch:15}, // 'Opened Date',
    {wch:20}, // 'Product',
    {wch:20}, // 'Type of Feedback',
    {wch:15}, // 'Area(s) Affected',
    {wch:15}, // 'Activity',
    {wch:15}, // 'Tags',
    {wch:150}, // 'Description',
    {wch:75}, // 'Notes',
    {wch:15}, // 'Feedback Loop Status',
    {wch:15}, // 'Journey',
    {wch:15}, // 'Experience',
    {wch:15}, // 'Magnitude',
    {wch:15}, // 'Requested By',
    {wch:25}, // 'Reporter\'s Name',
    {wch:25}, // 'Reporter\'s Email',
    {wch:15}, // 'Reporter\'s Global ID',
    {wch:25}, // 'Reporter\'s Phone',
    {wch:25}, // 'User\'s Name',
    {wch:25}, // 'User\'s Email',
    {wch:25}, // 'User\'s Phone',
    {wch:15}, // 'Master Customer #',
    {wch:15}, // 'Tier',
    {wch:15}, // 'KOB',
    {wch:15}, // 'Region',
    {wch:15}, // 'Logon/Branch',
    {wch:15}, // 'Device',
    {wch:15}  // 'Browser'
];
ws2['!cols'] = wscols2;
var ws_name2 = "Outstanding";
XLSX.utils.book_append_sheet(wb, ws2, ws_name2);

// create file 'in memory'
var wbout = new Buffer(XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
XLSX.writeFile(wb, argv.i+'-converted.xlsx'); // 'outputTest2.xlsx');
