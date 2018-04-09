var XLSX = require('xlsx');
var workbook = XLSX.readFile('example.xlsx');
/* Get worksheet */
var worksheet = workbook.Sheets['Opened Last Week'];
/* Find desired cell */
// var desired_cell = worksheet['A12'];
/* Get the value */
// var desired_value = (desired_cell ? desired_cell.v : undefined);
// console.log('desired_value: '+desired_value);

columnValues = {
  'A': 'Interaction ID',
  'C': 'Requested By',
  'D': 'Contact Medium',
  'E': 'Opened On',
  'G': 'Status',
  'H': 'Title',
  'I': 'Affected System',
  'J': 'Description'
}
/* data = [
  { "A":"324234" }
]; */

var data = []
for(z in worksheet) {
  var cell = worksheet[z];
  var value = (cell ? cell.v : undefined);
  console.log(z + ': ' + value);
  var
  // data2.push()
}





// console.log(columnValues.A)




/* create workbook & set props*/
   const wb = { SheetNames: [], Sheets: {} };
   wb.Props = {
      Title: "FOL / FEI Enhancement Request Review",
      Author: "Matthew Evert"
   };

/*create sheet data & add to workbook*/
var ws = XLSX.utils.json_to_sheet(data);
var ws_name = "Opened Last Week";
XLSX.utils.book_append_sheet(wb, ws, ws_name);

/* create file 'in memory' */
var wbout = new Buffer(XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
XLSX.writeFile(wb, 'outputTest1.xlsx');
