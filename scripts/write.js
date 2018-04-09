const XLSX = require('xlsx');
data = [
          { "agentNo":"324234", "subName":"30, Jul 2013 09:24 AM" },
          { "agentNo":"444443", "subName":"30, Jul 2013 09:24 AM" },
          { "agentNo2":"444444", "subName":"31, Jul 2013 09:24 AM" }
];

/* create workbook & set props*/
   const wb = { SheetNames: [], Sheets: {} };
   wb.Props = {
      Title: "Stats from app",
      Author: "John Doe"
   };

/*create sheet data & add to workbook*/
var ws = XLSX.utils.json_to_sheet(data);
var ws_name = "DataSheet 1";
XLSX.utils.book_append_sheet(wb, ws, ws_name);

/* create file 'in memory' */
var wbout = new Buffer(XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
XLSX.writeFile(wb, 'out.xlsb');
