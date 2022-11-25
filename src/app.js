const { GoogleSpreadsheet } = require('google-spreadsheet');
const creds = require('./secret.json');
const util = require('util');

const { formatObject } = require('./format-object');

const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1ZgxuK2-dEvCqWSBh9lwR9zdy3R9BkliFt4Pdf6VBtzI/edit#gid=0';
const id = SPREADSHEET_URL.replace('https://docs.google.com/spreadsheets/d/', '').replace('/edit#gid=0', '');

const accessSpreadSheet = async () => {
   const doc = new GoogleSpreadsheet(id);
   await doc.useServiceAccountAuth({
      client_email: creds.client_email,
      private_key: creds.private_key
   });
   await doc.loadInfo();
   const sheet = doc.sheetsByIndex[0];
   await sheet.loadHeaderRow();
   const rowsData = await sheet.getRows();
   let columns = sheet.headerValues.length, rows = rowsData.length + 1;
   const loadedCells = `A1:${String.fromCharCode(65 + columns - 1)}${rows}`;

   await sheet.loadCells(loadedCells);

   const ans = formatObject({ sheet, columns, rows });


   return ans;
}

accessSpreadSheet()

module.exports = { accessSpreadSheet };