const obj = { title: '', freeze: 'A1', styles: [], merges: [], rows: {}, cols: {}, validations: [] };
const componentToHex = (c) => {
   var hex = c.toString(16);
   return hex.length == 1 ? "0" + hex : hex;
}

const getBackGroundInHex = (red, green, blue) => {
   return "#" + componentToHex(red) + componentToHex(green) + componentToHex(blue);
}
const getFormat = (rawData) => {
   let format;
   if (rawData && rawData.effectiveValue) {

      if (Object.keys(rawData.effectiveValue)[0] === 'numberValue') {

         if (rawData.effectiveFormat.numberFormat && rawData.effectiveFormat.numberFormat.type === 'NUMBER') {
            format = 'numberNoDecimal'
         }
         else if (rawData.effectiveFormat.numberFormat && rawData.effectiveFormat.numberFormat.type === 'PERCENT') {
            format = 'percentNoDecimal'
         }
      }
      // else if (Object.keys(rawData.effectiveValue)[0] === 'stringValue') {
      //    format = 'string'
      // }
   }

   return format;

}
const getTextFormatting = (textFormat) => {
   const requiredTextFormat = {};
   const defaultTextFormat =
   {
      fontFamily: 'Calibri',
      fontSize: 12, bold: false,
      italic: false,
      strikethrough: false,
      underline: false,
      foregroundColor: {},
      foregroundColorStyle: { rgbColor: {} }
   }
   const fontFormatFields = Object.keys(defaultTextFormat);

   for (const field of fontFormatFields) {

      if (typeof defaultTextFormat[field] !== 'object') {
         if (textFormat[field] !== defaultTextFormat[field]) {
            requiredTextFormat[field] = textFormat[field];
         }
      }
      else {
         if (JSON.stringify(textFormat[field]) !== JSON.stringify(defaultTextFormat[field])) {
            requiredTextFormat[field] = textFormat[field];
         }
      }
   }
   return requiredTextFormat;

};

const getStyle = (color, text, textFormat, rawData) => {

   const { red, green, blue } = color;
   const bgcolor = getBackGroundInHex(red, blue, green);
   const requiredTextFormat = getTextFormatting(textFormat);

   const format = getFormat(rawData);


   let style = {};
   if (format) {

      style.format = format;
   }

   if (Object.keys(requiredTextFormat).length > 0) {
      style.font = requiredTextFormat;
   }
   style.bgcolor = bgcolor;


   return style;

}
const toLowerCaseText = (userEnteredValue) => {
   let text = '';
   if (userEnteredValue === undefined) {
      return text;
   }
   text = userEnteredValue[Object.keys(userEnteredValue)[0]];
   if (Object.keys(userEnteredValue)[0] === 'stringValue') {
      return text;
   }


   if (typeof text === 'string') {
      text = text.toLowerCase();
   }
   if (text === undefined) {
      text = '';
   }
   return text;
}



const formatObject = ({ sheet, columns, rows }) => {
   let mapSize = 0;
   let styleIndex = 0;
   let maxColumnCharacters = Array(columns).fill(0);
   const map = new Map();


   obj.title = sheet.title;
   for (let i = 0; i < rows; i++) {
      obj['rows'][i.toString()] = {};
      for (let j = 0; j < columns; j++) {
         if (j == 0) {
            obj['rows'][i.toString()]['cells'] = {};
         }
         const cellInfo = sheet.getCell(i, j);
         const color = cellInfo._rawData.effectiveFormat.backgroundColor;


         const textFormat = cellInfo._rawData.effectiveFormat.textFormat;
         const lowerCaseText = toLowerCaseText(cellInfo._rawData.userEnteredValue);
         const style = getStyle(color, lowerCaseText, textFormat, cellInfo._rawData);

         const styleInString = JSON.stringify(style);

         if (map.has(styleInString)) {
            styleIndex = map.get(styleInString);

         }
         else {
            map.set(styleInString, mapSize);

            styleIndex = mapSize;
            mapSize++;

         }
         if (cellInfo._rawData.formattedValue) {

            maxColumnCharacters[j] = Math.max(cellInfo._rawData.formattedValue.length, maxColumnCharacters[j])
         }


         obj['rows'][i.toString()]['cells'][j.toString()] = {
            text: lowerCaseText, style: styleIndex
         }

      }
   }
   for (let [key, value] of map) {
      obj.styles[value] = JSON.parse(key);
   }
   for (let i = 0; i < columns; i++) {
      obj.cols[i] = {
         width: maxColumnCharacters[i] * 8
      }
   }

   return obj;
}

module.exports = { formatObject }