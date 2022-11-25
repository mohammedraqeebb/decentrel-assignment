const { accessSpreadSheet } = require('./app');

//validation, merges,freeze,

test('testing final Object', async () => {
   const obj = await accessSpreadSheet();
   const { rows, cols, title, freeze, styles, merges, validations } = obj;

   expect(title).toBe('Sheet1');

   expect(Object.keys(rows).length).toBe(9);
   expect(Object.keys(cols).length).toBe(7);

   expect(styles).toEqual([
      { bgcolor: '#010101' },
      { font: { bold: true }, bgcolor: '#010101' },
      { format: 'numberNoDecimal', bgcolor: '#010101' },
      { format: 'percentNoDecimal', bgcolor: '#010101' }
   ]);

   expect(rows).toEqual({
      '0': {
         cells: {
            '0': { text: '', style: 0 },
            '1': { text: 'Common Stock', style: 1 },
            '2': { text: 'Preferred Stock', style: 1 },
            '3': { text: 'Stock Plan', style: 1 },
            '4': { text: 'Total Securities', style: 1 },
            '5': { text: 'Outstanding Ownership', style: 1 },
            '6': { text: 'Fully Diluted Ownership', style: 1 }
         }
      },
      '1': {
         cells: {
            '0': { text: 'Founders', style: 0 },
            '1': { text: 9250000, style: 2 },
            '2': { text: '', style: 0 },
            '3': { text: '', style: 0 },
            '4': { text: 9250000, style: 2 },
            '5': { text: '=e2/e5', style: 3 },
            '6': { text: '=e2/e9', style: 3 }
         }
      },
      '2': {
         cells: {
            '0': { text: 'Options outstanding', style: 0 },
            '1': { text: '', style: 0 },
            '2': { text: '', style: 0 },
            '3': { text: 300000, style: 2 },
            '4': { text: '=d3', style: 2 },
            '5': { text: '=e3/e5', style: 3 },
            '6': { text: '=e3/e9', style: 3 }
         }
      },
      '3': {
         cells: {
            '0': { text: 'Promised Options', style: 0 },
            '1': { text: '', style: 0 },
            '2': { text: '', style: 0 },
            '3': { text: 350000, style: 2 },
            '4': { text: '=d4', style: 2 },
            '5': { text: '=e4/e5', style: 3 },
            '6': { text: '=e4/e9', style: 3 }
         }
      },
      '4': {
         cells: {
            '0': { text: 'Total Outstanding', style: 1 },
            '1': { text: '=sum(b2:b4)', style: 2 },
            '2': { text: 0, style: 0 },
            '3': { text: '=sum(d2:d4)', style: 2 },
            '4': { text: '=sum(e2:e4)', style: 2 },
            '5': { text: '=sum(f2:f4)', style: 3 },
            '6': { text: '=sum(g2:g4)', style: 3 }
         }
      },
      '5': {
         cells: {
            '0': { text: '', style: 0 },
            '1': { text: '', style: 0 },
            '2': { text: '', style: 0 },
            '3': { text: '', style: 0 },
            '4': { text: '', style: 0 },
            '5': { text: '', style: 0 },
            '6': { text: '', style: 0 }
         }
      },
      '6': {
         cells: {
            '0': { text: 'Options Available', style: 0 },
            '1': { text: '', style: 0 },
            '2': { text: '', style: 0 },
            '3': { text: 100000, style: 2 },
            '4': { text: '=d7', style: 2 },
            '5': { text: '', style: 0 },
            '6': { text: '=e7/e9', style: 3 }
         }
      },
      '7': {
         cells: {
            '0': { text: '', style: 0 },
            '1': { text: '', style: 0 },
            '2': { text: '', style: 0 },
            '3': { text: '', style: 0 },
            '4': { text: '', style: 0 },
            '5': { text: '', style: 0 },
            '6': { text: '', style: 0 }
         }
      },
      '8': {
         cells: {
            '0': { text: 'Total Fully Diluted', style: 1 },
            '1': { text: '=b5', style: 2 },
            '2': { text: '', style: 0 },
            '3': { text: '=sum(d5:d8)', style: 2 },
            '4': { text: '=sum(e5:e8)', style: 2 },
            '5': { text: '', style: 0 },
            '6': { text: '=sum(g5:g8)', style: 3 }
         }
      }
   })

   expect(cols).toEqual({
      '0': { width: 152 },
      '1': { width: 96 },
      '2': { width: 120 },
      '3': { width: 80 },
      '4': { width: 128 },
      '5': { width: 168 },
      '6': { width: 184 }
   })


})