const XlsxPopulate = require('xlsx-populate');

function createXlsx(name, inputObj, columns) {
  if (!name || !inputObj) return;

  // Load a new blank workbook
  XlsxPopulate.fromBlankAsync()
    .then((workbook) => {
      const sheet1 = workbook.sheet('Sheet1');

      sheet1.column('A').width(24);
      sheet1.column('B').width(48);
      sheet1.column('C').width(48);

      const keys = Object.keys(inputObj);
      const values = Object.values(inputObj);

      console.log(`keys: ${keys}`);
      console.log(`values: ${values}`);

      const len = keys.length;
      console.log(`len: ${len}`);

      for (let i = 0; i < keys.length; i++) {
        console.log(`keys[i]: ${keys[i]}`);

        // Modify the workbook.
        sheet1.cell(`A${i + 1}`).value(keys[i]);

        const value = values[i];

        if (value[columns[0]]) {
          console.log(`value: ${JSON.stringify(value[columns[0]].join())}`);
          sheet1.cell(`B${i + 1}`).value(value[columns[0]].join());
        }

        if (value[columns[1]]) {
          sheet1.cell(`C${i + 1}`).value(value[columns[1]].join());
        }
      }

      // Write to file.
      return workbook.toFileAsync(`./${name}.xlsx`);
    })
    .catch((error) => {
      throw error;
    });
}

module.exports = createXlsx;
