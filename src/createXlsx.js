const XlsxPopulate = require('xlsx-populate');

function createXlsx(name, inputObj) {
  if (!name || !inputObj) return;

  console.log(`Create ./${name}.xlsx`);

  // Load a new blank workbook
  XlsxPopulate.fromBlankAsync()
    .then((workbook) => {
      const sheet1 = workbook.sheet('Sheet1');

      sheet1.column('B').width(24);

      const keys = Object.keys(inputObj);
      const values = Object.values(inputObj);

      console.log(`keys: ${keys}`);
      console.log(`values: ${values}`);

      const len = keys.length;
      console.log(`len: ${len}`);

      for (let i = 0; i < keys.length; i++) {
        console.log(`keys[i]: ${keys[i]}`);
        console.log(`values[i]: ${values[i]}`);
        // Modify the workbook.
        sheet1.cell(`A${i + 1}`).value(keys[i]);
        sheet1.cell(`B${i + 1}`).value(values[i]);
      }

      // Write to file.
      return workbook.toFileAsync(`./${name}.xlsx`);
    })
    .catch((error) => {
      throw error;
    });
}

module.exports = createXlsx;
