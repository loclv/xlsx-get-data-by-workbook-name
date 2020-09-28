const XlsxPopulate = require('xlsx-populate');

function createXlsx(name) {
  console.log(`Create ./${name}.xlsx`);

  // Load a new blank workbook
  XlsxPopulate.fromBlankAsync().then((workbook) => {
    const sheet1 = workbook.sheet('Sheet1');

    sheet1.column('B').width(24);

    for (let i = 1; i <= 10; i++) {
      // Modify the workbook.
      sheet1.cell(`A${i}`).value('This is key!');
      sheet1.cell(`B${i}`).value('This is value!');
    }

    // Write to file.
    return workbook.toFileAsync(`./${name}.xlsx`);
  });
}

module.exports = createXlsx;
