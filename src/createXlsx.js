const XlsxPopulate = require('xlsx-populate');

function createXlsx(name) {
  console.log(`Create ./${name}.xlsx`);

  // Load a new blank workbook
  XlsxPopulate.fromBlankAsync().then((workbook) => {
    // Modify the workbook.
    workbook.sheet('Sheet1').cell('A1').value('This is neat!');

    // Write to file.
    return workbook.toFileAsync(`./${name}.xlsx`);
  });
}

module.exports = createXlsx;
