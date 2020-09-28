const XlsxPopulate = require('xlsx-populate');

function readXlsx(inputName, searchingKey) {
  XlsxPopulate.fromFileAsync(`./${inputName}.xlsx`)
    .then((workbook) => {
      const outputObj = {};

      const sheets = workbook.sheets();

      const len = sheets.length;

      for (let i = 0; i < len; i++) {
        console.log(sheets[i].name());

        const regex = new RegExp(searchingKey, 'g');

        const cells = sheets[i].find(regex);

        if (cells && cells.length) {
          for (let j = 0; j < cells.length; j++) {
            console.log(cells[j].value());
          }
        }
      }

      return outputObj;
    })
    .catch((error) => {
      throw error;
    });
}

module.exports = readXlsx;
