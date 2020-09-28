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

        sheets[i].find(regex, (match) => {
          console.log(match);
        });
      }

      return outputObj;
    })
    .catch((error) => {
      throw error;
    });
}

module.exports = readXlsx;
