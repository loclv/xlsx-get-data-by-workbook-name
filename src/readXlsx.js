const XLSX = require('xlsx-populate');

function readXlsx(inputName) {
  XLSX.fromFileAsync(`./${inputName}.xlsx`)
    .then((workbook) => {
      const outputObj = {};

      const sheets = workbook.sheets();

      for (let i = 0; i < sheets.length; i++) {
        // console.log(JSON.stringify(sheets[i]));
      }

      return outputObj;
    })
    .catch((error) => {
      throw error;
    });
}

module.exports = readXlsx;
