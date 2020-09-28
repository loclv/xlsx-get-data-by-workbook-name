const XLSX = require('xlsx-populate');

function readXlsx(inputName) {
  XLSX.fromFileAsync(`./${inputName}.xlsx`).then((workbook) => {
    const outputObj = {};

    return outputObj;
  });
}

module.exports = readXlsx;
