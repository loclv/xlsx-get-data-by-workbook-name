const XlsxPopulate = require('xlsx-populate');

function readXlsx(inputName, searchingKey, createXlsx, outputName) {
  XlsxPopulate.fromFileAsync(`./${inputName}.xlsx`)
    .then((workbook) => {
      const outputData = [];

      const sheets = workbook.sheets();

      const len = sheets.length;

      for (let i = 0; i < len; i++) {
        console.log(sheets[i].name());

        const regex = new RegExp(searchingKey, 'g');

        const cells = sheets[i].find(regex);

        if (cells && cells.length) {
          for (let j = 0; j < cells.length; j++) {
            console.log(cells[j].value());
            const sheetValues = outputData[sheets[i].name()];
            if (sheetValues && sheetValues.length) {
              sheetValues.push(cells[j].value());
            } else {
              outputData[sheets[i].name()] = [cells[j].value()];
            }
          }
        }
      }

      console.log(outputData);
      createXlsx(outputName, outputData);
    })
    .catch((error) => {
      throw error;
    });
}

module.exports = readXlsx;
