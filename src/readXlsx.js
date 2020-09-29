const XlsxPopulate = require('xlsx-populate');

function getMatchedCells(searchingKey, sheet, originalData, propertyName) {
  const regex = new RegExp(searchingKey, 'g');
  const outputData = originalData;

  const cells = sheet.find(regex);

  if (cells && cells.length) {
    const cellsLen = cells.length;
    for (let j = 0; j < cellsLen; j++) {
      if (!outputData[sheet.name()]) {
        outputData[sheet.name()] = {
          [propertyName]: [cells[j].value()],
        };
      } else if (!outputData[sheet.name()][propertyName]) {
        outputData[sheet.name()][propertyName] = [cells[j].value()];
      } else {
        outputData[sheet.name()][propertyName].push(cells[j].value());
      }
    }
  }

  return outputData;
}

function readXlsx(
  inputName,
  searchingKeys,
  anotherKey,
  createXlsx,
  outputName,
) {
  XlsxPopulate.fromFileAsync(`./${inputName}.xlsx`)
    .then((workbook) => {
      const outputData = [];
      const columns = ['B', 'C'];

      const sheets = workbook.sheets();

      const len = sheets.length;

      for (let i = 0; i < len; i++) {
        const sheet = sheets[i];

        for (let k = 0; k < searchingKeys.length; k++) {
          getMatchedCells(searchingKeys[k], sheet, outputData, columns[0]);
        }

        // another searching keyword
        getMatchedCells(anotherKey, sheet, outputData, columns[1]);
      }

      createXlsx(outputName, outputData, columns);
    })
    .catch((error) => {
      throw error;
    });
}

module.exports = readXlsx;
