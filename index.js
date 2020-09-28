const dotenv = require('dotenv');

const createXlsx = require('./src/createXlsx');
const readXlsx = require('./src/readXlsx');

dotenv.config();

const inputName = process.env.INPUT_NAME;
const searchingKey = process.env.SEARCH_KEY;

const readObj = readXlsx(inputName, searchingKey);
console.log(readObj);

const outputName = process.env.OUTPUT_NAME;

const object = {a: 1, b: 2, c: 3};

createXlsx(outputName, object);
