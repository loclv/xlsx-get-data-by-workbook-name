const dotenv = require('dotenv');

const createXlsx = require('./src/createXlsx');
const readXlsx = require('./src/readXlsx');

dotenv.config();

const inputName = process.env.INPUT_NAME;

const readObj = readXlsx(inputName);
console.log(readObj);

const outputName = process.env.OUTPUT_NAME;

const object = {a: 1, b: 2, c: 3};

createXlsx(outputName, object);
