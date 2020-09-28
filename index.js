const dotenv = require('dotenv');

const createXlsx = require('./src/createXlsx');
const readXlsx = require('./src/readXlsx');

dotenv.config();

const inputName = process.env.INPUT_NAME;
const searchingKey = process.env.SEARCH_KEY;

const outputName = process.env.OUTPUT_NAME;

readXlsx(inputName, searchingKey, createXlsx, outputName);
