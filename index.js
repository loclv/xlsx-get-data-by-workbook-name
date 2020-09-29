const dotenv = require('dotenv');

const createXlsx = require('./src/createXlsx');
const readXlsx = require('./src/readXlsx');

dotenv.config();

const inputName = process.env.INPUT_NAME;
const searchingKeys = process.env.SEARCH_KEYS.split(', ');
const anotherSearchingKey = process.env.ANOTHER_SEARCH_KEY;
const outputName = process.env.OUTPUT_NAME;

readXlsx(inputName, searchingKeys, anotherSearchingKey, createXlsx, outputName);
