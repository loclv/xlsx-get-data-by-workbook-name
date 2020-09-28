const dotenv = require('dotenv');

const createXlsx = require('./src/createXlsx');
const readXlsx = require('./src/readXlsx');

dotenv.config();

const inputName = process.env.INPUT_NAME;
const searchingKeys = process.env.SEARCH_KEY.split(', ');
const outputName = process.env.OUTPUT_NAME;

readXlsx(inputName, searchingKeys, createXlsx, outputName);
