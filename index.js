const XLSX = require('xlsx-populate');
const dotenv = require('dotenv');
const createXlsx = require('./src/createXlsx');

dotenv.config();

const inputName = process.env.INPUT_NAME;
const outputName = process.env.OUTPUT_NAME;

createXlsx(outputName);
