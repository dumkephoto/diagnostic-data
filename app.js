'use strict';

import XLSX from 'xlsx';
import fs from 'fs';

/* https://www.npmjs.com/package/xlsx */

let preParseDir = './assets/pre-parse/';
let postParseDir = './assets/post-parse/'

let files = fs.readdirSync(preParseDir);

files.forEach(function(fileName) {
	if (fileName === '.gitignore' || fileName.indexOf('.xls') === -1) return;

	let preParseFilePath = preParseDir + fileName;
	let postParseFilePath = postParseDir + fileName;

	let workbook = XLSX.readFile(preParseFilePath);

	let sheetNames = workbook.SheetNames;
	sheetNames.forEach(function(sheetName) {
		let worksheet = workbook.Sheets[sheetName];
		console.log(XLSX.utils.sheet_to_json(worksheet));
	})
});