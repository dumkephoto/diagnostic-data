import XLSX from 'xlsx';
import fs from 'fs';

var preParseDir = './assets/pre-parse/';
var postParseDir = './assets/post-parse/'

var files = fs.readdirSync(preParseDir);
console.log(files);


files.forEach(function(file) {
	var preParseFilePath = preParseDir + file;
	var postParseFilePath = postParseDir + file;
	var workbook = XLSX.readFile(preParseFilePath);

	var sheetNames = workbook.SheetNames;
	console.log(sheetNames);
	var address_of_cell = 'A1';

	sheetNames.forEach(function(sheetName) {

		var worksheet = workbook.Sheets[sheetName];
		console.log('testing:',XLSX.utils.sheet_to_json(worksheet));
	})

});