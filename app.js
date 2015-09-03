'use strict';

import XLSX from 'xlsx';
import fs from 'fs';
import _ from 'lodash';

/* https://www.npmjs.com/package/xlsx */

function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

function sheet_from_array_of_arrays(data, opts) {
	var ws = {};
	var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = {v: data[R][C] };
			if(cell.v == null) continue;
			var cell_ref = XLSX.utils.encode_cell({c:C,r:R});
			
			cell.t = 's';
			
			ws[cell_ref] = cell;
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
	return ws;
}

function getDate(lines) {
	let diagnosticDataIndex = _.indexOf(lines, 'Diagnostic Data:');
	if (diagnosticDataIndex < 0) return null;

	return lines[diagnosticDataIndex+1].substring(0, lines[diagnosticDataIndex+1].length - 2);
}

function getUserHandle(lines) {
	let userHandle = _.find(lines, function(line) {
		return line.indexOf('UserHandle') === 0;
	})

	if (userHandle) return userHandle.substring(userHandle.indexOf(': ') + 2);
	return null;
}

function getPlaybackSpeed(lines) {
	let playbackSpeed = _.find(lines, function(line) {
		return line.indexOf('PlaybackSpeed') === 0;
	})

	if (playbackSpeed) return playbackSpeed.substring(playbackSpeed.indexOf(': ') + 2);
	return null;
}

function getTranscripts(lines) {
	let transcripts = _.find(lines, function(line) {
		return line.indexOf('Transcripts') === 0;
	})

	if (transcripts) return transcripts.substring(transcripts.indexOf(': ') + 2);
	return null;
}

function getVideoQuality(lines) {
	let videoQuality = _.find(lines, function(line) {
		return line.indexOf('VideoQuality') === 0;
	})

	if (videoQuality) return videoQuality.substring(videoQuality.indexOf(': ') + 2);
	return null;
}

function getUserAgent(lines) {
	let userAgent = _.find(lines, function(line) {
		return line.indexOf('UserAgent') === 0;
	})

	if (userAgent) return userAgent.substring(userAgent.indexOf(': ') + 2);
	return null;
}

function getRequestData(lines) {
	let requestData = _.find(lines, function(line) {
		return line.indexOf('Request Data') === 1;
	})

	if (requestData) {
		requestData = requestData.trim();
		requestData = requestData.substring(requestData.indexOf(': ') + 2);
		requestData = requestData.split(', ');

		let module = _.find(requestData, function(info) {
			return info.indexOf('m:"') === 0;
		})
		module = module.substring(3, module.length - 1);

		let course = _.find(requestData, function(info) {
			return info.indexOf('course:"') === 0;
		})
		course = course.substring(8, course.length - 1);

		let q = _.find(requestData, function(info) {
			return info.indexOf('q:"') === 0;
		})
		q = q.substring(3, q.length - 1);

		let responseData = {
			"course": course,
			"module": module,
			"q": q
		};

		return responseData;
	}

	return null;
}

function getClipSelected(lines) {
	let clipsSelected = [];

	let clipSelected = _.findIndex(lines, function(line) {
		return line.indexOf('Clip selected - ') === 0;
	});

	while(clipSelected != -1) {
		let clip = lines[clipSelected];
		let clipInfo = clip.substring(clip.indexOf(': ') + 2);
		let clipName = clipInfo.substring(0, clipInfo.indexOf(' '));
		let clipTitle = clipInfo.substring(clipInfo.indexOf(': ') + 2);

		let clipObject = {
			"name": clipName,
			"title": clipTitle
		}

		clipsSelected.push(clipObject.name);
		lines = lines.slice(clipSelected + 1);

		clipSelected = _.findIndex(lines, function(line) {
			return line.indexOf('Clip selected - ') === 0;
		});
	}

	if (clipsSelected) return clipsSelected.join(',\n');
	return null;
}

let preParseDir = './assets/pre-parse/';
let postParseDir = './assets/post-parse/'

let files = fs.readdirSync(preParseDir);

files.forEach(function(fileName) {
	if (fileName === '.gitignore' || fileName.indexOf('.xls') === -1) return;

	let preParseFilePath = preParseDir + fileName;
	let postParseFilePath = postParseDir + fileName;

	let workbook = XLSX.readFile(preParseFilePath);

	let sheetNames = workbook.SheetNames;
	let wb = new Workbook();
	wb.SheetNames.push('support.pluralsight.com');

	let newRows = [];
	newRows.push(['date', 'userHandle', 'course', 'module', 'clipSelected', 'q', 'playbackSpeed', 'transcripts', 'videoQuality', 'userAgent'])

	sheetNames.forEach(function(sheetName) {
		let worksheet = workbook.Sheets[sheetName];
		let rows = XLSX.utils.sheet_to_json(worksheet);

		rows.forEach(function(row) {

			let newRow = [];

			let text = row.Text;
			let lines = text.split('\n');
			//console.log(lines);

			let date = getDate(lines);
			if (!date) return;
			//console.log(date);

			let userHandle = getUserHandle(lines);
			if (!userHandle) return;
			//console.log(userHandle);

			let requestData = getRequestData(lines);
			if (!requestData) return;
			let course = requestData.course;
			let module = requestData.module;
			let q = requestData.q;
			//console.log(requestData);

			let clipSelected = getClipSelected(lines);
			if (!clipSelected) return;
			//console.log(clipSelected);


			let playbackSpeed = getPlaybackSpeed(lines);
			if (!playbackSpeed) return;
			//console.log(playbackSpeed);

			let transcripts = getTranscripts(lines);
			if (!playbackSpeed) return;
			//console.log(transcripts);

			let videoQuality = getVideoQuality(lines);
			if (!videoQuality) return;
			//console.log(videoQuality);

			let userAgent = getUserAgent(lines);
			if (!userAgent) return;
			//console.log(userAgent);

			newRow.push(date);
			newRow.push(userHandle);
			newRow.push(course);
			newRow.push(module);
			newRow.push(clipSelected);
			newRow.push(q);
			newRow.push(playbackSpeed);
			newRow.push(transcripts);
			newRow.push(videoQuality);
			newRow.push(userAgent);

			newRows.push(newRow);

		});
	});

	let ws = sheet_from_array_of_arrays(newRows)
	wb.Sheets['support.pluralsight.com'] = ws;
	//console.log(ws);

	XLSX.writeFile(wb, postParseFilePath);
});