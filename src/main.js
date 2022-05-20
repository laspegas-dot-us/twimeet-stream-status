const fs = require("fs");
const { EOL } = require("os");
const axios = require("axios");
const moment = require("moment");
const ExcelJS = require("exceljs");

const panelsScheduleFile = "../../twimeet-stream-status-test/schedule.xlsx";
var scheduleWorkbook = null;
var scheduleWorksheet = null;
var schedule = {};

class StatusFilePath {
	static panelNow   = "../../twimeet-stream-status-test/status.txt";
	static panelNext  = "../../twimeet-stream-status-test/status_next.txt";
	static panelBreak = "../../twimeet-stream-status-test/status_break.txt";
	static lpfmSong   = "../../twimeet-stream-status-test/status_song.txt";
}

// #=#=#=#=#=#=#=# #=#=#=#=#=#=#=# #=#=#=#=#=#=#=# #=#=#=#=#=#=#=# #=#=#=#=#=#=#=#

async function readSchedule() {
	
	if (scheduleWorkbook === null || scheduleWorksheet === null)
		return;

	let panels = {
		now:  {name: "", org: "", time: moment()},
		next: {name: "", org: "", time: moment()}
	}

	let closestBefore = Number.NEGATIVE_INFINITY;
	let closestAfter = Infinity;

	scheduleWorksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
		if (rowNumber < 2) return;

		let startTime = moment.utc(row.getCell("A").value);
		let panelName = row.getCell("B").value;
		let panelOrg  = row.getCell("C").value;

		let delta = startTime - moment();
		if (delta > 0 && delta < closestAfter) {
			closestAfter = delta;
			panels["next"] = {name: panelName, org: panelOrg, time: startTime};
			return;
		}

		if (delta < 0 && delta > closestBefore) {
			closestBefore = delta;
			panels["now"] = {name: panelName, org: panelOrg, time: startTime};
		}

	});

	return panels;

}

async function refreshPanels() {
	let upcoming = await readSchedule();

	if (upcoming.now.name === "") {
		putToFile(`-`, StatusFilePath.panelNow);
	} else {
		putToFile(`${upcoming.now.org}: ${upcoming.now.name}`, StatusFilePath.panelNow);
	}
	
	putToFile(`Następnie o ${upcoming.next.time.format("HH:mm")} => ${upcoming.next.name}`, StatusFilePath.panelNext);
	putToFile(`Zapraszamy już o ${upcoming.next.time.format("HH:mm")} na prelekcję pt.${EOL}${EOL}${upcoming.next.org}: ${upcoming.next.name}`, StatusFilePath.panelBreak);

	console.log("Pomyślnie odświeżono harmonogram atrakcji.");
}


async function refreshSchedule() {

	var workbook = new ExcelJS.Workbook();

	try {
		await workbook.xlsx.readFile(panelsScheduleFile);
	} catch (err) {
		console.error("Błąd podczas wczytywania harmonogramu!");
		console.error(err);
		return;
	}

	scheduleWorkbook = workbook;
	scheduleWorksheet = scheduleWorkbook.getWorksheet(1);

}


async function getSongStatus() {

	const songStatusApiUrl = "https://laspegas.us/api/now";

	let res = await axios.get(songStatusApiUrl);
	if (res.status !== 200) {
		console.error("Brak połączenia z Las Pegasus API!:" + EOL + res.data);
		return null;
	}

	console.log("Odświeżono tytuł utworu i wykonawcę: " + res.data.title);
	return `${res.data.artist} - ${res.data.title}`;

}


function putToFile(content, filePath) {

	try {
		fs.writeFileSync(filePath, content);
	} catch (err) {
		console.error(`Błąd zapisu pliku ${filePath}!`)
		console.error(err);
	}

}

async function main() {

	setInterval(() => getSongStatus().then(song => putToFile(song, StatusFilePath.lpfmSong)), 5 * 1000);
	setInterval(() => refreshSchedule(), 10 * 1000);
	setInterval(() => refreshPanels(), 4 * 1000);

	refreshSchedule();

}


main();
