/**
 * Callback for rendering the homepage card.
 * @return {CardService.Card}
 */
function onHomepage(e) {
	var introText = CardService.newTextParagraph()
		.setText('Open a spreadsheet to add events.');
	var section = CardService.newCardSection()
		.addWidget(introText);
	var card = CardService.newCardBuilder()
		.addSection(section)
		.build();
	return card;
}

function getRow(loc) {
	return parseInt(loc.replace(/[A-Z]/g, ''));
}

function getColumn(loc) {
	return loc.replace(/[0-9]/g, '');
}

function makeRange(c1, r1, c2, r2) {
	return `${c1}${r1}:${c2}${r2}`;
}

function parseDate(s: string) {
	return Utilities.parseDate(s.replace(/\(.*\)/, ''), "GMT", "yyyy-MM-dd");
}

function formatDate(d: Date) {
	return Utilities.formatDate(d, timeZone, "yyyy-MM-dd");
}

function formatTime(d: Date) {
	return Utilities.formatDate(d, timeZone, "HH:mm");
}

function formatDateTime(d: Date) {
	return Utilities.formatDate(d, timeZone, "yyyy-MM-dd HH:mm");
}

function combineDateTime(date, time: Date): Date {
	date = new Date(date);
	date.setHours(time.getHours());
	date.setMinutes(time.getMinutes());
	return date;
}

class CalendarEvent {
	title: string;
	startTime: Date;
	endTime: Date; 
	location: string;
	description: string;
	error: string;
}

interface TimeInterval {
	start: Date,
	end: Date
}

var timeZone = CalendarApp.getTimeZone();

function parseTime(time: string): Date {
	return Utilities.parseDate(time, timeZone, "HH:mm");
}

function parseTimeInterval(time: string): TimeInterval {

	// Utilities.parseDate cannot handle 'a.m.' and 'p.m.',
	// so remove the '.'
	time = time.replace(/\./g, '');

	// split time string into start and end times
	var startTime;
	var endTime;

	// check for dash
	var j = time.indexOf('-');
	var offset = 1;
	if (j == -1) {
		// check for en dash
		j = time.indexOf('–'); 
	}
	if (j == -1) {
		// check for 'to'
		j = time.indexOf('to');
		offset = 2;
	}

	if (j != -1) {
		startTime = time.substring(0, j).trim();
		endTime = time.substring(j + offset).trim();
	} else {
		startTime = time;
		endTime = '';
	}

	return {start: parseTime(startTime), end: parseTime(endTime)};
}

function doAddEvents(e) {
	var dateColumn = e.formInput.dateColumn;
	var timeColumn = e.formInput.timeColumn;
	var venueColumn = e.formInput.venueColumn;
	var titleColumn = e.formInput.titleColumn;
	var descriptionColumn = e.formInput.descriptionColumn;
	var lookupVenueColumn = e.formInput.lookupVenueColumn;
	var lookupAddressColumn = e.formInput.lookupAddressColumn;

	var eventsFirstRow = e.formInput.eventsFirstRow;
	var eventsLastRow = e.formInput.eventsLastRow;
	var lookupFirstRow = e.formInput.lookupFirstRow;
	var lookupLastRow = e.formInput.lookupLastRow;

	var nEvents = eventsLastRow - eventsFirstRow + 1;
	var events = new Array<CalendarEvent>(nEvents);
	for (let i = 0; i < events.length; ++i) {
		events[i] = new CalendarEvent();
	}

	var sheet = SpreadsheetApp.getActiveSheet();

	var range;
	var data;

	range = makeRange(titleColumn, eventsFirstRow, titleColumn, eventsLastRow);
	data = sheet.getRange(range).getValues();
	for (let i = 0; i < events.length; ++i) {
		events[i].title = data[i][0];
	}

	range = makeRange(descriptionColumn, eventsFirstRow, descriptionColumn, eventsLastRow);
	data = sheet.getRange(range).getValues();
	for (let i = 0; i < events.length; ++i) {
		events[i].description = data[i][0];
	}

	range = makeRange(dateColumn, eventsFirstRow, dateColumn, eventsLastRow);
	data = sheet.getRange(range).getValues();
	var dates = new Array<Date>(nEvents);
	for (let i = 0; i < events.length; ++i) {
		dates[i] = parseDate(data[i][0]);
	}

	range = makeRange(timeColumn, eventsFirstRow, timeColumn, eventsLastRow);
	data = sheet.getRange(range).getValues();
	for (let i = 0; i < events.length; ++i) {
		var interval = parseTimeInterval(data[i][0]);
		events[i].startTime = combineDateTime(dates[i], interval.start);
		events[i].endTime = combineDateTime(dates[i], interval.end);
	}

	range = makeRange(venueColumn, eventsFirstRow, venueColumn, eventsLastRow);
	data = sheet.getRange(range).getValues();
	for (let i = 0; i < events.length; ++i) {
		// TODO lookup
		events[i].location = data[i][0];
	}
	
	// TODO show events
	var section = CardService.newCardSection()
		.addWidget(
			CardService.newTextParagraph().setText(
				//eventsFirstRow + ' ' + eventsLastRow + ' ' + dateColumn + ' ' + dateRange
				events[0].title + ' ' + events[0].description + ' ' + 
					formatDateTime(events[0].startTime) + ' ' + formatDateTime(events[0].endTime) + ' ' +
					'\n...\n' +
				events[nEvents-1].title + ' ' + events[nEvents-1].description + ' ' + 
					formatDateTime(events[nEvents-1].startTime) + ' ' + formatDateTime(events[nEvents-1].endTime)
			)
		);

	var card = CardService.newCardBuilder()
		.addSection(section)
		.build();

	// TODO add events to calendar
	
	/*
	var date = Utilities.parseDate(e.formInput.date, "GMT", DATE_FORMAT);

	var timeZone = e.formInput.timeZone;
	var startTime = combineDateTime(
		date, Utilities.parseDate(e.formInput.startTime, timeZone, "h:mm a")
	);
	var endTime = combineDateTime(
		date, Utilities.parseDate(e.formInput.endTime, timeZone, "h:mm a")
	);
	var description = e.parameters.description;

	// create event
	var calendar = CalendarApp.getDefaultCalendar();
	var event = calendar.createEvent(
		title,
		startTime,
		endTime,
		{
			location: location,
			description: description
		}
	);
 */


	return card;
}

