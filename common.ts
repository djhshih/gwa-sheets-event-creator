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

function doGetEvents(e) {
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

	var venues = new Map<string, string>;

	var nEvents = eventsLastRow - eventsFirstRow + 1;
	var events = new Array<CalendarEvent>(nEvents);
	for (let i = 0; i < events.length; ++i) {
		events[i] = new CalendarEvent();
	}

	var sheet = SpreadsheetApp.getActiveSheet();

	var range;
	var data;

	// construct record for venue
	// assumes that lookup table consists of two columns: venue and address
	range = makeRange(lookupVenueColumn, lookupFirstRow, lookupAddressColumn, lookupLastRow);
	data = sheet.getRange(range).getValues();
	for (let i = 0; i < data.length; ++i) {
		venues.set(data[i][0], data[i][1]);
	}

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
		events[i].location = venues.get(data[i][0]) as string;
	}

	var index = 0;
	var event = events[index];

	var indexText = CardService.newTextParagraph()
		.setText(`Event ${String(index + 1)} of ${nEvents}`);

	var titleText = CardService.newTextInput()
		.setFieldName('title')
		.setTitle('Title')
		.setValue(event.title)
		;
	
	var startTimeText = CardService.newTextInput()
		.setFieldName('startTime')
		.setTitle('Start')
		.setValue(formatDateTime(event.startTime))
		;

	var endTimeText = CardService.newTextInput()
		.setFieldName('endTime')
		.setTitle('End')
		.setValue(formatDateTime(event.endTime))
		;

	var locationText = CardService.newTextInput()
		.setFieldName('location')
		.setTitle('Location')
		.setValue(event.location)
		;

	var descriptionText = CardService.newTextInput()
		.setFieldName('description')
		.setTitle('Description')
		.setValue(event.description)
		;

	// show events
	var section = CardService.newCardSection()
		.addWidget(indexText)
		.addWidget(titleText)
		.addWidget(startTimeText)
		.addWidget(endTimeText)
		.addWidget(locationText)
		.addWidget(descriptionText)
		;
	
	// Make button
	var actionOne = CardService.newAction()
		.setFunctionName('doAddEvent')
		.setParameters({index: String(index), events: JSON.stringify(events)});
	var actionAll = CardService.newAction()
		.setFunctionName('doAddEvents')
		.setParameters({index: String(index), events: JSON.stringify(events)});
	var addOneButton = CardService.newTextButton()
		.setText('Add')
		.setOnClickAction(actionOne)
		.setTextButtonStyle(CardService.TextButtonStyle.FILLED);
	var addAllButton = CardService.newTextButton()
		.setText('Add All')
		.setOnClickAction(actionAll)
		.setTextButtonStyle(CardService.TextButtonStyle.FILLED);
	var footer = CardService.newFixedFooter()
		.setPrimaryButton(addOneButton)
		.setSecondaryButton(addAllButton);

	var card = CardService.newCardBuilder()
		.addSection(section)
		.setFixedFooter(footer)
		.build();

	return card;
}

function doAddEvent(e) {
	var index = Number(e.parameters.index);
	var events = JSON.parse(e.parameters.events);
	if (index >= 0 && index < events.length) {
		addEvent(events[index]);
	} else {
		throw new RangeError('Event index is out of range')
	}

	// TODO populate next event
}

function doAddEvents(e) {
	var index = Number(e.parameters.index);
	var events = JSON.parse(e.parameters.events);
	if (index >= 0 && index < events.length) {
		for (let i = index; i < events.length; ++i) {
			addEvent(events[i]);
		}
	} else {
		throw new RangeError('Event index is out of range')
	}

	var calendar = CalendarApp.getDefaultCalendar();

	var nEvents = events.length - index;

	var text = CardService.newTextParagraph();
	if (nEvents > 1) {
		text.setText(
			`A total of ${events.length - index} events added to: ${calendar.getId()}`
		);
	} else {
		text.setText(
			`Event added to: ${calendar.getId()}`
		);
	}

	var section = CardService.newCardSection()
		.addWidget(text);

	var card = CardService.newCardBuilder()
		.addSection(section)
		.build();
	
	return card;
}

function addEvent(event: CalendarEvent) {

	// may need to manually convert because JSON.parse does not parse to Date
	if (!(event.startTime instanceof Date)) {
		event.startTime = new Date(event.startTime);
	}
	if (!(event.endTime instanceof Date)) {
		event.endTime = new Date(event.endTime);
	}

	// create event
	var calendar = CalendarApp.getDefaultCalendar();
	var event = calendar.createEvent(
		event.title,
		event.startTime,
		event.endTime,
		{
			location: event.location,
			description: event.description
		}
	);
}
