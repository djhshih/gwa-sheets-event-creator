function onSheetsHomepage(e) {
	var promptLabel = CardService.newTextParagraph()
		.setText('Select the header for each field:');

	var eventsTableLabel = CardService.newTextParagraph()
		.setText('Events table');
		
	var lookupTableLabel = CardService.newTextParagraph()
		.setText('Address table');

	var eventsFirstRowText = CardService.newTextInput()
		.setFieldName('eventsFirstRow')
		.setTitle('First row')
		.setValue('4')

	var eventsLastRowText = CardService.newTextInput()
		.setFieldName('eventsLastRow')
		.setTitle('Last row')
		.setValue('19')

	var dateColumnText = CardService.newTextInput()
		.setFieldName('dateColumn')
		.setTitle('Date')
		.setValue('A')
		;

	var timeColumnText = CardService.newTextInput()
		.setFieldName('timeColumn')
		.setTitle('Time')
		.setValue('B')
		;

	var venueColumnText = CardService.newTextInput()
		.setFieldName('venueColumn')
		.setTitle('Location')
		.setValue('D')	
		;

	var titleColumnText = CardService.newTextInput()
		.setFieldName('titleColumn')
		.setTitle('Title')
		.setValue('E')	
		;

	var descriptionColumnText = CardService.newTextInput()
		.setFieldName('descriptionColumn')
		.setTitle('Description')
		.setValue('H')	
		;

	var lookupFirstRowText = CardService.newTextInput()
		.setFieldName('lookupFirstRow')
		.setTitle('First row')
		.setValue('23')	
		;

	var lookupLastRowText = CardService.newTextInput()
		.setFieldName('lookupLastRow')
		.setTitle('Last row')
		.setValue('33')	
		;

	var lookupVenueColumnText = CardService.newTextInput()
		.setFieldName('lookupVenueColumn')
		.setTitle('Venue')
		.setValue('A')	
		;

	var lookupAddressColumnText = CardService.newTextInput()
		.setFieldName('lookupAddressColumn')
		.setTitle('Address')
		.setValue('B')	
		;

	
	var section = CardService.newCardSection()
		.addWidget(promptLabel)
		.addWidget(eventsTableLabel)
		.addWidget(eventsFirstRowText)
		.addWidget(eventsLastRowText)
		.addWidget(dateColumnText)
		.addWidget(timeColumnText)
		.addWidget(venueColumnText)
		.addWidget(titleColumnText)
		.addWidget(descriptionColumnText)
		.addWidget(lookupTableLabel)
		.addWidget(lookupFirstRowText)
		.addWidget(lookupLastRowText)
		.addWidget(lookupVenueColumnText)
		.addWidget(lookupAddressColumnText)
		;

	// Make button
	var action = CardService.newAction()
		.setFunctionName('doAddEvents');
	var addButton = CardService.newTextButton()
		.setText('Add')
		.setOnClickAction(action)
		.setTextButtonStyle(CardService.TextButtonStyle.FILLED);
	var footer = CardService.newFixedFooter()
		.setPrimaryButton(addButton);

	var peekHeader = CardService.newCardHeader()
		.setTitle('Event');

	var builder = CardService.newCardBuilder()
		.addSection(section)
		.setFixedFooter(footer)
		.setPeekCardHeader(peekHeader);
	
	return builder.build();
}

