function onMessage(event) {
  const message = event.message;

  if (message.slashCommand) {
    switch (message.slashCommand.commandId) {
      case 1: // Help command
        return createSelectionCard();
    }
  }
}

function createSelectionCard(){
  return {
    "cardsV2": [
      {
        "cardId": "2",
        "card": {
          "sections": [
            {
              "widgets": [
                {
                  "dateTimePicker": {
                    "label": "OOO Start Date",
                    "name": "startdate",
                    "type": "DATE_AND_TIME",
                    "valueMsEpoch": new Date().getTime() - (6 * 60 * 60 * 1000)
                  },
                  "horizontalAlignment": "START"
                },
                {
                  "dateTimePicker": {
                    "label": "OOO End Date",
                    "name": "enddate",
                    "type": "DATE_AND_TIME",
                    "valueMsEpoch": new Date().getTime() - (6 * 60 * 60 * 1000)
                  }
                },
                {
                  "textInput": {
                    "label": "oooSubject",
                    "type": "SINGLE_LINE",
                    "name": "oooSubject",
                    "hintText": "your OOO subject line",
                    "value": "'Name' OOO"
                  }
                },
                {
                  "textInput": {
                    "label": "oooMessage",
                    "type": "MULTIPLE_LINE",
                    "name": "oooMessage",
                    "hintText": "Your OOO email message",
                    "value": "Hello I am Currently out of office"
                  }
                },
                {
                  "textInput": {
                    "label": "Calendar Event Name",
                    "type": "SINGLE_LINE",
                    "name": "eventName",
                    "hintText": "Type your name then OOO",
                    "value": ""
                  }
                },
                {
                  "selectionInput": {
                    "type": "CHECK_BOX",
                    "label": "",
                    "name": "contacts",
                    "items": [
                      {
                        "text": "Only send a response to people in my Contacts",
                        "value": "1",
                        "selected": false
                      }
                    ]
                  },
                  "horizontalAlignment": "START"
                },
                {
                  "selectionInput": {
                    "type": "CHECK_BOX",
                    "label": "",
                    "name": "domain",
                    "items": [
                      {
                        "text": "Only send a response to people in {company}",
                        "value": "1",
                        "selected": false
                      },
                    ]
                  },
                  "horizontalAlignment": "START"
                },
                {
                  "buttonList": {
                    "buttons": [
                      {
                        "text": "Submit",
                        "onClick": {
                          "action": {
                            "function": "extractDateRange"
                          }
                        },
                        "altText": "",
                        "color": {
                          "red": 0,
                          "green": 0.486,
                          "blue": 0.106,
                          "alpha": 1
                        }
                      }
                    ]
                  },
                  "horizontalAlignment": "CENTER"
                }
              ]
            }
          ],
          "header": {
            "title": "Set your OOO time",
            "subtitle": "Hello, Thank you for using {company}'s OOO chat App, please fill in the options below.",
            "imageUrl": "https://lh3.googleusercontent.com/proxy/EQlbAg8dvCwq6d74687UNJ55fLtslnqSTHIhVt9V5YfPtQ53BGNfkmKzv6gMGFDEqjoXLcx0DUqBFYk6Ji5eUAa_JwA7oqnW8jhiNCCDZ4mEtgvcgkkTPJe4ZsWVJCfr2xrabSHRAw=s88-w88-h88-c-k-no",
            "imageType": "SQUARE",
            "imageAltText": ""
          }
        }
      }
    ]
  }
}

function onCardClick(event) {
    // Check if the action is associated with the "Submit" button
    if (event.action && event.action.actionMethodName === 'extractDateRange') {

      const startDate = event.common.formInputs.startdate[""].dateTimeInput.msSinceEpoch;
      const endDate = event.common.formInputs.enddate[""].dateTimeInput.msSinceEpoch;
      const subject = event.common.formInputs.oooSubject[""].stringInputs.value[0];
      const message = event.common.formInputs.oooMessage[""].stringInputs.value[0];
      const eventTitle = event.common.formInputs.eventName[""].stringInputs.value[0];
      var domain = false;
      var contacts = false;

      if (event.common.formInputs.domain) {
        domain = event.common.formInputs.domain[""].stringInputs.value[0] === '1';
      }
      if (event.common.formInputs.contacts) {
        contacts = event.common.formInputs.contacts[""].stringInputs.value[0] === '1';
      }

      var card = turnOnAutoResponder(startDate, endDate, subject, message, domain, contacts);

      if (card && card.cardsV2[0].card.sections[0].widgets[0].decoratedText.text.includes('successfully')) {
        card = cancelMeetingsDuringDateRange(startDate, endDate);
        if (card && card.cardsV2[0].card.sections[0].widgets[0].decoratedText.text.includes('successfully')) {
          card = addEventToCalendar(startDate, endDate, eventTitle);
          if (card && card.cardsV2[0].card.sections[0].widgets[0].decoratedText.text.includes('successfully')) {
            return card;
          }
          else{
            return card;
          }
        }
        else{
          return card;
        }
      }
      else{
        return card;
      }

    }
}

function turnOnAutoResponder(startDateMsEpoch, endDateMsEpoch, oooSubject, oooMessage, domain, contacts) {
  try {

    var startdate = startDateMsEpoch + 21600000;
    var enddate = endDateMsEpoch + 86400000 + 21600000;
    var currentDate = new Date()
    currentDate.setHours(0, 0, 0, 0);

    if(enddate > (startdate + (15 * 24 * 60 * 60 * 1000))){
      return createErrorResponseCard('you cannot exceed 15 days in the future');
    }

    if (!startDateMsEpoch || !endDateMsEpoch || !oooSubject || !oooMessage) {
      return createErrorResponseCard('You must fill out all fields');
    }

    if (startDateMsEpoch < currentDate || endDateMsEpoch < currentDate) {
      return createErrorResponseCard('Start date and end date must be greater than or equal to the current date and time.');
    }
    
    if (startDateMsEpoch >= endDateMsEpoch) {
      return createErrorResponseCard('Start date must be before the end date.');
    }

    // Turn on vacation responder
    Gmail.Users.Settings.updateVacation({
      enableAutoReply: true,
      responseSubject: oooSubject,
      responseBodyHtml: oooMessage,
      restrictToDomain: domain,
      restrictToContacts: contacts,
      startTime: startdate,
      endTime: enddate
    }, "me");

    console.log('Vacation responder turned on successfully.');
    return createResponseCard(`Your out of office has been successfully scheduled. start date: ${Utilities.formatDate(new Date(startdate), "GMT", "yyyy-MM-dd")} end date: ${Utilities.formatDate(new Date(enddate), "GMT", "yyyy-MM-dd")}`);
  } catch (error) {
    console.error('Error turning on vacation responder:', error.message);
  }
}

function cancelMeetingsDuringDateRange(startDate, endDate) {
  try {
    const calendarId = 'primary';

    startDate = startDate + 21600000;
    endDate = endDate + 21600000;

    const events = CalendarApp.getCalendarById(calendarId).getEvents(new Date(startDate), new Date(endDate));

    if (events && events.length > 0) {
      for (const event of events) {
        // Check if the event has guests before declining
        if (event.getGuestList().length > 0) {
          event.setMyStatus(CalendarApp.GuestStatus.NO);
          Utilities.sleep(1000); // Add a delay between operations to ensure proper processing
          console.log(`Successfully declined event: ${event.getId()}`);
        }
      }
    } else {
      console.log('No events found during the specified date range.');
    }

    return createResponseCard(`Meetings have been successfully canceled. Start date: ${Utilities.formatDate(new Date(startDate), 'GMT', 'yyyy-MM-dd')} End date: ${Utilities.formatDate(new Date(endDate), 'GMT', 'yyyy-MM-dd')}`);
  } catch (error) {
    console.error('Error:', error.message);
    return createErrorResponseCard('Sorry there was an error canceling meetings in the provided date range');
  }
}

function addEventToCalendar(startDateMsEpoch, endDateMsEpoch, eventTitle) {
  try {
    var calendarId = '{replace with out of office calendar}';
    var calendar = CalendarApp.getCalendarById(calendarId);

    // Convert epoch timestamps to Date objects
    var startdate = new Date(startDateMsEpoch + 21600000);
    var enddate = new Date(endDateMsEpoch + 21600000);

    // Create the event
    var event = calendar.createEvent(eventTitle, startdate, enddate);

    // Optional: Add any additional event details
    event.setDescription('Description of the event');
    event.setLocation('Event Location');
    
    console.log('Event added successfully!');
    return createResponseCard(`Your out of office has been successfully scheduled. start date: ${Utilities.formatDate(new Date(startdate), "GMT", "yyyy-MM-dd")} end date: ${Utilities.formatDate(new Date(enddate), "GMT", "yyyy-MM-dd")}`);

  } catch (error) {
    console.error('Error:', error.message);
    return createErrorResponseCard('Sorry there was an error adding your OOO event to the calendar');
  }
}

function createResponseCard(responseText) {
  return {
    "cardsV2": [
      {
        "cardId": "1",
        "card": {
          "sections": [
            {
              "widgets": [
                {
                  "decoratedText": {
                    "topLabel": "",
                    "text": responseText,
                    "startIcon": {
                      "knownIcon": "NONE",
                      "altText": "Task done",
                      "iconUrl": ""
                    },
                    "wrapText": true
                  }
                }
              ]
            }
          ],
          "header": {
            "title": "OOO app",
            "subtitle": "Helping you manage your OOO",
            "imageUrl": "",
            "imageType": "SQUARE"
          }
        }
      }
    ]
  }
}

function createErrorResponseCard(responseText) {
  return {
    "cardsV2": [
      {
        "cardId": "1",
        "card": {
          "sections": [
            {
              "widgets": [
                {
                  "decoratedText": {
                    "topLabel": "",
                    "text": responseText,
                    "startIcon": {
                      "knownIcon": "NONE",
                      "altText": "Task done",
                      "iconUrl": ""
                    },
                    "wrapText": true
                  }
                }
              ]
            }
          ],
          "header": {
            "title": "OOO app",
            "subtitle": "Helping you manage your OOO",
            "imageUrl": "",
            "imageType": "SQUARE"
          }
        }
      }
    ]
  }
}
