/* NOTES FOR END USERS: 
  This script has been updated to colorize default calendar events meeting the following criteria:

  Events Colorization Criteria
    1. Event Title is NOT in the "igoredEvents" list.
    2. Event Color is the default color (API returns this as blank / ''). Events that have already been customized will be ignored.
    3. Event is NOT an all day event.

  Colors by Event Parameters:
    - Personal Events = Green
    - Company Events (attendees >50 guests) = Blue
    - Internal Only Events (< 50 attendees) = Purple
    - External Events = Yellow

    See here for color options: https://developers.google.com/apps-script/reference/calendar/event-color. To change colors, modify the 'setColor()' values in the code

    TIP: Ctrl/Cmd + F and search for the event you're looking to adjust:

      - PERSONAL EVENTS
      - LARGE INTERNAL EVENTS (> 50 ATTENDEES)
      - SMALL INTERNAL EVENTS (< 50 ATTENDEES)
      - EXTERNAL EVENTS

  Event Lookup Timeframe:
    - Currently set to 30 days (i.e. events within the net 30 days will be colorized)
    - To adjust, change the number in this line 'nextweek.setDate(nextweek.getDate() + 30);' to the desired number of days


  Disclaimer(s): 
    - Will not warranty your calendar colors
    - No refunds or returns
    - I am not a financial advisor ($AMD TO THA MOON)*/

// function generateRandomInteger(min, max) { //Random number generator function to get a random color code when the number of events evaluated exceeds 11
//   return Math.floor(min + Math.random()*(max - min + 1))
// }

function doGet() {
  return ColorEvents()
}

function ColorEvents() {

  const today = new Date();
  const nextweek = new Date();
  nextweek.setDate(nextweek.getDate() + 30);
  Logger.log(today + " " + nextweek);

  const events = CalendarApp.getDefaultCalendar().getEvents(today, nextweek);
  // VVVVVV Insert your comma separated list of event titles to ignore here. Must be an EXACT match (check for trailing spaces or special characters if issues).
  const ignoredEvents = ["Lunch", "Bridge Club with Taylor ", "LE TSE Open Office Hour"] 

 let updatedCalendars = []
 
 for (let i = 0; i < events.length; i++) { 
  try {
    let event = events[i];
    let title = events[i].getTitle();
    let color = events[i].getColor();
    let eventDate = events[i].getStartTime();

    //If the event title is not in the ignored list, and no color has been set already, proceed
    if (!ignoredEvents.includes(title) && color == '') {
      // Ensuring event is not a full day event
      switch (event.isAllDayEvent()) {
        case false:
        
          // Finding events with no guest i.e. blocks.
          const guestlist = event.getGuestList();
          switch (guestlist.length) {
            
            case 0: 
              //PERSONAL EVENTS
              Logger.log(`${eventDate} Personal Event '${title}' - setting color to green`)
              event.setColor(CalendarApp.EventColor.PALE_GREEN)
              updatedCalendars.push(`Updated Event '${title}' to color green.`)
              break;

              default:
                // Determining internal vs. external event
                guestEmails = [] //Build list of guests
                guestlist.map(guest => {
                  guestEmails.push(guest.getEmail())
                });

                switch (guestEmails.every(email => email.includes("smartsheet.com") || email.includes("calendar.google.com"))) { //Need to verify that this won't scope in external calendar groups
                  case true: 
                    switch (guestEmails.length >= 50) { //How to solve for 'guest list too large' scenario?
                      //LARGE INTERNAL EVENTS (> 50 ATTENDEES)
                      case true: 
                        Logger.log(`${eventDate} Company Event '${title}' - setting color to blue`)
                        event.setColor(1);
                        updatedCalendars.push(`Updated Event '${title}' to color blue.`)
                        break;

                      case false:
                        //SMALL INTERNAL EVENTS (< 50 ATTENDEES)
                        Logger.log(`${eventDate} Internal Only Event '${title}' - setting color to purple`)
                        event.setColor(3);
                        updatedCalendars.push(`Updated Event '${title}' to color purple.`)
                        break;
                    }
                    break;
                    
                  case false:
                    //EXTERNAL EVENTS 
                    Logger.log(`${eventDate} External Event '${title}' - setting color to yellow`)
                    updatedCalendars.push(`Updated Event '${title}' to color yellow.`)
                    event.setColor(5);
                    break;
                  }
                }
              }
              // Logger.log({updatedCalendars})
              // return HtmlService.createHtmlOutputFromFile('Index');
            }
          }
          catch(err) {
            Logger.log(err)
            return HtmlService.createHtmlOutputFromFile(`<body>${err}</body>`)
          }
        }
      }
