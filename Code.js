// Get moment.js
eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.24.0/moment.min.js').getContentText());

function schedule() {
  var spreadsheetUi = SpreadsheetApp.getUi();

  var activeSpreadsheet = SpreadsheetApp.getActive();

  var ratioSpreadsheet = activeSpreadsheet.getSheetByName("Ratio");
  var activityNames = ratioSpreadsheet.getRange('A3:A').getValues().filter(String);
  var ratios = ratioSpreadsheet.getRange('B3:B').getValues().filter(String);
  var lastRow = activityNames.length+3;

  var weekSpreadsheet = activeSpreadsheet.getSheetByName("Week plan");
  var weekInfoRange = weekSpreadsheet.getRange("E3:V"+lastRow);

//  var variablesSpreadsheet = activeSpreadsheet.getSheetByName("Variables");
//  var hoursBufferDate = moment(variablesSpreadsheet.getRange('B2').getValue());
//  var hoursBuffer = hoursBufferDate.hours()+(hoursBufferDate.minutes()/60);

  var date = moment();
  var startOfWeek = date.clone().startOf('week').add(1, 'd');
  var endOfWeek = date.clone().endOf('week');
  var calendars = CalendarApp.getAllCalendars();
  var eventsThisWeek = CalendarApp.getDefaultCalendar().getEvents(startOfWeek.toDate(), endOfWeek.toDate()).filter(function(val) {return !val.isAllDayEvent();});

  var weekMinutes = [];
  for(var i = 0; i < 6; i++) {
    var today = startOfWeek.clone().add(i, 'd').toDate();
    var eventsForDay = CalendarApp.getDefaultCalendar().getEventsForDay(today);
    var minutesToday = [];

    for(var j = 0; j < 24*60; j++) {
      minutesToday[j] = true;
    }

    eventsForDay.forEach(function(val) {
      if(val.isAllDayEvent()) {
        return;
      }

      var startTime = val.getStartTime();
      var endTime = val.getEndTime();

      var minutesStart = (startTime.getHours()*60)+startTime.getMinutes();
      var minutesEnd = (endTime.getHours()*60)+endTime.getMinutes();



      var title = val.getTitle()

      if(title.indexOf("Hotschedules") !== -1) {
        // Give myself time to get ready for work
        minutesStart -= 40;
        // Give myself time to get back and chill
        minutesEnd += 40;
      }

      if(startTime.getDate() != endTime.getDate()) {
        if(startTime.getDate() != today.getDate()) {
          minutesStart = 0;
        } else if (endTime.getDate() != today.getDate()) {
          minutesEnd = (24*60);
        }
      }

      for(var j = minutesStart; j <= minutesEnd; j++) {
        minutesToday[j] = false;
      }
    });

    weekMinutes[i] = minutesToday;

    // spreadsheetUi.alert(weekMinutes[i]);
  }

  var minuteTotals = weekMinutes.map(function(minutes) {
    var minutesTotal = 0;
    for(var minute in minutes) {
      if(minutes[minute] == true) {
        minutesTotal++;
      }
    }
    return minutesTotal;
  });

  var hourTotals = minuteTotals.map(function(minuteTotal) {
    return minuteTotal/60/24;
  });

  for(var i in hourTotals) {
    weekInfoRange.getCell(1, 3+(i*3)).setValue(hourTotals[i]);
  }


  // This seperates the two parts of this big ugly function
  // return;

  // Scheduling

  var fluidCalendar = CalendarApp.getCalendarsByName('Fluid time')[0];

  if(!fluidCalendar) {
    fluidCalendar = CalendarApp.createCalendar('Fluid time');
    fluidCalendar = fluidCalendar.setTimeZone("America/New_York");
  }

  var weekInfoRows = weekInfoRange.getValues().slice(1);

  var weekActivityRatios = [];

  for(var i in weekMinutes) {
    var activitiesToday = [];
    var shiftedIndex = i*3;
    for(var j in activityNames) {
      var currentRow = weekInfoRows[j];
      activitiesToday[j] = {name: activityNames[j][0], order: currentRow[shiftedIndex], percent: currentRow[shiftedIndex+1]};
    }

    weekActivityRatios[i] = activitiesToday.sort(function(a, b) {
      if (a.order < b.order) {
        return -1;
      }
      if (a.order > b.order) {
        return 1;
      }
      return 0;
    });
  }

  for(var i = 0; i < weekMinutes.length; i++) {
    var dayMinutes = weekMinutes[i];
    var dayMinutesTotal = minuteTotals[i];
    var date = startOfWeek.clone().add(i, 'd');

    var minute = 0;
    var activityIndex = 0;

    var jsDate = date.toDate();

    var activitiesToday = weekActivityRatios[i];

    fluidCalendar.getEventsForDay(jsDate).forEach(function(event) {
      Utilities.sleep(100);
//      var startDate = event.getStartTime();
//      var endDate = event.getEndTime();
//      var title = event.getTitle();
      event.deleteEvent()
    });

    while(minute < 24*60 && activityIndex < activityNames.length) {
      var currentActivity = activitiesToday[activityIndex];
      var name = currentActivity.name;
      var percent = currentActivity.percent;

      var activityMinutesTotal = Math.floor(percent*dayMinutesTotal);
      var activityMinutesCount = 0;
      var activityRanges = [];

      while(activityMinutesCount < activityMinutesTotal) {
        var currentMinute = minute;

        if(dayMinutes[currentMinute]) {
          if(activityRanges.length == 0) {
            activityRanges[0] = {start: currentMinute, end: currentMinute};
          }

          activityRanges[activityRanges.length - 1].end = currentMinute;

          activityMinutesCount++;

          minute++;
        } else {

          if(activityRanges[activityRanges.length - 1]) {
            activityRanges[activityRanges.length - 1].end = currentMinute;
          }

          // Add minutes until true, then tie off current activity range and increment
          var greaterThanOrEqualToDayLength = false;

          while(!greaterThanOrEqualToDayLength && !dayMinutes[currentMinute]) {
            currentMinute++;

            if(currentMinute >= 24*60) {
              greaterThanOrEqualToDayLength = true;
            }
          }

          if(greaterThanOrEqualToDayLength) {
            break;
          }

          activityRanges.push({start: currentMinute, end: currentMinute});

          minute = currentMinute;
        }
      }

      for(var j in activityRanges) {

        var range = activityRanges[j];

        var rangeDuration = range.end-range.start;

        // Skip time ranges less than 10 minutes
        if(rangeDuration < 10) {
          minute -= rangeDuration;
          continue;
        }

        var startDate = date.clone().add(range.start-1, 'm').toDate();
        var endDate = date.clone().add(range.end, 'm').toDate();

        fluidCalendar.createEvent(name, startDate, endDate);

        Utilities.sleep(100);
      }

      activityIndex++;
    }
  }

  updateActualRatios();
}

function clearFluidWeek() {
  var fluidCalendar = CalendarApp.getCalendarsByName('Fluid time')[0];

  if(!fluidCalendar) {
    return;
  }

  var date = moment();
  var startOfWeek = date.clone().startOf('week').add(1, 'd');
  var endOfWeek = date.clone().endOf('week');

  var fluidThisWeek = fluidCalendar.getEvents(startOfWeek.toDate(), endOfWeek.toDate());

  fluidThisWeek.forEach(function(event) {event.deleteEvent(); Utilities.sleep(100)});
}

function updateActualRatios() {
  var spreadsheetUi = SpreadsheetApp.getUi();

  var activeSpreadsheet = SpreadsheetApp.getActive();

  var actualSpreadsheet = activeSpreadsheet.getSheetByName("Week actual");

  var activityNames = actualSpreadsheet.getRange('A4:A').getValues().filter(String);

  var lastRow = activityNames.length+3;

  var dayRatiosRange = actualSpreadsheet.getRange('D4:I'+(lastRow));
  var weekRatiosRange = actualSpreadsheet.getRange('C4:C');
  var weekTotalScheduledCell = actualSpreadsheet.getRange('C3');

  var date = moment();
  var startOfWeek = date.clone().startOf('week').add(1, 'd');
  var endOfWeek = date.clone().endOf('week');
  var calendars = CalendarApp.getAllCalendars();
  var eventsThisWeek = CalendarApp.getDefaultCalendar().getEvents(startOfWeek.toDate(), endOfWeek.toDate()).filter(function(val) {return !val.isAllDayEvent();});

  var weekActivities = [];

  var fluidCalendar = CalendarApp.getCalendarsByName('Fluid time')[0];

  if(!fluidCalendar) {
    fluidCalendar = CalendarApp.createCalendar('Fluid time');
  }

  var weekMinutesTotal = 0;

  for(var i = 0; i < 6; i++) {
    var today = startOfWeek.clone().add(i, 'd').toDate();
    var eventsForDay = fluidCalendar.getEventsForDay(today);

    var activities = activityNames.map(function(activity) { return {name: activity[0], minutes: 0, percent: 0} });

    var dayMinutesTotal = 0;

    for (var j in activities) {
      var currentActivity = activities[j];

      for(var k in eventsForDay) {
        var currentEvent = eventsForDay[k];

        if(currentEvent.getTitle() == currentActivity.name) {
          var startTime = currentEvent.getStartTime();
          var endTime = currentEvent.getEndTime();

          var minutesStart = (startTime.getHours()*60)+startTime.getMinutes();
          var minutesEnd = (endTime.getHours()*60)+endTime.getMinutes();

          if(startTime.getDate() != endTime.getDate()) {
            if(startTime.getDate() != today.getDate()) {
              minutesStart = 0;
            } else if (endTime.getDate() != today.getDate()) {
              minutesEnd = (24*60);
            }
          }

          var minutesDuration = minutesEnd-minutesStart;

          currentActivity.minutes += minutesDuration;
          dayMinutesTotal += minutesDuration;
        }
      }
    }

    // Calculate percent for each activity

    for (var j in activities) {
      var currentActivity = activities[j];

      currentActivity.percent = (currentActivity.minutes / dayMinutesTotal);
    }

    weekMinutesTotal += dayMinutesTotal;

    weekActivities.push({total: dayMinutesTotal, activities: activities});
  }

  var rotatedWeekActivities = [];

  for(var i in activityNames) {
    rotatedWeekActivities[i] = [];
    for(var j in weekActivities) {
      rotatedWeekActivities[i][j] = weekActivities[j].activities[i];
    }
  }

  var weekRatios = rotatedWeekActivities.map(function(activities) {
    return activities.map(function(activity) {
      return activity.percent;
    });
  });

  dayRatiosRange.setValues(weekRatios);

  weekTotalScheduledCell.setValue(weekMinutesTotal/60/24);
}

function addCalendarUpdateTrigger() {
  ScriptApp
      .newTrigger('updateActualRatios')
      .forUserCalendar(Session.getActiveUser().getEmail())
      .onEventUpdated()
      .create()
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Composition manager')
      .addItem('Push to calendar', 'schedule')
      .addItem('Clear fluid time this week', 'clearFluidWeek')
      .addItem('Update actual ratios', 'updateActualRatios')
      .addToUi();
}

