function main() {
  var calendars = CalendarApp.getAllCalendars();
  for (i = 0; i < calendars.length; i++) {
    console.log(calendars[i].getId(), calendars[i].getName());
  }
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
 
  //スプレッドシートのメニューにカスタムメニュー「カレンダー連携 > 実行」を作成
  var subMenus = [];
  subMenus.push({
    name: "Pull",
    functionName: "pullSchedule"  //実行で呼び出す関数を指定
  });
  subMenus.push({
    name: 'Push',
    functionName: "pushSchedule"
  })
  subMenus.push({
    name: "Clear sheet",
    functionName: "clearSheet"
  })
  subMenus.push({
    name: "Set Calendar ID",
    functionName: "setCalendarID"
  })
  subMenus.push({
    name: "Set Pull Range",
    functionName: "setPullRange"
  })
  ss.addMenu("Synchronize calendar", subMenus);
}
 
/**
 * 予定を作成する
 */
function pullSchedule() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Google Calendar Console");
  var backend = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backend");
  var calendarID = sheet.getRange(1, 2).getValue();
  var calendar = CalendarApp.getCalendarById(calendarID);

  backend.getRange(3,1,1000,10).setValue("");
  
  // var now = new Date();
  var pullStartTime = new Date(sheet.getRange(1, 4).getValue());
  var pullEndTime = new Date(sheet.getRange(1, 6).getValue());
  var events = calendar.getEvents(pullStartTime, pullEndTime);
  // Browser.msgBox("?", JSON.stringify(events), Browser.Buttons.OK);
  for (var i in events) {
    var eventID = events[i].getId();
    var color = events[i].getColor();
    var description = events[i].getDescription();
    var endTime = events[i].getEndTime();
    var location = events[i].getLocation();
    var startTime = events[i].getStartTime();
    var title = events[i].getTitle();

    sheet.getRange(parseInt(i)+3, 1).setValue(eventID)
    sheet.getRange(parseInt(i)+3, 2).setValue(startTime)
    sheet.getRange(parseInt(i)+3, 3).setValue(endTime)
    sheet.getRange(parseInt(i)+3, 4).setValue(title)
    sheet.getRange(parseInt(i)+3, 5).setValue(location)
    sheet.getRange(parseInt(i)+3, 6).setValue(description)
    sheet.getRange(parseInt(i)+3, 7).setValue(color)

    backend.getRange(parseInt(i)+3, 1).setValue(eventID)
    backend.getRange(parseInt(i)+3, 2).setValue(startTime)
    backend.getRange(parseInt(i)+3, 3).setValue(endTime)
    backend.getRange(parseInt(i)+3, 4).setValue(title)
    backend.getRange(parseInt(i)+3, 5).setValue(location)
    backend.getRange(parseInt(i)+3, 6).setValue(description)
    backend.getRange(parseInt(i)+3, 7).setValue(color)
  }
}

function isAllDayEvent(startTime, endTime) {
  var startIsZero = startTime.getHours() == 0 && startTime.getMinutes() == 0 && startTime.getSeconds() == 0;
  var endIsZero = endTime.getHours() == 0 && endTime.getMinutes() == 0 && endTime.getSeconds() == 0;
  // var isOneDay = endTime.getTime() - startTime.getTime() == 1000 * 60 * 60 * 24;
  return startIsZero && endIsZero;
}

function isN(value) {
  return value == "" || value == null;
}

function isNotUpdated(i, contents, backendContents) {
  const contentRow = contents[i]
  for (const row of backendContents) {
    // Browser.msgBox(row[0] + " == " + contentRow[0] + " " + String(row[0] == contentRow[0]))
    if (row[0] == contentRow[0]) {
      // Browser.msgBox("row[1]" + ((new Date(row[1])) - (new Date(contentRow[1])) == 0))
      // Browser.msgBox("row[2]" + ((new Date(row[2])) - (new Date(contentRow[2])) == 0))
      // Browser.msgBox("row[3]" + String(row[3] == contentRow[3]))
      // Browser.msgBox("row[4]" + String(row[4] == contentRow[4]))
      // Browser.msgBox("row[5]" + String(row[5] == contentRow[5]))
      // Browser.msgBox("row[6]" + String(row[6] == contentRow[6]))
      return ((new Date(row[1])) - (new Date(contentRow[1])) == 0) && ((new Date(row[2])) - (new Date(contentRow[2])) == 0) && (row[3] == contentRow[3]) && (row[4] == contentRow[4]) && (row[5] == contentRow[5]) && (row[6] == contentRow[6]);
    }
  }
  // Browser.msgBox("No ID matched")
  return false;
}

function pushSchedule() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Google Calendar Console");
  var backend = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backend");
  var calendarID = sheet.getRange(1, 2).getValue();
  var calendar = CalendarApp.getCalendarById(calendarID);

  var contents = sheet.getRange(3, 1, sheet.getLastRow(), 7).getValues();
  var backendContents = backend.getRange(3, 1, backend.getLastRow(), 7).getValues();
  // Browser.msgBox(backend.getRange(3, 1).getValue())

  let countDelete = 0;
  let countUpdated = 0;
  let countRemained = 0;
  let countCreated = 0;

  for (var i = 0; i < sheet.getLastRow()-2; i++) {
    var eventID = contents[i][0];
    var startTime = new Date(contents[i][1]);
    var endTime = new Date(contents[i][2]);
    var title = contents[i][3]
    var location = contents[i][4]
    var description = contents[i][5]
    var color = contents[i][6]

    var newEvent;

    if (isN(eventID) && isN(title)) {
      continue;
    }

    // Browser.msgBox("isNotUpdated = " + isNotUpdated(i, contents, backendContents))
    if (isNotUpdated(i, contents, backendContents)) {
      sheet.getRange(parseInt(i)+3, 1, 1, 7).setValue("");
      countRemained = countRemained + 1;
      // Browser.msgBox("Remained")
      continue;
    }

    // Browser.msgBox("UPDATED");
    if (isN(title)) {
      calendar.getEventById(eventID).deleteEvent();
      countDelete = countDelete + 1;
      // Browser.msgBox("Delete")
    } else {
      try {
        if (!(eventID == null || eventID == "") && calendar.getEventById(eventID) != null) {
          var event = calendar.getEventById(eventID);
          event.setColor(color)
          event.setDescription(description)
          event.setTime(startTime, endTime)
          event.setLocation(location)
          event.setTitle(title)
          countUpdated = countUpdated + 1;
          // Browser.msgBox("Updated")
        } else {
          if (isAllDayEvent(startTime, endTime)) {
            newEvent = calendar.createAllDayEvent(title, startTime, endTime);
          } else {
            newEvent = calendar.createEvent(title, startTime, endTime);
          }
          newEvent.setColor(color);
          newEvent.setDescription(description);
          newEvent.setLocation(location);
          countCreated = countCreated + 1;
          // Browser.msgBox("Created")
        }
      } catch(e) {
        Logger.log(e);
      }
    }
    sheet.getRange(parseInt(i)+3, 1, 1, 7).setValue("");
  }

  backend.getRange(3,1,1000,10).setValue("");
  Browser.msgBox("Events pushed", `Created: ${countCreated}, Updated: ${countUpdated}, Deleted: ${countDelete}, Remained: ${countRemained}`, Browser.Buttons.OK);
}

function setCalendarID() {
  var result = Browser.inputBox("Set Calendar ID", "What's your calendar ID?", Browser.Buttons.OK_CANCEL);
  if (result != "cancel"){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Google Calendar Console");
    sheet.getRange(1, 2).setValue(result)
    Browser.msgBox("Successfully set calendar ID: " + result);
    calendarID = result;
  } else {
    Browser.msgBox("Canceled");
  }
}

function setPullRange() {
  var startTime = Browser.inputBox("Set start time", "Please type in YYYY/MM/DD", Browser.Buttons.OK_CANCEL);
  if (startTime == "cancel") {
    Browser.msgBox("Canceled");
    return;
  }
  var endTime = Browser.inputBox("Set end time", "Please type in YYYY/MM/DD", Browser.Buttons.OK_CANCEL);
  if (endTime == "cancel") {
    Browser.msgBox("Canceled");
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Google Calendar Console");
  sheet.getRange(1, 4).setValue(startTime);
  sheet.getRange(1, 6).setValue(endTime);
  Browser.msgBox("Successfully the pull range: " + startTime + " ~ " + endTime);
}

function clearSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Google Calendar Console");
  sheet.getRange(3,1,1000, 10).setValue("");
}