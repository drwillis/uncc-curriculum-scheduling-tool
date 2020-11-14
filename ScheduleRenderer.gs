function ScheduleRenderer(semester) {
  this.currentSchedule = [];
  this.instructionalModeToString = {
    'Face to Face': 'F2F', 'Hybrid Synchronous': 'HYS', 'Hybrid Asynchronous': 'HYA', 'Online Synchronous': 'IS', 'Online Asynchronous': 'IA'
  };
  var timeListMWF = [];
  var roomListMWF = [];
  var timeListTR = [];
  var roomListTR = [];
  var timeListSU = [];
  var roomListSU = [];
  if (semester == "Summer") {
    this.dayTimeSlotLists = { 
      M: { times: timeListMWF, rooms: roomListMWF, startRow: 2, hide: false, time_room_column: 1 , output_column: 2 , cost_column: 7, title: 'Monday'},
      T: { times: timeListTR , rooms: roomListTR , startRow: 2, hide:  true, time_room_column: 1 , output_column: 3 , cost_column: 7, title: 'Tuesday' },
      W: { times: timeListMWF, rooms: roomListMWF, startRow: 2, hide:  true, time_room_column: 1 , output_column: 4 , cost_column: 7, title: 'Wednesday' },
      R: { times: timeListTR , rooms: roomListTR , startRow: 2, hide:  true, time_room_column: 1 , output_column: 5 , cost_column: 7, title: 'Thursday' },
      F: { times: timeListMWF, rooms: roomListMWF, startRow: 2, hide:  true, time_room_column: 1 , output_column: 6 , cost_column: 7, title: 'Friday' },
      S: { times: timeListSU , rooms: roomListSU , startRow: 2, hide:  true, time_room_column: 1 , output_column: 8 , cost_column: 7, title: 'Saturday' },
      U: { times: timeListSU , rooms: roomListSU , startRow: 2, hide:  true, time_room_column: 1 , output_column: 9 , cost_column: 7, title: 'Sunday' }};    
  } else {
    this.dayTimeSlotLists = { 
      M: { times: timeListMWF, rooms: roomListMWF, startRow: 2 , hide: false, time_room_column: 1 , output_column: 2 , cost_column:  5, title: 'Monday' },
      T: { times: timeListTR , rooms: roomListTR , startRow: 28, hide: false, time_room_column: 6 , output_column: 7 , cost_column:  9, title: 'Tuesday' },
      W: { times: timeListMWF, rooms: roomListMWF, startRow: 2 , hide: true , time_room_column: 1 , output_column: 3 , cost_column:  5, title: 'Wednesday' },
      R: { times: timeListTR , rooms: roomListTR , startRow: 28, hide: true , time_room_column: 6 , output_column: 8 , cost_column:  9, title: 'Thursday' },
      F: { times: timeListMWF, rooms: roomListMWF, startRow: 2 , hide: true , time_room_column: 1 , output_column: 4 , cost_column:  5, title: 'Friday' },
      S: { times: timeListSU , rooms: roomListSU , startRow: 2 , hide: false, time_room_column: 10, output_column: 11, cost_column: 13, title: 'Saturday'},
      U: { times: timeListSU , rooms: roomListSU , startRow: 2 , hide: true , time_room_column: 10, output_column: 12, cost_column: 13, title: 'Sunday'}};
  }
  this.destination_sheet = undefined;
  this.MAX_NUM_ROWS = 600;
  this.MAX_NUM_COLUMNS = 30;
}

ScheduleRenderer.prototype.setDestinationSheet = function(destination_sheet) {
  this.destination_sheet = destination_sheet;
}

ScheduleRenderer.prototype.setTimesToRender = function(timeList) {
  var keys = Object.keys(this.dayTimeSlotLists);
  for (var timeIntervalIdx = 0; timeIntervalIdx < timeList.length; timeIntervalIdx++) {
    for (var keyIdx = 0; keyIdx < keys.length; keyIdx++) {
      if (timeList[timeIntervalIdx].days.join().indexOf(keys[keyIdx]) != -1) { // true if the time includes this day
        var dayTimeSlotList = this.dayTimeSlotLists[keys[keyIdx]]['times'];
        if (this.findTimeIntervalIndex(dayTimeSlotList, timeList[timeIntervalIdx]) == -1) {
          dayTimeSlotList.push(timeList[timeIntervalIdx]);
          dayTimeSlotList.sort(function (a, b) { return a.start - b.start; });
        }
      }
    }
  }
  this.renderEmptySchedule();
}

ScheduleRenderer.prototype.setRoomsToRender = function(roomList) {
  var keys = Object.keys(this.dayTimeSlotLists);
  for (var roomIdx = 0; roomIdx < roomList.length; roomIdx++) {
    for (var keyIdx = 0; keyIdx < keys.length; keyIdx++) {
      var dayRoomList = this.dayTimeSlotLists[keys[keyIdx]]['rooms'];
      if (this.findRoomIndex(dayRoomList, roomList[roomIdx]) == -1) {
        dayRoomList.push(roomList[roomIdx]);
        dayRoomList.sort(function (a, b) { return a.getId() - b.getId(); });
      }
    }
  }
  this.renderEmptySchedule();
}

// Adds the passed course, and roomWithTimeInterval to the schedule
ScheduleRenderer.prototype.renderEmptySchedule = function() {
  this.destination_sheet.getRange(1, 1, this.MAX_NUM_ROWS, this.MAX_NUM_COLUMNS).clearContent();
  var keys = Object.keys(this.dayTimeSlotLists);
  var outputRows = [];
  for (var keyIdx = 0; keyIdx < keys.length; keyIdx++) {
    var courseDay = keys[keyIdx];
    outputRows[courseDay] = this.dayTimeSlotLists[courseDay]['startRow'];
  }
  
  // render the days of the week titles in appropriate columns of row 1 of spreadsheet
  var renderDaysColumnTitles = ['M','T','W','R','F'];
  for (var dayIdx = 0; dayIdx < renderDaysColumnTitles.length; dayIdx++) {
    var titleStr = this.dayTimeSlotLists[renderDaysColumnTitles[dayIdx]].title.toUpperCase();
    this.destination_sheet.getRange(1, this.dayTimeSlotLists[renderDaysColumnTitles[dayIdx]].output_column).setValue(titleStr);
  }
  
  for (var keyIdx = 0; keyIdx < keys.length; keyIdx++) {
    var courseDay = keys[keyIdx];
    var hideOutput = this.dayTimeSlotLists[courseDay]['hide'];
    if (hideOutput) {
      continue;
    }
    var dayTimeSlotList = this.dayTimeSlotLists[courseDay]['times'];
    var dayRoomList = this.dayTimeSlotLists[courseDay]['rooms'];
    var timeAndRoomColumn = this.dayTimeSlotLists[courseDay]['time_room_column'];
    var outputColumn = this.dayTimeSlotLists[courseDay]['output_column'];
    for (var timeIntervalIdx = 0; timeIntervalIdx < dayTimeSlotList.length; timeIntervalIdx++) {
      var timeInterval = dayTimeSlotList[timeIntervalIdx];
      var timeIntervalString = timeInterval.getTimeString(timeInterval.start) + " - " + timeInterval.getTimeString(timeInterval.end);
      this.destination_sheet.getRange(outputRows[courseDay], timeAndRoomColumn).setValue(timeIntervalString);
      outputRows[courseDay] = outputRows[courseDay] + 1;
      for (var roomsToRenderIdx = 0; roomsToRenderIdx < dayRoomList.length; roomsToRenderIdx++) {
        var room = dayRoomList[roomsToRenderIdx];
        this.destination_sheet.getRange(outputRows[courseDay], timeAndRoomColumn).setValue(room.getId());
        outputRows[courseDay] = outputRows[courseDay] + 1;
      }
    }
  }
}

// Adds the passed course, and roomWithTimeInterval to the schedule
ScheduleRenderer.prototype.addCourseToSchedule = function(newScheduledCourse) {
  this.currentSchedule.push(newScheduledCourse);
  //this.currentSchedule.sort(function (a, b) { return a.CourseTime.start - b.CourseTime.start; });
  if (this.destination_sheet == undefined) {
    return;
  }

  var slotsHaveChanged = false;
  var courseTime = newScheduledCourse.CourseTime;
  var courseRoom = newScheduledCourse.Room;

  for (var dayIdx = 0; dayIdx < courseTime.days.length; dayIdx++) {
    var courseDay = courseTime.days[dayIdx];
    var dayTimeSlotList = this.dayTimeSlotLists[courseDay]['times'];
    var dayRoomList = this.dayTimeSlotLists[courseDay]['rooms'];
      
    // check if the time interval is not in the list of candidate time intervals for this day, if not add it and re-draw schedule
    // this only checks for a time slot with a matching start time (it does not match start and end times)
    var timeIntervalIndex = this.findTimeIntervalIndex(dayTimeSlotList, courseTime);
    if (timeIntervalIndex == -1 && false) {
      //SpreadsheetApp.getUi().alert('Added new time interval ' +  courseTime.getId() + '.');  
      dayTimeSlotList.push(courseTime);
      dayTimeSlotList.sort(function (a, b) { return a.start - b.start; });
      slotsHaveChanged = true;
    }

    // check if the room is not in the list of candidate rooms for this day, if not add it and re-draw schedule
    var roomIndex = this.findRoomIndex(dayRoomList, courseRoom);
    if (roomIndex == -1) {
      dayRoomList.push(courseRoom);
      dayRoomList.sort(function (a, b) { return a.getId() - b.getId(); });
      slotsHaveChanged = true;
    }
  }
  
  if (slotsHaveChanged) { // clear the rendered schedule and re-render all slots on all days
    this.renderEmptySchedule();
  }
  
  var coursesToRender = [];
  if (slotsHaveChanged) {
    coursesToRender = this.currentSchedule;
  } else {
    coursesToRender.push(newScheduledCourse);
  }
  
  // add code for courses which span multiple time slots (add * to slot text)
  for (var courseIdx = 0; courseIdx < coursesToRender.length; courseIdx++) {
    var courseToRender = coursesToRender[courseIdx];
    var courseTime = courseToRender.CourseTime;
    var courseRoom = courseToRender.Room;
    var facultyNameList = courseToRender.FacultyCoursesAndPrefs;
    var courseInstructionalMode = courseToRender.Course.instructional_mode;
    //var numTimeSlots = courseToRender.CourseTime;
    // TODO: Track the credit hours of the course to only insert the number of classes
    // consistent with the credit hours.
    for (var dayIdx = 0; dayIdx < courseToRender.CourseTime.days.length; dayIdx++) {
      var courseDay = courseTime.days[dayIdx];
      
      var dayTimeSlotList = this.dayTimeSlotLists[courseDay]['times'];
      var dayRoomList = this.dayTimeSlotLists[courseDay]['rooms'];
      var numRoomsToRender = dayRoomList.length;
      var numTimeIntervalsToRender = dayTimeSlotList.length;      

      var outputColumn = this.dayTimeSlotLists[courseDay]['output_column'];
      var startRow = this.dayTimeSlotLists[courseDay]['startRow'];
      
      var roomIndex = this.findRoomIndex(dayRoomList, courseRoom);
      //var timeIntervalIndex = this.findTimeIntervalIndex(dayTimeSlotList, courseTime);
      var overlappingTimeIntervalIndices = this.getOverlappingTimeIntervalIndices(dayTimeSlotList, courseTime);
      
      for (var timeIntervalIdx = 0; timeIntervalIdx < overlappingTimeIntervalIndices.length; timeIntervalIdx++) {
        var timeIntervalIndex = overlappingTimeIntervalIndices[timeIntervalIdx];
        var multipleTimeSlots = (timeIntervalIdx > 0) ? "*" : "";
        var courseTimeAndRoomRow = timeIntervalIndex*(numRoomsToRender+1) + roomIndex + startRow + 1;
        var facultyNameList = courseToRender.FacultyCoursesAndPrefs;
        var courseInstructionalMode = courseToRender.Course.instructional_mode;
        //var daysStr = courseToRender.CourseTime.days.join('/');
        if (facultyNameList) {
          facultyNameList = courseToRender.FacultyCoursesAndPrefs.name.split(" ");
        }
        var instructionalModeAcronym = this.instructionalModeToString[courseInstructionalMode];
        var facultyName = "";
        var facultyprefCost = "";
        if (courseToRender.FacultyCoursesAndPrefs != undefined) {
          facultyName = " " + facultyNameList[0].replace(/,\s*$/, "") + " ";
          facultyprefCost = (courseToRender.costArray['faculty_time'] == undefined) ? 'X' : courseToRender.costArray['faculty_time'];
        }
        var slotString = //DEBUG courseTime.getId() + " " + courseRoom.getId() + " " + daysStr + " " +
            multipleTimeSlots + courseToRender.Course.getId() + " " + instructionalModeAcronym + facultyName + facultyprefCost;
        var curContent = this.destination_sheet.getRange(courseTimeAndRoomRow, outputColumn).getValue();
        if (curContent != "") {
          this.destination_sheet.getRange(courseTimeAndRoomRow, outputColumn).setValue(curContent + '\n' + slotString);
        } else {
          this.destination_sheet.getRange(courseTimeAndRoomRow, outputColumn).setValue(slotString);
        }
      }
    }
    // write the cost to the cost column
    var costColumn = this.dayTimeSlotLists[courseDay]['cost_column'];
    var curCost = this.destination_sheet.getRange( courseTimeAndRoomRow, costColumn).getValue();
    if (curCost != "") {
      this.destination_sheet.getRange( courseTimeAndRoomRow, costColumn).setValue(curCost + ', ' + courseToRender.cost.toFixed(2));
    } else {
      this.destination_sheet.getRange( courseTimeAndRoomRow, costColumn).setValue(courseToRender.cost.toFixed(2));
    }
  }
}

ScheduleRenderer.prototype.getOverlappingTimeIntervalIndices = function (timeIntervalArray, timeInterval) {
  var overlappingTimeIntervalIndices = []
  for (var i = 0; i < timeIntervalArray.length; i++) {
    if (timeIntervalArray[i].overlaps(timeInterval)) {
      overlappingTimeIntervalIndices.push(i);
    }
  }
  return overlappingTimeIntervalIndices;
}


ScheduleRenderer.prototype.findRoomIndex = function (roomArray, room) {
  for (var i = 0; i < roomArray.length; i++) {
    if (roomArray[i].getId() == room.getId()) {
      return i;
    }
  }
  return -1;
}

ScheduleRenderer.prototype.findTimeIntervalIndex = function (timeIntervalArray, timeInterval) {
  for (var i = 0; i < timeIntervalArray.length; i++) {
    if (timeIntervalArray[i].start - timeInterval.start === 0) {
      return i;
    }
  }
  return -1;
}

ScheduleRenderer.prototype.findDaysOfWeekCategoryIndex = function (timeInterval) {
  for (var categoryIdx = 0; categoryIdx < this.timeSlotCategories.length; categoryIdx++) {
    if (this.timeSlotCategories[categoryIdx].indexOf(timeInterval.days[0]) != -1) {      
      return categoryIdx;
    }
  }
  return -1;
}

