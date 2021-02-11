// 
// This file is part of the UNCC ECE Scheduling Software and distributed 
// as a Google script embedded as part of a Google Sheets Spreadsheet
// Copyright (c) 2019, 2020, 2021 Andrew Willis, All rights reserved.
//
// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions
// are met:
//
//   * Redistributions of source code must retain the above copyright
//     notice, this list of conditions and the following GPL license text.
//
// This program is free software: you can redistribute it and/or modify  
// it under the terms of the GNU General Public License as published by  
// the Free Software Foundation, version 3.
// 
//  This program is distributed in the hope that it will be useful, but 
//  WITHOUT ANY WARRANTY; without even the implied warranty of 
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU 
//  General Public License for more details.
// 
//  You should have received a copy of the GNU General Public License 
//  along with this program. If not, see <http://www.gnu.org/licenses/>.
//

//
// The ECE Course Scheduling Engine
// Andrew Willis
//      June 15, 2019
// September 29, 2020
// November  12, 2020
// January   22, 2021 

/////////// KNOWN SHORTCOMINGS ///////////
// NO SEARCH YET, i.e., GREEDY OPTIMIZATION, ASSIGNMENTS ARE SEQUENTIAL WITH NO BACKTRACKING
// NO VARIABILITY WHEN MULTIPLE OPTIONS HAVE SAME COST: UNDER-CONSTRAINED CHOICES SELECT TIMES AND ROOMS BY TAKING ARRAY INDEX 0
// CONSISTENCY IN FACULTY->COURSE ASSOCIATIONS ARE NOT ENFORCED (NO SCRIPT-DRIVEN DATA VALIDATION AT THE MOMENT) 
// CONSISTENCY SPECIFIC: CANNOT DETECT WHEN A COURSE ASSIGNED TO A FACULTY IS NOT IN THE SCHEDULE
// WHAT IS THE BEST ORDERING FOR SCHEDULING COURSES? PRIORITY-BASED ORDERING WORKS BUT IS THERE SOMETHING BETTER?
// IMPROVE EMAIL FACULTY FUNCTION: PROVIDE SELECTION DIALOG WITH SELECT ALL FUNCTIONALITY
//

// Create google sheet menu items
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Scheduling Engine')
    .addItem('Compute New Schedule', 'computeNewSchedule')
    .addItem('Revise Prior Schedule', 'revisePriorSchedule')
    .addSeparator()
    .addSubMenu(ui.createMenu('Import')
      .addItem('Pull Faculty Preferences from Import Preferences Tab', 'pullFacultyPreferencesFromImportTab')
      .addItem('Pull Faculty Course Assignments from Teaching Plan (not implemented)', 'pullCoursesAndAssignmentsFromTeachingPlan')
      .addItem('Pull Academic Affairs Built Courses and Validate', 'pullUniversityBuiltCoursesAndValidate')
      .addItem('Pull Banner Built Courses and Validate', 'pullBannerCoursesAndValidate'))
    .addSubMenu(ui.createMenu('Reporting')
      .addItem('Cross-Validate Scheduled Courses with Faculty Course Assignments (not implemented)', 'crossValidateScheduledCoursesAndFacultyCourseAssignments')
      .addItem('Create Faculty Course Assignment Cost Report (not implemented)', 'createFacultyCourseAssignmentCostReport'))
    .addSubMenu(ui.createMenu('Export')
      .addItem('University Course Building Sheet', 'exportToUniversityCourseBuildingFormat'))
    .addSubMenu(ui.createMenu('Faculty')
      .addItem('eMail Faculty: Confirm Course Assignments and Time Preferences (not implemented)', 'emailFacultyAssignments')
      .addItem('eMail Faculty: Notify of Scheduled Course and Time Assignments', 'emailFacultyScheduledCourseAndTimeAssignments')
      .addItem('eMail Faculty: Solicit Titles for Special Topics Courses', 'emailFacultySolicitSpecialTopicsCourseTitles')
      .addItem('Create Calendar Events For Teaching Assignments (not implemented)', 'createCalendarEventsForTeachingTimes'))
    .addSubMenu(ui.createMenu('Debug')
      .addItem('Show Logs (not implemented)', 'showLogs'))
    .addToUi();
}

var SHEET_NAME_CONFIGURATION = "Config";
var SHEET_NAME_COURSE_CONSTRAINTS = "Courses";
var SHEET_NAME_TIME_SLOTS = "Time Intervals";
var SHEET_NAME_FACULTY_COURSES_AND_PREFERENCES = "Faculty Prefs & Courses";
var SHEET_NAME_ROLLED_CRNS = "CRNs";
var SHEET_NAME_ROOM_SLOTS = "Rooms";
var SHEET_NAME_COURSE_AND_TIME_CONSTRAINTS = "Pre-Scheduled Courses";
var SHEET_NAME_OUTPUT_TEMPLATE = "Template";
var SHEET_NAME_OUTPUT_SHEET = "Schedule Output";
var SHEET_NAME_OUTPUT_SHEET_TEST = "Schedule Output 2";
var SHEET_NAME_PRIOR_SCHEDULE_SHEET = "Prior Schedule";
var SHEET_NAME_IMPORT_FACULTY_PREFERENCES = "Imported Faculty Prefs";
var SHEET_NAME_IMPORT_REGISTRAR_COURSES = "Imported Registrar Course List";
var SHEET_NAME_EXPORT_COURSE_BUILDING = "Exported Course Builds";
var SHEET_NAME_OUTPUT_SHEET_AA = "AA Schedule Output";
var SHEET_NAME_OUTPUT_SHEET_BANNER = "Banner Schedule Output";

var DEFAULT_DEPT_CODE = "ECGR";

// HARD-CODED CONSTANTS
var OUTPUT_SHEET_INDEX_ROW_COST_OUTPUT = 5;
var OUTPUT_SHEET_INDEX_COL_COST_OUTPUT = 7;
var OUTPUT_SHEET_INDEX_ROW_NUMCOURSES_OUTPUT = 7;
var OUTPUT_SHEET_INDEX_COL_NUMCOURSES_OUTPUT = 7;

var defaultScheduleRenderer = null;

function testBannerLogin() {
  if (false) {
    var options =
    {
      "method": "post",
      "payload": payload,
      "followRedirects": false
    };
    var headers = {
      'Connection': 'keep-alive',
      'Content-Type': 'application/json;charset=utf-8',
      'Accept': 'application/json, text/plain, */*',
      'Cookie': '...',
    }
    var login = UrlFetchApp.fetch("https://webauth.uncc.edu/idp/profile/SAML2/Redirect/SSO?execution=e3s1", options);
    var sessionDetails = login.getAllHeaders()['Set-Cookie'];
    var downloadPayload =
    {
      "__EVENTTARGET": 'ctl00$ActionsPlaceHolder$exportDownloadLink1',
    };// This is just an example it may or may not be needed, if needed u need to trace the values from the developer tools.
  }
  var payload =
  {
    "sel_subj": DEFAULT_DEPT_CODE,
  };
  var downloadCsv = UrlFetchApp.fetch("https://selfservice.uncc.edu/pls/BANPROD/bwckgens.p_proc_term_date");
  Logger.log(downloadCsv.getContentText())
}

function pullUniversityBuiltCoursesAndValidate() {
  // dialog by clicking the close button in its title bar.
  var DATA_FORMAT_REGISTRAR = false;
  var DATA_FORMAT_REPORTS_CENTRAL = false;
  var ui = SpreadsheetApp.getUi();
  var data_sheet_response = ui.alert('Please confirm:', 'Import Data from ' + SHEET_NAME_IMPORT_REGISTRAR_COURSES + '?', ui.ButtonSet.YES_NO_CANCEL);
  // Process the user's response.
  if (data_sheet_response == ui.Button.YES) {
    var data_format_response = ui.alert('Please confirm:', 'Click Yes to indicate the schedule data is formatted by the registrar (Nickcoy Findlater) and No if the data is formatted as a Reports Central Worksheet?', ui.ButtonSet.YES_NO);
    if (data_format_response == ui.Button.YES) {
      DATA_FORMAT_REGISTRAR = true;
    } else if (data_format_response == ui.Button.NO) {
      DATA_FORMAT_REPORTS_CENTRAL = true;
    } else {
      ui.alert('Canceling schedule data import.');
      return;
    }
  } else {
    ui.alert('Canceling schedule data import.');
    return;
  }
  if (DATA_FORMAT_REGISTRAR == true) {
    /* Reports from AA - Monique Wilson had this format
      var COLUMN_INDEX_SUBJECT = 5;
      var COLUMN_INDEX_COURSE_NUMBER = 6;
      var COLUMN_INDEX_SECTION = 7;
      var COLUMN_INDEX_CRN_NUMBER = 8;
      var COLUMN_INDEX_XLST_ID = 9;
      var COLUMN_INDEX_BILL_HR = 10;    // NOT USED
      var COLUMN_INDEX_CR = 11;
      var COLUMN_INDEX_MO = 12;
      var COLUMN_INDEX_TU = 13;
      var COLUMN_INDEX_WE = 14;
      var COLUMN_INDEX_TH = 15;
      var COLUMN_INDEX_FR = 16;
      var COLUMN_INDEX_SA = 17;
      var COLUMN_INDEX_SU = 18;
      var COLUMN_INDEX_PART_TERM = 19;  // NOT USED
      var COLUMN_INDEX_BEGINS = 20;
      var COLUMN_INDEX_ENDS = 21;
      var COLUMN_INDEX_FIRST_NAME = 22; // NOT USED
      var COLUMN_INDEX_INSTRCTR = 23;
      var COLUMN_INDEX_BUILDING = 24;
      var COLUMN_INDEX_ROOM = 25;  
      var ROW_INDEX_FIRST_COURSE = 1;
      */
    // Reports from AA - Nickcoy Findlater have this format
    // Reports from Registrar via Reports Central for No College Designation have this format
    var COLUMN_INDEX_START_DATE = 0;
    var COLUMN_INDEX_END_DATE = 1;
    var COLUMN_INDEX_SUBJECT = 2;
    var COLUMN_INDEX_COURSE_NUMBER = 3;
    var COLUMN_INDEX_SECTION = 4;
    var COLUMN_INDEX_TITLE = 5; // This field started appearing 12/1 for Summer 2021 schedules
    var COLUMN_INDEX_CRN_NUMBER = 6;
    var COLUMN_INDEX_XLST_ID = 7;
    var COLUMN_INDEX_BILL_HR = 8;    // NOT USED
    var COLUMN_INDEX_CR = 9;
    var COLUMN_INDEX_MO = 10;
    var COLUMN_INDEX_TU = 11;
    var COLUMN_INDEX_WE = 12;
    var COLUMN_INDEX_TH = 13;
    var COLUMN_INDEX_FR = 14;
    var COLUMN_INDEX_SA = 15;
    var COLUMN_INDEX_SU = 16;
    var COLUMN_INDEX_PART_TERM = 17;  // NOT USED
    var COLUMN_INDEX_MILITARY_BEGIN = 18;
    var COLUMN_INDEX_MILITARY_END = 19;
    var COLUMN_INDEX_INSTRUCTOR_FIRST_NAME = 20; // NOT USED
    var COLUMN_INDEX_INSTRUCTOR_LAST_NAME = 21;
    var COLUMN_INDEX_BUILDING = 22;
    var COLUMN_INDEX_ROOM = 23;
    var COLUMN_INDEX_MODE = 24;
    var COLUMN_INDEX_INSTRUCTIONAL_MODE = 25;
    var COLUMN_INDEX_CAMPUS = 26;
    var COLUMN_INDEX_ROOM_CAPACITY = 27;
    var COLUMN_INDEX_MAX = 28;
    var COLUMN_INDEX_ENR = 28;
    var COLUMN_INDEX_PROJ = 29;
    var COLUMN_INDEX_PRIOR = 30;
    var COLUMN_INDEX_STATUS = 31;
  } else if (DATA_FORMAT_REPORTS_CENTRAL == true) {
    // Reports from Registrar via Reports Central for Dept. Worksheet have this format
    var COLUMN_INDEX_PART_OF_TERM = 0;
    var COLUMN_INDEX_COURSE = 1;
    var COLUMN_INDEX_SUBJECT = 2;
    var COLUMN_INDEX_COURSE_NUMBER = 3;
    var COLUMN_INDEX_SECTION = 4;
    var COLUMN_INDEX_CRN_NUMBER = 5;
    var COLUMN_INDEX_STATUS = 6;
    var COLUMN_INDEX_GRADE_TYPE = 7;    // NOT USED
    var COLUMN_INDEX_COURSE_TITLE = 8;  // NOT USED
    var COLUMN_INDEX_LONG_TITLE = 9;    // NOT USED
    var COLUMN_INDEX_XLST_ID = 10;
    var COLUMN_INDEX_BILL_HR = 11;      // NOT USED
    var COLUMN_INDEX_CR = 12;
    var COLUMN_INDEX_MO = 13;
    var COLUMN_INDEX_TU = 14;
    var COLUMN_INDEX_WE = 15;
    var COLUMN_INDEX_TH = 16;
    var COLUMN_INDEX_FR = 17;
    var COLUMN_INDEX_SA = 18;
    var COLUMN_INDEX_SU = 19;
    var COLUMN_INDEX_PART_TERM = 20;    // NOT USED
    var COLUMN_INDEX_CLASS_BEGIN = 21;
    var COLUMN_INDEX_CLASS_END = 22;
    var COLUMN_INDEX_MILITARY_BEGIN = 23;
    var COLUMN_INDEX_MILITARY_END = 24;
    var COLUMN_INDEX_INSTRUCTOR_LAST_NAME = 25;
    var COLUMN_INDEX_INSTRUCTOR_FIRST_NAME = 26; // NOT USED
    var COLUMN_INDEX_INSTRUCTOR_ID = 27; // NOT USED
    var COLUMN_INDEX_INSTRUCTOR_EMAIL = 29; // NOT USED
    var COLUMN_INDEX_BUILDING = 31;
    var COLUMN_INDEX_ROOM = 32;
    var COLUMN_INDEX_ROOM_CAPACITY = 33;
  } else {
    ui.alert('No valid import format found, canceling schedule data import.');
    return;
  }
  var ROW_INDEX_FIRST_COURSE = 1;
  var NO_DATA_STRING = '.';
  var PART_TERM_DECODING_DICTIONARY = { 1: 'Full', H1: 'T1', H2: 'T2' };

  var schedule = SpreadsheetApp.getActiveSpreadsheet();
  // Load Time Interval data
  var time_slot_sheet = schedule.getSheetByName(SHEET_NAME_TIME_SLOTS);
  var time_interval_datarange = time_slot_sheet.getDataRange();

  // Load Faculty Course Assignments and Time Interval Preference data
  var faculty_sheet = schedule.getSheetByName(SHEET_NAME_FACULTY_COURSES_AND_PREFERENCES);
  var faculty_datarange = faculty_sheet.getDataRange();

  // Load Faculty Course Assignments and Time Interval Preference data
  var faculty_sheet = schedule.getSheetByName(SHEET_NAME_FACULTY_COURSES_AND_PREFERENCES);
  var faculty_datarange = faculty_sheet.getDataRange();

  // Load AA course data 
  // 1fUfalSYLBQO5iBr0LJMMPx0k3hlet7xBaI0lIjOIuws
  //var aa_course_spreadsheet = SpreadsheetApp.openById('1a-z3_fURib15mTKc_iJ0g68u6CAqqIh2cd8PUqq1-7E');//("202010.09.17.19");
  //var aa_course_spreadsheet = SpreadsheetApp.openById('1k6zzPr3rzWSdoQ0z8y2ZpcqJEjOYCTZAGNHwRR1Xo88');//("2-10-2020");
  //var aa_course_spreadsheet = SpreadsheetApp.openById('1B7a7X7w7HSicycuinwKFwv7bi4nyCMQsHpMwTKNr70k');//("2-10-2020");
  //var aa_course_sheet = aa_course_spreadsheet.getSheetByName('Sheet1');
  var aa_course_sheet = schedule.getSheetByName(SHEET_NAME_IMPORT_REGISTRAR_COURSES);

  var template_output_sheet = schedule.getSheetByName(SHEET_NAME_OUTPUT_TEMPLATE);
  var output_sheet_aa = schedule.getSheetByName(SHEET_NAME_OUTPUT_SHEET_AA);

  if (!time_slot_sheet || !faculty_sheet || !output_sheet_aa || !aa_course_sheet) {
    SpreadsheetApp.getUi().alert('Could not read sheet data.');
    throw Error("Exiting due to sheet data access error.");
  }

  // Copy template sheet to output sheet  
  var template_data_range = template_output_sheet.getDataRange().getDataRegion();
  var template_data_values = template_data_range.getValues();
  template_data_range.copyFormatToRange(output_sheet_aa, template_data_range.getColumn(),
    template_data_range.getColumn() + template_data_range.getWidth(),
    template_data_range.getRow(), template_data_range.getRow() + template_data_range.getHeight());
  template_data_range.copyValuesToRange(output_sheet_aa, template_data_range.getColumn(),
    template_data_range.getColumn() + template_data_range.getWidth(),
    template_data_range.getRow(), template_data_range.getRow() + template_data_range.getHeight());

  var output_sheet_data = output_sheet_aa.getDataRange();
  var timeIntervalList = getTimeIntervalsToSchedule(time_interval_datarange);
  var facultyCoursesAndPrefsList = getFacultyCoursesAndPreferencesToSchedule(faculty_datarange, timeIntervalList);

  // Get full range of data
  var aa_data_range = aa_course_sheet.getDataRange();
  // get the data values in range
  var aa_course_data = aa_data_range.getValues();

  var scheduledCourseList = [];
  var daysArr = ['M', 'T', 'W', 'R', 'F', 'S', 'U'];
  var tmp_course_time = new CourseTime('8:00 AM', '8:50 AM', 'M/W/F', 3, [], [], 0);
  for (var courseIdx = ROW_INDEX_FIRST_COURSE; courseIdx < aa_course_data.length; courseIdx++) {
    if (aa_course_data[courseIdx][COLUMN_INDEX_SUBJECT] == DEFAULT_DEPT_CODE || //false) { // ECGR course or ENGR course with section number starting with 'E'
      (aa_course_data[courseIdx][COLUMN_INDEX_SUBJECT] == 'ENGR' && aa_course_data[courseIdx][COLUMN_INDEX_SECTION].toString().substring(0, 1) == 'E')) {
      var courseLogMsg = 'Reading row ' + (courseIdx + 1) + ' : ' + aa_course_data[courseIdx][COLUMN_INDEX_SUBJECT] + ' ' +
        aa_course_data[courseIdx][COLUMN_INDEX_COURSE_NUMBER] + '-' + aa_course_data[courseIdx][COLUMN_INDEX_SECTION];
      Logger.log(courseLogMsg);
      var aa_builtcourse = aa_course_data[courseIdx];
      var daysList = [];
      try {
        for (var daysColIdx = 0; daysColIdx <= 6; daysColIdx++) {
          if (aa_builtcourse[COLUMN_INDEX_MO + daysColIdx] == daysArr[daysColIdx]) {
            daysList.push(daysArr[daysColIdx]);
          }
        }
      } catch (e) {
        // Logs an ERROR message.
        Logger.log('Error Parsing days data: ' + e);
        continue;
      }

      try {
        var course_time = new CourseTime(tmp_course_time.timeStringToDate(aa_builtcourse[COLUMN_INDEX_MILITARY_BEGIN], '24hr'),
          tmp_course_time.timeStringToDate(aa_builtcourse[COLUMN_INDEX_MILITARY_END], '24hr'),
          daysList.join('/'),
          aa_builtcourse[COLUMN_INDEX_CR] / daysList.length,
          [], [], 0);
      } catch (e) {
        // Logs an ERROR message.
        Logger.log('Error Parsing creating course time object: ' + e);
        continue;
      }
      try {
        //var room = parseRoomListElement(aa_builtcourse[COLUMN_INDEX_BUILDING] + ' ' + aa_builtcourse[COLUMN_INDEX_ROOM]);
        var room = parseRoomListElement(aa_builtcourse[COLUMN_INDEX_BUILDING] + ' ' + aa_builtcourse[COLUMN_INDEX_ROOM]);
        if (aa_builtcourse[COLUMN_INDEX_ROOM] != room.number) {
          Logger.log('Error room number ' + aa_builtcourse[COLUMN_INDEX_ROOM] + ' and number component of building ' + aa_builtcourse[COLUMN_INDEX_BUILDING] + ' disagree.');
        }
      } catch (e) {
        // Logs an ERROR message.
        Logger.log('Error Parsing creating room object: ' + e);
        continue;
      }

      try {

        var aaScheduledCourse = new Course(aa_builtcourse[COLUMN_INDEX_SUBJECT],
          aa_builtcourse[COLUMN_INDEX_COURSE_NUMBER],
          //aa_builtcourse[COLUMN_INDEX_SECTION].replace('00',''),
          aa_builtcourse[COLUMN_INDEX_SECTION],
          aa_builtcourse[COLUMN_INDEX_CRN_NUMBER],
          PART_TERM_DECODING_DICTIONARY[aa_builtcourse[COLUMN_INDEX_PART_TERM]],
          course_time.days.length * course_time.credit_hours_per_day);

      } catch (e) {
        // Logs an ERROR message.
        Logger.log('Error Parsing creating course object: ' + e);
        continue;
      }

      var roomWithTimeInterval = new RoomWithTimeInterval(room, course_time);
      //var teacherForThisCourse = undefined;
      //var teacherForThisCourse = new FacultyCoursesAndPrefs(aa_builtcourse[COLUMN_INDEX_INSTRUCTOR_FIRST_NAME] + ' ' + aa_builtcourse[COLUMN_INDEX_INSTRUCTOR_LAST_NAME]);
      var firstInitial = aa_builtcourse[COLUMN_INDEX_INSTRUCTOR_FIRST_NAME].substring(0, 1);
      var teacherForThisCourse = new FacultyCoursesAndPrefs(aa_builtcourse[COLUMN_INDEX_INSTRUCTOR_LAST_NAME] + ', ' + firstInitial + '.');
      // DEBUG CODE
      //if (aa_builtcourse[COLUMN_INDEX_INSTRCTR] == 'Junior') {
      //  var aa = 1;
      //}
      var newScheduledCourse = new ScheduledCourse(aaScheduledCourse, room, course_time, 0, teacherForThisCourse, []);
      newScheduledCourse.cross_list_id = (aa_builtcourse[COLUMN_INDEX_XLST_ID] == NO_DATA_STRING) ? undefined : aa_builtcourse[COLUMN_INDEX_XLST_ID];
      scheduledCourseList.push(newScheduledCourse);
    }
  }

  // search all course pairs for courses in the same time slots these should be only cross-listed courses
  for (var courseAIdx = 0; courseAIdx < scheduledCourseList.length; courseAIdx++) {
    var aaScheduledCourseA = scheduledCourseList[courseAIdx];
    for (var courseBIdx = courseAIdx + 1; courseBIdx < scheduledCourseList.length; courseBIdx++) {
      // merge sections at the same time intervals here
      var aaScheduledCourseB = scheduledCourseList[courseBIdx];
      if (aaScheduledCourseA.cross_list_id != undefined && aaScheduledCourseA.cross_list_id == aaScheduledCourseB.cross_list_id &&
        aaScheduledCourseA.cross_list_id != '' &&
        aaScheduledCourseA.CourseTime.getId() == aaScheduledCourseB.CourseTime.getId() &&
        aaScheduledCourseA.Room.equals(aaScheduledCourseB.Room)) {
        if (aaScheduledCourseA.Course.section == aaScheduledCourseB.Course.section) {
          Logger.log('Cross Listed ' + aaScheduledCourseA.Course.getId() + ' with ' + aaScheduledCourseB.Course.getId());
          // merge and delete aaScheduledCourseB from the list
          aaScheduledCourseA.Course.numbers = aaScheduledCourseA.Course.numbers.concat(aaScheduledCourseB.Course.numbers);
          //if (aaScheduledCourseA.Course.section != aaScheduledCourseB.Course.section) {
          //  Logger.log('Merged result ' +  aaScheduledCourseA.Course.numbers.join('/') + ' has different section numbers ' + 
          //    aaScheduledCourseA.Course.section + ' != ' + aaScheduledCourseB.Course.section);
          //}
          scheduledCourseList.splice(courseBIdx, 1);
          courseBIdx--;
        }
      }
    }
  }

  for (var courseIdx = 0; courseIdx < scheduledCourseList.length; courseIdx++) {
    var aaScheduledCourse = scheduledCourseList[courseIdx];
    // Find the faculty member teaching this course (aaScheduledCourse) store the data in teacherForThisCourse      
    for (var facultyIdx = 0; teacherForThisCourse == undefined && facultyIdx < facultyCoursesAndPrefsList.length; facultyIdx++) {
      var faculty = facultyCoursesAndPrefsList[facultyIdx];
      for (var courseIdx2 = 0; courseIdx2 < faculty.courseList.length; courseIdx2++) {
        if (aaScheduledCourse.Course.sameNumber(faculty.courseList[courseIdx2])) {
          if (aaScheduledCourse.Course.sameSection(faculty.courseList[courseIdx2])) {
            //if(aaScheduledCourse.section == faculty.courseList[courseIdx2].section) {
            teacherForThisCourse = faculty;
            var aaNameArr = aaScheduledCourse.FacultyCoursesAndPrefs.name.split(' ');
            var eceNameArr = teacherForThisCourse.name.split(' ');
            if (eceNameArr[eceNameArr.length - 1] != aaNameArr[aaNameArr.length - 1]) {
              Logger.log('Instructor last name does not agree AA = ' + aaNameArr[aaNameArr.length - 1] + ' ECE = ' + eceNameArr[eceNameArr.length - 1]);
            }
            break;
          }
        }
      }
    }
  }

  var numScheduledCourses = 0;
  var numErrors = 0;
  while (scheduledCourseList.length > 0) {

    // get the next course from the prioritySortedCourseList
    var aaScheduledCourse = scheduledCourseList[0];

    Logger.log('ScheduleEngine: Scheduling ' + aaScheduledCourse.Course.getId());

    try {
      addCourseToSchedule(output_sheet_data, aaScheduledCourse);
      numScheduledCourses++;
    } catch (e) {
      numErrors++;
    }
    // remove course from prioritySortedCourseList
    scheduledCourseList.splice(scheduledCourseList.indexOf(aaScheduledCourse), 1);

    // remove rooms with overlapping time intervals from roomWithTimeIntervalList
    //removeFromList(availableRoomWithTimeInterval, roomWithTimeIntervalList, hasOverlappingRoomTimes);

    //output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_COST_OUTPUT, OUTPUT_SHEET_INDEX_COL_COST_OUTPUT).setValue(cost.toFixed(2));    
    output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_NUMCOURSES_OUTPUT, OUTPUT_SHEET_INDEX_COL_NUMCOURSES_OUTPUT).setValue(numScheduledCourses);
    Logger.log('ScheduleEngine: Scheduled ' + scheduledCourseList.length + " courses " + scheduledCourseList.length + " courses remain to be scheduled.");
  }
}

function pullBannerCoursesAndValidate() {
  var COLUMN_INDEX_CRN_NUMBER = 1;
  var COLUMN_INDEX_SUBJECT = 2;
  var COLUMN_INDEX_COURSE_NUMBER = 3;
  var COLUMN_INDEX_SECTION = 4;
  var COLUMN_INDEX_CMP = 5;
  var COLUMN_INDEX_CR = 6;
  var COLUMN_INDEX_TITLE = 7;
  var COLUMN_INDEX_DAYS = 8;
  var COLUMN_INDEX_TIME = 9;
  var COLUMN_INDEX_CAP = 10;
  var COLUMN_INDEX_ACT = 11;
  var COLUMN_INDEX_REM = 12;
  var COLUMN_INDEX_WL_CAP = 13;
  var COLUMN_INDEX_WL_ACT = 14;
  var COLUMN_INDEX_WL_REM = 15;
  var COLUMN_INDEX_XL_CAP = 16;
  var COLUMN_INDEX_XL_ACT = 17;
  var COLUMN_INDEX_XL_REM = 18;
  var COLUMN_INDEX_INSTRUCTOR = 19;  // NOT USED
  var COLUMN_INDEX_DATE = 20;
  var COLUMN_INDEX_SESSION = 21;
  var COLUMN_INDEX_LOCATION = 22; // NOT USED
  var ROW_INDEX_FIRST_COURSE = 1;

  var schedule = SpreadsheetApp.getActiveSpreadsheet();
  // Load Time Interval data
  var time_slot_sheet = schedule.getSheetByName(SHEET_NAME_TIME_SLOTS);
  var time_interval_datarange = time_slot_sheet.getDataRange();

  // Load Faculty Course Assignments and Time Interval Preference data
  var faculty_sheet = schedule.getSheetByName(SHEET_NAME_FACULTY_COURSES_AND_PREFERENCES);
  var faculty_datarange = faculty_sheet.getDataRange();

  // Load Faculty Course Assignments and Time Interval Preference data
  var faculty_sheet = schedule.getSheetByName(SHEET_NAME_FACULTY_COURSES_AND_PREFERENCES);
  var faculty_datarange = faculty_sheet.getDataRange();

  // Load Banner course data 
  var banner_course_spreadsheet = SpreadsheetApp.openById('118co6sAPtne6zIXTb-TDOzmXI67b-fCK0p3E9YlwJ1g');//("202010.09.17.19");
  var banner_course_sheet = banner_course_spreadsheet.getSheetByName('Sheet1');

  var template_output_sheet = schedule.getSheetByName(SHEET_NAME_OUTPUT_TEMPLATE);
  var output_sheet_banner = schedule.getSheetByName(SHEET_NAME_OUTPUT_SHEET_BANNER);

  if (!time_slot_sheet || !faculty_sheet || !output_sheet_banner || !banner_course_sheet) {
    SpreadsheetApp.getUi().alert('Could not read sheet data.');
    throw Error("Exiting due to sheet data access error.");
  }

  // Copy template sheet to output sheet  
  var template_data_range = template_output_sheet.getDataRange().getDataRegion();
  var template_data_values = template_data_range.getValues();
  template_data_range.copyFormatToRange(output_sheet_banner, template_data_range.getColumn(),
    template_data_range.getColumn() + template_data_range.getWidth(),
    template_data_range.getRow(), template_data_range.getRow() + template_data_range.getHeight());
  template_data_range.copyValuesToRange(output_sheet_banner, template_data_range.getColumn(),
    template_data_range.getColumn() + template_data_range.getWidth(),
    template_data_range.getRow(), template_data_range.getRow() + template_data_range.getHeight());

  var output_sheet_data = output_sheet_banner.getDataRange();
  var timeIntervalList = getTimeIntervalsToSchedule(time_interval_datarange);
  var facultyCoursesAndPrefsList = getFacultyCoursesAndPreferencesToSchedule(faculty_datarange, timeIntervalList);

  // Get full range of data
  var banner_data_range = banner_course_sheet.getDataRange();
  // get the data values in range
  var banner_course_data = banner_data_range.getValues();

  for (var courseIdx = ROW_INDEX_FIRST_COURSE; courseIdx < banner_course_data.length; courseIdx++) {
    if (courseIdx > ROW_INDEX_FIRST_COURSE && banner_course_data[courseIdx][COLUMN_INDEX_SUBJECT] == '' && banner_course_data[courseIdx][COLUMN_INDEX_COURSE_NUMBER] == '') {
      for (var columnIdx = 0; columnIdx < banner_course_data[courseIdx].length; columnIdx++) {
        if (banner_course_data[courseIdx][columnIdx] == '') {
          banner_course_data[courseIdx][columnIdx] = banner_course_data[courseIdx - 1][columnIdx];
        }
      }
    }
  }

  var scheduledCourseList = [];
  var tmp_course_time = new CourseTime('8:00 AM', '8:50 AM', 'M/W/F', 3, [], [], 0);
  for (var courseIdx = ROW_INDEX_FIRST_COURSE; courseIdx < banner_course_data.length; courseIdx++) {
    if (banner_course_data[courseIdx][COLUMN_INDEX_SUBJECT] == DEFAULT_DEPT_CODE || //false) { // ECGR course or ENGR course with section number starting with 'E'
      (banner_course_data[courseIdx][COLUMN_INDEX_SUBJECT] == 'ENGR' && banner_course_data[courseIdx][COLUMN_INDEX_SECTION].toString().substring(0, 1) == 'E')) {
      var courseLogMsg = 'Reading row ' + (courseIdx + 1) + ' : ' + banner_course_data[courseIdx][COLUMN_INDEX_SUBJECT] + ' ' +
        banner_course_data[courseIdx][COLUMN_INDEX_COURSE_NUMBER] + '-' + banner_course_data[courseIdx][COLUMN_INDEX_SECTION];
      Logger.log(courseLogMsg);
      var banner_builtcourse = banner_course_data[courseIdx];
      var daysList = banner_builtcourse[COLUMN_INDEX_DAYS].split('');
      if (banner_builtcourse[COLUMN_INDEX_TIME] == 'TBA') {
        Logger.log('Course ' + banner_course_data[courseIdx][COLUMN_INDEX_SUBJECT] + ' ' + banner_course_data[courseIdx][COLUMN_INDEX_COURSE_NUMBER] + ' has time = TBA. Skipping this course.');
        continue;
      }
      try {
        var start_stop_timeStringArr = banner_builtcourse[COLUMN_INDEX_TIME].split('-');
        var course_time = new CourseTime(tmp_course_time.timeStringToDate(start_stop_timeStringArr[0]),
          tmp_course_time.timeStringToDate(start_stop_timeStringArr[1]),
          daysList.join('/'),
          banner_builtcourse[COLUMN_INDEX_CR] / daysList.length,
          [], [], 0);
      } catch (e) {
        // Logs an ERROR message.
        Logger.log('Error Parsing creating course time object: ' + e);
        continue;
      }
      try {
        //var room = parseRoomListElement(banner_builtcourse[COLUMN_INDEX_LOCATION]);
        var room = parseRoomListElement(banner_builtcourse[COLUMN_INDEX_LOCATION]);
        //if (banner_builtcourse[COLUMN_INDEX_ROOM] != room.number) {
        //  Logger.log('Error room number ' + banner_builtcourse[COLUMN_INDEX_ROOM] + ' and number component of building ' + banner_builtcourse[COLUMN_INDEX_BUILDING] + ' disagree.');
        //}
      } catch (e) {
        // Logs an ERROR message.
        Logger.log('Error Parsing creating room object: ' + e);
        continue;
      }

      try {
        var bannerScheduledCourse = new Course(banner_builtcourse[COLUMN_INDEX_SUBJECT],
          banner_builtcourse[COLUMN_INDEX_COURSE_NUMBER],
          banner_builtcourse[COLUMN_INDEX_SECTION],
          banner_builtcourse[COLUMN_INDEX_CRN_NUMBER],
          "Full",
          course_time.days.length * course_time.credit_hours_per_day);

      } catch (e) {
        // Logs an ERROR message.
        Logger.log('Error Parsing creating course object: ' + e);
        continue;
      }

      var roomWithTimeInterval = new RoomWithTimeInterval(room, course_time);
      var teacherStr = banner_builtcourse[COLUMN_INDEX_INSTRUCTOR];
      teacherStr = teacherStr.substring(0, teacherStr.lastIndexOf(" "));
      var teacherForThisCourse = new FacultyCoursesAndPrefs(teacherStr);
      // DEBUG CODE
      //if (aa_builtcourse[COLUMN_INDEX_INSTRCTR] == 'Junior') {
      //  var aa = 1;
      //}
      var newScheduledCourse = new ScheduledCourse(bannerScheduledCourse, room, course_time, 0, teacherForThisCourse, []);
      //newScheduledCourse.cross_list_id = (aa_builtcourse[COLUMN_INDEX_XLST_ID] == NO_DATA_STRING) ? undefined : aa_builtcourse[COLUMN_INDEX_XLST_ID];
      scheduledCourseList.push(newScheduledCourse);
    }
  }

  // search all course pairs for courses in the same time slots these should be only cross-listed courses
  for (var courseAIdx = 0; courseAIdx < scheduledCourseList.length; courseAIdx++) {
    var aaScheduledCourseA = scheduledCourseList[courseAIdx];
    for (var courseBIdx = courseAIdx + 1; courseBIdx < scheduledCourseList.length; courseBIdx++) {
      // merge sections at the same time intervals here
      var aaScheduledCourseB = scheduledCourseList[courseBIdx];
      if (aaScheduledCourseA.CourseTime.getId() == aaScheduledCourseB.CourseTime.getId() &&
        aaScheduledCourseA.Room.equals(aaScheduledCourseB.Room)) {
        Logger.log('Cross Listed ' + aaScheduledCourseA.Course.getId() + ' with ' + aaScheduledCourseB.Course.getId());
        // merge and delete aaScheduledCourseB from the list
        aaScheduledCourseA.Course.numbers = aaScheduledCourseA.Course.numbers.concat(aaScheduledCourseB.Course.numbers);
        if (aaScheduledCourseA.Course.section != aaScheduledCourseB.Course.section) {
          Logger.log('Merged result ' + aaScheduledCourseA.Course.numbers.join('/') + ' has different section numbers ' +
            aaScheduledCourseA.Course.section + ' != ' + aaScheduledCourseB.Course.section);
        }
        scheduledCourseList.splice(courseBIdx, 1);
        courseBIdx--;
      }
    }
  }

  for (var courseIdx = 0; courseIdx < scheduledCourseList.length; courseIdx++) {
    var aaScheduledCourse = scheduledCourseList[courseIdx];
    // Find the faculty member teaching this course (aaScheduledCourse) store the data in teacherForThisCourse      
    for (var facultyIdx = 0; teacherForThisCourse == undefined && facultyIdx < facultyCoursesAndPrefsList.length; facultyIdx++) {
      var faculty = facultyCoursesAndPrefsList[facultyIdx];
      for (var courseIdx2 = 0; courseIdx2 < faculty.courseList.length; courseIdx2++) {
        if (aaScheduledCourse.Course.sameNumber(faculty.courseList[courseIdx2])) {
          if (aaScheduledCourse.Course.sameSection(faculty.courseList[courseIdx2])) {
            //if(aaScheduledCourse.section == faculty.courseList[courseIdx2].section) {
            teacherForThisCourse = faculty;
            var aaNameArr = aaScheduledCourse.FacultyCoursesAndPrefs.name.split(' ');
            var eceNameArr = teacherForThisCourse.name.split(' ');
            if (eceNameArr[eceNameArr.length - 1] != aaNameArr[aaNameArr.length - 1]) {
              Logger.log('Instructor last name does not agree AA = ' + aaNameArr[aaNameArr.length - 1] + ' ECE = ' + eceNameArr[eceNameArr.length - 1]);
            }
            break;
          }
        }
      }
    }
  }

  var numScheduledCourses = 0;
  var numErrors = 0;
  while (scheduledCourseList.length > 0) {

    // get the next course from the prioritySortedCourseList
    var aaScheduledCourse = scheduledCourseList[0];

    Logger.log('ScheduleEngine: Scheduling ' + aaScheduledCourse.Course.getId());

    try {
      addCourseToSchedule(output_sheet_data, aaScheduledCourse);
      numScheduledCourses++;
    } catch (e) {
      numErrors++;
    }
    // remove course from prioritySortedCourseList
    scheduledCourseList.splice(scheduledCourseList.indexOf(aaScheduledCourse), 1);

    // remove rooms with overlapping time intervals from roomWithTimeIntervalList
    //removeFromList(availableRoomWithTimeInterval, roomWithTimeIntervalList, hasOverlappingRoomTimes);

    //output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_COST_OUTPUT, OUTPUT_SHEET_INDEX_COL_COST_OUTPUT).setValue(cost.toFixed(2));    
    output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_NUMCOURSES_OUTPUT, OUTPUT_SHEET_INDEX_COL_NUMCOURSES_OUTPUT).setValue(numScheduledCourses);
    Logger.log('ScheduleEngine: Scheduled ' + scheduledCourseList.length + " courses " + scheduledCourseList.length + " courses remain to be scheduled.");
  }
}


function pullCoursesAndAssignmentsFromTeachingPlan() {
}

function pullFacultyPreferencesFromImportTab() {
  var schedule = SpreadsheetApp.getActiveSpreadsheet();
  // Load Time Interval data
  var time_slot_sheet = schedule.getSheetByName(SHEET_NAME_TIME_SLOTS);
  var time_interval_datarange = time_slot_sheet.getDataRange();

  // Load faculty preference data for import
  var import_faculty_prefs_sheet = schedule.getSheetByName(SHEET_NAME_IMPORT_FACULTY_PREFERENCES);
  var import_faculty_prefs_data = import_faculty_prefs_sheet.getDataRange();

  // Load Faculty Course Assignments and Time Interval Preference data
  var faculty_sheet = schedule.getSheetByName(SHEET_NAME_FACULTY_COURSES_AND_PREFERENCES);
  var faculty_datarange = faculty_sheet.getDataRange();

  // Load Prior Schedule data data
  var prior_schedule_datarange = faculty_sheet.getDataRange();

  if (!schedule || !time_slot_sheet || !import_faculty_prefs_sheet || !faculty_sheet) {
    SpreadsheetApp.getUi().alert('Could not read sheet data.');
    throw Error("Exiting due to sheet data access error.");
  }

  var timeIntervalList = getTimeIntervalsToSchedule(time_interval_datarange);
  var importedFacultyPreferenceList = importFacultyPreferences(import_faculty_prefs_data, timeIntervalList);
  var facultyCoursesAndPrefsList = getFacultyCoursesAndPreferencesToSchedule(faculty_datarange, timeIntervalList);

  // match email addresses to determine destination of new preferences
  for (var importIdx = 0; importIdx < importedFacultyPreferenceList.length; importIdx++) {
    for (var exportIdx = 0; exportIdx < facultyCoursesAndPrefsList.length; exportIdx++) {
      if (importedFacultyPreferenceList[importIdx].email == facultyCoursesAndPrefsList[exportIdx].email) {
        var srcData = importedFacultyPreferenceList[importIdx];
        var dstData = facultyCoursesAndPrefsList[exportIdx];
        dstData.courses_on_same_days = srcData.courses_on_same_days;
        dstData.hours_between_courses = srcData.hours_between_courses;
        for (var key in srcData.timeIntervalCostMap) {
          var val = dstData.timeIntervalCostMap[key];
          var newval = srcData.timeIntervalCostMap[key];
          dstData.timeIntervalCostMap[key] = srcData.timeIntervalCostMap[key];
        }
      }
    }
  }

  putFacultyCoursesAndPreferences(faculty_datarange, timeIntervalList, facultyCoursesAndPrefsList);
}

function crossValidateScheduledCoursesAndFacultyCourseAssignments() {
}

function createFacultyCourseAssignmentCostReport() {
}

function createCalendarEventsForTeachingAssignments() {
}

function emailFacultyConfirmCourseAssignmentsAndTimePreferences() {
}

function emailFacultySolicitSpecialTopicsCourseTitles() {
  emailFaculty('special_topics_course_titles');
}

function emailFacultyScheduledCourseAndTimeAssignments() {
  emailFaculty('teaching_schedule');
}

function emailFaculty(functionName) {
  if (functionName != 'teaching_schedule' && functionName != 'special_topics_course_titles') {
    SpreadsheetApp.getUi().alert('emailFaculty() invoked with invalid functionName = ' + functionName + '.');
    throw Error("Exiting due to script error.");
  }
  var schedule = SpreadsheetApp.getActiveSpreadsheet();
  var time_slot_sheet = schedule.getSheetByName(SHEET_NAME_TIME_SLOTS);
  var faculty_sheet = schedule.getSheetByName(SHEET_NAME_FACULTY_COURSES_AND_PREFERENCES);
  var prior_schedule_sheet = schedule.getSheetByName(SHEET_NAME_PRIOR_SCHEDULE_SHEET);

  if (!schedule || !faculty_sheet || !prior_schedule_sheet) {
    SpreadsheetApp.getUi().alert('Could not read sheet data.');
    throw Error("Exiting due to sheet data access error.");
  }

  // Load Time Interval data
  var time_interval_datarange = time_slot_sheet.getDataRange();
  var timeIntervalList = getTimeIntervalsToSchedule(time_interval_datarange);

  // Load Faculty Course Assignments and Time Interval Preference data
  var faculty_datarange = faculty_sheet.getDataRange();
  var facultyCoursesAndPrefsList = getFacultyCoursesAndPreferencesToSchedule(faculty_datarange, timeIntervalList);

  // create a list of the scheduled courses from the "Prior Schedule" sheet
  var scheduledCourseList = [];
  transferScheduledCourses('READ', prior_schedule_sheet, scheduledCourseList, facultyCoursesAndPrefsList);

  // emails will not be sent to email addresses in the exemptEmail array
  //var exemptEmail = ['nbousaba@uncc.edu','brodri17@uncc.edu'];
  var exemptEmail = [];

  // Form letter email constants that change each semester / year  
  var config_sheet = schedule.getSheetByName(SHEET_NAME_CONFIGURATION);
  var SEMESTER = config_sheet.getRange(8, 3, 1, 1).getValue();
  var YEAR = config_sheet.getRange(9, 3, 1, 1).getValue();
  var SCHEDULE_MANAGER = config_sheet.getRange(10, 3, 1, 1).getValue();
  var SCHEDULE_MANAGER_EMAIL = config_sheet.getRange(11, 3, 1, 1).getValue();
  var SCHEDULE_MANAGER_PHONE_NUMBER = config_sheet.getRange(12, 3, 1, 1).getValue();

  var TEST_EMAIL = false;
  var MASS_DISTRIBUTION_EMAIL = false;
  // dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var response_test = ui.alert('Please confirm:', 'Send a test email to ' + SCHEDULE_MANAGER + '?', ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (response_test == ui.Button.YES) {
    TEST_EMAIL = true;
  } else {
    var response_all = ui.alert('Please confirm:', 'Send an email to ALL teaching faculty?', ui.ButtonSet.YES_NO);
    if (response_all == ui.Button.YES) {
      MASS_DISTRIBUTION_EMAIL = true;
    } else {
      ui.alert('Canceling email distribution to faculty.');
      return;
    }
  }

  // construct and send out emails
  var subjectStr;
  var replyToStr = SCHEDULE_MANAGER_EMAIL;
  var messagePreLine1;
  var messagePreLine2;
  var messagePostLine1
  var mgmtMessage = "Log for emailFaculty " + functionName + ":\n";
  if (functionName == 'teaching_schedule') {
    var DEADLINE_DATE = config_sheet.getRange(17, 3, 1, 1).getValue();
    subjectStr = SEMESTER + ' ' + YEAR + ' Course Teaching Schedule (PLEASE REPLY)';
    //replyToStr = SCHEDULE_MANAGER_EMAIL;
    messagePreLine1 = ['Dear ', undefined, ',', '\n'];
    messagePreLine2 = ['\n', 'As a final step to course scheduling for ' + SEMESTER + ' ' + YEAR + ' we are asking faculty to review their data within the current teaching schedule.',
      ' The following table indicates the schedule details for your ' + SEMESTER + ' ' + YEAR + ' course(s):', '\n'];
    messagePostLine1 = ['\n', 'Please confirm your schedule by responding to this email as soon as possible or, at latest, by close of business ' + DEADLINE_DATE + '.', '\n', '\n',
      'If you have any questions regarding your schedule please contact ' + SCHEDULE_MANAGER + ' (' + SCHEDULE_MANAGER_EMAIL + ') by replying to this email or calling '
      + SCHEDULE_MANAGER_PHONE_NUMBER + '.', '\n'];
  } else if (functionName == 'special_topics_course_titles') {
    var DEADLINE_DATE = config_sheet.getRange(20, 3, 1, 1).getValue();
    //SCHEDULE_MANAGER = 'Andrew Willis';
    //SCHEDULE_MANAGER_EMAIL = 'arwillis@uncc.edu';
    //SCHEDULE_MANAGER_PHONE_NUMBER = '704-687-8420';
    //DEADLINE_DATE = 'Saturday, September 7';
    subjectStr = SEMESTER + ' ' + YEAR + ' Special Topics Course Titles';
    //replyToStr = SCHEDULE_MANAGER_EMAIL;
    messagePreLine1 = ['Dear ', undefined, ',', '\n'];
    messagePreLine2 = ['\n', 'The schedule indicates that you will be teaching one or more special topics courses in ' + SEMESTER + ' ' + YEAR + '.',
      ' The scheduling system requires a title for each special topics course to add these courses to the schedule.',
      ' The following table indicates the schedule details for your ' + SEMESTER + ' ' + YEAR + ' special topics course(s):', '\n'];
    messagePostLine1 = ['\n', 'Please respond to this email indicating a final title for your special topics courses as soon as possible or, at latest, by close of business ' + DEADLINE_DATE + '.', '\n', '\n',
      'If you have any questions regarding this request please contact ' + SCHEDULE_MANAGER + ' (' + SCHEDULE_MANAGER_EMAIL + ') by replying to this email or calling '
      + SCHEDULE_MANAGER_PHONE_NUMBER + '.', '\n'];//,'\n',
    //'Apologies in advance if you have specified this information in other correspondence.',
    //' This email is part of the new effort to streamline the course scheduling process. Thank you for understanding.'];    
  }

  var logLine;
  var SENT_TEST_EMAIL = false;
  var term2Strings = { 'Full': 'Full-Term', 'T1': 'Term-1', 'T2': 'Term-2' };

  while (scheduledCourseList.length > 0) {
    // take the first course
    var scheduledCourse = scheduledCourseList[0];
    // remove the course (facultyScheduledCourses[0]) from the list of courses which require email notifications to be send
    //scheduledCourseList.splice( scheduledCourseList.indexOf(scheduledCourse), 1);
    // find the faculty email if available
    var facultyScheduledCourses = [];
    if (scheduledCourse.FacultyCoursesAndPrefs != undefined && scheduledCourse.FacultyCoursesAndPrefs.email != undefined &&
      exemptEmail.indexOf(scheduledCourse.FacultyCoursesAndPrefs.email) < 0 && isValidEmailAddress(scheduledCourse.FacultyCoursesAndPrefs.email)) {
      var faculty = scheduledCourse.FacultyCoursesAndPrefs;
      var dstEmailAddress = scheduledCourse.FacultyCoursesAndPrefs.email;
      // Find all other courses this faculty member will teach
      // start search from list end so indices don't change when array elements are deleted
      for (var courseIdx = scheduledCourseList.length - 1; courseIdx >= 0; courseIdx--) {
        if (scheduledCourseList[courseIdx].FacultyCoursesAndPrefs != undefined && scheduledCourseList[courseIdx].FacultyCoursesAndPrefs.email == dstEmailAddress) {
          if (functionName == 'teaching_schedule') {
            facultyScheduledCourses.push(scheduledCourseList[courseIdx]);
          } else if (functionName == 'special_topics_course_titles' && scheduledCourseList[courseIdx].Course.isSpecialTopics()) {
            facultyScheduledCourses.push(scheduledCourseList[courseIdx]);
          }
          // remove the course (scheduledCourseList[courseIdx]) from the list of courses which require email notifications to be send
          scheduledCourseList.splice(scheduledCourseList.indexOf(scheduledCourseList[courseIdx]), 1);
        }
      }
      // do not send notifications to faculty that have no courses scheduled relevant to functionName (no course assignments/no special topics courses)
      if (facultyScheduledCourses.length == 0) {
        continue;
      }
      // output a row in the email containing the following: Course Dept. Code, Course Number(s), Section, Time Slot, Days of the Week, Building, Room Number
      var messageSchedule = ['\n'];
      var caveatNONETBA = false;
      //messageSchedule.push('COURSE');
      //messageSchedule.push('NUMBER');
      //messageSchedule.push('SECTION');
      //messageSchedule.push('TIME INTERVAL');
      //messageSchedule.push('DAYS OF WEEK');
      //messageSchedule.push('BUILDING');
      //messageSchedule.push('ROOM NUMBER');
      //messageSchedule.push('\n');
      for (var courseIdx = 0; courseIdx < facultyScheduledCourses.length; courseIdx++) {
        var scheduledCourse = facultyScheduledCourses[courseIdx];
        messageSchedule.push(scheduledCourse.Course.dept_code);
        var str_course_number = scheduledCourse.Course.numbers.join('/') + '-' + scheduledCourse.Course.section
        while (str_course_number.length < 20) {
          str_course_number = str_course_number + ' ';
        }
        messageSchedule.push(str_course_number);
        var str_course_instructional_mode = scheduledCourse.Course.instructional_mode;
        while (str_course_instructional_mode.length < 20) {
          str_course_instructional_mode = str_course_instructional_mode + ' ';
        }
        messageSchedule.push(str_course_instructional_mode);
        var str_course_term = term2Strings[scheduledCourse.Course.term];
        while (str_course_term.length < 11) {
          str_course_term = str_course_term + ' ';
        }
        messageSchedule.push(str_course_term);
        var str_course_time = scheduledCourse.CourseTime.getId();
        while (str_course_time.length < 30) {
          str_course_time = str_course_time + ' ';
        }
        messageSchedule.push(str_course_time);
        messageSchedule.push(scheduledCourse.Room.building);
        messageSchedule.push(scheduledCourse.Room.number);
        messageSchedule.push('\n');
        if (scheduledCourse.Room.building == 'NONE' && scheduledCourse.Room.number == 'TBA') {
          caveatNONETBA = true;
        }
      }
      if (caveatNONETBA && functionName == 'teaching_schedule') {
        // output message explaining a custom NON-ECE room has been scheduled for clarification on the specific room contact jmconrad@uncc.edu
        messageSchedule.push('\nYour schedule includes a room listed as \'NONE TBA\'. This indicates the room for this course is scheduled into a custom location.');
        messageSchedule.push('For clarification on the specific location of these courses contact ' + SCHEDULE_MANAGER_EMAIL + '.\n');
      }

      messagePreLine1[1] = faculty.name;
      var mailMessage = messagePreLine1.join('') + messagePreLine2.join('') + messageSchedule.join(' ') + messagePostLine1.join(''); // Second column
      if ((TEST_EMAIL == true && SENT_TEST_EMAIL == false) || (TEST_EMAIL == false && MASS_DISTRIBUTION_EMAIL == true)) {
        if (TEST_EMAIL == true) {
          dstEmailAddress = SCHEDULE_MANAGER_EMAIL;
          SENT_TEST_EMAIL = true;
        }
        logLine = 'emailFaculty ' + functionName + ': Sent notification to ' + dstEmailAddress + '.';
        mgmtMessage += logLine + '\n' + functionName + ': ' + messageSchedule.join('     ');
        Logger.log(logLine);
        // UNCOMMENT THE LINE BELOW TO SEND OUT EMAILS
        MailApp.sendEmail(dstEmailAddress, replyToStr, subjectStr, mailMessage);
      }
    } else {
      scheduledCourseList.splice(scheduledCourseList.indexOf(scheduledCourse), 1);
      logLine = 'emailFacultyScheduledCourseAndTimeAssignment: No notification sent for course ' + scheduledCourse.Course.getId() + ' at time ' + scheduledCourse.CourseTime.getId();
      //mgmtMessage += logLine + '\n';
      Logger.log(logLine);
    }
  }
  MailApp.sendEmail(SCHEDULE_MANAGER_EMAIL, 'emailFaculty ' + functionName + ' Logs', mgmtMessage);
}

function isValidEmailAddress(email) {
  return /\S+@\S+\.\S+/.test(email)
}

function exportToUniversityCourseBuildingFormat() {
  var CROSS_LIST_CODE_PREFIX = 'Y';
  var schedule = SpreadsheetApp.getActiveSpreadsheet();
  var time_slot_sheet = schedule.getSheetByName(SHEET_NAME_TIME_SLOTS);
  var faculty_sheet = schedule.getSheetByName(SHEET_NAME_FACULTY_COURSES_AND_PREFERENCES);
  var room_slot_sheet = schedule.getSheetByName(SHEET_NAME_ROOM_SLOTS);
  var rolled_crns_sheet = schedule.getSheetByName(SHEET_NAME_ROLLED_CRNS);
  var prior_schedule_sheet = schedule.getSheetByName(SHEET_NAME_PRIOR_SCHEDULE_SHEET);
  var export_coursebuilding_sheet = schedule.getSheetByName(SHEET_NAME_EXPORT_COURSE_BUILDING);

  if (!schedule || !faculty_sheet || !prior_schedule_sheet || !export_coursebuilding_sheet) {
    SpreadsheetApp.getUi().alert('Could not read sheet data.');
    throw Error("Exiting due to sheet data access error.");
  }

  // Load Time Interval data
  var time_interval_datarange = time_slot_sheet.getDataRange();
  var timeIntervalList = getTimeIntervalsToSchedule(time_interval_datarange);

  // Load Faculty Course Assignments and Time Interval Preference data
  var faculty_datarange = faculty_sheet.getDataRange();
  var facultyCoursesAndPrefsList = getFacultyCoursesAndPreferencesToSchedule(faculty_datarange, timeIntervalList);

  // Load Room data 
  var room_data = room_slot_sheet.getDataRange();
  var roomList = getRoomsToSchedule(room_data);

  // Load rolled CRN data 
  var crn_data = rolled_crns_sheet.getDataRange();
  var crnList = getRolledCRNs(crn_data);

  // create a list of the scheduled courses from the "Prior Schedule" sheet
  var scheduledCourseList = [];
  transferScheduledCourses('READ', prior_schedule_sheet, scheduledCourseList, facultyCoursesAndPrefsList);
  var built_courses_data = [];
  var xlst_id = 1;
  export_coursebuilding_sheet.getRange(2, 1, 500, 52).clearContent();

  var config_sheet = schedule.getSheetByName(SHEET_NAME_CONFIGURATION);
  var SEMESTER = config_sheet.getRange(8, 3, 1, 1).getValue();

  var attendance_method_list = ['A1 1-25% Online', 'A2 26-50% Online', 'A3 51-75% Online', 'A4 76-99% Online', 'S1 1-25% Online', 'S2 26-50% Online', 'S3 51-75% Online', 'S4 76-99% Online'];
  var PART_TERM_CODING_DICTIONARY = { Full: '1', T1: 'H1', T2: 'H2' };

  // splits courses listed as instructional mode Hybrid Synchronous or Hybrid Asynchronous
  // create a section for each face-to-face meeting time 
  for (var courseIdx = scheduledCourseList.length - 1; courseIdx >= 0; courseIdx--) {
    var scheduledCourse = scheduledCourseList[courseIdx];
    if (scheduledCourse.Course.isHybrid()) {
      // remove this course from the list
      scheduledCourseList.splice(courseIdx, 1);
      for (var inPersonDayIdx = 0; inPersonDayIdx < scheduledCourse.CourseTime.days.length; inPersonDayIdx++) {
        var face2faceScheduledCourse = scheduledCourse.cloneMe();
        var pctFace2Face = Math.round((100.0 * (scheduledCourse.CourseTime.days.length - 1)) / scheduledCourse.CourseTime.days.length);
        var attendanceMethod = undefined;
        if (pctFace2Face < 26) {
          attendanceMethod = (scheduledCourse.Course.isSynchronous()) ? attendance_method_list[4] : attendance_method_list[0];
        } else if (pctFace2Face < 51) {
          attendanceMethod = (scheduledCourse.Course.isSynchronous()) ? attendance_method_list[5] : attendance_method_list[1];
        } else if (pctFace2Face < 76) {
          attendanceMethod = (scheduledCourse.Course.isSynchronous()) ? attendance_method_list[6] : attendance_method_list[2];
        } else {
          attendanceMethod = (scheduledCourse.Course.isSynchronous()) ? attendance_method_list[7] : attendance_method_list[3];
        }
        face2faceScheduledCourse.Course.attendance_method = attendanceMethod;
        face2faceScheduledCourse.CourseTime.days = [scheduledCourse.CourseTime.days[inPersonDayIdx]];
        face2faceScheduledCourse.Course.section = face2faceScheduledCourse.Course.section + scheduledCourse.CourseTime.days[inPersonDayIdx];
        scheduledCourseList.splice(courseIdx, 0, face2faceScheduledCourse);
        // if a HYS/HYA course meets one time per week do not create online synchronous meeting times
        var onlineScheduledCourse = scheduledCourse.cloneMe();
        onlineScheduledCourse.CourseTime.days.splice(inPersonDayIdx, 1);
        if (onlineScheduledCourse.CourseTime.days.length > 0) {
          onlineScheduledCourse.Course.section = onlineScheduledCourse.Course.section + scheduledCourse.CourseTime.days[inPersonDayIdx];
          onlineScheduledCourse.Room.building = "ONLINE";
          onlineScheduledCourse.Room.number = "INTERNET";
          onlineScheduledCourse.Course.instructional_mode = (onlineScheduledCourse.Course.isSynchronous()) ? "Online Synchronous" : "Online Asynchronous";
          scheduledCourseList.splice(courseIdx, 0, onlineScheduledCourse);
        }
        //for (var dayIdx = 0 ; dayIdx < scheduledCourse.CourseTime.days.length; dayIdx++) {
        //  if (dayIdx == inPersonDayIdx) {
        //  } else {
        //  }
        //}
      }
    }
  }

  for (var courseIdx = 0; courseIdx < scheduledCourseList.length; courseIdx++) {
    var scheduledRoomWithTimeInterval = scheduledCourseList[courseIdx];
    var roomWithTimeInterval = new RoomWithTimeInterval(scheduledRoomWithTimeInterval.Room, scheduledRoomWithTimeInterval.CourseTime);
    // find the room for this course
    for (var roomIdx = 0; roomIdx < roomList.length; roomIdx++) {
      if (scheduledRoomWithTimeInterval.Room.getId() == roomList[roomIdx].getId()) {
        break;
      }
    }
    var scheduledRoom = (roomIdx >= roomList.length) ? undefined : roomList[roomIdx];
    for (var crossListingsIdx = 0; crossListingsIdx < scheduledRoomWithTimeInterval.Course.numbers.length; crossListingsIdx++) {
      var built_course = [];
      built_course.push("");                                                             // COURSE
      built_course.push(scheduledRoomWithTimeInterval.Course.dept_code);                 // SUBJECT
      var course_number = scheduledRoomWithTimeInterval.Course.numbers[crossListingsIdx];
      built_course.push(course_number); // COURSE_NUMBER
      var section = scheduledRoomWithTimeInterval.Course.section.toString();
      while (section.length < 3) {
        section = '0' + section;
      }
      // set second digit of section to '9' value for courses that start 5:00pm or later ==> 'Evening Courses' designation      
      if (scheduledRoomWithTimeInterval.CourseTime.start.getHours() >= 17 && section.substring(1, 2) == '0') {
        section = section.substring(0, 1) + '9' + section.substring(2);
      }
      // set second digit of section to '8' value for Summer online courses ==> 'Distant Education Courses' only      
      if (SEMESTER == "Summer" && scheduledRoomWithTimeInterval.Course.instructional_mode.indexOf('Online') >= 0) {
        section = section.substring(0, 1) + '8' + section.substring(2);
      }
      built_course.push(section);                                                        // SECTION
      //var instructionalModeAcronym = defaultScheduleRenderer.instructionalModeToString[courseInstructionalMode];
      built_course.push(scheduledRoomWithTimeInterval.Course.instructional_mode);        // INSTR_MODE      
      built_course.push(scheduledRoomWithTimeInterval.Course.hasOwnProperty('attendance_method') ? scheduledRoomWithTimeInterval.Course.attendance_method : "");         // ATTENDANCE METHOD
      var crn = scheduledRoomWithTimeInterval.Course.crn;
      // find course in list of CRNs (crnList)
      for (var crnIdx = 0; crnIdx < crnList.length; crnIdx++) {
        if (crnList[crnIdx].course_number.toString() == course_number && crnList[crnIdx].section == section) {
          crn = crnList[crnIdx].number;
        }
      }
      built_course.push(crn);                       // CRN
      if (scheduledRoomWithTimeInterval.Course.numbers.length > 1) {
        built_course.push(CROSS_LIST_CODE_PREFIX + xlst_id);                             // XLST ID
      } else {
        built_course.push("");                                                           // XLST ID
      }
      built_course.push("");                                                             // BILLHR
      built_course.push(scheduledRoomWithTimeInterval.Course.credit_hours);              // CR
      built_course.push((scheduledRoomWithTimeInterval.CourseTime.days.indexOf('M') < 0) ? " " : "M"); // MO
      built_course.push((scheduledRoomWithTimeInterval.CourseTime.days.indexOf('T') < 0) ? " " : "T"); // TU
      built_course.push((scheduledRoomWithTimeInterval.CourseTime.days.indexOf('W') < 0) ? " " : "W"); // WE
      built_course.push((scheduledRoomWithTimeInterval.CourseTime.days.indexOf('R') < 0) ? " " : "R"); // TH
      built_course.push((scheduledRoomWithTimeInterval.CourseTime.days.indexOf('F') < 0) ? " " : "F"); // FR
      built_course.push((scheduledRoomWithTimeInterval.CourseTime.days.indexOf('S') < 0) ? " " : "S"); // SA
      built_course.push((scheduledRoomWithTimeInterval.CourseTime.days.indexOf('U') < 0) ? " " : "U"); // SU
      var part_term = PART_TERM_CODING_DICTIONARY[scheduledRoomWithTimeInterval.Course.term];
      built_course.push(part_term);                                                      // PART_TERM
      built_course.push(scheduledRoomWithTimeInterval.CourseTime.getTimeString(scheduledRoomWithTimeInterval.CourseTime.start, '24hr'));   // BEGINS
      built_course.push(scheduledRoomWithTimeInterval.CourseTime.getTimeString(scheduledRoomWithTimeInterval.CourseTime.end, '24hr'));     // ENDS
      built_course.push("");                                                             // FIRST_NAME      
      built_course.push((scheduledRoomWithTimeInterval.FacultyCoursesAndPrefs == undefined) ? "TBD" : scheduledRoomWithTimeInterval.FacultyCoursesAndPrefs.name.split(',')[0]);         // INSTRCTR
      //built_course.push(scheduledRoomWithTimeInterval.Room.building);                    // BLDG      
      //built_course.push(scheduledRoomWithTimeInterval.Room.number);                      // RM      
      built_course.push(scheduledRoom == undefined ? undefined : scheduledRoom.building);            // BLDG      
      built_course.push(scheduledRoom == undefined ? undefined : scheduledRoom.number);              // RM      
      built_course.push("");                                                             // MODE
      built_course.push("");                                                             // INSTRUCTIONAL_METHOD
      built_course.push("");                                                             // ATTENDANCE_METHOD
      built_course.push("");                                                             // CAMPUS
      built_course.push("");                                                             // RM_CAPACITY
      built_course.push(scheduledRoom == undefined ? undefined : scheduledRoom.maximum_capacity);    // MAX
      built_course.push("");                                                             // ENR
      built_course.push("");                                                             // PROJ
      built_course.push("");                                                             // PRIOR
      built_course.push("");                                                             // WAIT_CAPACITY
      built_course.push("");                                                             // WAIT_COUNT
      built_course.push("");                                                             // WAIT_AVAIL
      built_course.push("");                                                             // STATUS
      built_course.push("");                                                             // STATUS_DESC
      built_course.push("");                                                             // AVAIL
      built_course.push(0);                                                              // MAX/RM_CAP
      built_course.push(0);                                                              // ENR/RM_CAP
      built_course.push(0);                                                              // ENR/MAX
      built_courses_data.push(built_course);
    }
    if (scheduledRoomWithTimeInterval.Course.numbers.length > 1) {
      xlst_id++;
    }
    //addCourseToSchedule(output_sheet_data, scheduledRoomWithTimeInterval);

    //Logger.log('ScheduleEngine: Scheduled ' + scheduledRoomWithTimeInterval.getId() + " in " + scheduledRoomWithTimeInterval.Room.getId() + 
    // " at " + scheduledRoomWithTimeInterval.CourseTime.getId() + ' with cost ' + scheduledRoomWithTimeInterval.cost); 

    // TODO: WHEN ONLY A PART OF A TIMESLOT IS OCCUPIED THE ENTIRE TIMESLOT IS NOT REMOVED BUT ONLY THE TIMESLOTS OCCUPIED ARE REMOVED
    // remove room with time interval from roomWithTimeIntervalList
    //removeFromList(scheduledRoomWithTimeInterval, roomWithTimeInterval, roomWithTimeIntervalList, RoomWithTimeInterval.prototype.isEqualRoomAndTimeInterval);

    //cost += scheduledRoomWithTimeInterval.cost;
    // add to scheduled course list
    //output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_COST_OUTPUT, OUTPUT_SHEET_INDEX_COL_COST_OUTPUT).setValue(cost.toFixed(2));
    //output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_NUMCOURSES_OUTPUT, OUTPUT_SHEET_INDEX_COL_NUMCOURSES_OUTPUT).setValue(courseIdx+1);
    //Logger.log('ScheduleEngine: Scheduled ' + scheduledCourseList.length + " courses.");
  }
  var built_courses_range = export_coursebuilding_sheet.getRange(2, 1, built_courses_data.length, built_courses_data[0].length);
  built_courses_range.setValues(built_courses_data);
}

function showLogs() {
}

// Course class constructor and supporting functions
function Course(dept_code, numbers, section, crn, term, credit_hours, priority, expected_enrollment, not_simultaneous_courses, simultaneous_courses, required_rooms, excluded_rooms, instructional_mode) {
  this.dept_code = dept_code;
  this.numbers = numbers.toString().split("/");
  this.section = section;
  this.crn = crn;
  this.term = term;
  this.credit_hours = credit_hours;
  this.priority = priority;
  this.expected_enrollment = expected_enrollment;
  this.not_simultaneous_courses = not_simultaneous_courses;
  this.simultaneous_courses = simultaneous_courses;
  this.required_rooms = required_rooms;
  this.excluded_rooms = excluded_rooms;
  this.instructional_mode = instructional_mode;
}

// Course::getId() function
Course.prototype.getId = function () {
  //return this.dept_code + " " + this.number + "-" + FormatNumberLength(this.section, 2);
  var term_str = (this.isFullTerm()) ? "" : ("-" + this.term);
  return this.dept_code + " " + this.numbers.join("/") + "-" + this.section + term_str;
  //return this.dept_code + " " + this.numbers.join("/") + "-" + this.section;
}

var COURSE_NUMBER_SPECIAL_TOPICS_UNDERGRAD = 4090;
var COURSE_NUMBER_SPECIAL_TOPICS_GRAD = 6090;

// Course::equals() function
Course.prototype.equals = function (other_course) {
  return this.sameNumber(other_course) && this.sameSection(other_course) && this.term == other_course.term;
}

// Course::sameNumber() function
Course.prototype.sameNumber = function (other_course) {
  if (this.numbers[0] == COURSE_NUMBER_SPECIAL_TOPICS_UNDERGRAD ||
    this.numbers[0] == COURSE_NUMBER_SPECIAL_TOPICS_GRAD) {
    return this.dept_code == other_course.dept_code && this.numbers[0] == other_course.numbers[0] && this.section == other_course.section;
  } else {
    return this.dept_code == other_course.dept_code && this.numbers[0] == other_course.numbers[0];
  }
}

// Course::sameSection() function
Course.prototype.sameSection = function (other_course) {
  // remove leading zeros so sections '1','01', and '001 are equal
  //if (this.section == '001') {
  //  var aaa=1;
  //}
  var thisSection = (typeof this.section == 'string' || this.section instanceof String) ? this.section.replace(/^0+/, '') : this.section;
  var otherSection = (typeof other_course.section == 'string' || other_course.section instanceof String) ? other_course.section.replace(/^0+/, '') : other_course.section;
  return thisSection == otherSection;
  //return this.section == other_course.section;
}

// boolean to detect special topics courses
Course.prototype.isSpecialTopics = function () {
  return /[3,4,5,6,8]090/.test(this.getId());
  //return this.getId().indexOf('090') >= 0;
}

// boolean to detect courses with hybrid instructional mode
Course.prototype.isHybrid = function () {
  return this.instructional_mode.indexOf("Hybrid") > -1;
}

// boolean to detect courses with synchronous instructional mode
Course.prototype.isSynchronous = function () {
  return this.instructional_mode.indexOf("Synchronous") > -1;
}

// boolean to detect courses with asynchronous instructional mode
Course.prototype.isAsynchronous = function () {
  return this.instructional_mode.indexOf("Asynchronous") > -1;
}

// boolean to detect full-term and part-term courses
Course.prototype.isFullTerm = function () {
  return this.term == "Full";
}

// Course::notSimultaneousConflict()
// detects courses that conflict, i.e., course pairs that are preferred to NOT be taught simultaneously
Course.prototype.notSimultaneousConflict = function (course_to_check) {
  var conflict = false;
  for (var notSimultaneousIdx = 0; conflict == false && notSimultaneousIdx < this.not_simultaneous_courses.length; notSimultaneousIdx++) {
    // do a regex pattern match here. For example, 3159's conflicts are 3???, 4???? indicating any 3XXX, 4XXX course is a conflict
    if (this.not_simultaneous_courses[notSimultaneousIdx].numbers[0].indexOf('?') >= 0) {
      var regexp_str = this.not_simultaneous_courses[notSimultaneousIdx].numbers[0].replace(/\?/g, '[0-9]?');
      var regexp = new RegExp(regexp_str);
      //var strt = course_to_check.numbers[0].toString();
      conflict = regexp.test(course_to_check.numbers[0]);
    } else {
      conflict = this.not_simultaneous_courses[notSimultaneousIdx].sameNumber(course_to_check);
    }
  }
  return conflict;
}

// Course::simultaneousPreference()
// detects course pairs that are preferred to be taught simultaneously
Course.prototype.simultaneousPreference = function (course_to_check) {
  var simultaneous = false;
  for (var simultaneousIdx = 0; simultaneous == false && simultaneousIdx < this.simultaneous_courses.length; simultaneousIdx++) {
    simultaneous = this.simultaneous_courses[simultaneousIdx].sameNumber(course_to_check);
  }
  return simultaneous;
}

// Course::clone implementation
Course.prototype.cloneMe = function () {
  return new Course(this.dept_code,
    this.numbers.join("/"),
    this.section,
    this.crn,
    this.term,
    this.credit_hours,
    this.priority,
    this.expected_enrollment,
    this.not_simultaneous_courses == undefined ? this.not_simultaneous_courses : this.not_simultaneous_courses.slice(),
    this.simultaneous_courses == undefined ? this.simultaneous_courses : this.simultaneous_courses.slice(),
    this.required_rooms == undefined ? this.required_rooms : this.required_rooms.slice(),
    this.excluded_rooms == undefined ? this.excluded_rooms : this.excluded_rooms.slice(),
    this.instructional_mode);
}

// Room class constructor and supporting functions
function Room(building, number, maximum_capacity, cost) {
  this.building = building;
  this.number = number;
  this.maximum_capacity = maximum_capacity;
  this.cost = cost;
}

// Room::getId() function
Room.prototype.getId = function () {
  return this.building + " " + this.number;
}

// Room::equals() function
Room.prototype.equals = function (room_to_check) {
  return this.building == room_to_check.building && this.number == room_to_check.number;
}

// Room::clone implementation
Room.prototype.cloneMe = function () {
  return new Room(this.building, this.number, this.maximum_capacity, this.cost);
}

// CourseTime class constructor and supporting functions
function CourseTime(start, end, days, credit_hours_per_day, exclude_course_list, include_course_list, cost) {
  if (start instanceof Date) {
    this.start = start;
  } else {
    // convert to Date
    this.start = this.timeStringToDate(start);
  }
  if (end instanceof Date) {
    this.end = end;
  } else {
    // convert to Date
    this.end = this.timeStringToDate(end);
  }
  this.days = days.split('/');
  this.credit_hours_per_day = credit_hours_per_day;
  this.exclude_course_list = exclude_course_list;
  this.include_course_list = include_course_list;
  this.term = "Full";  // all times are assumed to be viable for the Full term, they are split when a part-term course is scheduled
  this.cost = cost;
}

// CourseTime::getId() function
CourseTime.prototype.getId = function () {
  var start_time_str = this.getTimeString(this.start);
  var end_time_str = this.getTimeString(this.end);
  var days_str = this.days.join("/");
  var term_str = (this.isFullTerm()) ? "" : ("-" + this.term);
  return start_time_str + '-' + end_time_str + ' ' + days_str + term_str;
}

CourseTime.prototype.durationMinutes = function () {
  return (this.end - this.start) / (1000 * 60);
}

// CourseTime::equals() function
CourseTime.prototype.equals = function (other_course_time) {
  // exact equality of start and end 
  return this.start - other_course_time.start === 0 && this.end - other_course_time.end === 0 && this.term == other_course_time.term;
}

// boolean to detect full-term and part-term courses
CourseTime.prototype.isFullTerm = function () {
  return this.term == "Full";
}

CourseTime.prototype.overlaps = function (other_course_time) {
  // We do not use <= comparators here to make a time interval ending at 8:00 AM NOT overlap with a time interval starting at 8:00 AM
  return this.start < other_course_time.end && other_course_time.start < this.end && (this.isFullTerm() || this.term == other_course_time.term);
  //return Math.max(this.start, other_course_time.start) < Math.min(this.end, other_course_time.end);
  //if (this.start < other_course_time.start && other_course_time.start < this.end) return true; // b starts in a
  //if (this.start < other_course_time.end   && other_course_time.end   < this.end) return true; // b ends in a
  //if (other_course_time.start < this.start && this.end  <  other_course_time.end) return true; // a in b
  //return false; 
}

// CourseTime::getTimeString
// Creates a time string for a time interval
CourseTime.prototype.getTimeString = function (date, format) {
  var hours = date.getHours();
  var minutes = date.getMinutes();
  var strTime = '';
  if (format != undefined && format == '24hr') {
    hours = hours < 10 ? '0' + hours : hours;
    minutes = minutes < 10 ? '0' + minutes : minutes;
    strTime = hours.toString() + minutes.toString();
  } else {
    var ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours ? hours : 12; // the hour '0' should be '12'
    minutes = minutes < 10 ? '0' + minutes : minutes;
    strTime = hours + ':' + minutes + ' ' + ampm;
  }
  return strTime;
}

// HARD-CODED CONSTANTS
var DATE_CONSTANT_YEAR = 2020;
var DATE_CONSTANT_MONTH = 01;
var DATE_CONSTANT_DAY = 01;
var DATE_CONSTANT_FIRST_HOUR = 8;
var DATE_CONSTANT_FIRST_MINUTE = 00;
var DATE_CONSTANT_SEC = 00;
var DATE_CONSTANT_MSEC = 00;

// CourseTime::getTimeString
// Creates a time string for a time interval
CourseTime.prototype.timeStringToDate = function (timeString, format) {
  timeString = timeString.toString();
  // replace '/' with '-'
  timeString = timeString.replace(/\//ig, '-');
  // delete '.'
  timeString = timeString.replace(/\./g, '');

  if (timeString.length < 1) {
    return new Date(DATE_CONSTANT_YEAR, DATE_CONSTANT_MONTH, DATE_CONSTANT_DAY, 0, 0, DATE_CONSTANT_SEC, DATE_CONSTANT_MSEC);
  }
  var H, i, AMorPM;
  if (format != undefined && format == '24hr') {
    var hrmin = /([\d]{4}|[\d]{3})/.exec(timeString)[1];
    H = hrmin.substring(0, hrmin.length - 2);
    i = hrmin.substring(hrmin.length - 2, hrmin.length);
  } else {
    /**
    * extracting the year, month, day, hours and minutes.
    * the month, day and hours can be 1 or 2 digits(the leading zero is optional).
    * i.e: '4/3/2022 2:18 AM' is the same as '04/03/2022 02:18 AM' => Notice the absence of the leading zero.
    **/
    //var y = /\-([\d]{4})/.exec(timeString)[1],
    //m = /\-([\d]{2}|[\d])/.exec(timeString)[1],
    //d = /([\d]{2}|[\d])\-/.exec(timeString)[1],
    //H = /\s([\d]{2}|[\d]):/.exec(timeString)[1],
    H = parseInt(/([\d]{2}|[\d]):/.exec(timeString)[1], 10),
      i = parseInt(/:([\d]{2})/.exec(timeString)[1], 10),
      AMorPM = /(AM|PM|am|pm)/.exec(timeString)[1];
    if ((AMorPM == "PM" || AMorPM == "pm") && H != 12) {
      H += 12;
    }
  }
  // return a Date instance.
  //return new Date(y,m,d,H,i,00,00);
  return new Date(DATE_CONSTANT_YEAR, DATE_CONSTANT_MONTH, DATE_CONSTANT_DAY, H, i, DATE_CONSTANT_SEC, DATE_CONSTANT_MSEC);
}

// CourseTime::daysInCommon
// Input - Another CourseTime object
// Output - The intersection of the days offered arrays for the this object and the other_course_time object
CourseTime.prototype.daysInCommon = function (other_course_time) {
  var daysInCommon = [];
  for (var dayAIdx = 0; dayAIdx < this.days.length; dayAIdx++) {
    for (var dayBIdx = 0; dayBIdx < other_course_time.days.length; dayBIdx++) {
      if (this.days[dayAIdx] == other_course_time.days[dayBIdx]) {
        daysInCommon.push(this.days[dayAIdx]);
      }
    }
  }
  return daysInCommon;
}

CourseTime.prototype.hoursBetween = function (other_course_time) {
  if (this.overlaps(other_course_time)) {
    return 0;
  } else {
    if (this.term != other_course_time.term) {
      return 24 * 40; // approximate length in hours of the interval between part-term start times
    }
    var hrInterval = Math.max(this.start - other_course_time.end, other_course_time.start - this.end) / (1000 * 60 * 60);
    return hrInterval;
  }
}

// CourseTime::hasTimeIntervalAndDaysConflictWith
// detects conflicting times using time interval and days criteria
//CourseTime.prototype.hasTimeIntervalAndDaysConflictWith = function(other_course_time) {
//  return (this.overlaps(other_course_time) && this.daysInCommon(other_course_time).length > 0);
//}

// CourseTime::clone implementation
CourseTime.prototype.cloneMe = function () {
  return new CourseTime(this.start,
    this.end,
    this.days.join('/'),
    this.credit_hours_per_day,
    this.exclude_course_list == undefined ? this.exclude_course_list : this.exclude_course_list.slice(),
    this.include_course_list == undefined ? this.include_course_list : this.include_course_list.slice(),
    this.cost);
}

// RoomWithTimeInterval class constructor and supporting functions
function RoomWithTimeInterval(room, course_time) {
  this.Room = room;
  this.CourseTime = course_time;
  this.cost = room.cost + course_time.cost;
}

RoomWithTimeInterval.prototype.resetCost = function () {
  this.cost = this.Room.cost + this.CourseTime.cost;
}

RoomWithTimeInterval.prototype.getId = function () {
  return this.Room.getId() + ' at ' + this.CourseTime.getId();
}

function hasOverlappingRoomTimes(room_with_time_intervalA, room_with_time_intervalB) {
  // no code reuse from hasTimeIntervalAndDaysConflictWith should be fixed in externalized function
  var daysInCommon = [];
  for (var dayAIdx = 0; dayAIdx < room_with_time_intervalA.CourseTime.days.length; dayAIdx++) {
    for (var dayBIdx = 0; dayBIdx < room_with_time_intervalB.CourseTime.days.length; dayBIdx++) {
      if (room_with_time_intervalA.CourseTime.days[dayAIdx] == room_with_time_intervalB.CourseTime.days[dayBIdx]) {
        daysInCommon.push(room_with_time_intervalA.CourseTime.days[dayAIdx]);
      }
    }
  }
  return room_with_time_intervalA.Room.equals(room_with_time_intervalB.Room) &&
    room_with_time_intervalA.CourseTime.equals(room_with_time_intervalB.CourseTime) &&
    daysInCommon.length > 0;
}

RoomWithTimeInterval.prototype.isEqualRoomAndTimeInterval = function (room_with_time_intervalA, room_with_time_intervalB) {
  return room_with_time_intervalA.Room.equals(room_with_time_intervalB.Room) &&
    room_with_time_intervalA.CourseTime.equals(room_with_time_intervalB.CourseTime);
}

// RoomWithTimeInterval::clone implementation
RoomWithTimeInterval.prototype.cloneMe = function () {
  return new RoomWidthTimeInterval(this.Room.cloneMe(), this.CourseTime.cloneMe());
}

// FacultyCoursesAndPrefs class constructor and supporting functions
function FacultyCoursesAndPrefs(name, email, courseList, timeIntervalCostMap, courses_on_same_days, hours_between_courses) {
  this.name = name;
  this.email = email;
  this.courseList = courseList;
  this.timeIntervalCostMap = timeIntervalCostMap;
  this.courses_on_same_days = courses_on_same_days;
  this.hours_between_courses = hours_between_courses;
}

// FacultyCoursesAndPrefs::getId() function
FacultyCoursesAndPrefs.prototype.getId = function () {
  return this.name;
}

// FacultyCoursesAndPrefs::clone implementation
FacultyCoursesAndPrefs.prototype.cloneMe = function () {
  var copyOfTimeIntervalCostMap = {};
  Object.assign(copyOfTimeIntervalCostMap, this.timeIntervalCostMap);
  return new FacultyCoursesAndPrefs(this.name,
    this.email,
    this.courseList.slice(),
    copyOfTimeIntervalCostMap,
    this.courses_on_same_days,
    this.hours_between_courses);
}

// ScheduledCourse class constructor and supporting functions
function ScheduledCourse(course, room, course_time, cost, instructor, costArray) {
  this.Course = course;
  this.Room = room;
  this.CourseTime = course_time;
  this.cost = cost;
  this.FacultyCoursesAndPrefs = instructor;
  this.costArray = costArray;
}

// ScheduledCourse::getId() function
ScheduledCourse.prototype.getId = function () {
  var facultyName = (this.FacultyCoursesAndPrefs != undefined) ? this.FacultyCoursesAndPrefs.name : undefined;
  return this.Course.getId() + ' in ' + this.Room.getId() + ' at ' + this.CourseTime.getId() + ' instructed by ' + facultyName;
}

// ScheduledCourse::clone implementation
ScheduledCourse.prototype.cloneMe = function () {
  var copyOfCostArray = {};
  Object.assign(copyOfCostArray, this.costArray);
  return new ScheduledCourse(this.Course.cloneMe(),
    this.Room.cloneMe(),
    this.CourseTime.cloneMe(),
    this.cost,
    this.FacultyCoursesAndPrefs == undefined ? undefined : this.FacultyCoursesAndPrefs.cloneMe(),
    copyOfCostArray);
}

function CRN(number, dept_code, course_number, section) {
  this.number = number;
  this.dept_code = dept_code;
  this.course_number = course_number;
  this.section = section;
}

// CRN::clone implementation
CRN.prototype.cloneMe = function () {
  return new CRN(this.number,
    this.dept_code,
    this.course_number,
    this.section);
}


function computeNewSchedule() {
  var schedule = SpreadsheetApp.getActiveSpreadsheet();
  var config_sheet = schedule.getSheetByName(SHEET_NAME_CONFIGURATION);
  var SCHEDULE_DISTRIBUTED = config_sheet.getRange(2, 1, 1, 1).getValue();
  if (SCHEDULE_DISTRIBUTED) {
    SpreadsheetApp.getUi().alert('This function is disabled since the final schedule has been distributed to faculty.');
    return;
  }
  RunSchedulingEngine(false);
}

function revisePriorSchedule() {
  var schedule = SpreadsheetApp.getActiveSpreadsheet();
  var config_sheet = schedule.getSheetByName(SHEET_NAME_CONFIGURATION);
  var SCHEDULE_BUILT = config_sheet.getRange(3, 1, 1, 1).getValue();
  if (SCHEDULE_BUILT) {
    SpreadsheetApp.getUi().alert('This function is disabled since the final revised schedule has been exported for entry to the registrar.');
    return;
  }
  RunSchedulingEngine(true);
}

function RunSchedulingEngine(usePriorSchedule) {
  var schedule = SpreadsheetApp.getActiveSpreadsheet();
  var config_sheet = schedule.getSheetByName(SHEET_NAME_CONFIGURATION);
  var course_constraint_sheet = schedule.getSheetByName(SHEET_NAME_COURSE_CONSTRAINTS);
  var time_slot_sheet = schedule.getSheetByName(SHEET_NAME_TIME_SLOTS);
  var faculty_sheet = schedule.getSheetByName(SHEET_NAME_FACULTY_COURSES_AND_PREFERENCES);
  var room_slot_sheet = schedule.getSheetByName(SHEET_NAME_ROOM_SLOTS);
  var course_and_time_constraint_sheet = schedule.getSheetByName(SHEET_NAME_COURSE_AND_TIME_CONSTRAINTS);
  var template_output_sheet = schedule.getSheetByName(SHEET_NAME_OUTPUT_TEMPLATE);
  var output_sheet = schedule.getSheetByName(SHEET_NAME_OUTPUT_SHEET);
  var output_sheet_test = schedule.getSheetByName(SHEET_NAME_OUTPUT_SHEET_TEST);
  var prior_schedule_sheet = schedule.getSheetByName(SHEET_NAME_PRIOR_SCHEDULE_SHEET);

  if (!schedule || !course_constraint_sheet || !room_slot_sheet || !course_and_time_constraint_sheet
    || !template_output_sheet || !output_sheet || !faculty_sheet || !prior_schedule_sheet) {
    SpreadsheetApp.getUi().alert('Could not read sheet data.');
    throw Error("Exiting due to sheet data access error.");
  }
  // INPUTS
  // List of courses constraints: 
  //    course to schedule, list of courses that CANNOT be scheduled at simultaneous time slots, list of courses that CAN be scheduled at simultaneous time slots
  // List of hard-set courses and times
  //    course, time (no room?)
  // List of room availability
  //    room numbers, max seats
  // List of time slots
  //    times and days

  // Set up the destination sheet for the schedule renderer
  var config_sheet = schedule.getSheetByName(SHEET_NAME_CONFIGURATION);
  var SEMESTER = config_sheet.getRange(8, 3, 1, 1).getValue();
  defaultScheduleRenderer = new ScheduleRenderer(SEMESTER);
  defaultScheduleRenderer.setDestinationSheet(output_sheet_test);

  // Copy template sheet to output sheet  
  var template_data_range = template_output_sheet.getDataRange().getDataRegion();
  var template_data_values = template_data_range.getValues();
  template_data_range.copyFormatToRange(output_sheet, template_data_range.getColumn(),
    template_data_range.getColumn() + template_data_range.getWidth(),
    template_data_range.getRow(), template_data_range.getRow() + template_data_range.getHeight());
  template_data_range.copyValuesToRange(output_sheet, template_data_range.getColumn(),
    template_data_range.getColumn() + template_data_range.getWidth(),
    template_data_range.getRow(), template_data_range.getRow() + template_data_range.getHeight());

  var output_sheet_data = output_sheet.getDataRange();

  // Load Course data 
  var course_data = course_constraint_sheet.getDataRange();
  var unscheduledCourseList = getCoursesToSchedule(course_data);

  // Load Room data 
  var room_data = room_slot_sheet.getDataRange();
  var roomList = getRoomsToSchedule(room_data);
  var roomsToRender = roomList.filter(function (room) {
    return true;
  });
  defaultScheduleRenderer.setRoomsToRender(roomsToRender);

  // Load Time Interval data
  var time_interval_datarange = time_slot_sheet.getDataRange();
  var timeIntervalList = getTimeIntervalsToSchedule(time_interval_datarange);

  var timesToRender = timeIntervalList.filter(function (timeSlot) {
    if (SEMESTER == "Summer") {
      // FOR SUMMER SCHEDULING
      // render all time slots 2.0 credit hours or less (100 mins or less)
      return timeSlot.credit_hours_per_day <= 2.0;
    } else {
      // FOR FALL and SPRING SCHEDULING
      // render all time slots 1.5 credit hours or less (75 mins or less)
      return timeSlot.credit_hours_per_day <= 1.5;
    }
  });
  defaultScheduleRenderer.setTimesToRender(timesToRender);

  // Load Faculty Course Assignments and Time Interval Preference data
  var faculty_datarange = faculty_sheet.getDataRange();
  var facultyCoursesAndPrefsList = getFacultyCoursesAndPreferencesToSchedule(faculty_datarange, timeIntervalList);

  // Load Prior Schedule data data
  //var prior_schedule_datarange = prior_schedule_sheet.getDataRange();
  //var priorScheduledCourseList = getPriorScheduledCourses(prior_schedule_datarange, timeIntervalList);

  // Sort all lists by local costs to speed up 1-dimensional searches
  roomList.sort(function (a, b) { return a.cost - b.cost; });

  // Create sorted list of rooms with times and their local costs to speed up 1-dimensional searches
  var roomWithTimeIntervalList = [];
  for (var roomIdx = 0; roomIdx < roomList.length; roomIdx++) {
    for (var timeIdx = 0; timeIdx < timeIntervalList.length; timeIdx++) {
      var copyRoom = roomList[roomIdx].cloneMe();
      var copyCourseTime = timeIntervalList[timeIdx].cloneMe();
      roomWithTimeIntervalList.push(new RoomWithTimeInterval(copyRoom, copyCourseTime));
    }
  }
  roomWithTimeIntervalList.sort(function (a, b) { return a.cost - b.cost; });

  // Track global cost and current list of scheduled courses
  var cost = 0;
  var scheduledCourseList = [];

  // DEBUGGING CODE
  //var roomTestA = new Room("EPIC", "SD1");
  //var courseTimeTestA = new CourseTime("7:00 PM","8:15 PM", "M/W/F", [], [], 0);
  //var roomTestB = new Room("EPIC", "SD1");
  //var courseTimeTestB = new CourseTime("8:30 PM","9:45 PM", "T/R", [], [], 0);
  //if (roomTestA.building == roomTestB.building && roomTestA.number == roomTestB.number && courseTimeTestA.hasTimeIntervalConflictWith(courseTimeTestB)) {
  //  var aaa=true;
  //}

  // Populate the schedule using the information in the "Prior Schedule" tab
  if (usePriorSchedule == true) {
    // create a list of the scheduled courses from the "Prior Schedule" sheet
    transferScheduledCourses('READ', prior_schedule_sheet, scheduledCourseList, facultyCoursesAndPrefsList);
    for (var courseIdx = 0; courseIdx < scheduledCourseList.length; courseIdx++) {
      var scheduledRoomWithTimeInterval = scheduledCourseList[courseIdx];
      var roomWithTimeInterval = new RoomWithTimeInterval(scheduledRoomWithTimeInterval.Room, scheduledRoomWithTimeInterval.CourseTime);

      // render course to schedule
      addCourseToSchedule(output_sheet_data, scheduledRoomWithTimeInterval);
      defaultScheduleRenderer.addCourseToSchedule(scheduledRoomWithTimeInterval);

      Logger.log('ScheduleEngine: Scheduled ' + scheduledRoomWithTimeInterval.getId() + " in " + scheduledRoomWithTimeInterval.Room.getId() +
        " at " + scheduledRoomWithTimeInterval.CourseTime.getId() + ' with cost ' + scheduledRoomWithTimeInterval.cost);

      // TODO: WHEN ONLY A PART OF A TIMESLOT IS OCCUPIED THE ENTIRE TIMESLOT IS NOT REMOVED BUT ONLY THE TIMESLOTS OCCUPIED ARE REMOVED
      // remove room with time interval from roomWithTimeIntervalList
      removeFromList(scheduledRoomWithTimeInterval, roomWithTimeInterval, roomWithTimeIntervalList, RoomWithTimeInterval.prototype.isEqualRoomAndTimeInterval);

      cost += scheduledRoomWithTimeInterval.cost;
      // add to scheduled course list
      output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_COST_OUTPUT, OUTPUT_SHEET_INDEX_COL_COST_OUTPUT).setValue(cost.toFixed(2));
      output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_NUMCOURSES_OUTPUT, OUTPUT_SHEET_INDEX_COL_NUMCOURSES_OUTPUT).setValue(courseIdx + 1);
      Logger.log('ScheduleEngine: Scheduled ' + scheduledCourseList.length + " courses.");
    }
    return;
    cost += insertPreScheduledCourses(course_and_time_constraint_sheet, output_sheet_data, roomWithTimeIntervalList, facultyCoursesAndPrefsList, scheduledCourseList);
  }

  // Start a new schedule and insert pre-scheduled courses into the new schedule
  cost += insertPreScheduledCourses(course_and_time_constraint_sheet, output_sheet_data, roomWithTimeIntervalList, facultyCoursesAndPrefsList, scheduledCourseList);

  // Start of the scheduling engine code
  // Greedy Optimization (Proof of Concept)
  //      --> select a course by priority
  //      --> find a room and time-slot cost by lowest total cost
  //      --> schedule course for the best (lowest total cost) room and time-slot
  //

  // A better greedy optimization strategy:
  // combinatorically look at all possible triplets of (course, room, time interval) to find best 1-move cost
  //
  var prioritySortedCourseList = unscheduledCourseList;
  prioritySortedCourseList.sort(function (a, b) { return a.priority - b.priority; });

  // schedule all courses until there is a scheduling error or all courses have been scheduled
  while (prioritySortedCourseList.length > 0) {

    // get the next course from the prioritySortedCourseList
    var nextCourse = prioritySortedCourseList[0];


    Logger.log('ScheduleEngine: Scheduling ' + nextCourse.getId());

    // find the best room and time interval from the roomWithTimeIntervalList to offer nextCourse given the scheduledCourseList and the facultyCoursesAndPrefsList
    // and return the scheduledRoomWithTimeInterval result 
    var scheduledRoomWithTimeInterval = findRoomWithTimeInterval(roomWithTimeIntervalList, nextCourse, scheduledCourseList, facultyCoursesAndPrefsList);

    // ERROR: This course cannot be scheduled without conflicts
    if (scheduledRoomWithTimeInterval == undefined) {
      SpreadsheetApp.getUi().alert('It is not possible to schedule ' + nextCourse.getId() + ' without conflicts.');
      throw Error("Exiting due to scheduling error.");
    }

    cost += scheduledRoomWithTimeInterval.cost;

    Logger.log('ScheduleEngine: Scheduled ' + nextCourse.getId() + " in " + scheduledRoomWithTimeInterval.Room.getId() +
      " at " + scheduledRoomWithTimeInterval.CourseTime.getId() + ' with cost ' + scheduledRoomWithTimeInterval.cost);
    Logger.log('ScheduleEngine: Deleting course+time slot ' + roomWithTimeIntervalList[scheduledRoomWithTimeInterval.index].getId());

    // DEBUG
    //if (nextCourse.equals(new Course(DEFAULT_DEPT_CODE, 3121, 3))) {
    //  var test = 1;
    //}

    // remove the scheduledRoomWithTimeInterval from the roomWithTimeIntervalList (the index to remove is set in the findRoomWithTimeInterval() function)

    // When the building is 'NONE' or 'ONLINE', e.g., 'NONE TBA' or 'ONLINE INTERNET' we do not modify the room availability list (roomWithTimeIntervalList) 
    // as this room is virtual and has an inexaustible supply of space.
    if (scheduledRoomWithTimeInterval.Room.building != 'NONE' || scheduledRoomWithTimeInterval.Room.building != 'ONLINE') {
      // for *REAL* classrooms, e.g., those having BUILDING != 'NONE' and BUILDING != 'ONLINE' we must remove the space resource and track remaining space
      roomWithTimeIntervalList.splice(scheduledRoomWithTimeInterval.index, 1);

      // see if this was a part-term course going into a full-term course time. If so, split the course-time and add the available part-term time to the roomWithTimeInterval list.
      if (!nextCourse.isFullTerm() && scheduledRoomWithTimeInterval.CourseTime.isFullTerm()) {
        Logger.log('Split part-term course ' + nextCourse.getId() + '.');
        var oRoom = scheduledRoomWithTimeInterval.Room.cloneMe();
        var oCourseTime = scheduledRoomWithTimeInterval.CourseTime.cloneMe();
        oCourseTime.term = (nextCourse.term == 'T1') ? 'T2' : 'T1';
        var excessRoomWithTimeInterval = new RoomWithTimeInterval(oRoom, oCourseTime);
        Logger.log('ScheduleEngine: Adding unused part-term course+time slot ' + excessRoomWithTimeInterval.getId());
        roomWithTimeIntervalList.push(excessRoomWithTimeInterval);
      }

      // if only part of the times available for the scheduledRoomWithTimeInterval were used, re-insert the remaining time-slots for potential use to schedule other courses
      var scheduledRoomWithTimeIntervalTotalCreditHours = scheduledRoomWithTimeInterval.CourseTime.credit_hours_per_day * scheduledRoomWithTimeInterval.CourseTime.days.length;
      // If scheduledRoomWithTimeIntervalTotalCreditHours > 3 the code here will not work correctly - breaks for summer scheduling
      // TODO: FIX CREDIT HOUR SPLITS OF TIMESLOTS FOR SUMMER
      if (scheduledRoomWithTimeIntervalTotalCreditHours <= 3.0) {
        var excessDays = (scheduledRoomWithTimeIntervalTotalCreditHours - nextCourse.credit_hours) / scheduledRoomWithTimeInterval.CourseTime.credit_hours_per_day;
        if (excessDays >= 1) {
          var oRoom = scheduledRoomWithTimeInterval.Room;
          var copyRoom = new Room(oRoom.building, oRoom.number, oRoom.maximum_capacity, oRoom.cost);
          // NOTE: We have to clone this CourseTime object since many roomWithTimeInterval objects in the roomWithTimeIntervalList hold references to a single CourseTime object
          // Hence we create (2) new CourseTime objects due to the split in days; neither of which can alter the referenced original un-split CourseTime object
          // copy (1) goes back into the scheduledRoomWithTimeInterval object copy (2) is part of the excessRoomWithTimeInterval object
          var oCourseTime = scheduledRoomWithTimeInterval.CourseTime;
          var copyCourseTime = new CourseTime(oCourseTime.start, oCourseTime.end, oCourseTime.days.join("/"), oCourseTime.credit_hours_per_day,
            oCourseTime.exclude_course_list, oCourseTime.include_course_list, oCourseTime.cost);
          scheduledRoomWithTimeInterval.CourseTime = copyCourseTime; // copy (1) goes back into scheduledRoomWithTimeInterval
          var excessCourseTimeIntervals = splitCourseTimeByDay(scheduledRoomWithTimeInterval.CourseTime, excessDays); // creates copy (2) and returns it and modifies the days of copy (1) accordingly
          var excessRoomWithTimeInterval = new RoomWithTimeInterval(copyRoom, excessCourseTimeIntervals);
          Logger.log('ScheduleEngine: Adding unused course+time slot ' + excessRoomWithTimeInterval.getId());
          roomWithTimeIntervalList.push(excessRoomWithTimeInterval);
        }
      }
      // partially used course times have been split and remaining excessRoomWithTimeInterval has been re-inserted to the available roomWithTimeIntervalList
      // Returns an array of time intervals to enter for the single course into the standard time interval schedule
      Logger.log('ScheduleEngine: Deleted index ' + scheduledRoomWithTimeInterval.index +
        ' from available course+time slot list. There are ' + roomWithTimeIntervalList.length + ' slots available.');
    } else {
      Logger.log('ScheduleEngine: Scheduled index ' + scheduledRoomWithTimeInterval.index +
        ' into \'NONE\' building course+time slot list. The number of slots available is unchanged at ' + roomWithTimeIntervalList.length + '.');
    }

    // add to scheduled course list
    var newScheduledCourse = new ScheduledCourse(nextCourse, scheduledRoomWithTimeInterval.Room, scheduledRoomWithTimeInterval.CourseTime,
      scheduledRoomWithTimeInterval.cost, scheduledRoomWithTimeInterval.instructor, scheduledRoomWithTimeInterval.costArr);
    scheduledCourseList.push(newScheduledCourse);
    addCourseToSchedule(output_sheet_data, newScheduledCourse);
    defaultScheduleRenderer.addCourseToSchedule(newScheduledCourse);

    // remove course from prioritySortedCourseList
    prioritySortedCourseList.splice(prioritySortedCourseList.indexOf(nextCourse), 1);

    // remove rooms with overlapping time intervals from roomWithTimeIntervalList
    //removeFromList(availableRoomWithTimeInterval, roomWithTimeIntervalList, hasOverlappingRoomTimes);

    output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_COST_OUTPUT, OUTPUT_SHEET_INDEX_COL_COST_OUTPUT).setValue(cost.toFixed(2));
    output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_NUMCOURSES_OUTPUT, OUTPUT_SHEET_INDEX_COL_NUMCOURSES_OUTPUT).setValue(scheduledCourseList.length);
    Logger.log('ScheduleEngine: Scheduled ' + scheduledCourseList.length + " courses " + prioritySortedCourseList.length + " courses remain to be scheduled.");
  }  // scheduling complete 

  scheduledCourseList.sort(function (a, b) { return a.Course.numbers[0] - b.Course.numbers[0]; });
  transferScheduledCourses('WRITE', prior_schedule_sheet, scheduledCourseList);
} // exit engine

/* CONVERT TABS TO CLASSES
function PriorScheduleTab() {
  this.COLUMN_INDEX_ADD_TO_SCHEDULE = 0;
  this.COLUMN_INDEX_DEPARTMENT_CODE = 1;
  this.COLUMN_INDEX_COURSE_NUMBERS = 2;
  this.COLUMN_INDEX_SECTION = 3;
  this.COLUMN_INDEX_INSTRUCTIONAL_MODE = 4;
  this.COLUMN_INDEX_CRN_NUMBER = 5;
  this.COLUMN_INDEX_START_TIME = 6;
  this.COLUMN_INDEX_END_TIME = 7;
  this.COLUMN_INDEX_CREDIT_HOURS_PER_DAY = 8;
  this.COLUMN_INDEX_DAYS_OF_WEEK = 9;
  this.COLUMN_INDEX_BUILDING = 10;
  this.COLUMN_INDEX_ROOM = 11;
  this.COLUMN_INDEX_INSTRUCTOR = 12;
  this.COLUMN_INDEX_INSTRUCTOR_EMAIL = 13;
  this.COLUMN_INDEX_COST = 14;
  this.COLUMN_INDEX_COST_ARRAY = 15;
  this.COLUMN_INDEX_EMAIL_SCHEDULE = 16;
  this.COLUMN_INDEX_SCHEDULE_CONFIRMED = 17;
  this.COLUMN_INDEX_EMAIL_SPC_TOPICS_TITLE = 18;
  this.ROW_INDEX_FIRST_COURSE = 1;
  this.NUM_COLUMNS_PER_COURSE = 14;
  this.MAX_COURSES = 200
  // TODO: THE FOLLOWING 2 CONSTANTS APPEAR IN 3 DIFFERENT PLACES IN THE CODE AND MUST BE THE SAME DECLARED VALUES!
  this.COST_FACULTY_SAME_DAYS_PREFERENCE = 2.0;
  this.COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE = 0.5;  
}

PriorScheduleTab.prototype.transferScheduledCourses = function(transferMode, prior_schedule_sheet, 
                                                               scheduledCourseList, facultyCoursesAndPrefsList) {

}*/

// write scheduled courses to a sheet
function transferScheduledCourses(transferMode, prior_schedule_sheet, scheduledCourseList, facultyCoursesAndPrefsList) {
  var COLUMN_INDEX_ADD_TO_SCHEDULE = 0;
  var COLUMN_INDEX_DEPARTMENT_CODE = 1;
  var COLUMN_INDEX_COURSE_NUMBERS = 2;
  var COLUMN_INDEX_SECTION = 3;
  var COLUMN_INDEX_INSTRUCTIONAL_MODE = 4;
  var COLUMN_INDEX_CRN_NUMBER = 5;
  var COLUMN_INDEX_TERM = 6;
  var COLUMN_INDEX_START_TIME = 7;
  var COLUMN_INDEX_END_TIME = 8;
  var COLUMN_INDEX_CREDIT_HOURS_PER_DAY = 9;
  var COLUMN_INDEX_DAYS_OF_WEEK = 10;
  var COLUMN_INDEX_BUILDING = 11;
  var COLUMN_INDEX_ROOM = 12;
  var COLUMN_INDEX_INSTRUCTOR = 13;
  var COLUMN_INDEX_INSTRUCTOR_EMAIL = 14;
  var COLUMN_INDEX_COST = 15;
  var COLUMN_INDEX_COST_ARRAY = 16;
  var COLUMN_INDEX_EMAIL_SCHEDULE = 17;
  var COLUMN_INDEX_SCHEDULE_CONFIRMED = 18;
  var COLUMN_INDEX_EMAIL_SPC_TOPICS_TITLE = 19;
  var ROW_INDEX_FIRST_COURSE = 1;
  var NUM_COLUMNS_PER_COURSE = 15;
  var MAX_COURSES = 200
  // TODO: THE FOLLOWING 2 CONSTANTS APPEAR IN 3 DIFFERENT PLACES IN THE CODE AND MUST BE THE SAME DECLARED VALUES!
  var COST_FACULTY_SAME_DAYS_PREFERENCE = 2.0;
  var COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE = 0.5;

  if (transferMode == 'WRITE') {
    prior_schedule_sheet.getRange(ROW_INDEX_FIRST_COURSE + 1, 2, MAX_COURSES, 2 + NUM_COLUMNS_PER_COURSE - 1).clearContent();
    var schedule_output_datarange = prior_schedule_sheet.getRange(ROW_INDEX_FIRST_COURSE + 1, 2, scheduledCourseList.length, 2 + NUM_COLUMNS_PER_COURSE - 1);
    var prior_schedule_data = [];
    for (var courseIdx = 0; courseIdx < scheduledCourseList.length; courseIdx++) {
      var course_data = [];
      course_data[COLUMN_INDEX_DEPARTMENT_CODE - 1] = scheduledCourseList[courseIdx].Course.dept_code;
      course_data[COLUMN_INDEX_COURSE_NUMBERS - 1] = scheduledCourseList[courseIdx].Course.numbers.join("/");
      course_data[COLUMN_INDEX_SECTION - 1] = scheduledCourseList[courseIdx].Course.section;
      course_data[COLUMN_INDEX_CRN_NUMBER - 1] = scheduledCourseList[courseIdx].Course.crn;
      course_data[COLUMN_INDEX_TERM - 1] = scheduledCourseList[courseIdx].Course.term;
      course_data[COLUMN_INDEX_START_TIME - 1] = scheduledCourseList[courseIdx].CourseTime.getTimeString(scheduledCourseList[courseIdx].CourseTime.start);
      course_data[COLUMN_INDEX_END_TIME - 1] = scheduledCourseList[courseIdx].CourseTime.getTimeString(scheduledCourseList[courseIdx].CourseTime.end);
      course_data[COLUMN_INDEX_CREDIT_HOURS_PER_DAY - 1] = scheduledCourseList[courseIdx].CourseTime.credit_hours_per_day;
      course_data[COLUMN_INDEX_DAYS_OF_WEEK - 1] = scheduledCourseList[courseIdx].CourseTime.days.join("/");
      course_data[COLUMN_INDEX_BUILDING - 1] = scheduledCourseList[courseIdx].Room.building;
      course_data[COLUMN_INDEX_ROOM - 1] = scheduledCourseList[courseIdx].Room.number;
      if (scheduledCourseList[courseIdx].FacultyCoursesAndPrefs != undefined) {
        course_data[COLUMN_INDEX_INSTRUCTOR - 1] = scheduledCourseList[courseIdx].FacultyCoursesAndPrefs.name;
        course_data[COLUMN_INDEX_INSTRUCTOR_EMAIL - 1] = scheduledCourseList[courseIdx].FacultyCoursesAndPrefs.email;
      } else {
        course_data[COLUMN_INDEX_INSTRUCTOR - 1] = "undefined";
        course_data[COLUMN_INDEX_INSTRUCTOR_EMAIL - 1] = "undefined";
      }
      course_data[COLUMN_INDEX_INSTRUCTIONAL_MODE - 1] = scheduledCourseList[courseIdx].Course.instructional_mode;
      course_data[COLUMN_INDEX_COST - 1] = scheduledCourseList[courseIdx].cost.toFixed(4);
      var ostr = "";
      if (scheduledCourseList[courseIdx].costArray != undefined) {
        for (var key in scheduledCourseList[courseIdx].costArray) {
          ostr += key + "," + scheduledCourseList[courseIdx].costArray[key].toFixed(2) + " + ";
        }
        if (ostr.length > 3) {
          ostr = ostr.substring(0, ostr.length - 3);
        }
      }
      course_data[COLUMN_INDEX_COST_ARRAY - 1] = ostr;
      prior_schedule_data.push(course_data);
    }
    schedule_output_datarange.setValues(prior_schedule_data);
  } else if (transferMode == 'READ') {
    var schedule_input_datarange = prior_schedule_sheet.getDataRange().getValues();
    var cost;
    for (var i = ROW_INDEX_FIRST_COURSE; i < schedule_input_datarange.length; i++) {
      if (schedule_input_datarange[i][COLUMN_INDEX_ADD_TO_SCHEDULE]) {
        cost = 0;
        var course_time = new CourseTime(schedule_input_datarange[i][COLUMN_INDEX_START_TIME],
          schedule_input_datarange[i][COLUMN_INDEX_END_TIME],
          schedule_input_datarange[i][COLUMN_INDEX_DAYS_OF_WEEK],
          schedule_input_datarange[i][COLUMN_INDEX_CREDIT_HOURS_PER_DAY],
          [], [], 0);

        var room = new Room(schedule_input_datarange[i][COLUMN_INDEX_BUILDING], schedule_input_datarange[i][COLUMN_INDEX_ROOM], undefined, 0);

        var preScheduledCourse = new Course(schedule_input_datarange[i][COLUMN_INDEX_DEPARTMENT_CODE],
          schedule_input_datarange[i][COLUMN_INDEX_COURSE_NUMBERS],
          schedule_input_datarange[i][COLUMN_INDEX_SECTION],
          schedule_input_datarange[i][COLUMN_INDEX_CRN_NUMBER],
          schedule_input_datarange[i][COLUMN_INDEX_TERM],
          // TODO: BROKEN - Computing Course Credit hours from the hours scheduled per week - not correct
          course_time.days.length * course_time.credit_hours_per_day);
        preScheduledCourse.instructional_mode = schedule_input_datarange[i][COLUMN_INDEX_INSTRUCTIONAL_MODE];

        var roomWithTimeInterval = new RoomWithTimeInterval(room, course_time);
        roomWithTimeInterval.cost = 0;
        cost += roomWithTimeInterval.cost;

        var teacherForThisCourse = undefined;
        var teacher_scheduledCourseList = [];
        var costDelta = 0;

        // Find the faculty member teaching this course (nextCourse) store the data in teacherForThisCourse      
        for (var facultyIdx = 0; teacherForThisCourse == undefined && facultyIdx < facultyCoursesAndPrefsList.length; facultyIdx++) {
          var faculty = facultyCoursesAndPrefsList[facultyIdx];
          for (var courseIdx = 0; courseIdx < faculty.courseList.length; courseIdx++) {
            if (preScheduledCourse.sameNumber(faculty.courseList[courseIdx])) {
              if (preScheduledCourse.sameSection(faculty.courseList[courseIdx])) {
                teacherForThisCourse = faculty;
                break;
              }
            }
          }
        }

        // SAME DAYS Faculty Preference 
        // Find the 3hr courses already scheduled for this faculty member and store them in teacher_scheduledCourseList
        for (var idx = 0; idx < scheduledCourseList.length; idx++) {
          var scheduledCourse = scheduledCourseList[idx];
          if (teacherForThisCourse != undefined && teacherForThisCourse.name != undefined &&
            scheduledCourse.FacultyCoursesAndPrefs != undefined && scheduledCourse.FacultyCoursesAndPrefs.name != undefined &&
            teacherForThisCourse.name == scheduledCourse.FacultyCoursesAndPrefs.name && scheduledCourse.Course.credit_hours == 3) {
            teacher_scheduledCourseList.push(scheduledCourse);
          }
        }
        var costArr = {};
        //if (roomWithTimeInterval.Room.cost != 0) {
        //  costArr['room'] = candidateRoomWithTimeInterval.Room.cost;
        //}
        //if (roomWithTimeInterval.CourseTime.cost != 0) {
        //  costArr['time'] = candidateRoomWithTimeInterval.CourseTime.cost;
        //}

        if (teacherForThisCourse != undefined) {
          costDelta = teacherForThisCourse.timeIntervalCostMap[roomWithTimeInterval.CourseTime.getId()];
          if (costDelta != undefined) {
            cost += costDelta;
            costArr['faculty_time'] = costDelta;
          }
        }

        // SAME DAYS Faculty Preference 
        // if the faculty is already scheduled for a 3hr time slot
        // check if the same_days preference is true, if so, check the candidateRoomWithTimeInterval.CourseTime
        // if the candidateRoomWithTimeInterval.CourseTime is not on overlapping days add COST_FACULTY_SAME_DAYS_PREFERENCE to cost
        if (teacherForThisCourse != undefined && teacher_scheduledCourseList.length > 0) {
          for (var idx2 = 0; idx2 < teacher_scheduledCourseList.length; idx2++) {
            var numDaysInCommon = roomWithTimeInterval.CourseTime.daysInCommon(teacher_scheduledCourseList[idx2].CourseTime).length;
            if (numDaysInCommon == 0) {
              if (teacherForThisCourse.courses_on_same_days == true) {  // course does not have overlapping day and faculty prefers overlapping days
                costDelta = COST_FACULTY_SAME_DAYS_PREFERENCE;
                cost += costDelta;
                costArr['same_days'] = costDelta;
                break;
              }
            } else {
              if (teacherForThisCourse.courses_on_same_days == false) {  // course has overlapping days and faculty does not prefer overlapping days
                costDelta = COST_FACULTY_SAME_DAYS_PREFERENCE;
                cost += costDelta;
                costArr['same_days'] = costDelta;
                break;
              }
            }
          }
        }

        // Minimum Time Between Classes Preference 
        // if the faculty is already scheduled for a 3hr time slot, check in the candidateRoomWithTimeInterval.CourseTime
        // same days and time between slots is less than the value sepecified add COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE to cost
        if (teacherForThisCourse != undefined && teacher_scheduledCourseList.length > 0) {
          for (var idx2 = 0; idx2 < teacher_scheduledCourseList.length; idx2++) {
            var numDaysInCommon = roomWithTimeInterval.CourseTime.daysInCommon(teacher_scheduledCourseList[idx2].CourseTime).length;
            if (numDaysInCommon > 0) {
              var hoursBetweenCourses = roomWithTimeInterval.CourseTime.hoursBetween(teacher_scheduledCourseList[idx2].CourseTime);
              //var err = Math.abs(teacherForThisCourse.hours_between_courses - hoursBetweenCourses);
              if (hoursBetweenCourses < teacherForThisCourse.hours_between_courses) {
                costDelta = COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE;
                cost += costDelta;
                costArr['hrs_between_courses'] = costDelta;
                break;
              } //else if (hoursBetweenCourses > teacherForThisCourse.hours_between_courses) {
              //costDelta = 0.25*(hoursBetweenCourses - teacherForThisCourse.hours_between_courses)*COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE;
              //cost += costDelta;
              //costArr['hrs_between_courses'] = costDelta;
              //break;
              //}
            }
          }
        }

        var newScheduledCourse = new ScheduledCourse(preScheduledCourse, room, course_time, cost, teacherForThisCourse, costArr);
        scheduledCourseList.push(newScheduledCourse);
      }
    }
  }
}

// inserts the Pre-Scheduled courses into the schedule
function insertPreScheduledCourses(course_and_time_constraint_sheet, output_sheet_data, roomWithTimeIntervalList, facultyCoursesAndPrefsList, scheduledCourseList) {
  // HARD-CODED CONSTANTS
  var COLUMN_INDEX_ADD_TO_SCHEDULE = 0;
  var COLUMN_INDEX_DEPARTMENT_CODE = 1;
  var COLUMN_INDEX_COURSE_NUMBERS = 2;
  var COLUMN_INDEX_SECTION = 3;
  var COLUMN_INDEX_INSTRUCTIONAL_MODE = 4;
  var COLUMN_INDEX_CRN_NUMBER = 5;
  var COLUMN_INDEX_TERM = 6;
  var COLUMN_INDEX_START_TIME = 7;
  var COLUMN_INDEX_END_TIME = 8;
  var COLUMN_INDEX_CREDIT_HOURS_PER_DAY = 9;
  var COLUMN_INDEX_DAYS_OF_WEEK = 10;
  var COLUMN_INDEX_BUILDING = 11;
  var COLUMN_INDEX_ROOM = 12;
  var ROW_INDEX_FIRST_COURSE = 2;

  // TODO: THE FOLLOWING 2 CONSTANTS APPEAR IN 3 DIFFERENT PLACES IN THE CODE AND MUST BE THE SAME DECLARED VALUES!
  var COST_FACULTY_SAME_DAYS_PREFERENCE = 2.0;
  var COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE = 0.5;

  var ONLINE_ROOM_CONSTANT = new Room('ONLINE', 'INTERNET');

  var hardconstraints_data = course_and_time_constraint_sheet.getDataRange().getValues();
  var costTotal = 0;

  for (var i = ROW_INDEX_FIRST_COURSE; i < hardconstraints_data.length; i++) {
    if (hardconstraints_data[i][COLUMN_INDEX_ADD_TO_SCHEDULE]) {
      var cost = 0;
      var course_time = new CourseTime(hardconstraints_data[i][COLUMN_INDEX_START_TIME],
        hardconstraints_data[i][COLUMN_INDEX_END_TIME],
        hardconstraints_data[i][COLUMN_INDEX_DAYS_OF_WEEK],
        hardconstraints_data[i][COLUMN_INDEX_CREDIT_HOURS_PER_DAY],
        [], [], 0);

      var room = new Room(hardconstraints_data[i][COLUMN_INDEX_BUILDING], hardconstraints_data[i][COLUMN_INDEX_ROOM], undefined, 0);

      var hardConstraintCourse = new Course(hardconstraints_data[i][COLUMN_INDEX_DEPARTMENT_CODE],
        hardconstraints_data[i][COLUMN_INDEX_COURSE_NUMBERS],
        hardconstraints_data[i][COLUMN_INDEX_SECTION],
        hardconstraints_data[i][COLUMN_INDEX_CRN_NUMBER],
        hardconstraints_data[i][COLUMN_INDEX_TERM],
        // TODO: BROKEN - Computing Course Credit hours from the hours scheduled per week - not correct
        course_time.days.length * course_time.credit_hours_per_day);
      hardConstraintCourse.instructional_mode = hardconstraints_data[i][COLUMN_INDEX_INSTRUCTIONAL_MODE];

      // If instructional mode is either 'Online Synchronous' or 'Online Asynchronous' the required room is 'ONLINE INTERNET'
      if (hardConstraintCourse.instructional_mode == 'Online Synchronous' ||
        hardConstraintCourse.instructional_mode == 'Online Asynchronous') {
        room = ONLINE_ROOM_CONSTANT;
      }

      var roomWithTimeInterval = new RoomWithTimeInterval(room, course_time);
      roomWithTimeInterval.cost = 0;
      cost += roomWithTimeInterval.cost;

      var teacherForThisCourse = undefined;
      var teacher_scheduledCourseList = [];
      var costDelta = 0;

      // Find the faculty member teaching this course (nextCourse) store the data in teacherForThisCourse      
      for (var facultyIdx = 0; teacherForThisCourse == undefined && facultyIdx < facultyCoursesAndPrefsList.length; facultyIdx++) {
        var faculty = facultyCoursesAndPrefsList[facultyIdx];
        for (var courseIdx = 0; courseIdx < faculty.courseList.length; courseIdx++) {
          if (hardConstraintCourse.sameNumber(faculty.courseList[courseIdx])) {
            if (hardConstraintCourse.sameSection(faculty.courseList[courseIdx])) {
              teacherForThisCourse = faculty;
              break;
            }
          }
        }
      }

      // SAME DAYS Faculty Preference 
      // Find the 3hr courses already scheduled for this faculty member and store them in teacher_scheduledCourseList
      for (var idx = 0; idx < scheduledCourseList.length; idx++) {
        var scheduledCourse = scheduledCourseList[idx];
        if (teacherForThisCourse != undefined && teacherForThisCourse.name != undefined &&
          scheduledCourse.FacultyCoursesAndPrefs != undefined && scheduledCourse.FacultyCoursesAndPrefs.name != undefined &&
          teacherForThisCourse.name == scheduledCourse.FacultyCoursesAndPrefs.name && scheduledCourse.Course.credit_hours == 3) {
          teacher_scheduledCourseList.push(scheduledCourse);
        }
      }
      var costArr = {};
      //if (roomWithTimeInterval.Room.cost != 0) {
      //  costArr['room'] = candidateRoomWithTimeInterval.Room.cost;
      //}
      //if (roomWithTimeInterval.CourseTime.cost != 0) {
      //  costArr['time'] = candidateRoomWithTimeInterval.CourseTime.cost;
      //}

      if (teacherForThisCourse != undefined) {
        costDelta = teacherForThisCourse.timeIntervalCostMap[roomWithTimeInterval.CourseTime.getId()];
        if (costDelta != undefined) {
          cost += costDelta;
          costArr['faculty_time'] = costDelta;
        }
      }

      // SAME DAYS Faculty Preference 
      // if the faculty is already scheduled for a 3hr time slot
      // check if the same_days preference is true, if so, check the candidateRoomWithTimeInterval.CourseTime
      // if the candidateRoomWithTimeInterval.CourseTime is not on overlapping days add COST_FACULTY_SAME_DAYS_PREFERENCE to cost
      if (teacherForThisCourse != undefined && teacher_scheduledCourseList.length > 0) {
        for (var idx2 = 0; idx2 < teacher_scheduledCourseList.length; idx2++) {
          var numDaysInCommon = roomWithTimeInterval.CourseTime.daysInCommon(teacher_scheduledCourseList[idx2].CourseTime).length;
          if (numDaysInCommon == 0) {
            if (teacherForThisCourse.courses_on_same_days == true) {  // course does not have overlapping day and faculty prefers overlapping days
              costDelta = COST_FACULTY_SAME_DAYS_PREFERENCE;
              cost += costDelta;
              costArr['same_days'] = costDelta;
              break;
            }
          } else {
            if (teacherForThisCourse.courses_on_same_days == false) {  // course has overlapping days and faculty does not prefer overlapping days
              costDelta = COST_FACULTY_SAME_DAYS_PREFERENCE;
              cost += costDelta;
              costArr['same_days'] = costDelta;
              break;
            }
          }
        }
      }

      // Minimum Time Between Classes Preference 
      // if the faculty is already scheduled for a 3hr time slot, check in the candidateRoomWithTimeInterval.CourseTime
      // same days and time between slots is less than the value sepecified add COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE to cost
      if (teacherForThisCourse != undefined && teacher_scheduledCourseList.length > 0) {
        for (var idx2 = 0; idx2 < teacher_scheduledCourseList.length; idx2++) {
          var numDaysInCommon = roomWithTimeInterval.CourseTime.daysInCommon(teacher_scheduledCourseList[idx2].CourseTime).length;
          if (numDaysInCommon > 0) {
            var hoursBetweenCourses = roomWithTimeInterval.CourseTime.hoursBetween(teacher_scheduledCourseList[idx2].CourseTime);
            //var err = Math.abs(teacherForThisCourse.hours_between_courses - hoursBetweenCourses);
            if (hoursBetweenCourses < teacherForThisCourse.hours_between_courses) {
              costDelta = COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE;
              cost += costDelta;
              costArr['hrs_between_courses'] = costDelta;
              break;
            } //else if (hoursBetweenCourses > teacherForThisCourse.hours_between_courses) {
            //costDelta = 0.25*(hoursBetweenCourses - teacherForThisCourse.hours_between_courses)*COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE;
            //cost += costDelta;
            //costArr['hrs_between_courses'] = costDelta;
            //break;
            //}
          }
        }
      }

      var newScheduledCourse = new ScheduledCourse(hardConstraintCourse, room, course_time, cost, teacherForThisCourse, costArr);
      scheduledCourseList.push(newScheduledCourse);

      // render the course to a sheet
      addCourseToSchedule(output_sheet_data, newScheduledCourse);
      defaultScheduleRenderer.addCourseToSchedule(newScheduledCourse);

      Logger.log('ScheduleEngine: Scheduled ' + hardConstraintCourse.getId() + ' in ' + room.getId() +
        " at " + course_time.getId() + ' with cost ' + roomWithTimeInterval.cost);

      // TODO: WHEN ONLY A PART OF A TIMESLOT IS OCCUPIED THE ENTIRE TIMESLOT IS NOT REMOVED BUT ONLY THE TIMESLOTS OCCUPIED ARE REMOVED
      // remove room with time interval from roomWithTimeIntervalList
      removeFromList(newScheduledCourse, roomWithTimeInterval, roomWithTimeIntervalList, RoomWithTimeInterval.prototype.isEqualRoomAndTimeInterval);

      costTotal += cost;
      // add to scheduled course list
      output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_COST_OUTPUT, OUTPUT_SHEET_INDEX_COL_COST_OUTPUT).setValue(costTotal.toFixed(2));
      output_sheet_data.getCell(OUTPUT_SHEET_INDEX_ROW_NUMCOURSES_OUTPUT, OUTPUT_SHEET_INDEX_COL_NUMCOURSES_OUTPUT).setValue(scheduledCourseList.length);
      Logger.log('ScheduleEngine: Scheduled ' + scheduledCourseList.length + " courses.");
    }
  }
  return costTotal;
}

function removeFromList(scheduledRoomWithTimeInterval, roomWithTimeInterval, roomWithTimeIntervalList, meetsRemovalCriteria) {
  // When the building is 'NONE' or 'ONLINE', e.g., 'NONE TBA' or 'ONLINE INTERNET' we do not modify the room availability list (roomWithTimeIntervalList) 
  // as this room is virtual and has an inexaustible supply of space      
  if (scheduledRoomWithTimeInterval.Room.building == 'NONE' || scheduledRoomWithTimeInterval.Room.building == 'ONLINE') {
    return;
  }

  var foundIdxs = [];
  for (var idx = 0; idx < roomWithTimeIntervalList.length; idx++) {
    var query = roomWithTimeIntervalList[idx];
    if (meetsRemovalCriteria(roomWithTimeInterval, query) && roomWithTimeInterval.CourseTime.daysInCommon(query.CourseTime).length > 0) {
      foundIdxs.push(idx);
    }
  }
  // remove from end of array first so indices don't change due to removals
  foundIdxs.sort(function (a, b) { return b - a; });

  // remove timeslot
  for (var delIdx = 0; delIdx < foundIdxs.length; delIdx++) {

    var scheduledRoomWithTimeIntervalTotalCreditHours = roomWithTimeIntervalList[foundIdxs[delIdx]].CourseTime.credit_hours_per_day * roomWithTimeIntervalList[foundIdxs[delIdx]].CourseTime.days.length;
    // If scheduledRoomWithTimeIntervalTotalCreditHours > 3 the code here will not work correctly - breaks for summer scheduling
    // TODO: FIX CREDIT HOUR SPLITS OF TIMESLOTS FOR SUMMER
    if (scheduledRoomWithTimeIntervalTotalCreditHours <= 3.0) {
      var excessDays = (scheduledRoomWithTimeIntervalTotalCreditHours - scheduledRoomWithTimeInterval.Course.credit_hours) / scheduledRoomWithTimeInterval.CourseTime.credit_hours_per_day;
      var occupiedCourseTime = roomWithTimeIntervalList[foundIdxs[delIdx]].CourseTime;
      Logger.log('ScheduleEngine: Deleting course+time slot ' + roomWithTimeIntervalList[foundIdxs[delIdx]].getId());
      roomWithTimeIntervalList.splice(foundIdxs[delIdx], 1);

      if (excessDays >= 1) {
        var oRoom = scheduledRoomWithTimeInterval.Room;
        var copyRoom = new Room(oRoom.building, oRoom.number, oRoom.maximum_capacity, oRoom.cost);
        // NOTE: We have to clone this CourseTime object since many roomWithTimeInterval objects in the roomWithTimeIntervalList hold references to a single CourseTime object
        // Hence we create (2) new CourseTime objects due to the split in days; neither of which can alter the referenced original un-split CourseTime object
        // copy (1) goes back into the scheduledRoomWithTimeInterval object copy (2) is part of the excessRoomWithTimeInterval object
        var oCourseTime = scheduledRoomWithTimeInterval.CourseTime;
        var copyCourseTime = new CourseTime(oCourseTime.start, oCourseTime.end, oCourseTime.days.join("/"), oCourseTime.credit_hours_per_day,
          oCourseTime.exclude_course_list, oCourseTime.include_course_list, oCourseTime.cost);
        scheduledRoomWithTimeInterval.CourseTime = copyCourseTime; // copy (1) goes back into scheduledRoomWithTimeInterval
        var excessCourseTimeIntervals = splitCourseTimeByDay(scheduledRoomWithTimeInterval.CourseTime, excessDays, occupiedCourseTime); // creates copy (2) and returns it and modifies the days of copy (1) accordingly
        //if (excessCourseTimeIntervals != undefined) {
        var excessRoomWithTimeInterval = new RoomWithTimeInterval(copyRoom, excessCourseTimeIntervals);
        Logger.log('ScheduleEngine: Adding unused course+time slot ' + excessRoomWithTimeInterval.getId());
        //roomWithTimeIntervalList.push(excessRoomWithTimeInterval);
        //}
      }
    }
  }
  Logger.log('ScheduleEngine: Deleted ' + foundIdxs.length + ' available course+time slots. There are ' + roomWithTimeIntervalList.length + ' slots available.');
}

// retrieve Course data
function getCoursesToSchedule(course_datarange) {
  // HARD-CODED CONSTANTS
  var COLUMN_INDEX_ADD_TO_SCHEDULE = 0;
  var COLUMN_INDEX_DEPARTMENT_CODE = 1;
  var COLUMN_INDEX_COURSE_NUMBERS = 2;
  var COLUMN_INDEX_COURSE_SECTION = 3;
  var COLUMN_INDEX_COURSE_INSTRUCTIONAL_MODE = 4;
  var COLUMN_INDEX_COURSE_CRN_NUMBER = 5;
  var COLUMN_INDEX_COURSE_TERM = 6;
  var COLUMN_INDEX_COURSE_CREDIT_HOURS = 7;
  var COLUMN_INDEX_COURSE_PRIORITY = 8;
  var COLUMN_INDEX_COURSE_EXPECTED_ENROLLMENT = 9;
  var COLUMN_INDEX_COURSE_NOT_SIMULTANEOUS_COURSES = 10;
  var COLUMN_INDEX_COURSE_SIMULTANEOUS_COURSES = 11;
  var COLUMN_INDEX_REQUIRED_ROOM_IDS = 12;
  var COLUMN_INDEX_EXCLUDED_ROOM_IDS = 13;
  var ROW_INDEX_FIRST_COURSE = 1;

  var course_data = course_datarange.getValues();
  var courseList = [];

  for (var rowIdx = ROW_INDEX_FIRST_COURSE; rowIdx < course_datarange.getHeight(); rowIdx++) {
    var scheduleThisCourse = course_data[rowIdx][COLUMN_INDEX_ADD_TO_SCHEDULE];
    if (scheduleThisCourse && isValidDeptCode(course_data[rowIdx][COLUMN_INDEX_DEPARTMENT_CODE]) &&
      isValidCourseNumber(course_data[rowIdx][COLUMN_INDEX_COURSE_NUMBERS])) { // careful for whitespace!
      var dept_code = course_data[rowIdx][COLUMN_INDEX_DEPARTMENT_CODE];
      var numbers = course_data[rowIdx][COLUMN_INDEX_COURSE_NUMBERS];
      var section = course_data[rowIdx][COLUMN_INDEX_COURSE_SECTION];
      var instructional_mode = course_data[rowIdx][COLUMN_INDEX_COURSE_INSTRUCTIONAL_MODE];
      var crn = course_data[rowIdx][COLUMN_INDEX_COURSE_CRN_NUMBER];
      var term = course_data[rowIdx][COLUMN_INDEX_COURSE_TERM];
      var credit_hours = course_data[rowIdx][COLUMN_INDEX_COURSE_CREDIT_HOURS];
      var priority = course_data[rowIdx][COLUMN_INDEX_COURSE_PRIORITY];
      var expected_enrollment = course_data[rowIdx][COLUMN_INDEX_COURSE_EXPECTED_ENROLLMENT];
      var not_simultaneous_courses = parseCommaSeparatedArray(course_data[rowIdx][COLUMN_INDEX_COURSE_NOT_SIMULTANEOUS_COURSES].toString(), "courses");
      var simultaneous_courses = parseCommaSeparatedArray(course_data[rowIdx][COLUMN_INDEX_COURSE_SIMULTANEOUS_COURSES].toString(), "courses");
      var these_rooms_only_ids = parseCommaSeparatedArray(course_data[rowIdx][COLUMN_INDEX_REQUIRED_ROOM_IDS].toString(), "rooms");
      var excluded_rooms = parseCommaSeparatedArray(course_data[rowIdx][COLUMN_INDEX_EXCLUDED_ROOM_IDS].toString(), "rooms");
      var course = new Course(dept_code, numbers, section, crn, term, credit_hours, priority, expected_enrollment,
        not_simultaneous_courses, simultaneous_courses, these_rooms_only_ids, excluded_rooms, instructional_mode);
      courseList.push(course);
    }
  }
  return courseList;
}

// convert (NOT SIMULTANEOUS COURSES, SIMULTANEOUS COURSES, REQUIRED ROOM IDS, EXCLUDED ROOM IDS) lists to arrays
function parseCommaSeparatedArray(str, filter) {
  var array = [];
  if (filter == "rooms") {
    array = str.split(",");     // split into string array using ',' delimiter
    for (var strIdx = array.length - 1; strIdx >= 0; strIdx--) {
      if (array[strIdx] != "") {
        array[strIdx] = parseRoomListElement(array[strIdx]);
      } else {
        array.splice(strIdx, 1); // remove empty cell elements       
      }
    }
  } else if (filter == "courses") {
    str = str.replace(/ /g, ''); // remove space characters
    array = str.split(",");     // split into string array using ',' delimiter
    for (var strIdx = array.length - 1; strIdx >= 0; strIdx--) {
      if (array[strIdx] != "") {
        array[strIdx] = parseCourseListElement(array[strIdx]);
      } else {
        array.splice(strIdx, 1); // remove empty cell elements
      }
    }
  }
  // DEBUG
  //if (array == undefined) {
  //  var test = 1;
  //}
  return array;
}

function parseRoomListElement(str) {
  // HARD-CODED CONSTANTS
  var DEFAULT_BUILDING = "EPIC";
  // remove spaces in the string
  //str = str.replace(/\s/g, '');
  //var building = /[A-Z]{4}/.exec(str); // 4 capital letters,, if null assume DEFAULT_BUILDING
  str = str.trim();
  str = str.split(' ');
  var building = (str.length == 1) ? undefined : str[0];
  var number = (str.length == 1) ? str[0] : str[1];
  if (building == undefined) {
    building = DEFAULT_BUILDING;
  }// else {
  // building = building[0];
  //}
  //var number = str.substring(4,str.length);
  //var number = /[0-9]{4}/.exec(str)[0]; // return 1st match of 4 digits (required)
  return new Room(building, number);
}

function parseCourseListElement(str) {
  // HARD-CODED CONSTANTS
  //var DEFAULT_DEPT_CODE = DEFAULT_DEPT_CODE;

  var dept_code = /[A-Z]{4}/.exec(str); // 4 capital letters, if null assume DEFAULT_DEPT_CODE
  if (dept_code == null) {
    dept_code = DEFAULT_DEPT_CODE;
  } else {
    dept_code = dept_code[0];
  }
  str = str.split("-");
  var number = /[0-9,?]{4}/.exec(str[0])[0]; // 4 digits (required)
  var section = 1; // default section if none found
  if (str.length > 1) {  // 2 optional capital letter or digit and 1 (required) additional digit
    //section = /[A-Z]?[\d]{1,2}/.exec(str[1]);
    section = /[A-Z0-9]{2}?[0-9]{1}/.exec(str[1]);
  }
  return new Course(dept_code, number, section);
}

// retrieve Room data
function getRoomsToSchedule(room_datarange) {
  // HARD-CODED CONSTANTS
  var COLUMN_INDEX_USE_ROOM = 0;
  var COLUMN_INDEX_BUILDING = 1;
  var COLUMN_INDEX_ROOM_NUMBER = 2;
  var COLUMN_INDEX_MAXIMUM_CAPACITY = 3;
  var COLUMN_INDEX_COST = 4;
  var ROW_INDEX_FIRST_ROOM = 1;

  var room_data = room_datarange.getValues();
  var roomList = [];

  for (var rowIdx = ROW_INDEX_FIRST_ROOM; rowIdx < room_datarange.getHeight(); rowIdx++) {
    var building = room_data[rowIdx][COLUMN_INDEX_BUILDING];
    if (room_data[rowIdx][COLUMN_INDEX_USE_ROOM] == false || building.trim() == "") {
      continue;
    }
    var number = room_data[rowIdx][COLUMN_INDEX_ROOM_NUMBER];
    var maximum_capacity = room_data[rowIdx][COLUMN_INDEX_MAXIMUM_CAPACITY];
    var cost = room_data[rowIdx][COLUMN_INDEX_COST];
    var room = new Room(building, number, maximum_capacity, cost);
    roomList.push(room);
  }
  return roomList;
}

// retrieve CourseTime time interval data
function getTimeIntervalsToSchedule(time_interval_datarange) {
  // HARD-CODED CONSTANTS
  var COLUMN_INDEX_USE_THIS_TIME_INTERVAL = 0;
  var COLUMN_INDEX_START_TIME = 1;
  var COLUMN_INDEX_END_TIME = 2;
  var COLUMN_INDEX_CREDIT_HOURS_PER_DAY = 3;
  var COLUMN_INDEX_DAYS = 4;
  var COLUMN_INDEX_EXCLUDE_COURSE_LIST = 5;
  var COLUMN_INDEX_INCLUDE_COURSE_LIST = 6;
  var COLUMN_INDEX_COST = 7;
  var ROW_INDEX_FIRST_TIME_INTERVAL = 1;

  var time_interval_data = time_interval_datarange.getValues();
  var timeIntervalList = [];

  for (var rowIdx = ROW_INDEX_FIRST_TIME_INTERVAL; rowIdx < time_interval_datarange.getHeight(); rowIdx++) {
    var use_this_time_interval = time_interval_data[rowIdx][COLUMN_INDEX_USE_THIS_TIME_INTERVAL];
    if (use_this_time_interval == true) {
      var start = time_interval_data[rowIdx][COLUMN_INDEX_START_TIME];
      var end = time_interval_data[rowIdx][COLUMN_INDEX_END_TIME];
      var credit_hours_per_day = time_interval_data[rowIdx][COLUMN_INDEX_CREDIT_HOURS_PER_DAY];
      var days = time_interval_data[rowIdx][COLUMN_INDEX_DAYS];
      var exclude_course_list = parseCommaSeparatedArray(time_interval_data[rowIdx][COLUMN_INDEX_EXCLUDE_COURSE_LIST].toString(), "courses");
      var include_course_list = parseCommaSeparatedArray(time_interval_data[rowIdx][COLUMN_INDEX_INCLUDE_COURSE_LIST].toString(), "courses");
      var cost = time_interval_data[rowIdx][COLUMN_INDEX_COST];
      var course_time = new CourseTime(start, end, days, credit_hours_per_day, exclude_course_list, include_course_list, cost);
      timeIntervalList.push(course_time);
    }
  }
  return timeIntervalList;
}

// retrieve FacultyCoursesAndPreferences data
function getFacultyCoursesAndPreferencesToSchedule(faculty_datarange, courseTimeList) {
  // HARD-CODED CONSTANTS
  var COLUMN_INDEX_FACULTY_NAME = 0;
  var COLUMN_INDEX_FACULTY_EMAIL = 1;
  var COLUMN_INDEX_COURSE_ASSIGNMENTS_RANGE = [2, 10];
  var COLUMN_INDEX_AT_RISK = 11;
  var COLUMN_INDEX_PREFS_SAME_DAYS = 12;
  var COLUMN_INDEX_PREFS_HOURS_BETWEEN_COURSES = 13;
  var COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS = [15, 39];
  var DEFAULT_PREFS_TIME_INTERVAL_COST = 0;
  var ROW_INDEX_FIRST_FACULTY = 3;

  var faculty_courses_and_prefs_data = faculty_datarange.getValues();
  var facultyCoursesAndPrefsList = [];
  var warningShown = false;

  for (var rowIdx = ROW_INDEX_FIRST_FACULTY; rowIdx < faculty_datarange.getHeight(); rowIdx++) {
    var name = faculty_courses_and_prefs_data[rowIdx][COLUMN_INDEX_FACULTY_NAME];
    var email = faculty_courses_and_prefs_data[rowIdx][COLUMN_INDEX_FACULTY_EMAIL];
    if (name != "") {
      var courseList = [];
      for (var colIdx = COLUMN_INDEX_COURSE_ASSIGNMENTS_RANGE[0]; colIdx <= COLUMN_INDEX_COURSE_ASSIGNMENTS_RANGE[1]; colIdx++) {
        if (faculty_courses_and_prefs_data[rowIdx][colIdx] != "") {
          // not super efficient here but gets the job done
          courseList.push(parseCommaSeparatedArray(faculty_courses_and_prefs_data[rowIdx][colIdx].toString(), "courses")[0]);
        }
      }
      var courses_on_same_days = faculty_courses_and_prefs_data[rowIdx][COLUMN_INDEX_PREFS_SAME_DAYS];
      var hours_between_courses = faculty_courses_and_prefs_data[rowIdx][COLUMN_INDEX_PREFS_HOURS_BETWEEN_COURSES];
      // WE DEPEND ON THE TIME INTERVAL LIST PASSED (courseTimeList) TO HAVE THE SAME SEQUENCING AS THE COLUMNS OF THE FACULTY TIME INTERVAL COLUMNS
      // i.e., The preference columns and time interval list elements are collocated to be corresponding!
      // THIS IMPLIES THAT DATA ROWS IN THE TIME INTERVAL TAB MUST BE IN THE SAME ORDER AS DATA COLUMNS FOR TIME INTERVALS IN THE FACULTY PREFERENCES TAB
      // A map is used for faculty preference time interval costs, keys are CourseTime IDs, i.e., time intervals as strings, and values are costs
      var timeIntervalCostMap = {}; // <-- Efficiency: could be a map with timeinterval keys storing only non-default values
      for (var colIdx = COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]; colIdx <= COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[1]; colIdx++) {
        if (courseTimeList[colIdx - COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]] != undefined) {
          if (faculty_courses_and_prefs_data[rowIdx][colIdx] != "") {
            timeIntervalCostMap[courseTimeList[colIdx - COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]].getId()] = faculty_courses_and_prefs_data[rowIdx][colIdx];
          } else {
            timeIntervalCostMap[courseTimeList[colIdx - COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]].getId()] = DEFAULT_PREFS_TIME_INTERVAL_COST;
          }
        } else {
          if (!warningShown) {
            // Preferences for this time interval are not available show warning dialog          
            //SpreadsheetApp.getUi().alert('Warning: Faculty time preferences not available for ' + courseTimeList[colIdx - COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]] + '.');
            Logger.log('Warning: Faculty time preferences not available for ' + courseTimeList[colIdx - COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]] + '.');
            warningShown = true;
          }
        }
      }
      var faculty_courses_and_prefs = new FacultyCoursesAndPrefs(name, email, courseList, timeIntervalCostMap, courses_on_same_days, hours_between_courses);
      facultyCoursesAndPrefsList.push(faculty_courses_and_prefs);
    }
  }
  return facultyCoursesAndPrefsList;
}

function putFacultyCoursesAndPreferences(faculty_datarange, courseTimeList, facultyCoursesAndPrefsList) {
  // HARD-CODED CONSTANTS
  var COLUMN_INDEX_FACULTY_NAME = 0;
  var COLUMN_INDEX_FACULTY_EMAIL = 1;
  var COLUMN_INDEX_COURSE_ASSIGNMENTS_RANGE = [2, 10];
  var COLUMN_INDEX_AT_RISK = 11;
  var COLUMN_INDEX_PREFS_SAME_DAYS = 12;
  var COLUMN_INDEX_PREFS_HOURS_BETWEEN_COURSES = 13;
  var COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS = [15, 39];
  var DEFAULT_PREFS_TIME_INTERVAL_COST = 0;
  var ROW_INDEX_FIRST_FACULTY = 3;

  var faculty_courses_and_prefs_data = faculty_datarange.getValues();

  for (var facultyIdx = 0; facultyIdx < facultyCoursesAndPrefsList.length; facultyIdx++) {
    for (var rowIdx = ROW_INDEX_FIRST_FACULTY; rowIdx < faculty_datarange.getHeight(); rowIdx++) {
      //faculty_courses_and_prefs_data[rowIdx][COLUMN_INDEX_FACULTY_NAME];
      var sheet_email = faculty_courses_and_prefs_data[rowIdx][COLUMN_INDEX_FACULTY_EMAIL];
      if (sheet_email != "" && sheet_email != undefined && sheet_email == facultyCoursesAndPrefsList[facultyIdx].email) {
        var srcData = facultyCoursesAndPrefsList[facultyIdx];
        // set courses on same days preference
        faculty_courses_and_prefs_data[rowIdx][COLUMN_INDEX_PREFS_SAME_DAYS] = srcData.courses_on_same_days;
        // set hours between courses preference
        faculty_courses_and_prefs_data[rowIdx][COLUMN_INDEX_PREFS_HOURS_BETWEEN_COURSES] = srcData.hours_between_courses;
        // set time slot preference cost values
        // WE DEPEND ON THE TIME INTERVAL LIST PASSED (courseTimeList) TO HAVE THE SAME SEQUENCING AS THE COLUMNS OF THE FACULTY TIME INTERVAL COLUMNS
        // i.e., The preference columns and time interval list elements are collocated to be corresponding!
        // THIS IMPLIES THAT DATA ROWS IN THE TIME INTERVAL TAB MUST BE IN THE SAME ORDER AS DATA COLUMNS FOR TIME INTERVALS IN THE FACULTY PREFERENCES TAB
        // A map is used for faculty preference time interval costs, keys are CourseTime IDs, i.e., time intervals as strings, and values are costs
        // use the loaded coursetime list to convert timeslot string keys into time slot column indices
        for (var cTimeIdx = 0; cTimeIdx <= COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[1] - COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]; cTimeIdx++) {
          var timeSlotKey = courseTimeList[cTimeIdx].getId();
          faculty_courses_and_prefs_data[rowIdx][COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0] + cTimeIdx] = srcData.timeIntervalCostMap[timeSlotKey];
        }
        break;
      }
    }
  }
  faculty_datarange.setValues(faculty_courses_and_prefs_data);
}

// retrieve Faculty teaching preferences for import
function importFacultyPreferences(import_faculty_prefs_datarange, courseTimeList) {
  // HARD-CODED CONSTANTS
  var COLUMN_INDEX_IMPORT_THIS_ENTRY = 0;
  var COLUMN_INDEX_TIMESTAMP = 1;
  var COLUMN_INDEX_FACULTY_EMAIL = 2;
  var COLUMN_INDEX_PREFS_SAME_DAYS = 3;
  var COLUMN_INDEX_PREFS_HOURS_BETWEEN_COURSES = 4;
  var COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS = [5, 28];
  var ROW_INDICES_EXCLUDED_TIME_INTERVALS = [13];
  var DEFAULT_PREFS_TIME_INTERVAL_COST = 0;
  var ROW_INDEX_FIRST_FACULTY_PREFERENCE = 1;

  var MAP_SameDayString_To_Value = {};
  MAP_SameDayString_To_Value['Yes'] = true;
  MAP_SameDayString_To_Value['No'] = false;

  var MAP_HoursBetweenCoursesString_To_Value = {};
  MAP_HoursBetweenCoursesString_To_Value['No preference, I can teach courses back-to-back.'] = 0;
  MAP_HoursBetweenCoursesString_To_Value['I prefer 1 hour between courses on the same day.'] = 1;
  MAP_HoursBetweenCoursesString_To_Value['I prefer 2 hours between courses on the same day.'] = 2;
  MAP_HoursBetweenCoursesString_To_Value['I prefer 3 hours between courses on the same day.'] = 3;

  var import_faculty_prefs_data = import_faculty_prefs_datarange.getValues();
  var importedFacultyPrefsList = [];

  for (var rowIdx = ROW_INDEX_FIRST_FACULTY_PREFERENCE; rowIdx < import_faculty_prefs_datarange.getHeight(); rowIdx++) {
    var import_this_entry = import_faculty_prefs_data[rowIdx][COLUMN_INDEX_IMPORT_THIS_ENTRY];
    if (import_this_entry == true) {
      var email = import_faculty_prefs_data[rowIdx][COLUMN_INDEX_FACULTY_EMAIL];
      var courses_on_same_days = MAP_SameDayString_To_Value[import_faculty_prefs_data[rowIdx][COLUMN_INDEX_PREFS_SAME_DAYS]];
      var hours_between_courses = MAP_HoursBetweenCoursesString_To_Value[import_faculty_prefs_data[rowIdx][COLUMN_INDEX_PREFS_HOURS_BETWEEN_COURSES]];
      var timeIntervalCostMap = {};
      var skipIdx = 0; // skip indices corresponding to times are not available for import, e.g., the 11:30-12:45 T/R slot.
      for (var colIdx = COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]; colIdx <= COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[1]; colIdx++) {
        // the code below only works if ROW_INDICES_EXCLUDED_TIME_INTERVALS does not include sequential indices which is OK for now
        // TODO: use this code:
        // while(ROW_INDICES_EXCLUDED_TIME_INTERVALS.includes(colIdx+skipIdx) == true) { skipIdx++; }
        // when the includes() function becomes available for google script (or implement it separately here)
        for (var excludedIdx = 0; excludedIdx < ROW_INDICES_EXCLUDED_TIME_INTERVALS.length; excludedIdx++) {
          if (ROW_INDICES_EXCLUDED_TIME_INTERVALS[excludedIdx] == colIdx + skipIdx - COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]) {
            skipIdx++;
            break;
          }
        }
        if (import_faculty_prefs_data[rowIdx][colIdx] != "") {
          prefStringArray = import_faculty_prefs_data[rowIdx][colIdx].split(","); // if more than one preference was specified
          var avgCost = 0;
          for (var strIdx = 0; strIdx < prefStringArray.length; strIdx++) {
            avgCost += convertTimePreferenceStringToCost(prefStringArray[strIdx].trim());
          }
          avgCost /= prefStringArray.length;
          timeIntervalCostMap[courseTimeList[colIdx + skipIdx - COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]].getId()] = avgCost;
        } else {
          timeIntervalCostMap[courseTimeList[colIdx + skipIdx - COLUMN_INDEX_PREFS_TIME_INTERVAL_COSTS[0]].getId()] = DEFAULT_PREFS_TIME_INTERVAL_COST;
        }
      }
      var importedFacultyPrefs = new FacultyCoursesAndPrefs(undefined, email, undefined, timeIntervalCostMap, courses_on_same_days, hours_between_courses);
      importedFacultyPrefsList.push(importedFacultyPrefs);
    }
  }
  return importedFacultyPrefsList;
}

// retrieve CourseTime time interval data
function getRolledCRNs(crns_datarange) {
  // HARD-CODED CONSTANTS
  var COLUMN_INDEX_USE_THIS_CRN = 0;
  var COLUMN_INDEX_DEPT_CODE = 1;
  var COLUMN_INDEX_COURSE_NUMBER = 2;
  var COLUMN_INDEX_SECTION_NUMBER = 3;
  var COLUMN_INDEX_CRN = 4;
  var ROW_INDEX_FIRST_CRN = 1;

  var crn_data = crns_datarange.getValues();
  var crnList = [];

  for (var rowIdx = ROW_INDEX_FIRST_CRN; rowIdx < crns_datarange.getHeight(); rowIdx++) {
    var use_this_crn = crn_data[rowIdx][COLUMN_INDEX_USE_THIS_CRN];
    if (use_this_crn == true) {
      var dept_code = crn_data[rowIdx][COLUMN_INDEX_DEPT_CODE];
      var course_number = crn_data[rowIdx][COLUMN_INDEX_COURSE_NUMBER];
      var section = crn_data[rowIdx][COLUMN_INDEX_SECTION_NUMBER].toString();
      while (section.length < 3) {
        section = '0' + section;
      }
      var number = crn_data[rowIdx][COLUMN_INDEX_CRN];
      var crn = new CRN(number, dept_code, course_number, section);
      crnList.push(crn);
    }
  }
  return crnList;
}

// identifies rooms and time intervals where nextCourse can be offered under the constraints of the existing scheduledCourseList
function findRoomWithTimeInterval(roomWithTimeIntervalList, nextCourse, scheduledCourseList, facultyCoursesAndPrefsList) {
  // HARD-CODED CONSTANTS
  var COST_AVOID_SIMULTANEOUS_COURSE = +2.0;
  var COST_PREFERRED_SIMULTANEOUS_COURSE = -0.25;
  var TARGET_ROOM_OCCUPANCY_PCT = 0.8;
  var COST_ROOM_OCCUPANCY_PCT_MULTIPLIER = 1.0;
  var COST_ROOM_SIZE_EXCEEDED = 5.0;
  var COST_SLOT_CREDIT_HOUR_MULTIPLIER = 2.0;
  var COST_SLOT_UNOCCUPIED_MULTIPLIER = 2.0;

  var USE_TRAVEL_TIME = true;
  var COST_TRAVEL_TIME = 0.1;
  var MIN_TRAVEL_TIME = 30; // at least MIN_TRAVEL_TIME minutes must be available when classes are back to back in different campus buildings

  // TODO: ARE THESE 3 BOOLEANS STILL NECESSARY? SEEMS LIKE WE ALWAYS WANT TO CONSIDER "ALL" THE DATA
  var USE_FACULTY_TIME_PREFERENCES = true;
  var USE_FACULTY_SAME_TIME_PREFERENCE = true;
  var USE_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE = true;
  // TODO: THE FOLLOWING 2 CONSTANTS APPEAR IN 3 DIFFERENT PLACES IN THE CODE AND MUST BE THE SAME DECLARED VALUES!
  var COST_FACULTY_SAME_DAYS_PREFERENCE = 2.0;
  var COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE = 0.5;
  var ONLINE_ROOM_CONSTANT = new Room('ONLINE', 'INTERNET');

  var candidateRoomWithTimeIntervalList = [];

  var teacherForThisCourse;
  var teacher_scheduledCourseList = [];

  // If instructional mode is either 'Online Synchronous' or 'Online Asynchronous' the required room is 'ONLINE INTERNET'
  if (nextCourse.instructional_mode == 'Online Synchronous' ||
    nextCourse.instructional_mode == 'Online Asynchronous') {
    var hasOnlineRoomRequired = false;
    for (var roomIdx = 0; roomIdx < nextCourse.required_rooms.length; roomIdx++) {
      if (!nextCourse.required_rooms[roomIdx].equals(ONLINE_ROOM_CONSTANT)) {
        //SpreadsheetApp.getUi().alert('Course ' + nextCourse.getId() + ' is set to ONLINE but has required room '+
        //                             nextCourse.required_rooms[roomIdx].getId());  
        //throw Error( "Exiting due to data entry error." );
        Logger.log('ScheduleEngine: WARNING: Course ' + nextCourse.getId() + ' is set to ONLINE but has required room ' + nextCourse.required_rooms[roomIdx].getId() + '.');
        nextCourse.required_rooms.splice(roomIdx, 1);
        roomIdx--;
      } else {
        hasOnlineRoomRequired = true;
      }
    }
    // add the online 'ONLINE INTERNET' room as a required room if it has not already been entered into the 'Course' tab data already
    if (!hasOnlineRoomRequired) {
      nextCourse.required_rooms.push(ONLINE_ROOM_CONSTANT);
    }
  }

  // Find the faculty member teaching this course (nextCourse) store the data in teacherForThisCourse
  if (USE_FACULTY_TIME_PREFERENCES) {
    for (var facultyIdx = 0; teacherForThisCourse == undefined && facultyIdx < facultyCoursesAndPrefsList.length; facultyIdx++) {
      var faculty = facultyCoursesAndPrefsList[facultyIdx];
      for (var courseIdx = 0; courseIdx < faculty.courseList.length; courseIdx++) {
        if (nextCourse.sameNumber(faculty.courseList[courseIdx])) {
          if (nextCourse.sameSection(faculty.courseList[courseIdx])) {
            teacherForThisCourse = faculty;
            break;
          }
        }
      }
    }
  }

  // SAME DAYS Faculty Preference 
  // Find the 3hr courses already scheduled for this faculty member and store them in teacher_scheduledCourseList
  if (USE_FACULTY_TIME_PREFERENCES && USE_FACULTY_SAME_TIME_PREFERENCE) {
    for (var idx = 0; idx < scheduledCourseList.length; idx++) {
      var scheduledCourse = scheduledCourseList[idx];
      if (teacherForThisCourse != undefined && teacherForThisCourse.name != undefined &&
        scheduledCourse.FacultyCoursesAndPrefs != undefined && scheduledCourse.FacultyCoursesAndPrefs.name != undefined &&
        teacherForThisCourse.name == scheduledCourse.FacultyCoursesAndPrefs.name && scheduledCourse.Course.credit_hours == 3) {
        teacher_scheduledCourseList.push(scheduledCourse);
      }
    }
  }

  // search all options 
  for (var idx = 0; idx < roomWithTimeIntervalList.length; idx++) {
    var simultaneousConflictList = [];
    var conflicts = 0;
    var idxIsOK = true;
    var candidateRoomWithTimeInterval = roomWithTimeIntervalList[idx];
    candidateRoomWithTimeInterval.resetCost();
    var cost = candidateRoomWithTimeInterval.cost; // room and time interval costs are the starting cost
    var costDelta = 0;
    var costArr = {};
    if (candidateRoomWithTimeInterval.Room.cost != 0) {
      costArr['room'] = candidateRoomWithTimeInterval.Room.cost;
    }
    if (candidateRoomWithTimeInterval.CourseTime.cost != 0) {
      costArr['time'] = candidateRoomWithTimeInterval.CourseTime.cost;
    }

    // TODO: For multi-hour courses (>1.5 hours on T/R or >1 hour M/W/F), e.g., 3hr courses, we should average cost across all occupied time slots
    // DEBUG
    //if (nextCourse.numbers[0] == 4144) {
    //  var test = true;
    //}

    if (USE_FACULTY_TIME_PREFERENCES && teacherForThisCourse != undefined) {
      costDelta = teacherForThisCourse.timeIntervalCostMap[candidateRoomWithTimeInterval.CourseTime.getId()];
      if (costDelta != undefined) {
        cost += costDelta;
        costArr['faculty_time'] = costDelta;
      }
    }

    // SAME DAYS Faculty Preference 
    // if the faculty is already scheduled for a 3hr time slot
    // check if the same_days preference is true, if so, check the candidateRoomWithTimeInterval.CourseTime
    // if the candidateRoomWithTimeInterval.CourseTime is not on overlapping days add COST_FACULTY_SAME_DAYS_PREFERENCE to cost
    if (USE_FACULTY_TIME_PREFERENCES && USE_FACULTY_SAME_TIME_PREFERENCE) {
      if (teacherForThisCourse != undefined && teacher_scheduledCourseList.length > 0) {
        for (var idx2 = 0; idx2 < teacher_scheduledCourseList.length; idx2++) {
          var numDaysInCommon = candidateRoomWithTimeInterval.CourseTime.daysInCommon(teacher_scheduledCourseList[idx2].CourseTime).length;
          if (numDaysInCommon == 0) {
            if (teacherForThisCourse.courses_on_same_days == true) {  // course does not have overlapping day and faculty prefers overlapping days
              costDelta = COST_FACULTY_SAME_DAYS_PREFERENCE;
              cost += costDelta;
              costArr['same_days'] = costDelta;
              break;
            }
          } else {
            if (teacherForThisCourse.courses_on_same_days == false) {  // course has overlapping days and faculty does not prefer overlapping days
              costDelta = COST_FACULTY_SAME_DAYS_PREFERENCE;
              cost += costDelta;
              costArr['same_days'] = costDelta;
              break;
            }
          }
        }
      }
    }

    // Minimum Time Between Classes Preference 
    // if the faculty is already scheduled for a 3hr time slot, check in the candidateRoomWithTimeInterval.CourseTime
    // same days and time between slots is less than the value sepecified add COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE to cost
    if (USE_FACULTY_TIME_PREFERENCES && USE_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE) {
      if (teacherForThisCourse != undefined && teacher_scheduledCourseList.length > 0) {
        for (var idx2 = 0; idx2 < teacher_scheduledCourseList.length; idx2++) {
          var numDaysInCommon = candidateRoomWithTimeInterval.CourseTime.daysInCommon(teacher_scheduledCourseList[idx2].CourseTime).length;
          if (numDaysInCommon > 0) {
            var hoursBetweenCourses = candidateRoomWithTimeInterval.CourseTime.hoursBetween(teacher_scheduledCourseList[idx2].CourseTime);
            //var err = Math.abs(teacherForThisCourse.hours_between_courses - hoursBetweenCourses);
            if (hoursBetweenCourses < teacherForThisCourse.hours_between_courses) {
              costDelta = COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE;
              cost += costDelta;
              costArr['hrs_between_courses(too small)'] = costDelta;
              break;
            } //else if (hoursBetweenCourses > teacherForThisCourse.hours_between_courses) {
            //costDelta = 0.25*(hoursBetweenCourses - teacherForThisCourse.hours_between_courses)*COST_FACULTY_TIME_BETWEEN_CLASSES_PREFERENCE;
            //cost += costDelta;
            //costArr['hrs_between_courses(too large)'] = costDelta;
            //break;
            //}
          }
        }
      }
    }

    // Building-to-Building Travel Time Between Classes Constraint
    // If the faculty is already scheduled for classes, check the candidateRoomWithTimeInterval.CourseTime to see if teaching is on the
    // same days and, if the classes are in different *REAL* buildings, we add COST_TRAVEL_TIME cost if this slot starts or ends less then MIN_TRAVEL_TIME (30 mins)
    // before or after the teacher's existing class.
    if (teacherForThisCourse != undefined && teacher_scheduledCourseList.length > 0) {
      for (var idx2 = 0; idx2 < teacher_scheduledCourseList.length; idx2++) {
        var numDaysInCommon = candidateRoomWithTimeInterval.CourseTime.daysInCommon(teacher_scheduledCourseList[idx2].CourseTime).length;
        if (numDaysInCommon > 0) {
          if (candidateRoomWithTimeInterval.Room.building != teacher_scheduledCourseList[idx2].Room.building) {
            if (candidateRoomWithTimeInterval.Room.building != 'NONE' && teacher_scheduledCourseList[idx2].Room.building != 'NONE') {
              var minutesBetweenCourses = candidateRoomWithTimeInterval.CourseTime.hoursBetween(teacher_scheduledCourseList[idx2].CourseTime) * 60;
              //var err = Math.abs(teacherForThisCourse.hours_between_courses - hoursBetweenCourses);
              if (minutesBetweenCourses < teacherForThisCourse.hours_between_courses) {
                costDelta = COST_TRAVEL_TIME;
                cost += costDelta;
                costArr['travel_time'] = costDelta;
                break;
              }
            }
          }
        }
      }
    }

    // check if the course time available includes the term or part-term the course will be offered
    if (idxIsOK) {
      idxIsOK = false;
      if (nextCourse.term == candidateRoomWithTimeInterval.CourseTime.term ||
        (candidateRoomWithTimeInterval.CourseTime.isFullTerm() && !nextCourse.isFullTerm())) {
        idxIsOK = true;
      }
    }

    // check if room includes only a specific list of courses, if so, make sure this course is OK or skip this candidateRoomWithTimeInterval
    if (idxIsOK) {
      if (candidateRoomWithTimeInterval.CourseTime.include_course_list.length > 0) {
        idxIsOK = false;
        var include_course_list = candidateRoomWithTimeInterval.CourseTime.include_course_list;
        for (var include_course_idx = 0; include_course_idx < include_course_list.length; include_course_idx++) {
          if (nextCourse.sameNumber(include_course_list[include_course_idx])) {
            idxIsOK = true;
            break; // this candidateRoomWithTimeInterval can include this course
          }
        }
      }
    }

    // check if room excludes a list of specific courses, if so, make sure this course is OK or skip this candidateRoomWithTimeInterval
    if (idxIsOK) {
      // by checking this after the Room.include_course_list the exclude course list has precedence when a course is both included and excluded from a room
      if (candidateRoomWithTimeInterval.CourseTime.exclude_course_list.length > 0) {
        var exclude_course_list = candidateRoomWithTimeInterval.CourseTime.exclude_course_list;
        for (var exclude_course_idx = 0; exclude_course_idx < exclude_course_list.length; exclude_course_idx++) {
          if (nextCourse.sameNumber(exclude_course_list[exclude_course_idx])) {
            idxIsOK = false;
            break; // this candidateRoomWithTimeInterval cannot include this course
          }
        }
      }
    }

    // check if room expected enrollment < room.maximum_capacity, if not skip this room and time interval
    if (idxIsOK) {
      // CLASSES SCHEDULED AS HYBRID CAPACITY HAVE A LARGER CAPACITY
      if (nextCourse.isHybrid() && nextCourse.isSynchronous() &&
        nextCourse.expected_enrollment <= candidateRoomWithTimeInterval.Room.maximum_capacity * candidateRoomWithTimeInterval.CourseTime.days.length) {
        // add cost on the percent difference of target occupancy percent and occupancy percent if expected number of students enroll
        costDelta = COST_ROOM_OCCUPANCY_PCT_MULTIPLIER * (Math.abs(TARGET_ROOM_OCCUPANCY_PCT - (nextCourse.expected_enrollment / candidateRoomWithTimeInterval.Room.maximum_capacity)));
        cost += costDelta;
        costArr['enrollment'] = costDelta;
      } else if (nextCourse.expected_enrollment <= candidateRoomWithTimeInterval.Room.maximum_capacity) {
        // add cost on the percent difference of target occupancy percent and occupancy percent if expected number of students enroll
        costDelta = COST_ROOM_OCCUPANCY_PCT_MULTIPLIER * (Math.abs(TARGET_ROOM_OCCUPANCY_PCT - (nextCourse.expected_enrollment / candidateRoomWithTimeInterval.Room.maximum_capacity)));
        cost += costDelta;
        costArr['enrollment'] = costDelta;
      } else {
        idxIsOK = false;
      }
    }

    // check if course hours match with time interval
    if (idxIsOK) {
      var slotTotalCreditHours = candidateRoomWithTimeInterval.CourseTime.credit_hours_per_day * candidateRoomWithTimeInterval.CourseTime.days.length;
      if (slotTotalCreditHours >= nextCourse.credit_hours) {
        // penalize by adding cost for excess credit hours of scheduled room time
        costDelta = COST_SLOT_CREDIT_HOUR_MULTIPLIER * (slotTotalCreditHours - nextCourse.credit_hours);
        // If this is not a perfect fit a split may occur. Prefer: Split 1 hr/day courses for a 1 hr section and 1.5 hr/day slots for a 1.5 hr section
        if (costDelta > 0) {
          var days_needed = nextCourse.credit_hours / candidateRoomWithTimeInterval.CourseTime.credit_hours_per_day;
          var hrs_unoccupied = COST_SLOT_UNOCCUPIED_MULTIPLIER * Math.ceil(days_needed) - days_needed;
          costDelta = hrs_unoccupied;
          costArr['hrs_unoccupied'] = costDelta;
        }
        cost += costDelta;
        if (costDelta != 0) {
          costArr['credit_hours'] = costDelta;
        }
      } else {
        idxIsOK = false;
      }
    }

    // check if this course requires a specific list of rooms. if so, check for room list to ensure this candidateRoomWithTimeInterval is OK
    for (var reqRoomIdx = 0; idxIsOK && reqRoomIdx < nextCourse.required_rooms.length; reqRoomIdx++) {
      idxIsOK = false;
      //if (candidateRoomWithTimeInterval.Room.getId() == nextCourse.required_room_ids[reqRoomIdx]) {
      if (candidateRoomWithTimeInterval.Room.equals(nextCourse.required_rooms[reqRoomIdx])) {
        idxIsOK = true;
        break; // this is one of the required rooms!
      }
    }

    // check if this course excludes a specific list of rooms. if so, check for room list to ensure this candidateRoomWithTimeInterval is OK
    for (var excludeRoomIdx = 0; idxIsOK && excludeRoomIdx < nextCourse.excluded_rooms.length; excludeRoomIdx++) {
      //if (candidateRoomWithTimeInterval.Room.getId() == nextCourse.excluded_room_ids[excludeRoomIdx]) {
      if (candidateRoomWithTimeInterval.Room.equals(nextCourse.excluded_rooms[excludeRoomIdx])) {
        idxIsOK = false;
        break;  // this is one of the excluded rooms!
      }
    }

    // allows us to keep cost incentive for simultaneous and preferred classes limited to only one preference counted per course
    var foundSimultaneousPreference = false;
    // get courses in this timeInterval from scheduledCourseList
    for (var idx2 = 0; idxIsOK && idx2 < scheduledCourseList.length; idx2++) {
      // penalize cost for timeslots having this course in the not_simultaneous_courses list
      // reward cost for timeslots having this course in the simultaneous_courses list
      var query = scheduledCourseList[idx2];
      // DEBUG
      //if (query.Course.numbers[0] != 2156 && nextCourse.numbers[0] == 2156 && nextCourse.section == "L02") {
      //  var test = true;
      //}

      // detect offering of an overlapping time slot on overlapping days with the same instructor, if so, skip this room and time interval
      if (query.CourseTime.overlaps(candidateRoomWithTimeInterval.CourseTime) && query.CourseTime.daysInCommon(candidateRoomWithTimeInterval.CourseTime).length > 0 &&
        teacherForThisCourse != undefined && teacherForThisCourse.name != undefined && query.FacultyCoursesAndPrefs != undefined && query.FacultyCoursesAndPrefs.name != undefined &&
        teacherForThisCourse.name == query.FacultyCoursesAndPrefs.name) {
        idxIsOK = false;
      }

      // detect offering of the same course in the same time slot on overlapping days, if so, skip this room and time interval
      if (query.Course.sameNumber(nextCourse) && query.CourseTime.equals(candidateRoomWithTimeInterval.CourseTime) && query.CourseTime.daysInCommon(candidateRoomWithTimeInterval.CourseTime).length > 0) {
        idxIsOK = false;
      }

      //DEBUG
      //  if (query.Course.numbers[0] == 2111 && nextCourse.numbers[0] == 2155) {
      //    var test = true;
      //  }

      // detect courses held simultaneously for this time slot, if it is simultaneous, determine if this course has "not simultaneous" or "simultaneous" preference
      if (idxIsOK && query.CourseTime.overlaps(candidateRoomWithTimeInterval.CourseTime) && query.CourseTime.daysInCommon(candidateRoomWithTimeInterval.CourseTime).length > 0) {

        // detect if course has a "not simultaneous" preference when paired with the queried course and time slot, if so, skip this room and time interval
        if (nextCourse.notSimultaneousConflict(query.Course)) {
          //Logger.log('ScheduleEngine: Found simultaneous slot conflict: ' + query.Course.getId() + " in " + query.Room.getId() + " at " + query.CourseTime.getId());    
          simultaneousConflictList.push(query);  // conflict with scheduled course
          conflicts++;
          costDelta = COST_AVOID_SIMULTANEOUS_COURSE;
          cost += costDelta;
          costArr['not_simultaneous'] = costDelta;
          idxIsOK = false;
        }

        // SIMULTANEOUS HERE REQUIRES OVERLAPPING DAYS AND TIME INTERVALS USED FOR MULTI-DAY COURSES (not 3 hr courses) 
        // detect if course has a "simultaneous" preference when paired with the queried course and time slot, if so, lower the cost of this time slot
        if (idxIsOK && nextCourse.simultaneousPreference(query.Course) && !foundSimultaneousPreference) {
          foundSimultaneousPreference = true;
          costDelta = COST_PREFERRED_SIMULTANEOUS_COURSE;
          cost += costDelta; // found preferred simultaneous course time interval
          costArr['simultaneous'] = costDelta;
        }
      }
    }

    // if idxIsOK == true this is a valid time slot and cost is already computed
    if (idxIsOK == true) {
      candidateRoomWithTimeInterval.cost = cost;
      candidateRoomWithTimeInterval.costArr = costArr;
      candidateRoomWithTimeInterval.index = idx;
      candidateRoomWithTimeInterval.instructor = teacherForThisCourse;
      candidateRoomWithTimeIntervalList.push(candidateRoomWithTimeInterval);
    } // consideration of this candidateRoomWithTimeInterval is complete

  } // search complete and a candidateRoomWithTimeIntervalList has been constructed

  // sort the candidateRoomWithTimeIntervalList by increasing cost
  candidateRoomWithTimeIntervalList.sort(function (a, b) { return a.cost - b.cost; });

  // DEBUG
  //if (nextCourse.equals(new Course(DEFAULT_DEPT_CODE,"4290"))) {
  //  var test = 1;
  //}

  // choose topmost, i.e., "best," candidateRoomWithTimeInterval from the list (lowest cost with tie-breaking by natural ordering)
  var scheduledRoomWithTimeInterval = candidateRoomWithTimeIntervalList[0];
  return scheduledRoomWithTimeInterval;
}

// splits this collection of times into two time intervals having disjoint days, e.g., "M/W/F" -> ["M/W","F"] at a location determined by excessDays
function splitCourseTimeByDay(course_time, excessDays, occupiedCourseTime) {
  if (occupiedCourseTime != undefined) {
    // if occupiedCourseTime is not undefined the task is to remove the specific days scheduled (occupiedCourseTime.days) and place remaining times in a newly generated CourseTime object
    var unchangedDaysArray = course_time.days;
    var setSubtractedDaysArray = occupiedCourseTime.days.slice(0); // clone the array of days do not modify the passed object
    for (var dayIdx = 0; dayIdx < unchangedDaysArray.length; dayIdx++) { // maybe there's a removeAll(unchangedDaysArray)
      setSubtractedDaysArray.splice(setSubtractedDaysArray.indexOf(unchangedDaysArray[dayIdx]), 1);
    }
    return new CourseTime(course_time.start, course_time.end, setSubtractedDaysArray.join('/'), course_time.credit_hours_per_day,
      course_time.exclude_course_list, course_time.include_course_list, course_time.cost);
  } else {
    // if occupiedCourseTime is undefined the program decides the specific days to remove when splitting a course time. This uses default processing to split times into separate days.
    // Default processing creates a new roomWithTimeInterval with the required scheduled days removed. We remove days from the begining, e.g., M/W/F, has M removed first then W and so on.
    var days_str_new = "";
    var daysArray = course_time.days;
    for (var dayIdx = 1; dayIdx <= excessDays; dayIdx++) {
      days_str_new += ((dayIdx + 1 > excessDays) ? "" : "/") + daysArray[daysArray.length - 1];
      daysArray.splice(daysArray.length - 1, 1);
    }
    return new CourseTime(course_time.start, course_time.end, days_str_new, course_time.credit_hours_per_day,
      course_time.exclude_course_list, course_time.include_course_list, course_time.cost);
  }
}

var instructionalModeToString = {
  'Face to Face': 'F2F', 'Hybrid Synchronous': 'HYS', 'Hybrid Asynchronous': 'HYA', 'Online Synchronous': 'IS', 'Online Asynchronous': 'IA'
};

// Adds the passed course, and roomWithTimeInterval to the schedule
function addCourseToSchedule(output_sheet_data, newScheduledCourse) {
  // HARD-CODED CONSTANTS
  var OUTPUT_SHEET_ROOM_INDEX_ROW_OFFSET = 1;
  var OUTPUT_SHEET_COST_INDEX_COLUMN_OFFSET = (/[M|W|F]/.exec(newScheduledCourse.CourseTime.days) != null) ? 6 : 10;

  // convert the U,M,T,W,R,F,S day strings to output sheet column indices 
  var days_of_week_colIdxs = convertDaysOfWeekToSlotColumnIndices(newScheduledCourse.CourseTime.days);
  var rooms_colIdxs = convertDaysOfWeekToRoomColumnIndices(newScheduledCourse.CourseTime.days);
  if (newScheduledCourse.CourseTime.durationMinutes() < 5) {
    Logger.log('Error: Cannot add course ' + newScheduledCourse.Course.getId() + ' to the schedule at time ' + newScheduledCourse.CourseTime.getId() + ' it is shorter than 5 minutes in length.');
    return;
  }
  var timeinterval_rowRange = convertTimeIntervalToRowRange(newScheduledCourse.CourseTime.start, newScheduledCourse.CourseTime.end);
  timeinterval_rowRange[1] += 2; // two slots for the 10 minutes between class periods on output sheet before next time interval
  var output_cost_column_index = OUTPUT_SHEET_COST_INDEX_COLUMN_OFFSET;
  var hoursScheduled = 0;
  var roomRowIdxs = [];
  // this checking on credit hours isn't totally necessary and could be removed without harm 
  // this actually may be needed for when a course with a small (say 1) credit hour need schedules into a slot having more (say 3) credit hours
  var total_credit_hours = (newScheduledCourse.Course.credit_hours == undefined) ? 10 : newScheduledCourse.Course.credit_hours;

  for (var slotColIdx = 0; hoursScheduled < total_credit_hours && slotColIdx < days_of_week_colIdxs.length; slotColIdx++) {
    if (slotColIdx == 0) {
      roomRowIdxs = findRoomRowsInOutputSheet(output_sheet_data, newScheduledCourse.Room.getId(), rooms_colIdxs[slotColIdx],
        timeinterval_rowRange[0] + OUTPUT_SHEET_ROOM_INDEX_ROW_OFFSET, timeinterval_rowRange[1] + OUTPUT_SHEET_ROOM_INDEX_ROW_OFFSET);
    }
    for (var slotRowIdx = 0; slotRowIdx < roomRowIdxs.length; slotRowIdx++) {
      var courseToRender = newScheduledCourse;
      var instructionalModeAcronym = instructionalModeToString[courseToRender.Course.instructional_mode];
      var facultyName = "";
      var facultyprefCost = "";
      if (courseToRender.FacultyCoursesAndPrefs != undefined) {
        var facultyNameList = courseToRender.FacultyCoursesAndPrefs.name.split(" ");
        facultyName = " " + facultyNameList[0].replace(/,\s*$/, "") + " ";
        facultyprefCost = (courseToRender.costArray['faculty_time'] == undefined) ? 'X' : courseToRender.costArray['faculty_time'];
      }
      if (instructionalModeAcronym == undefined) {
        instructionalModeAcronym = '?';
      }
      if (facultyprefCost == undefined) {
        facultyprefCost = '?';
      }
      var slotString = courseToRender.Course.getId() + " " + instructionalModeAcronym + facultyName + facultyprefCost;
      if (slotRowIdx > 0) {
        slotString = '*' + slotString;
      }
      var curValue = output_sheet_data.getCell(roomRowIdxs[slotRowIdx], days_of_week_colIdxs[slotColIdx]).getValue();
      if (curValue != "") {
        output_sheet_data.getCell(roomRowIdxs[slotRowIdx], days_of_week_colIdxs[slotColIdx]).setValue(curValue + '\n' + slotString);
      } else {
        output_sheet_data.getCell(roomRowIdxs[slotRowIdx], days_of_week_colIdxs[slotColIdx]).setValue(slotString);
      }
    }
    hoursScheduled += newScheduledCourse.CourseTime.credit_hours_per_day;
  }
  if (roomRowIdxs.length > 0) {
    var curCost = output_sheet_data.getCell(roomRowIdxs[0], output_cost_column_index).getValue();
    if (curCost != "") {
      output_sheet_data.getCell(roomRowIdxs[0], output_cost_column_index).setValue(curCost + ', ' + newScheduledCourse.cost.toFixed(2));
    } else {
      output_sheet_data.getCell(roomRowIdxs[0], output_cost_column_index).setValue(newScheduledCourse.cost.toFixed(2));
    }
  } else {
    Logger.log('Error: Could not add ' + newScheduledCourse.Course.getId() + ' to the schedule in room ' + newScheduledCourse.Room.getId() + ' at time ' + newScheduledCourse.CourseTime.getId() + '.');
  }
}

// Input room_id - starting (row,col) index search proceeds over rows
// returns the row index of the room in the output sheet
function findRoomRowsInOutputSheet(output_sheet_data, room_id, columnIdx, start_rowIdx, end_rowIdx) {
  var roomRowIdxs = [];
  for (var qrowIdx = start_rowIdx; qrowIdx < end_rowIdx; qrowIdx++) {
    var cellValue = output_sheet_data.getCell(qrowIdx, columnIdx).getValue();
    if (cellValue === room_id) { // careful for whitespace!
      roomRowIdxs.push(qrowIdx);
    }
  }
  return roomRowIdxs;
}

// department code is valid if it is "ENGR" or DEFAULT_DEPT_CODE
function isValidDeptCode(cellValue) {
  // HARD-CODED CONSTANTS
  if (cellValue == "ENGR" || cellValue == DEFAULT_DEPT_CODE) {
    return true;
  }
  return false;
}

function FormatNumberLength(num, length) {
  var r = "" + num;
  while (r.length < length) {
    r = "0" + r;
  }
  return r;
}

// number is valid if it is a number X satisfying  1000  <= X < 9000
function isValidCourseNumber(cellValue) {
  return true;
}

function convertTimeIntervalToRowRange(start_time, end_time) {
  var rowIndices = [];
  rowIndices.push(convertTimeToRowIndex(start_time));
  rowIndices.push(convertTimeToRowIndex(end_time));
  return rowIndices;
}

function convertTimeToRowIndex(timeval) {

  // HARD-CODED CONSTANTS --> TODO FIX
  var MINS_PER_ROW = 5
  var START_ROW = 3
  // 8:00 AM START TIME FOR CLASSES
  var DAY_START_TIME = new Date(DATE_CONSTANT_YEAR, DATE_CONSTANT_MONTH, DATE_CONSTANT_DAY,
    DATE_CONSTANT_FIRST_HOUR, DATE_CONSTANT_FIRST_MINUTE, DATE_CONSTANT_SEC, DATE_CONSTANT_MSEC);
  // COMPUTE START AND END TIME DIFFERENCE IN MINUTES --> CONVERT TO ROW INDICES
  var query_time = new Date(DATE_CONSTANT_YEAR, DATE_CONSTANT_MONTH, DATE_CONSTANT_DAY,
    timeval.getHours(), timeval.getMinutes(), DATE_CONSTANT_SEC, DATE_CONSTANT_MSEC);
  var time_diff_ms = (query_time - DAY_START_TIME); // Difference in milliseconds.
  var rowIndex = START_ROW + (time_diff_ms / ((1000 * 60) * MINS_PER_ROW));
  return rowIndex;
}

// column index of course offering for this day
function convertDaysOfWeekToSlotColumnIndices(day_string_array) {
  //var day_string_array = days_of_week.split("/");
  var columnIdxs = [];
  for (var i = 0; i < day_string_array.length; i++) {
    var idx = convertDayOfWeekToSlotColumnIndex(day_string_array[i]);
    if (idx > 0) {
      columnIdxs.push(idx);
    }
  }
  return columnIdxs;
}

// column index of rooms for a day 
function convertDaysOfWeekToRoomColumnIndices(day_string_array) {
  //var day_string_array = days_of_week.split("/");
  var columnIdxs = [];
  for (var i = 0; i < day_string_array.length; i++) {
    var idx = convertDayOfWeekToRoomColumnIndex(day_string_array[i]);
    if (idx > 0) {
      columnIdxs.push(idx);
    }
  }
  return columnIdxs;
}

// Maps days strings to slot column indices in output sheet
// HARD-CODED CONSTANTS
var MAP_DayOfWeek_To_OutputColumnIndex = {};
MAP_DayOfWeek_To_OutputColumnIndex['M'] = 3;
MAP_DayOfWeek_To_OutputColumnIndex['T'] = 8;
MAP_DayOfWeek_To_OutputColumnIndex['W'] = 4;
MAP_DayOfWeek_To_OutputColumnIndex['R'] = 9;
MAP_DayOfWeek_To_OutputColumnIndex['F'] = 5;
MAP_DayOfWeek_To_OutputColumnIndex['S'] = -1;
MAP_DayOfWeek_To_OutputColumnIndex['U'] = -1;
function convertDayOfWeekToSlotColumnIndex(day_string) {
  return MAP_DayOfWeek_To_OutputColumnIndex[day_string];
}

// Maps days strings to associated time/room column in output sheet
// HARD-CODED CONSTANTS
var MAP_DayOfWeek_To_RoomColumnIndex = {};
MAP_DayOfWeek_To_RoomColumnIndex['M'] = 2;
MAP_DayOfWeek_To_RoomColumnIndex['T'] = 7;
MAP_DayOfWeek_To_RoomColumnIndex['W'] = 2;
MAP_DayOfWeek_To_RoomColumnIndex['R'] = 7;
MAP_DayOfWeek_To_RoomColumnIndex['F'] = 2;
MAP_DayOfWeek_To_RoomColumnIndex['S'] = -1;
MAP_DayOfWeek_To_RoomColumnIndex['U'] = -1;
function convertDayOfWeekToRoomColumnIndex(day_string) {
  return MAP_DayOfWeek_To_RoomColumnIndex[day_string];
}

// Maps faculty preference time-slot preference words to associated costs for automated scheduling
// HARD-CODED CONSTANTS
var MAP_TimePreferenceString_To_Cost = {};
MAP_TimePreferenceString_To_Cost['Best'] = 0.0;
MAP_TimePreferenceString_To_Cost['Acceptable'] = 1.0;
MAP_TimePreferenceString_To_Cost['Not Preferred'] = 2.0;
MAP_TimePreferenceString_To_Cost['Not Possible'] = 3.0;
function convertTimePreferenceStringToCost(preference_string) {
  return MAP_TimePreferenceString_To_Cost[preference_string];
}
