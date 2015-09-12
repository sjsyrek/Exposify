/**
 * Exposify
 *
 * Copyright 2015 Steven J. Syrek
 * https://github.com/sjsyrek/Exposify
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * @fileoverview Exposify is a Google Sheets add-on that automates a variety
 * of tasks related to the teaching of expository writing courses. Key features
 * include automatic setup of grade books, attendance records, and folder
 * hierarchies in Google Drive for organizing course sections and the return of
 * graded assignments; batch word counts of student assignments; differential
 * comparison of paper drafts (e.g. rough versus final); and various formatting
 * and administrative tasks.
 * @author steven.syrek@gmail.com (Steven Syrek)
 */

/**
 * NOTE: Exposify will not work out-of-the-box if you just copy and paste the code
 * into a Google Scripts Editor. You will need to activate the Drive API in both
 * Resources : Advanced Google services and the Developers Console for the project.
 * If you have not already set up a project in the Developers Console, you will
 * need to create one, associate it with the script project, and, in the Script
 * Editor, create a Script Property called DEVELOPER_KEY with your API key as the
 * value. In addition, you will need to enable the Google Picker API in the
 * Developers Console in order for the file picking user interface to function.
 */

//TODO: publish Exposify to GitHub
//TODO: use sheet.appendRow() instead of getLastRow() where possible
//TODO: add end comment to every function
//TODO: add all functions to Exposify.prototype
//TODO: make sure every function has error checking blocks
//TODO: make sure all alerts actually call alert();
//TODO: generalize word counts so user can enter a custom value (default 1700 is ok)
//TODO: folder structure, folder sharing, collecting and returning assignments
//TODO: finish paper comparison diff function
//TODO: automatically create templates for Docs
//TODO: automatically produce warning rosters and final gradebooks
//TODO: autotmatically add students to Contacts
//TODO: make error messages more informative
//TODO: make sure error logging is correct format
//TODO: fix JSDoc param comments
//TODO: make it possible to specify the maximum number of students in a course
//TODO: make sure internal functions that reference this actually work
//TODO: add titles to alerts


/**
 * Create an interface to the Exposify framework without polluting the global namespace,
 * in the event other scripts are attached to this spreadsheet or Exposify's functionality
 * is extended. I don't know if this is completely necessary, but it seemed like a worthy
 * practice to follow for programming in JavaScript.
 */
(function() {
  var expos = new Exposify();
  this.expos = expos;
})(); // end self-executing anonymous function


// CONSTANTS


var EMAIL_DOMAIN = '@scarletmail.rutgers.edu'; // default email domain for students
var STYLESHEET = 'Stylesheet.html'; // can't include a css stylesheet, so we put styles here and concatenate the html pages later
var TIMEZONE = 'America/New_York'; // default timezone
var MAX_STUDENTS = 22; // maximum number of students in any course (not a good idea to change this)

var ALERT_TITLE_DEFAULT = 'Exposify'; // UI constant aliases
var OK = 'ok';
var OK_CANCEL = 'okCancel';
var YES_NO = 'yesNo';
var PROMPT = 'prompt';

var GRADED_PAPERS_FOLDER_NAME = 'Graded Papers'; // name of the folder for Graded Papers
var ATTENDANCE_SHEET_COLUMN_WIDTH = 25; // width of columns in the attendance record part of the gradebook, 25 is the minimum recommended if you want all the dates to be visible
var COLOR_BLANK = '#ffffff'; // #ffffff is white
var COLOR_SHADED = '#ededed'; // #ededed is light grey, a nice color for contrast and also a pun on the purpose of this application
var FONT = 'verdana,sans,sans-serif'; // font for the gradebook, with fallbacks

var DAYS = { // an enum for days of the week; do not ever change these values or the whole thing will blow up
  'Sunday': 0,
  'Monday': 1,
  'Tuesday': 2,
  'Wednesday': 3,
  'Thursday': 4,
  'Friday': 5,
  'Saturday': 6
};

var EMAIL_REGEX = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/; // a regular expression for validating email addresses

var MIME_TYPE_CSV = 'text/csv'; // file MIME types for reading file data
var MIME_TYPE_GOOGLE_SHEET = 'application/vnd.google-apps.spreadsheet';

/**
 * Default text to display for various alert messages throughout the application.
 */
var ALERT_INSTALL_THANKS = 'Thanks for installing Exposify! Add a new section by selecting \"Setup New Expos Section\" in the Exposify menu.';
var ALERT_SETUP_ADD_STUDENTS_SUCCESS = '\'$\' successfully imported! You should double-check the spreadsheet to make sure it\'s correct.';
var ALERT_SETUP_CREATE_CONTACTS_SUCCESS = 'New contact group successfully created for $.';
var ALERT_SETUP_CREATE_FOLDER_STRUCTURE = 'This command will create or update the folder structure for your Expos section, including a shared coursework folder and individual folders for each student based on this sheet. Do you wish to proceed?';
var ALERT_SETUP_CREATE_FOLDER_STRUCTURE_NOTHING_NEW = 'Nothing to update.';
var ALERT_SETUP_SHARE_FOLDERS = 'This will share the course folder with all students and each student folder with that student, respectively. Do you wish to proceed?';
var ALERT_SETUP_SHARE_FOLDERS_MISSING_COURSE_FOLDER = 'There is no course folder for this course. Use the Create Folder Structure command to create one before executing this command.';
var ALERT_SETUP_SHARE_FOLDERS_MISSING_GRADED_FOLDER = 'There is no folder for graded papers for this course. Use the Create Folder Structure command to create one before executing this command.';
var ALERT_SETUP_SHARE_FOLDERS_SUCCESS = 'The folders were successfully shared!';
var ALERT_SETUP_NEW_GRADEBOOK_ALREADY_EXISTS = 'A gradebook for section $ already exists. If you want to overwrite it, make it the active spreadsheet and try again.';
var ALERT_SETUP_NEW_GRADEBOOK_SUCCESS = 'New gradebook created for $.';
var ALERT_NO_GRADEBOOK = 'You have not set up a gradebook yet for this sheet. Do that before anything else.';

var TOAST_DISPLAY_TIME = 10; // how long should the little toast window linger before disappearing
var TOAST_TITLE = 'Success!' // toast window title

var ERROR_INSTALL = 'There was a problem with installation. Please try again.';
var ERROR_FORMAT_SET_SHADED_ROWS = 'There was a problem formatting the sheet. Please try again.';
var ERROR_FORMAT_SWITCH_STUDENT_NAMES = 'There was a problem formatting the sheet. Please try again.';
var ERROR_SETUP_NEW_GRADEBOOK_FORMAT = 'There was a problem formatting the page. Try again.';
var ERROR_SETUP_ADD_STUDENTS = 'There was a problem reading the file.';
var ERROR_SETUP_ADD_STUDENTS_EMPTY = 'I could not find any students in the file \"$\". Make sure you didn\'t modify it after downloading it from Sakai.';
var ERROR_SETUP_ADD_STUDENTS_INVALID = '\'$\' is not a valid CSV or Google Sheets file. Please try again.';
var ERROR_SETUP_CREATE_FOLDER_STRUCTURE = 'There was a problem creating the folder structure. Please try again.';
var ERROR_SETUP_SHARE_FOLDERS = 'There was a problem sharing the folders. Please try again.';
var ERROR_ASSIGNMENTS_CALC_WORD_COUNTS = 'There is no course folder for this course. Use the Create Folder Structure command to create one before executing this command.';

/* The COURSE_FORMATS object literal contains the basic data used to format Exposify gradebooks,
 * depending on the course selected. Altering these could have unpredictable effects on the application,
 * though new course formats can be added (use the 'O' object, for 'Other' courses, as a model). The template
 * is used to format the height of rows, the width of columns, to set the course name and column headings,
 * and to apply data validations to the sheet in order to enforce the usage of specific grade entries.
 * There are two types of grade validation: numeric and non-numeric. Numeric validations require values
 * between 0–100. Non-numeric grades require values supplied by a list of possible grades. There is also
 * help text that will appear when a user hovers over a cell with a grade validation applied to it.
 */
var COURSE_FORMATS = {
  '0': {
    name: 'Course Name', // name of the course
    rows: [30, 30, 40, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20], // heights of the first 25 rows
    columns: [215, 85, 55, 55, 55, 55, 55, 55, 55, 55, 55, 55, 55], // widths of the first 13 columns
    columnHeadings: ['Student Name', 'Student ID', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'Final Grade'] // column headings, of which there should be the same number as the length of the column property
  },
  '101': {
    name: 'Expository Writing',
    rows: [30, 30, 40, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20],
    columns: [215, 85, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 50],
    columnHeadings: ['Student Name', 'Student ID', 'RD 1', 'FD1 (L)', 'FD1 (I)', 'FD1 Grade', 'RD 2', 'FD2 (L)', 'FD2 (I)', 'FD2 Grade', 'MT', 'RD 3', 'FD3 (L)', 'FD3 (I)', 'FD3 Grade', 'RD 4', 'FD4 (L)', 'FD4 (I)', 'FD 4 Grade', 'RD 5', 'FD5 (L)', 'FD5 (I)', 'FD5 Grade', 'FE', 'Final Grade'],
    gradeValidations: {
      paperGrades: {
        requiredValues: ['A', 'B+', 'B', 'C+', 'C', 'NP'],
        helpText: 'Enter A, B+, B, C+, C, or NP',
        rangeToValidate: ['F4:F25', 'J4:J25', 'O4:O25', 'S4:S25', 'W4:W25']
      },
      examGrades: {
        requiredValues: ['P', 'NP'],
        helpText: 'Enter P or NP',
        rangeToValidate: ['K4:K25', 'X4:X25']
      },
      finalGrades: {
        requiredValues: ['A', 'B+', 'B', 'C+', 'C', 'NC', 'F', 'TF', 'TZ'],
        helpText: 'Enter A, B+, B, C+, C, NC, F, TF, or TZ',
        rangeToValidate: ['Y4:Y25']
      },
      roughDraftStatus: {
        requiredValues: ['X'],
        helpText: 'Enter X if this assignment is complete',
        rangeToValidate: ['C4:C25', 'G4:G25', 'L4:L25', 'P4:P25', 'T4:T25']
      },
      lateFinalStatus: {
        requiredValues: ['L'],
        helpText: 'Enter L if this assignment is late',
        rangeToValidate: ['D4:D25', 'H4:H25', 'M4:M25', 'Q4:Q25', 'U4:U25']
      },
      incompleteFinalStatus: {
        requiredValues: ['I'],
        helpText: 'Enter I if this assignment is incomplete',
        rangeToValidate: ['E4:E25', 'I4:I25', 'N4:N25', 'R4:R25', 'V4:V25']
      },
      /**
							* Package and return grade validation data as an object with one field for the non-numeric
							* data validations specified for this course. I could probably generalize this.
							* @return {Object}
							* @return {Array} Object.nonNumeric - The non-numeric grade validations.
							*/
      getGradeValidations: function() {
        var nonNumeric = [this.paperGrades, this.examGrades, this.finalGrades, this.roughDraftStatus, this.lateFinalStatus, this.incompleteFinalStatus];
        return {nonNumeric: nonNumeric}; // package and return validation data
      }
    }
  },
  '103': {
    name: 'Exposition & Argument',
    rows: [30, 30, 40, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20],
    columns: [215, 85, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 45, 55, 45],
    columnHeadings: ['Student Name', 'Student ID', 'RD 1', 'FD1 (L)', 'FD1 (I)', 'FD1 Grade', 'RD 2', 'FD2 (L)', 'FD2 (I)', 'FD2 Grade', 'RD 3', 'FD3 (L)', 'FD3 (I)', 'FD3 Grade', 'PR', 'RP RD1', 'RP RD2', 'RP Grade', 'RP (L)', 'RP (I)', 'Partici-\npation', 'Number Grade', 'Final Grade'],
    gradeValidations: {
      roughDraftStatus: {
        requiredValues: ['X'],
        helpText: 'Enter X if this assignment is complete',
        rangeToValidate: ['C4:C25', 'G4:G25', 'K4:K25', 'P4:P25', 'Q4:Q25']
      },
      lateFinalStatus: {
        requiredValues: ['L'],
        helpText: 'Enter L if this assignment is late',
        rangeToValidate: ['D4:D25', 'H4:H25', 'L4:L25', 'S4:S25']
      },
      incompleteFinalStatus: {
        requiredValues: ['I'],
        helpText: 'Enter I if this assignment is incomplete',
        rangeToValidate: ['E4:E25', 'I4:I25', 'M4:M25', 'T4:T25']
      },
      proposalGrade: {
        requiredValues: ['P', 'NP'],
        helpText: 'Enter P or NP',
        rangeToValidate: ['O4:O25']
      },
      numericGrades: {
        helpText: 'Enter a numeric grade from 0–100',
        rangeToValidate: ['F4:F25', 'J4:J25', 'N4:N25', 'R4:R25', 'U4:V25']
      },
      finalGrades: {
        requiredValues: ['A', 'B+', 'B', 'C+', 'C', 'NC', 'F', 'TF', 'TZ'],
        helpText: 'Enter A, B+, B, C+, C, NC, F, TF, or TZ',
        rangeToValidate: ['W4:W25']
      },
      /**
							* Package and return grade validation data as an object with two fields for both the non-numeric
							* and numeric data validations specified for this course.
							* @return {Object}
							* @return {Array} Object.nonNumeric - The non-numeric grade validations.
							* @return {Array} Object.numeric - The numeric grade validations.
							*/
      getGradeValidations: function() {
        var nonNumeric = [this.roughDraftStatus, this.lateFinalStatus, this.incompleteFinalStatus, this.proposalGrade, this.finalGrades];
        var numeric = [this.numericGrades];
        return {nonNumeric: nonNumeric, numeric: numeric}; // package and return validation data
      }
    },
    finalGradeFormula: '((((F$ + J$ + N$) / 300) * .45) + ((R$ / 100) * .40) + ((U$ / 100) * .15)) * 100',
    finalGradeFormulaRange: 'V4:V25'
  },
  '201': {
    name: 'Research 201',
    rows: [30, 30, 40, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20],
    columns: [215, 85, 55, 55, 55, 55, 55, 55, 55, 55, 55, 55, 55],
    columnHeadings: ['Student Name', 'Student ID', 'Analytic Essay', 'LR1', 'LR2', 'LR3', 'LR4', 'LR5', 'Oral', 'FD', 'Partici-\npation', 'Number Grade', 'Final Grade'],
    gradeValidations: {
      numericGrades: {
        helpText: 'Enter a numeric grade from 0–100',
        rangeToValidate: ['C4:L25']
      },
      finalGrades: {
        requiredValues: ['A', 'B+', 'B', 'C+', 'C', 'NC', 'F', 'TF', 'TZ'],
        helpText: 'Enter A, B+, B, C+, C, NC, F, TF, or TZ',
        rangeToValidate: ['M4:M25']
      },
      /**
							* Package and return grade validation data as an object with two fields for both the non-numeric
							* and numeric data validations specified for this course.
							* @return {Object}
							* @return {Array} Object.nonNumeric - The non-numeric grade validations.
							* @return {Array} Object.numeric - The numeric grade validations.
							*/
      getGradeValidations: function() {
        var nonNumeric = [this.finalGrades];
        var numeric = [this.numericGrades];
        return {nonNumeric: nonNumeric, numeric: numeric}; // package and return validation data
      }
    },
    finalGradeFormula: '(((C$ / 100) * .1) + ((D$ / 100) * .02) + ((E$ / 100) * .02) + ((F$ / 100) * .02) + ((G$ / 100) * .02) + ((H$ / 100) * .02) + ((I$ / 100) * .1) + ((J$ / 100) * .6) + ((K$ / 100) * .1)) * 100', // calculate final grade using official 201 weighting formula
    finalGradeFormulaRange: 'L4:L25'
  }
};

var OTHER_COURSE_NUMBER = '0'; // dummy course number for when 'Other' is selected from the setupNewGradebook dialog

var COURSE_SHORT_TITLES = {
  Expository: 'Expos',
  Exposition: 'Expos',
  Research: 'Research',
  Default: 'Course'
};

/* The SUMMER_SESSIONS object literal is a slightly obtuse way of storing information about the slightly obtuse summer session schedule.
 * The order is: day of the week the session starts (0–6), month the session starts (0–12), date counting from 1 on which the session
 * would start if the first day of the month were the same day of the week as the day the session starts, how long the course is in weeks
 */
var SUMMER_SESSIONS = {
  'A': [2, 4, 22, 4], // i.e. day of the week section A starts is: Tuesday = 2, month is May = 4, starts on the 22nd if the first of May is a Tuesday, course is 4 weeks long
  'B': [2, 4, 22, 6],
  'C': [2, 4, 22, 8],
  'D': [1, 5, 22, 5], // June and July sections start on Monday
  'E': [1, 5, 22, 6],
  'F': [1, 5, 22, 8],
  'G': [1, 6, 1, 5],
  'H': [1, 6, 1, 6],
  'J': [1, 6, 15, 4],
  'M': [2, 4, 8, 14],
  'R': [1, 6, 1, 2],
  'S': [4, 6, 29, 3], // except this section starts on a Thursday
  'T': [2, 4, 22, 12],
  'V': [2, 4, 22, 10],
};

/**
 * These templates are used for the alerts and dialogs displayed when a user selects the associated menu options
 */
var DIALOG_SETUP_NEW_GRADEBOOK = {
  alert: {
    alertType: YES_NO,
    msg: 'This will replace all data on this sheet. Are you sure you wish to proceed?',
    title: 'Setup New Gradebook'
  },
  dialog: {
    title: 'Setup New Section',
    html: 'setupNewGradebookDialog.html',
    width: 525,
    height: 450
  },
  error_msg: 'There was a problem with the setup process. Please try again.'
};
var DIALOG_SETUP_ADD_STUDENTS = {
  alert: {
    alertType: YES_NO,
    msg: 'This will replace any students currently listed in this gradebook. Are you sure you wish to proceed?',
    title: 'Add Students to Section'
  },
  dialog: {
    title: 'Add students to section',
    html: 'addStudentsFilePickerDialog.html',
    width: 800,
    height: 600
  },
  error_msg: 'There was a problem accessing your Drive. Please try again.'
};
var DIALOG_SETUP_CREATE_CONTACTS = {
  alert: {
    alertType: YES_NO,
    msg: 'This will create or update contacts and create a contact group based on the students on this spreadsheet. Students no longer in this course will be removed from the contact group if it already exists, but the contacts themselves will not be deleted. Is this what you want to do?',
    title: 'Create Contact Group'
  },
  command: 'setupCreateContacts',
  error_msg: 'Unable to access contacts. Please try again.'
};

/**
 * These templates are used for the sidebars displayed when a user selects the associated menu options
 */
var SIDEBAR_ASSIGNMENTS_CALC_WORD_COUNTS = {
  title: 'Word counts',
  html: 'assignmentsCalcWordCounts.html'
};
var SIDEBAR_HELP = {
  title: 'Exposify Help',
  html: 'Exposify_help.html'
};

/**
 * Toggles for whether or not I am tracking errors and installs and the names of the associated spreadsheets.
 */
var ERROR_TRACKING = true; // determines whether errors are sent to the error tracking spreadsheet
var ERROR_TRACKING_SHEET_NAME = 'Errors';
var INSTALL_TRACKING = true; // determine whether errors are sent to the install tracking spreadsheet
var INSTALL_TRACKING_SHEET_NAME = 'Installs';


// TRIGGER FUNCTIONS


/**
 * Execute as a trigger whenever the application is installed as an add-on to a Google Spreadsheet.
 */
function onInstall(e) {
  try {
    onOpen(e); // setup the custom menu, which is really the only important thing this function does
    expos.alert({msg: ALERT_INSTALL_THANKS})();
    expos.logInstall(); // tell me when someone has installed the add-on, for my records
  } catch(e) {
    expos.alert({msg: ERROR_INSTALL})();
    expos.logError('onInstall', e); // tell me when something goes wrong, so I can fix things
  }
} // end onInstall


/**
 * Execute as a trigger whenever the attached Google Spreadsheet is opened. Add the custom
 * Exposify menu to the menu bar.
 */
function onOpen() {
  var ui = expos.getUi();
  var menu = expos.getMenu();
  try {
    menu
      .addSubMenu(ui.createMenu('Setup')
        .addItem('New gradebook...', 'exposifySetupNewGradebook')
        .addItem('Add students to gradebook...', 'exposifySetupAddStudents')
        .addItem('Create contact group for students...', 'exposifySetupCreateContacts')
        .addItem('Create or update folder structure for this section...', 'exposifySetupCreateFolderStructure')
        .addItem('Share folders with students...', 'exposifySetupShareFolders'))
      .addSubMenu(ui.createMenu('Assignments')
        .addItem('Copy assignments for grading...', 'exposifyAssignmentsCopy')
        .addItem('Return graded assignments...', 'exposifyAssignmentsReturn')
        .addItem('Calculate word counts...', 'exposifyAssignmentsCalcWordCounts')
        .addItem('Compare rough and final drafts for selected student...', 'exposifyAssignmentsCompareDrafts'))
      .addSubMenu(ui.createMenu('Format')
        .addItem('Switch order of student names', 'exposifyFormatSwitchStudentNames')
        .addItem('Refresh shading of alternating rows', 'exposifyFormatSetShadedRows'))
      .addSeparator()
      .addItem('Help...', 'exposifyHelp')
      .addToUi();
  } catch(e) {
    expos.logError('onOpen', e);
  }
} // end onOpen


// CONSTRUCTORS


/**
 * When a user sets up a new gradebook, there are various options that can be selected to customize it.
 * This constructor turns those options into a Course object, adding default values for everything else
 * from the COURSE_FORMATS object literal, which stores various templates for course set-ups (i.e. all
 * values that are not supplied by the user).
 * @constructor
 */
function Course(courseInfo) {
  this.name = COURSE_FORMATS[courseInfo.course].name;
  this.number = courseInfo.course;
  this.section = courseInfo.section.toUpperCase(); // just in case it doesn't work on the client side?
  this.nameSection = this.name + ':' + this.section // if we need the name of the course followed by the section, as in the gradebook heading
  this.numberSection = (this.number === OTHER_COURSE_NUMBER ? this.section : this.number + ':' + this.section); // if course number is an empty string, just return the section
  this.semester = expos.getSemesterYearString(courseInfo.semester); // automatically determines the current year and adds that to the semester (Fall, Spring, or Summer)
  this.meetingDays = courseInfo.meetingDays;
  this.rows = COURSE_FORMATS[courseInfo.course].rows;
  this.columns = COURSE_FORMATS[courseInfo.course].columns;
  this.columnHeadings = COURSE_FORMATS[courseInfo.course].columnHeadings;
  this.gradeValidations = expos.doMakeGradeValidations(courseInfo.course); // this is complicated, so I do the work in a separate function; initialized to a GradeValidationSet
}; // end Course


/**
 * Constructs a set of grade validations for use by the Course constructor. Grade validations are simply
 * spreadsheet data validations that are defined according to the COURSE_FORMATS templstes. For example,
 * if a course only allows numeric grades to be given for a particular assignment, the spreadsheet will
 * enforce that requirement in accredence with the rules specified in COURSE_FORMATS.
 * @constructor
 */
function GradeValidationSet() {
  this.validations = [];
  this.ranges = [];
}; // end GradeValidationSet


/**
 * This is a simple student record, containing the student's name and netid, which can be computed into
 * a valid email address. It assumes all emails have the same domain, but this can be modified for edge
 * cases using the {@code setEmail()} method.
 * @constructor
 */
function Student(name, netid) {
  this.name = name;
  this.netid = netid;
  this.email = netid + EMAIL_DOMAIN;
}; // end Student


/**
 * Creates a folder object with information about a folder stored in the user's Google Drive.
 * @constructor
 */
function Folder(name, parent, path) {
  this.name = name;
  this.parent = parent;
  this.path = path;
}; // end Folder


var FolderStructure = function(semesterTitle, courseTitle) {
  this.rootFolder = DriveApp.getRootFolder();
  this.semesterTitle = semesterTitle;
  this.courseTitle = courseTitle;
  this.semesterFolder = this.getSemesterFolder();
  this.courseFolder = this.getCourseFolder();
  this.gradedFolder = this.getGradedFolder();
  this.studentFolders = this.getStudentFolders();
}; // end FolderStructure


/**
 * This is the main Exposify constructor, the namespace for most of the methods and properties of the add-on.
 * @constructor
 */
function Exposify() {
  // Private properties
  /**
   * Stores a reference to the active Spreadsheet object, which shouldn't vary after the object is created.
   * @private {Spreadsheet}
   */
  var spreadsheet_ = SpreadsheetApp.getActiveSpreadsheet();
  /**
   * Stores a reference to the Ui object for this spreadsheet, which shouldn't vary after the object is created.
   * @private {Ui}
   */
  var ui_ = SpreadsheetApp.getUi();
  /**
   * Stores a reference to the Menu object for this spreadsheet.
   * @private {Menu}
   */
  var menu_ = ui_.createMenu('Exposify');
  /**
   * Stores references to common UI button sets, so I don't have to look them up at runtime.
   * @private {ButtonSet}
   */
  var ok_ = ui_.ButtonSet.OK;
  var okCancel_ = ui_.ButtonSet.OK_CANCEL;
  var yes_ = ui_.Button.YES;
  var yesNo_ = ui_.ButtonSet.YES_NO;
  // Protected methods
  /**
   * Return the active Spreadsheet object.
   * @protected
   * @return {Spreadsheet} spreadsheet_ - A Google Apps Spreadsheet object.
   */
  this.getSpreadsheet = function() { return spreadsheet_; };
  /**
   * Return the Ui object for this spreadsheet.
   * @protected
   * @return {Ui} ui_ - The Ui object for the Spreadsheet object to which Exposify is attached.
   */
  this.getUi = function() { return ui_; };
  /**
   * Return the Menu object for this spreadsheet.
   * @protected
   * @return {Menu} menu_ - The Menu object for the Spreadsheet object to which Exposify is attached.
   */
  this.getMenu = function() {return menu_; };
  /**
   * Set the default time zone for the spreadsheet. Return the spreadsheet for chaining.
   * @protected
   * @param {string} timezone - A string representing a timezone in "long" format, as listed by Joda.org
   * @return {Spreadsheet} spreadsheet_ - The Spreadsheet object to which Exposify is attached.
   */
  this.setTimezone = function(timezone) {
    spreadsheet_.setSpreadsheetTimeZone(timezone);
    return spreadsheet_;
  this.menu = ui_.createMenu('Exposify');
  };
  /**
   * Display a dialog box to the user.
   * @protected
   * @param {HtmlOutput} htmlDialog - The sanitized html to display as a dialog box to the user.
   * @param {string} title - The title of the dialog box.
   */
  this.showModalDialog = function(htmlDialog, title) {
    ui_.showModalDialog(htmlDialog, title);
  };
  /**
   * Display a sidebar to the user.
   * @protected
   * @param {HtmlOutput} htmlSidebar - The sanitized html to display as a sidebar to the user.
   */
  this.showSidebar = function(htmlSidebar) {
    ui_.showSidebar(htmlSidebar);
  };
  // Protected properties
  /**
   * This is a simple interface for accessing the built-in UI alert controls.
   * @protected {Object}
   */
  this.alertUi = {
    ok: ok_,
    okCancel: okCancel_,
    yes: yes_,
    yesNo: yesNo_
    };
  // Initialization procedures
  spreadsheet_.setSpreadsheetTimeZone(TIMEZONE); // sets the default time zone to the value stored by TIMEZONE
}; //end Exposify


// MENU COMMANDS


/**
 * Since menu commands have to call functions in the global namespace, I can't call methods defined on
 * the Exposify prototype. These are the menu functions, and I set them up to pass control to the
 * function {@code checkSheetStatus()} if they shouldn't be executed if no gradebook has been created,
 * or function {@code executeMenuCommand()} otherwise. The latter function processes a constant object
 * literal into a confirmation alert and HTML dialog box and displays them to the user. The object
 * literals contain the alert message to display first as a confirmation, followed by the details of the
 * dialog box.
 */
function exposifySetupNewGradebook() { expos.executeMenuCommand.call(expos, DIALOG_SETUP_NEW_GRADEBOOK); }
function exposifySetupAddStudents() { expos.checkSheetStatus.call(expos, DIALOG_SETUP_ADD_STUDENTS); }
function exposifySetupCreateContacts() { expos.checkSheetStatus.call(expos, DIALOG_SETUP_CREATE_CONTACTS); }
function exposifyAssignmentsCalcWordCounts() { return expos.checkSheetStatus.call(expos, {command: 'assignmentsCalcWordCounts'}); }
function exposifyFormatSwitchStudentNames() { return expos.checkSheetStatus.call(expos, {command: 'formatSwitchStudentNames', error_msg: ERROR_FORMAT_SWITCH_STUDENT_NAMES}); }
function exposifyFormatSetShadedRows() { return expos.checkSheetStatus.call(expos, {command: 'formatSetShadedRows', error_msg: ERROR_FORMAT_SET_SHADED_ROWS}); }
function exposifyHelp() { return expos.executeMenuCommand.call(expos, {command: 'help'}); }

// old functions to be replaced
function exposifySetupCreateFolderStructure() { return expos.setupCreateFolderStructure(); }
function exposifySetupSharedFolders() { return expos.setupSharedFolders(); }
function exposifyAssignmentsCopy() { return expos.assignmentsCopy(); }
function exposifyAssignmentsReturn() { return expos.assignmentsReturn(); }
function exposifyAssignmentsCompareDrafts() { return expos.assignmentsCompareDrafts(); }


// CALLBACKS


/**
 * As with the menu commands, callbacks from user interfaces (from the client side) that use the
 * Google API must call functions in the global namespace, at least as far as I can tell. These
 * functions pass control to other functions that do the actual work. In the future, I may replace
 * these functions with a single callback handler, as with the menu functions, if it's possible.
 */
function assignmentsCalcWordCountsCallback(params) { return expos.assignmentsCalcWordCounts(params); }
function assignmentsCalcWordCountsCallbackGetTitle() { return expos.assignmentsCalcWordCountsGetTitle(); }
function getOAuthToken() { return expos.getOAuthToken(); }
function setupNewGradebookCallback(courseInfo) { expos.setupNewGradebook(courseInfo); }
function setupAddStudentsCallback(id) { expos.setupAddStudents(id); }


// EXPOSIFY FUNCTIONS


/**
 * Check that an incoming request to make an alert has the correct parameters and raise an
 * exception if it does not. The parameter is an object with two fields, one containing the type of alert
 * and one containing the message to be displayed to the user. The available alert types are OK, OK_CANCEL,
 * and YES_NO. These are defined as constart values. This function returns another function, which can be
 * executed to display the dialog box, i.e. with {@code alert(confirmation)();}.
 * @param {Object} confirmation - The object holding the arguments to the function.
 * @param {string} confirmation.string - Alert type of the alert dialog, which determines the buttons to display.
 * @param {string} confirmation.msg - The message to display in the alert dialog.
 * @return {Function} alertFunction - A function that can be executed to display the alert.
 */
Exposify.prototype.alert = function(confirmation) {
  try {
    if (!confirmation.hasOwnProperty('alertType')) {
      return this.doMakeAlert({alertType: OK, msg: confirmation.msg}); // A simple alert with an OK button is the default
    } else if (!this.alertUi.hasOwnProperty(confirmation.alertType)) {
      var e = 'Alert type ' + confirmation.alertType + 'is not defined on Exposify.';
      throw e // Throw an exception if the alert type doesn't exist, probably superfluous error checking
    } else {
      var alertFunction = this.doMakeAlert(confirmation); // Factor out the alert composition
      return alertFunction;
    }
  } catch(e) { this.logError('Exposify.prototype.alert', e); }
} // end Exposify.prototype.alert


/**
 * Check whether an array contains a specified item. Modified code from http://stackoverflow.com/a/237176.
 * @param {Array} arr - The array to check.
 * @param {*} item - Any value.
 * @return {boolean}
 */
Exposify.prototype.arrayContains = function(arr, item) {
  var i = arr.length;
  while (i--) {
    if (arr[i] === item) {
      return true;
    }
  }
  return false;
} // end Exposify.prototype.arrayContains


/**
 * Return word counts for the set of files specified by the user.
 * @param {Object} params - An object containing the parameters passed to the callback function.
 * @param {string} params.students - Either "selected" or "all," the option to search all papers or a subset.
 * @param {string} params.filter - The filter for the files, optionally entered by the user.
 */
Exposify.prototype.assignmentsCalcWordCounts = function(params) {
  try {
    var sheet = this.getActiveSheet();
    var students = params.students;
    var filter = params.filter;
    var counts = null;
    if (students === 'selected') {
      var counts = this.doCalcWordCountsSelected(sheet, filter);
    } else if (students === 'all') {
      var counts = this.doCalcWordCountsAll(sheet, filter);
    }
    counts.sort(function (a, b) {
      if (a.document > b.document) {
              return 1;
      }
      if (a.document < b.document) {
              return -1;
      }
      return 0;
      });
    return counts;
  } catch(e) { this.logError('Exposify.prototype.assignmentsCalcWordCounts', e); }
} // end Exposify.prototype.assignmentsCalcWordCounts


/**
 * Retrieve the title of the course and number of students enrolled for use in the
 * word counts sidebar.
 * @return {string} title - The course title and enrollment figure.
 */
Exposify.prototype.assignmentsCalcWordCountsGetTitle = function() {
  try {
    var sheet = this.getActiveSheet();
    var courseTitle = this.getCourseTitle(sheet);
    var enrollment = this.getStudentCount(sheet);
    var title = courseTitle + ' (' + enrollment + ' students)';
    return title;
  } catch(e) { this.logError('Exposify.prototype.assignmentsCalcWordCountsGetTitle', e); }
} // end Exposify.prototype.assignmentsCalcWordCountsGetTitle


/**
 * Check whether a gradebook has already been set up for this sheet. If so, pass
 * control to the {@code executeMenuCommand()} function. Return false otherwise.
 * @param {Object} params - The parameters passed to the original menu command.
 */
Exposify.prototype.checkSheetStatus = function(params) {
  try {
    var sheet = this.getActiveSheet();
    var check = this.getSheetStatus(sheet);
    if (check === false) {
      var alert = this.alert({msg: ALERT_NO_GRADEBOOK});
      alert();
      return false
    } else {
      this.executeMenuCommand(params);
    }
  } catch(e) { this.logError('Exposify.prototype.checkSheetStatus', e); }
} // end Exposify.prototype.checkSheetStatus


/**
 * Create a dialog box to display to the user using information stored as a template an object literal
 * constant. The html field of the argument object should be an HTML file.
 * @param {Object} dialog - The template containing the specification of the dialog box.
 * @param {string} dialog.title - The title of the dialog box to display.
 * @param {string} dialog.html - The html file to process and display as the dialog box content.
 * @param {number} dialog.width - The width of the dialog box.
 * @param {number} dialog.height - The height of the dialog box.
 * @return {HtmlOutput} htmlDialog - The sanitized html dialog box, ready to be displayed to the user.
 */
Exposify.prototype.createHtmlDialog = function(dialog) {
  try {
    var stylesheet = this.getHtmlOutputFromFile(STYLESHEET);
    var body = this.getHtmlOutputFromFile(dialog.html).getContent(); // Sanitize the HTML file
    var page = stylesheet.append(body).getContent(); // Combine the style sheet with the body
    var htmlDialog = this.getHtmlOutput(page)
      .setWidth(dialog.width)
      .setHeight(dialog.height);
    return htmlDialog;
  } catch(e) { this.logError('Exposify.prototype.createHtmlDialog', e); }
} // end Exposify.prototype.createHtmlDialog


/**
 * Insert student names into the spreadsheet.
 * @param {Array} students - A list of Student objects containing the data to add to the spreadsheet.
 * @param {Sheet} sheet - A Google Apps Sheet object, the sheet to which names will be added.
 */
Exposify.prototype.doAddStudents = function (students, sheet) {
  try {
    var studentList = [];
    var fullRange = sheet.getRange(4, 1, MAX_STUDENTS, 2);
    fullRange.clearContent(); // erase whatever data is already on the sheet where we put student names
    var range = sheet.getRange(4, 1, students.length, 2); // get a range of two columns and a number of rows equal to the number of students to insert
    students.forEach(function(student) { studentList.push([student.name, student.netid]); } ); // add a row to the temporary studentList array for each student
    range.setValues(studentList); // set the value of the whole range at once, so I don't call the API more than necessary
  } catch(e) { this.logError('Exposify.prototype.doAddStudents', e); }
} // end Exposify.prototype.doAddStudents


/**
 * Format a spreadsheet sheet for use as a gradebook for a specified course.
 * @param {Object} newCourse - Information about the new course on which to base the formatting.
 * @param {Course} newCourse.course - The course information requested from the user.
 * @param {Sheet} newCourse.sheet - The sheet to format.
 */
 Exposify.prototype.doFormatSheet = function(newCourse) {
  var course = newCourse.course;
  var sheet = newCourse.sheet;
  try {
    var section = course.section; // create a series of variables from the Course object passed in, for legibility
    var semester = course.semester;
    var courseNumber = course.number;
    var courseFormat = COURSE_FORMATS[courseNumber];
    var rows = course.rows;
    var lastRow = rows.length;
    var columns = course.columns;
    var lastColumn = columns.length;
    var courseTitle = course.nameSection;
    var columnHeadings = course.columnHeadings;
    var gradeValidations = course.gradeValidations;
    var headingRange = sheet.getRange(3, 1, 1, columnHeadings.length); // cell range for gradebook column headings
    var centerRange = sheet.getRange(3, 3, lastRow, lastColumn); // cell range for central part of gradebook, where grade data is actually entered
    var topRowsRange = sheet.getRange(1, 1, 3, lastColumn); // rows to keep at the top of the spreadsheet view
    var titleRange = sheet.getRange('A1:A2'); // course name and semester titles
    var mergeTitleRange = sheet.getRange('A1:B2'); // we want to merge each of these with the following cell to create a bigger space for the titles
    var mergeRange = sheet.getRange(1, 3, 2, lastColumn - 2); // merge the empty columns in the top rows so it looks nicer
    var cornerRange = sheet.getRange('A3:B3'); // where the frozen rows and columns intersect
    var fullRange = sheet.getRange(1, 1, lastRow, lastColumn); // range of the entire gradebook
    var maxRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()); // range of the entire visible sheet
    sheet.clear(); // clear all values and formatting
    maxRange.clearDataValidations(); // this has to be done separately
    sheet.setFrozenRows(0); // make sure only the correct rows and columns are frozen when formatting is complete
    sheet.setFrozenColumns(0);
    fullRange.breakApart(); // break apart any merged cells
    for (i = 1; i <= lastRow; i += 1) { // set row heights
      sheet.setRowHeight(i, rows[i-1]);
    }
    for (i = 1; i <= lastColumn; i += 1) { // set column widths
      sheet.setColumnWidth(i, columns[i-1]);
    }
    fullRange.setFontFamily([FONT]); // set font
    fullRange.setFontSize(11); // student names and grades font size
    titleRange.setFontSizes([[16],[14]]); // titles font size
    headingRange.setFontSize(9); // headings font size
    cornerRange.setHorizontalAlignment('center'); // set text alignments
    cornerRange.setVerticalAlignment('middle');
    centerRange.setHorizontalAlignment('center');
    centerRange.setVerticalAlignment('middle');
    fullRange.setBorder(true, true, true, true, true, true); // set cell borders
    titleRange.setValues([[courseTitle],[semester]]); // set titles
    mergeTitleRange.mergeAcross(); // merge title cells
    mergeRange.mergeAcross(); // merge other cells in the first two rows
    headingRange.setValues([columnHeadings]); // set column headings
    headingRange.setWrap(true); // set word wrapping
    topRowsRange.setBackground(COLOR_SHADED); // set background color of frozen rows
    sheet.setFrozenRows(3); // freeze first three rows (sorry, magic number)
    sheet.setFrozenColumns(2); // freeze first two columns (sorry, another magic number)
    if (gradeValidations !== undefined) {
      this.setGradeValidations(sheet, gradeValidations); // set data validations for grades
    }
    if (courseFormat.hasOwnProperty('finalGradeFormulaRange')) {
      this.doSetFormulas(sheet, courseNumber); // apply final grade formula to this range
    }
    if (course.meetingDays.length !== 0) {
      this.doFormatSheetAddAttendanceRecord(course, sheet); // add an attendance sheet if the user asked for it
    }
    this.doSetShadedRows(sheet); // set alternating color of student rows
    sheet.setName(course.numberSection); // name sheet with section number
    return true;
  } catch(e) { this.logError('Exposify.prototype.doFormatSheet', e); }
} // end Exposify.prototype.doFormatSheet


/**
 * Add an attendance record to a newly formatted gradebook.
 * @param {Course} course - A Course object containing formatting information for the course.
 * @param {Sheet} sheet - The Google Apps Sheet object to format.
 */
Exposify.prototype.doFormatSheetAddAttendanceRecord = function(course, sheet) {
  try {
    var courseData = this.getCourseData(course); // get an object representing the date the semester begins, which days of the week it meets, and the duration in weeks it meets for
    var semesterBeginsDate = courseData.semesterBeginsDate;
    var meetingDays = courseData.meetingDays;
    var meetingWeeks = courseData.meetingWeeks;
    var schedule = this.doMakeSchedule(semesterBeginsDate, meetingDays, meetingWeeks); // calculate a schedule of actual days from this information to insert into the spreadsheet
    var width = schedule.length; // number of meetings
    var begin = sheet.getLastColumn() + 1; // place to insert attendance sheet in spreadsheet
    var end = begin + width - 1; // end of attendance sheet
    var maxColumns = sheet.getMaxColumns();
    var columnsToAdd = schedule.length - (maxColumns - (begin - 1)); // extend the length of the spreadsheet, if necessary
    if (columnsToAdd > 0) {
      sheet.insertColumnsAfter(maxColumns, columnsToAdd);
    }
    var attendanceRange = sheet.getRange(3, begin, 1, width); // headings for dates, one row in height
    var mergeTitleRange = sheet.getRange(1, begin, 2, width); // merge top rows to look nicer
    var shadedRange = sheet.getRange(1, begin, 3, width); // we need to add alternate row shading for the attendance sheet, too
    var lastRow = course.rows.length;
    var borderRange = sheet.getRange(1, begin, lastRow, width); // and add borders
    for (i = begin; i <= end; i += 1) { // set column widths
      sheet.setColumnWidth(i, ATTENDANCE_SHEET_COLUMN_WIDTH);
    }
    attendanceRange.setFontFamily([FONT]); // set font
    attendanceRange.setFontSize(9); // set font size
    attendanceRange.setHorizontalAlignment('left'); // set text alignments
    attendanceRange.setVerticalAlignment('middle');
    borderRange.setBorder(true, true, true, true, true, true); // set cell borders
    attendanceRange.setValues([schedule]); // set titles, this call does the important work
    mergeTitleRange.merge(); // merge title cells
    shadedRange.setBackground(COLOR_SHADED); // set background color of top rows
  } catch(e) { this.logError('Exposify.prototype.doFormatSheetAddAttendanceRecord', e); }
} // end Exposify.prototype.doFormatSheetAddAttendanceRecord


/**
 * Create an alert dialog box to be displayed to the user. The alert is comprised of an alert type, which should be
 * OK, OK_CANCEL, or YES_NO, and a message to print in the dialog box. The alert types are constant values. This
 * function returns another function that can be executed to display the dialog box.
 * @param {Object}
 * @param {string} Object.alertType - The type of alert to display.
 * @param {string} Object.msg - The message to display in the alert.
 * @return {Function} dialog - The function for displaying the alert dialog.
 */
Exposify.prototype.doMakeAlert = function(confirmation) {
  try {
    var alertType = confirmation.alertType;
    var msg = confirmation.msg;
    var ui = this.getUi();
    var title = (confirmation.hasOwnProperty('title') ? confirmation.title : ALERT_TITLE_DEFAULT);
    var alertUi = this.alertUi;
    var ok = alertUi.ok;
    var yes = alertUi.yes;
    var okCancel = alertUi.okCancel;
    var yesNo = alertUi.yesNo;
    var alerts = { // Map alert functions to different alert types
      ok: function() { return ui.alert(title, msg, ok); },
      okCancel: function() { return (ui.alert(title, msg, okCancel)) === ok ? true : false; },
      yesNo: function() { return (ui.alert(title, msg, yesNo)) === yes ? true : false; },
      prompt: function() {
        var response = ui.prompt(title, msg, okCancel);
        return response.getSelectedButton() === ok ? response.getResponseText() : false;
      }
    };
    var dialog = alerts[alertType]; // Create a function using the closures stored in the {@code alerts} variable.
    return dialog; // Return the function without executing it.
  } catch(e) { this.logError('Exposify.prototype.doMakeAlert', e); }
} // end Exposify.prototype.doMakeAlert


/**
 * Make Google Apps Data Validation objects for applying grade validation to a new gradebook.
 * @param {number} courseNumber - The course number, used to pull data from the course template object literal.
 * @return {GradeValidationSet} gradeValidations - The grade validations to apply and the ranges to apply them to.
 */
Exposify.prototype.doMakeGradeValidations = function(courseNumber) {
  try {
    var gradeValidations = new GradeValidationSet();
    var courseFormats = COURSE_FORMATS[courseNumber];
    if (courseFormats.hasOwnProperty('gradeValidations')) {
      var gradeValidationInfo = courseFormats.gradeValidations.getGradeValidations();
      var nonNumeric = gradeValidationInfo.nonNumeric;
      nonNumeric.forEach(function(validationSet) {
        var newDataValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList(validationSet.requiredValues, (validationSet.requiredValues.length > 1 ? true : false))
        .setAllowInvalid(false)
        .setHelpText(validationSet.helpText)
        .build();
        gradeValidations.validations.push(newDataValidation);
        gradeValidations.ranges.push(validationSet.rangeToValidate);
      });
      if (gradeValidationInfo.hasOwnProperty('numeric')) {
        var numeric = gradeValidationInfo.numeric;
        numeric.forEach(function(validationSet) {
          var newDataValidation = SpreadsheetApp.newDataValidation()
          .requireNumberBetween(0, 100)
          .setAllowInvalid(false)
          .setHelpText(validationSet.helpText)
          .build();
          gradeValidations.validations.push(newDataValidation);
          gradeValidations.ranges.push(validationSet.rangeToValidate);
        });
      }
      return gradeValidations; // neat and tidy package
    }
  } catch(e) { this.logError('Exposify.prototype.doMakeGradeValidations', e); }
} // end Exposify.prototype.doMakeGradeValidations


/**
 * Calculate a schedule for a course, which is complicated so I don't know if it will always be 100% accurate but probably good enough.
 * I am going to burn in hell for writing this function.
 * @param {Date} semesterBeginsDate - The first day of the semester.
 * @param {Array} meetingDays - A list of the days of the week when the course normally meets.
 * @param {number} meetingWeeks - The number of weeks for which a course meets (default is 15).
 * @return {Array} daysToMeet - A list of text dates to be inserted into the spreadsheet.
 */
Exposify.prototype.doMakeSchedule = function(semesterBeginsDate, meetingDays, meetingWeeks) {
  try {
    var day = 1;
    var month = semesterBeginsDate.getMonth();
    var year = semesterBeginsDate.getFullYear();
    var firstDayOfClass = semesterBeginsDate.getDate();
    var lastDay = this.getLastDayOfMonth(month, year);
    var daysToMeet = [];
    var firstDayOfSpringBreak = this.getFirstDayOfSpringBreak(year); // get first day of Spring Break, so we don't include dates for that week
    var tuesdayOfThanksgivingWeek = this.getTuesdayOfThanksgivingWeek(year); // get Tuesday of Thanksgiving week, so we can take change of day designations into account
    var alternateDesignationYear = this.getAlternateDesignationYearStatus(year); // except on some years, when 9/1 is a Tuesday, the designation days are different
    for (day = firstDayOfClass, week = 1; day < lastDay + 1 && week < meetingWeeks + 1; day += 1) { // check every single day in the semester to see if it belongs in the course schedule
      if (month === 2 && day === firstDayOfSpringBreak) { // if the day we're checking is the first day of Spring Break, just skip 9 days
        day += 9;
        week += 2;
      }
      var dayToCheck = (new Date(year, month, day)).getDay();
      if (month === 8 && day === 8 && alternateDesignationYear) { // pretend today is Monday, September 1 if we're checking September 8 and September 1 was a Tuesday
        dayToCheck = 1;
      }
      if (month === 10 && (day >= tuesdayOfThanksgivingWeek && day <= tuesdayOfThanksgivingWeek + 5)) { // changes of designation during Thanksgiving week
        if (day === tuesdayOfThanksgivingWeek) {
          dayToCheck = (alternateDesignationYear ? 2 : 4); // Tuesday becomes Thursday, unless it's one of those weird years (2015 and 2020) when it stays a Tuesday
        } else if (day === (tuesdayOfThanksgivingWeek + 1)) { // Wednesday becomes Friday
          dayToCheck = 5;
        } else {
          day += 4; // skip the rest
          dayToCheck = 1; // pretend it's Monday
          week += 1;
          if (day > 30) { // make sure we didn't go over 30 days for November by skipping 4 days at the end of the month
            day = day - 30;
            month += 1;
            lastDay = this.getLastDayOfMonth(month, year);
          }
        }
      }
      if (meetingDays.some(function(meetingDay) { return dayToCheck === DAYS[meetingDay]; })) { // if the day we're checking is one of the days the class meets, add it to a running list of meeting days
        daysToMeet.push((month + 1) + '/\n' + day); // create the actual text that will appear in the spreadsheet for each meeting day, i.e. 9/1 with a carriage return after the forward slash to look nice and avoid automatic date formatting
      }
      if (day === lastDay) { // if we're at the last day of the month, reset the day counter to 0, increase the month counter, and calculate the last day of the new month
        day = 0;
        month += 1;
        lastDay = this.getLastDayOfMonth(month, year);
      }
      if (dayToCheck === 6) { // if the day we're checking is Saturday, increment the week counter
        week += 1;
      }
    }
    return daysToMeet; // an array of text dates ready to be inserted directly into the spreadsheet
  } catch(e) { this.logError('Exposify.prototype.doMakeSchedule', e); }
} // end Exposify.prototype.doMakeSchedule


/**
 * Extract student names and ids from the 'participant data file compatible with Microsoft Excel' downloadable
 * from the Site Info page of a Sakai course site. This function will only work if that file has been unmodified.
 * This function works whether or not the file has been converted from csv (comma separated values) format
 * into Google Sheets format. Returns an array of Student objects with which to populate the spreadsheet gradebook.
 * @param {string} id - The id of the Google Apps File object to open.
 * @param {string} mimeType - The mimeType of the Google Apps File object to open.
 * @return {Array} students - A list of Student objects, the students to add to the gradebook.
 */
Exposify.prototype.doParseSpreadsheet = function(id, mimeType) {
  try {
    var students = [];
    var sheet = this.getActiveSheet();
    var section = this.getSectionTitle(sheet);
    if (mimeType === 'application/vnd.google-apps.spreadsheet') {
      var file = SpreadsheetApp.openById(id); // open file to retrieve data
      var page = file.getSheets()[0];
      var lastRow = page.getLastRow();
      var range = page.getRange(2, 1, lastRow - 1, 6).getValues();
      for (row = 0; row < lastRow - 1; row += 1) {
        var fileSection = range[row][2].substr(18, 2); // extract section from course code
        if (range[row][0] && range[row][4] === 'Student' && fileSection === section) {
          students.push(new Student(range[row][0], range[row][1])); // create list of Student objects from spreadsheet
        }
      }
    } else if (mimeType === 'text/csv') {
      var file = DriveApp.getFileById(id);
      var data = file.getAs('text/csv').getDataAsString(); // convert file data into a string (can't open csv files in Google Drive)
      var csv = Utilities.parseCsv(data);
      var length = csv.length;
      for (row = 1; row < length; row += 1) {
        var fileSection = csv[row][2].substr(18, 2); // extract section from course code
        if (csv[row][0] && csv[row][4] === 'Student' && fileSection === section) {
          students.push(new Student(csv[row][0], csv[row][1])); // create list of Student objects from csv file
        }
      }
    }
    return students;
  } catch(e) { this.logError('Exposify.prototype.doParseSpreadsheet', e); }
} // end Exposify.prototype.doParseSpreadsheet


/**
 * Set formulas for automatically calculating final grades in gradebooks that require it.
 * @param {Sheet} sheet - The Google Apps Sheet object to modify.
 * @param {number} courseNumber - The course number, used to pull data from the course template object literal.
 */
Exposify.prototype.doSetFormulas = function(sheet, courseNumber) {
  try {
    var courseFormat = COURSE_FORMATS[courseNumber];
    var calcRange = sheet.getRange(courseFormat.finalGradeFormulaRange); // get the Range object representing the cells to which to apply the formulas
    var formula = courseFormat.finalGradeFormula;
    var formulas = [];
    for (i = 4; i < 26; i += 1) { // 22 students maximum
      formulas.push([formula.replace('$', i, 'g')]); // substitle '$' wildcard with the appropriate row number for each cell to which we are applying the final grade formula
    }
    calcRange.setFormulas(formulas); // apply the formulas to the Range object
  } catch(e) { this.logError('Exposify.prototype.doSetFormulas', e); }
} // end Exposify.prototype.doFormatSheetSpecialRules


/**
 * Make the gradebook easier to read by setting alternating shaded and unshaded rows.
 * @param {Sheet} sheet - The Google Apps Sheet object with the gradebook to modify.
 */
 Exposify.prototype.doSetShadedRows = function (sheet) {
  try {
    var lastRow = MAX_STUDENTS + 3; // maximum number of students plus three to account for title rows
    var lastColumn = sheet.getLastColumn();
    var rows = lastRow - 3;
    var shadedRange = sheet.getRange(4, 1, rows, lastColumn);
    var blankColor = COLOR_BLANK;
    var shadedColor = COLOR_SHADED;
    var blankRow = [];
    var shadedRow = [];
    var newRows = [];
    for (i = 0; i < lastColumn; i += 1) { // generate array of alternating colors of the correct length
      blankRow.push(blankColor);
      shadedRow.push(shadedColor);
    }
    for (i = 4; i <= lastRow; i += 1) {
      i % 2 === 0 ? newRows.push(blankRow) : newRows.push(shadedRow); // generate array of alternating shaded and blank rows so I only have to call setBackgrounds once
    }
    shadedRange.setBackgrounds(newRows); // set row backgrounds
  } catch(e) { this.logError('Exposify.prototype.doSetShadedRows', e); }
} // end Exposify.prototype.doSetShadedRows


/**
 * Switch student name order from last name first to first name last or vice versa.
 * @param {Sheet} sheet - The Google Apps Sheet object with the gradebook to modify.
 */
Exposify.prototype.doSwitchStudentNames = function (sheet) {
  try {
    var students = this.getStudents(sheet);
    for (i = 0; i < students.length; i += 1) {
      var name = students[i].name;
      if (name.match(/.+,.+/)) { // match student names against a regular expression pattern to determine whether or not the name strings contain commas... I hope there aren't any people whose names actually contain commas
        students[i].name = this.getNameFirstLast(name);
      } else {
        students[i].name = this.getNameLastFirst(name);
      }
    }
    this.doAddStudents(students, sheet); // repopulate the sheet with the student names
    sheet.sort(1); // sort sheet alphabetically
    this.doSetShadedRows(sheet); // because the sort will probably mess them up
  } catch(e) { this.logError('Exposify.prototype.doSwitchStudentNames', e); }
} // end Exposify.prototype.doSwitchStudentNames


/**
 * Generate a regular expression for counting the words of all the documents in a course folder.
 * @param {Sheet} sheet - A Google Apps Sheet object containing the gradebook to check.
 * @param {string} filter - The file search filter supplied by the user.
 * @return {Array} counts - The array of word count information returned by {@code getWordCounts()}.
 */
Exposify.prototype.doCalcWordCountsAll = function(sheet, filter) {
  try {
    var studentList = this.getStudentNames(sheet);
    var regex = '(.*';
    studentList.forEach( function(student, index) {
      regex += (student + (index === studentList.length - 1 && filter === '' ? '.*)' : '.*|.*')); // I am a bad person
    });
    if (filter !== '') {
      regex += (filter + '.*)+(.*' + filter + '.*|.*'); // I'm sorry
      studentList.forEach( function(student, index) {
        regex += (student + (index === studentList.length - 1 ? '.*)' : '.*|.*')); // Seriously
      });
    }
    var re = new RegExp(regex);
    var counts = this.getWordCounts(sheet, re);
    return counts;
  } catch(e) { this.logError('Exposify.prototype.doCalcWordCountsAll', e); }
} // end Exposify.prototype.doCalcWordCountsAll


/**
 * Generate a regular expression for counting the words of a specific student's documents.
 * @param {Sheet} sheet - A Google Apps Sheet object containing the gradebook to check.
 * @param {string} filter - The file search filter supplied by the user.
 * @return {Array} counts - The array of word count information returned by {@code getWordCounts()}.
 */
Exposify.prototype.doCalcWordCountsSelected = function(sheet, filter) {
  try {
    var cellValue = sheet.getActiveCell().getValue();
    if (cellValue === '') {
      return [];
    }
    var regex = (filter === '' ? '.*' + cellValue + '.*' : '(.*' + cellValue + '.*|.*' + filter + '.*)+(.*' + filter + '.*|.*' + cellValue + '.*)'); // I mean it
    var re = new RegExp(regex);
    var counts = this.getWordCounts(sheet, re);
    return counts;
  } catch(e) { this.logError('Exposify.prototype.doCalcWordCountsSelected', e); }
} // end Exposify.prototype.doCalcWordCountsSelected


/**
 * Execute a menu command selected by the user, first displaying an alert and then an
 * HTML dialog box, both provided as arguments and based on object literal constants.
 * @param {Object} params - The parameters passed to the function.
 * @param {Object=} alert - An optional alert to display before the command is executed.
 * @param {Object=} dialog - An optional dialog box to display to collect user input.
 * @param {string=} error_msg - An optional error message to display if something goes wrong.
 */
Exposify.prototype.executeMenuCommand = function(params) {
  try {
    if (params.hasOwnProperty('alert')) { // show the alert, if there is one
      var alert = this.alert(params.alert);
      var response = alert();
    }
    if (params.hasOwnProperty('dialog') && response === true) { // show the dialog, if there is one, and if the alert response is true
      var dialog = params.dialog;
      var htmlDialog = this.createHtmlDialog(dialog); // "sanitize" the HTML so Google will display it
      this.showModalDialog(htmlDialog, dialog.title); // to limit the number of times I reference Ui
    }
    if (params.hasOwnProperty('command')) { // execute whatever other command the user is requesting, if no alert or dialog is needed
      var that = this; // have I told you lately that I love JavaScript?
      var commands = {
        setupCreateContacts: function() { that.setupCreateContacts(that.getActiveSheet()); },
        assignmentsCalcWordCounts: function() { that.showHtmlSidebar(SIDEBAR_ASSIGNMENTS_CALC_WORD_COUNTS); },
        formatSwitchStudentNames: function() { that.doSwitchStudentNames(that.getActiveSheet()); },
        formatSetShadedRows: function() { that.doSetShadedRows(that.getActiveSheet()); },
        help: function() { that.showHtmlSidebar(SIDEBAR_HELP); }
      };
      var command = commands[params.command];
      command();
    }
  } catch(e) {
    if (params.hasOwnProperty('error_msg')) {
      this.alert({msg: params.error_msg})();
    }
    this.logError('Exposify.prototype.executeMenuCommand', e);
  }
} // end Exposify.prototype.executeMenuCommand


/**
 * Get the Sheet object that represents the sheet the user is currently working with.
 * @return {Sheet}
 */
Exposify.prototype.getActiveSheet = function() {
  return this.getSpreadsheet().getActiveSheet();
}; // end Exposify.prototype.getActiveSheet


/**
 * Get the Spreadsheet object that represents the spreadsheet to which Exposify is attached.
 * @return {Spreadsheet}
 */
Exposify.prototype.getActiveSpreadsheet = function() {
  return this.getSpreadsheet();
}; // end Exposify.prototype.getActiveSpreadsheet


/**
 * Determine whether this year uses a different schedule, only if September 1 falls on a Tuesday.
 * @param {number} year - The year to check.
 * @return {boolean}
 */
Exposify.prototype.getAlternateDesignationYearStatus = function(year) { // change in designation days are different if September 1 is a Tuesday (see http://senate.rutgers.edu/RLBAckS1003AAcademicCalendarPart2.pdf)
  var firstDayOfSeptember = (new Date(year, 8, 1)).getDay();
  return firstDayOfSeptember === 2 ? true : false; // return true if the first day of September of the year being checked is a Tuesday and false otherwise
} // end Exposify.prototype.getAlternateDesignationYearStatus


/**
 * Parse a Course object into a new data object for use in creating a schedule for an attendance sheet,
 * mostly by calculating the date the course begins—a complicated enough operation that I refactored it
 * into a separate function.
 * @param {Course} course - The Course object to analyze.
 */
Exposify.prototype.getCourseData = function(course) {
  try {
    var semester = course.semester; // the semester string, i.e. 'Fall 2015'
    var semesterYear = semester.match(/\d+/)[0]; // the semester string with the season removed, i.e. '2015'
    var semesterSeason = semester.match(/\D+/)[0].trim(); // the semester string with the year removed, i.e. 'Fall'
    var meetingDays = course.meetingDays;
    var meetingWeeks = 15; // spring and fall courses meet for 15 weeks (yay magic numbers)
    switch (semesterSeason) {
      case 'Spring':
        var semesterMonth = 0; // January = 0
        var semesterMonthFirstDay = (new Date(semesterYear, semesterMonth, 1)).getDay();
        var semesterBeginsDay = 15; // if the first day of January is a Tuesday, the semester begins on January 15th
        if (semesterMonthFirstDay > 2) {
            semesterBeginsDay = (7 - (semesterMonthFirstDay - 3)) + 14; // if the first day of January is after Tuesday, add 14 days to the number of days between January 1st and the following Tuesday
        } else if (semesterMonthFirstDay < 2) {
            semesterBeginsDay = (2 - semesterMonthFirstDay) + 15; // if the first day of January is before Tuesday, add 15 days to the number of days between January 1st and the next Tuesday
        }
        break;
      case 'Fall':
        var semesterMonth = 8; // September = 8
        var semesterMonthFirstDay = (new Date(semesterYear, semesterMonth, 1)).getDay();
        var semesterBeginsDay = 1; // if the first day of September is a Tuesday, the semester begins on September 1st
        if (semesterBeginsDay > 2) {
          semesterBeginsDay = (7 - (semesterMonthFirstDay - 3)); // if the first day of September is after Tuesday, calculate the date of the following Tuesday
        } else if (semesterMonthFirstDay < 2) {
            semesterBeginsDay = (2 - semesterMonthFirstDay) + 1; // if the first day of September is before Tuesday, calculate the date of the next Tuesday
        }
        break;
      case 'Summer':
        var summerSection = course.section.match(/\D/)[0]; // use a regular expression to determine the section of the course, A–V
        var session = SUMMER_SESSIONS[summerSection];
        var semesterMonth = session[1]; // check which month the course starts in
        var semesterMonthFirstDay = (new Date(semesterYear, semesterMonth, 1)).getDay();
        var semesterBeginsDay = session[2]; // check which day of the week the course starts
        meetingWeeks = session[3]; // check the duration, in weeks, of the course
        if (semesterMonthFirstDay > session[0]) {
            semesterBeginsDay = (7 - (semesterMonthFirstDay - (session[0] + 1)) + (semesterBeginsDay - 1)); // calculate the first day of the course (works assuming there is predictability to when summer courses begin)
        } else if (semesterMonthFirstDay < session[0]) {
            semesterBeginsDay = (session[0] - semesterMonthFirstDay) + semesterBeginsDay;
        }
        break;
    }
    var semesterBeginsDate = new Date(semesterYear, semesterMonth, semesterBeginsDay);
    return {semesterBeginsDate: semesterBeginsDate,
      meetingDays: meetingDays,
      meetingWeeks: meetingWeeks};
  } catch(e) {
    this.logError('Exposify.prototype.getCourseData', e);
  }
} // end Exposify.prototype.getCourseData


/**
 * Get the course folder for the gradebook on the active spreadsheet.
 * @param {Sheet} sheet - A Google Apps Sheet object, the gradebook for which we want the associated course folder
 */
Exposify.prototype.getCourseFolder = function(sheet) {
  try {
    var courseTitle = this.getCourseTitle(sheet);
    var folderIter = DriveApp.getFoldersByName(courseTitle);
    return folderIter.hasNext() ? folderIter.next() : null; // return the first match found
  } catch(e) { this.logError('Exposify.prototype.getCourseFolder', e); }
} // end Exposify.prototype.getCourseFolder


/**
 * Return the name of the course for a given gradebook.
 * @param {Sheet} sheet - The Google Apps Sheet object from which to retrieve the course name.
 * @return {string} courseTitle - The name of the course, with section number appended.
 */
Exposify.prototype.getCourseTitle = function(sheet) {
  try {
    var title = sheet.getRange('A1').getValue(); // the name of the course, from the gradebook
    var courseTitle = title.replace(/(\s\d+)?:/, ' '); // string manipulation to get a folder name friendly version of the course name and section code
    if (courseTitle === undefined) {
      var e = 'courseTitle is undefined on Exposify.prototype.getCourseTitle';
      throw e;
    }
    return courseTitle;
  } catch(e) { this.logError('Exposify.prototype.getCourseTitle', e); }
} // Exposify.prototype.getCourseTitle


/**
 * Get the API key for this script for use in client side HTML. The key is stored as a script property,
 * because we don't want end users to be able to see it.
 * @return {string} key - The API key for this app.
 */
Exposify.prototype.getDeveloperKey = function() {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('DEVELOPER_KEY');
    return key;
  } catch(e) { this.logError('Exposify.prototype.getDeveloperKey', e); }
} // end Exposify.prototype.getDeveloperKey


/**
 * Return the day on which Spring Break begins for a given year.
 * @param {number} year - The year to check.
 * @return {number}
 */
Exposify.prototype.getFirstDayOfSpringBreak = function(year) {
  var firstDayOfMarch = new Date(year, 3, 1).getDay();
  return firstDayOfMarch + (6 - firstDayOfMarch) + 7; // Spring Break starts the second Saturday of March, so find out the first day of March, add days to get to Saturday, and add 7 to that
} // end Exposify.prototype.getFirstDayOfSpringBreak


/**
 * Sanitize HTML text and return an HtmlOutput object that can be displayed to the user.
 * @param {string} html - The HTML to sanitize.
 * @return {HtmlOutput} output - The sanitized HTML object.
 */
Exposify.prototype.getHtmlOutput = function(html) {
  try {
    var output = HtmlService.createHtmlOutput(html)
      .setSandboxMode(HtmlService
      .SandboxMode.IFRAME);
    return output;
  } catch(e) { this.logError('Exposify.prototype.getHtmlOutput', e); }
} // end Exposify.prototype.getHtmlOutput


/**
 * Sanitize HTML from a file and return an HtmlOutput object that can be displayed to the user.
 * It is probably possible to generalize these two functions.
 * @param {string} filename - The name of the HTML file to sanitize.
 * @return {HtmlOutput} output - The sanitized HTML object.
 */
Exposify.prototype.getHtmlOutputFromFile = function(filename) {
  try {
    var output = HtmlService.createHtmlOutputFromFile(filename)
      .setSandboxMode(HtmlService
      .SandboxMode.IFRAME);
    return output;
  } catch(e) { this.logError('Exposify.prototype.getHtmlOutputFromFile', e); }
}; // end Exposify.prototype.getHtmlOutputFromFile


/**
 * Return the last day of a month for a given year.
 * @param {number} month - A month.
 * @param {year} year - A year.
 * @return {number}
 */
Exposify.prototype.getLastDayOfMonth = function(month, year) {
   month += 1;
   return month === 2 ? year & 3 || !(year % 25) && year & 15 ? 28 : 29 : 30 + (month + (month >> 3 ) & 1); // do some bit twiddling to figure out the last day of any given month, hard to read code courtesy of http://jsfiddle.net/TrueBlueAussie/H89X3/22/
} // end Exposify.prototype.getLastDayOfMonth


/**
 * Switch a name from "last, first" to "first last" order.
 * @param {string} name - A name in last, first order.
 * @return {string} newName - The name in first last order.
 */
Exposify.prototype.getNameFirstLast = function(name) {
  var names = name.split(','); // if name string contains a comma, assume they are in last, first order and split them at the comma
  if (names.length > 1) { // if, for some reason, the name is already first name first
    var newName = names[1].trim() + ' ' + names[0].trim(); // remove leading and trailing whitespace but add a space between them
  } else {
    var newName = names[0];
  }
  return newName;
} // end Exposify.prototype.getNameFirstLast


/**
 * Switch a name from "first last" to "last, first" order.
 * @param {string} name - A name in first last order.
 * @return {string} newName - The name in last, first order.
 */
Exposify.prototype.getNameLastFirst = function(name) {
  var names = name.split(' '); // if names are in first last order, split them at the space
  var newName = names.pop() + ', ' + names.join(' '); // insert commas between the names and add a space
  return newName;
} // end Exposify.prototype.getNameLastFirst


/**
 * Get authorization for Drive access from client side code by calling a dummy function, just in case
 * the user needs to authenticate, and then returning the necessary OAuth token.
 * @return {Object}
 * @return {string} Object.token - An OAuth token.
 * @return {string} Object.key - The Developer API key for this application.
 */
Exposify.prototype.getOAuthToken = function() {
  try {
    DriveApp.getRootFolder();
    var token = ScriptApp.getOAuthToken();
    var key = this.getDeveloperKey();
    return {token: token, key: key};
  } catch(e) { this.logError('Exposify.prototype.getOAuthToken', e); }
} // end Exposify.prototype.getOAuthToken


/**
 * Return the name of the course for a given gradebook.
 * @param {Sheet} sheet - The Google Apps Sheet object from which to retrieve the course section.
 * @return {string} courseSection - The section of the course.
 */
Exposify.prototype.getSectionTitle = function(sheet) {
  try {
    var title = sheet.getRange('A1').getValue(); // the name of the course, from the gradebook
    var re = /.+:/;
    var courseSection = title.replace(re, '');
    return courseSection;
  } catch(e) { this.logError('Exposify.prototype.getSectionTitle', e); }
} // Exposify.prototype.getSectionTitle


/**
 * Return the semester for which a given gradebook is used.
 * @param {Sheet} sheet - The Google Apps Sheet object from which to retrieve the semester.
 * @return {string} semesterTitle - The semester.
 */
Exposify.prototype.getSemesterTitle = function(sheet) {
  var semesterTitle = sheet.getRange('A2').getValue(); // the semester, from the gradebook
  return semesterTitle;
} // end Exposify.prototype.getSemesterTitle


/**
 * Return a string that concatenates a given semester with the current year.
 * @param {string} semester - A semester.
 * @return {string}
 */
Exposify.prototype.getSemesterYearString = function(semester) {
  var year = new Date().getFullYear(); // assume any given gradebook is being created for the current year (not sure if that's a good idea, but it seems likely in the vast majority of cases)
  return semester + ' ' + year; // create a string from the semester and the current year, i.e. 'Fall 2015'
} // end Exposify.prototype.getSemesterYearString


/**
 * Check whether a gradebook has been set up for a Sheet. Return true if so, false otherwise.
 * @param {Sheet} sheet - The Google Apps Sheet object to check.
 * @return {boolean}
 */
Exposify.prototype.getSheetStatus = function(sheet) {
  try {
    var props = PropertiesService.getDocumentProperties();
    var key = sheet.getName();
    var status = props.getProperty(key);
    return status === 'active' ? true : false;
  } catch(e) { this.logError('Exposify.prototype.getSheetStatus', e); }
} // end Exposify.prototype.getSheetStatus


/**
 * Retrieve student data from the gradebook and convert it into an array of Student objects.
 * @param {Sheet} sheet - The Google Apps Sheet object from which to retrieve student data.
 * @return {Array} students - A list of Student objects containing student names and email addresses.
 */
Exposify.prototype.getStudents = function(sheet) {
  try {
    var studentData = sheet.getRange(4, 1, MAX_STUDENTS, 2).getValues();
    var students = [];
    studentData.forEach(function(student) {
      if (student[0] !== '') {
        var name = student[0];
        var netid = student[1];
        students.push(new Student(name, netid));
      }
    });
    return students;
  } catch(e) { this.logError('Exposify.prototype.getStudents', e); }
} // end Exposify.prototype.getStudents


/**
 * Return the number of students in the active gradebook.
 * @param {Sheet} sheet - A Google Apps Sheet object with the gradebook to count.
 * @return {number} count - The number of students.
 */
Exposify.prototype.getStudentCount = function(sheet) {
  try {
    var count = sheet.getLastRow() - 3; // the number of students is equal to the number of rows with content minus the three header rows
    return count;
  } catch (e) { this.logError('Exposify.prototype.getStudentCount', e); }
} // end Exposify.prototype.getStudentCount


/**
 * Retrieve student names from the gradebook, in first name first order.
 * @param {Sheet} sheet - The Google Apps Sheet object from which to retrieve student data.
 * @return {Array} studentNames - A list of student names.
 */
Exposify.prototype.getStudentNames = function(sheet) {
  try {
    var students = this.getStudents(sheet);
    var studentNames = [];
    var that = this;
    students.forEach(function(student) {
      var studentName = student.name.match(/.+,.+/) ? that.getNameFirstLast(student.name) : student.name;
      studentNames.push(studentName);
    });
    return studentNames;
  } catch(e) { this.logError('Exposify.prototype.getStudentNames', e); }
} // end Exposify.prototype.getStudentNames


/**
 * Return the date on which Thanksgiving falls in November for a given year.
 * @param {number} year - A year.
 * @return {number}
 */
Exposify.prototype.getTuesdayOfThanksgivingWeek = function(year) {
  var firstDayOfNovember = new Date(year, 10, 1).getDay();
  var firstThursdayOfNovember = 1; // if first day of November is a Thursday
  if (firstDayOfNovember < 4) {
    firstThursdayOfNovember = (5 - firstDayOfNovember); // if it's before Thursday, calculate the date of the next Thursday
  } else if (firstDayOfNovember > 4) {
    firstThursdayOfNovember = 7 - (firstDayOfNovember - 5); // if it's after Thursday, calculate the date of the following Thursday
  }
  function findThanksgiving(day) {
    return day + 7 > 30 ? day : findThanksgiving(day + 7); // use recursion to continue adding seven to the memoized day variable until doing so would result in a value greater than 30, thus we have the last Thursday in November
  }
  return findThanksgiving(firstThursdayOfNovember) - 2; // the Tuesday of Thanksgiving week is the value of Thanksgiving Day minus 2 days
} // end Exposify.prototype.getTuesdayOfThanksgivingWeek


/**
 * Count the words in a set of documents according to a supplied regular expression.
 * @param {Sheet} sheet - A Google Apps Sheet object.
 * @param {RegExp} re - A regular expression to use for matching filenames.
 * @return {Array} counts - An array of objects containing the word count data.
 */
Exposify.prototype.getWordCounts = function(sheet, re) {
  try {
    var courseFolder = this.getCourseFolder(sheet);
    if (courseFolder === null || courseFolder.isTrashed()) {
      var e = 'No course folder could be found for this gradebook.';
      throw e; // throw an exception if there's no course folder present
    }
    var filesIter = courseFolder.getFiles();
    var filtered = [];
    while (filesIter.hasNext()) {
      var file = filesIter.next();
      var match = file.getName().match(re); // match each filename found in the course folder against the supplied regular expression
      if (match !== null && file.getMimeType() === 'application/vnd.google-apps.document') { // make sure the document is a Google Doc
        filtered.push(file);
      }
    }
    var counts = [];
    filtered.forEach( function(file) {
      var doc = DocumentApp.openById(file.getId()).getBody().getText();
      count = doc.split(/\s+/g).length; // simple word count
      var lastUpdated = file.getLastUpdated(); // last time the file was updated, useful to know
      var formattedDate = lastUpdated.getMonth() + '/' + lastUpdated.getDate() + '/' + lastUpdated.getFullYear();
      counts.push({document: file.getName(), count: count, lastUpdated: formattedDate});
    });
    return counts;
  } catch(e) { this.logError('Exposify.prototype.getWordCounts', e); }
} // end Exposify.prototype.getWordCounts


/**
 * Log a function and exception caught by another function to a spreadsheet on my Google Drive
 * so I can check into it. This is my primitive form of error tracking, which I presume is better
 * than nothing. This function requires the name of the calling function and the error message
 * caught by the exception handling code block. The latter is displayed to the user for reporting
 * back to me. Error tracking can be turned off by setting the ERROR_TRACKING constant to false.
 * @param {string} callingFunction - The name of the function that is logging the error.
 * @param {string} traceback - The runtime error message to record in the error log.
 */
Exposify.prototype.logError = function(callingFunction, traceback) {
  if (ERROR_TRACKING === true) {
    var spreadsheet = this.getActiveSpreadsheet();
    var logFileId = PropertiesService.getScriptProperties().getProperty('LOG_FILE_ID');
    var logs = SpreadsheetApp.openById(logFileId);
    var errorLogSheet = logs.getSheetByName(ERROR_TRACKING_SHEET_NAME);
    var date = new Date();
    var timestamp = date.toDateString() + ' ' + date.toTimeString();
    var email = spreadsheet.getOwner().getEmail();
    var id = spreadsheet.getId();
    var info = [timestamp, email, id, callingFunction, traceback];
    var pasteRange = errorLogSheet.getRange((errorLogSheet.getLastRow() + 1), 1, 1, 5);
    pasteRange.setValues([info]);
  }
  var msg = 'You can tell Steve you saw this error message, and maybe he can fix it:\n(' + errorLogSheet.getLastRow() + ') ' + traceback;
  this.alert({msg: msg})();
} // end Exposify.prototype.logError


/**
 * Record the email address of someone who installs Exposify and the Google Docs
 * spreadsheet id number of the document to which it is attached. This is intended
 * for communication and updating purposes only. It can be turned off by setting the
 * INSTALL_TRACKING constant to false.
 */
Exposify.prototype.logInstall = function() {
  if (INSTALL_TRACKING === true) {
    var spreadsheet = this.getActiveSpreadsheet();
    var logFileId = PropertiesService.getScriptProperties().getProperty('LOG_FILE_ID');
    var logs = SpreadsheetApp.openById(logFileId);
    var installLogSheet = logs.getSheetByName(INSTALL_TRACKING_SHEET_NAME);
    var date = new Date();
    var timestamp = date.toDateString() + ' ' + date.toTimeString();
    var email = spreadsheet.getOwner().getEmail();
    var id = spreadsheet.getId();
    var info = [timestamp, email, id];
    var pasteRange = installLogSheet.getRange((installLogSheet.getLastRow() + 1), 1, 1, 3);
    pasteRange.setValues([info]);
  }
} // end Exposify.prototype.logInstall


/**
 * Set a list of given grade validations as data validations on a given Sheet object.
 * @param {Sheet} sheet - A Google Apps Sheet object on which to set the data validations.
 * @param {GradeValidationSet} gradeValidations - A GradeValidationSet object containing the validation data.
 */
Exposify.prototype.setGradeValidations = function(sheet, gradeValidations) {
  try {
    gradeValidations.ranges.forEach(function(rangeList, index) {
      rangeList.forEach(function(range) { sheet.getRange(range).setDataValidation(gradeValidations.validations[index]); }); // set data validations
    });
  } catch(e) { this.logError('Exposify.prototype.setGradeValidations', e); }
} // end Exposify.prototype.setGradeValidations


/**
 * Set a property for this document indicating that a gradebook has been set up on
 * this sheet. Return true if successful, false otherwise.
 * @param {Sheet} sheet - Google Apps Sheet object for which to set the property.
 * @return {boolean}
 */
Exposify.prototype.setSheetStatus = function(sheet) {
  try {
    var props = PropertiesService.getDocumentProperties();
    var key = sheet.getName();
    var status = props.setProperty(key, 'active'); // set property as a key/value pair; the name of the Sheet is the key
    var check = this.getSheetStatus(sheet); // make sure the property was actually set
    return check === true ? true : false;
  } catch(e) { this.logError('Exposify.prototype.setSheetStatus', e); }
} // end Exposify.prototype.setSheetStatus


/**
 * Convert a CSV or Google Sheets file into a list of student names and add them to the
 * gradebook.
 * @param {string} id - The file id of the file from which to extract student names.
 */
Exposify.prototype.setupAddStudents = function(id) {
  try {
    var spreadsheet = this.getActiveSpreadsheet();
    var sheet = this.getActiveSheet();
    var file = DriveApp.getFileById(id);
    var mimeType = file.getMimeType(); // Google Sheets or csv format
    var filename = file.getName();
    var students = [];
    if (mimeType === MIME_TYPE_GOOGLE_SHEET) {
      students = this.doParseSpreadsheet(id, MIME_TYPE_GOOGLE_SHEET);
    } else if (mimeType === MIME_TYPE_CSV) {
      students = this.doParseSpreadsheet(id, MIME_TYPE_CSV);
    } else {
      this.alert({msg: ERROR_SETUP_ADD_STUDENTS_INVALID.replace('$', filename)})(); // '$' is a wildcard value that is replaced with the filename
      return;
    }
    if (students.length === 0) {
      this.alert({msg: ERROR_SETUP_ADD_STUDENTS_EMPTY.replace('$', filename)})();
    } else {
      this.doAddStudents(students, sheet);
      spreadsheet.toast(ALERT_SETUP_ADD_STUDENTS_SUCCESS.replace('$', filename), TOAST_TITLE, TOAST_DISPLAY_TIME);
    }
  } catch(e) {
    this.alert({msg: ERROR_SETUP_ADD_STUDENTS})();
    this.logError('Exposify.prototype.setupAddStudents', e);
  }
} // end Exposify.prototype.setupAddStudents


/**
 * Create a Google Contacts contact group for the students listed in the active gradebook.
 * @param {Sheet} sheet - The Google Apps Sheet object from which to retrieve student names.
 */
Exposify.prototype.setupCreateContacts = function(sheet) {
  try {
    var students = this.getStudents(sheet);
    var allContacts = ContactsApp.getContactsByEmailAddress(EMAIL_DOMAIN);
    var allContactsEmails = allContacts.map(function(contact) { return contact.getEmails()[0].getAddress(); }); // try to save time by reducing API calls
    var contactGroupTitle = this.getCourseTitle(sheet);
    var contactGroup = ContactsApp.getContactGroup(contactGroupTitle); // does this contact group already exist?
    var that = this; // thanks, JavaScript
    if (contactGroup !== null) { // if the group already exists, delete it
      contactGroup.deleteGroup();
    }
    contactGroup = ContactsApp.createContactGroup(contactGroupTitle); // create a new group
    students.forEach(function(student) { // for each Student object passed in the argument array
      var contactExists = allContactsEmails.indexOf(student.email);
      if (contactExists === -1) { // if not, create a new contact
        var name = that.getNameFirstLast(student.name).split(' ');
        var contact = ContactsApp.createContact(name[0], name[1], student.email);
      } else {
        var contact = ContactsApp.getContact(student.email);
      }
      contactGroup.addContact(contact); // this is slow :(
    });
    var msg = ALERT_SETUP_CREATE_CONTACTS_SUCCESS.replace('$', contactGroupTitle); // '$' is replaced with the name of the contact group
    this.alert({msg: msg})();
  } catch(e) { this.logError('Exposify.prototype.setupCreateContacts', e); }
} // end Exposify.prototype.setupCreateContacts


/**
 * Convert user input, collected from a dialog box, into a newly formatted gradebook.
 * @param {Object} courseInfo - User input collected into an object.
 * @param {string} courseInfo.course - A course number.
 * @param {string} courseInfo.section - A section code.
 * @param {string} courseInfo.semester - A semester name.
 * @param {Array} courseInfo.meetingDays - A list of days of the week when the class meets.
 */
Exposify.prototype.setupNewGradebook = function(courseInfo) {
  var spreadsheet = this.getActiveSpreadsheet();
  var sheet = this.getActiveSheet();
  var newName = courseInfo.course === OTHER_COURSE_NUMBER ? courseInfo.section : courseInfo.course + ':' + courseInfo.section; // only show the course number if it's real
  var exists = spreadsheet.getSheetByName(newName);
  if (exists !== null && sheet.getName() === newName) {
    var msg = ALERT_SETUP_NEW_GRADEBOOK_ALREADY_EXISTS.replace('$', newName);
    this.alert({msg: msg})(); // avoid creating a new sheet with the same name as an existing sheet
    return;
  }
  var newCourse = new Course(courseInfo); // create new Course object with information collected from the user by the dialog box
  try {
    var check = this.doFormatSheet({course: newCourse, sheet: sheet}); // do the actual work, probably in a way that I should further refactor
    if (check === true) {
      var checkStatus = this.setSheetStatus(sheet);
    }
    if (checkStatus === true) {
      spreadsheet.toast(ALERT_SETUP_NEW_GRADEBOOK_SUCCESS.replace('$', newCourse.nameSection), TOAST_TITLE, TOAST_DISPLAY_TIME); // cute pop-up window
    } else {
      this.alert({msg: ERROR_SETUP_NEW_GRADEBOOK_FORMAT})();
    }
  } catch(e) {
    this.alert({msg: ERROR_SETUP_NEW_GRADEBOOK_FORMAT})();
    this.logError('Exposify.prototype.setupNewGradebook', e);
  }
} // end Exposify.prototype.setupNewGradebook


/**
 * Display a sidebar to the user.
 * @param {Object} sidebar - An object literal constant containing data for building the sidebar.
 */
Exposify.prototype.showHtmlSidebar = function(sidebar) {
  try {
    var stylesheet = this.getHtmlOutputFromFile(STYLESHEET);
    var body = this.getHtmlOutputFromFile(sidebar.html).getContent(); // Sanitize the HTML file
    var page = stylesheet.append(body).getContent(); // Combine the style sheet with the body
    var htmlSidebar = this.getHtmlOutput(page)
      .setTitle(sidebar.title);
    this.showSidebar(htmlSidebar);
  } catch(e) { this.logError('Exposify.prototype.showHtmlSidebar', e); }
} // end Exposify.prototype.showHtmlSidebar


// FOLDERSTRUCTURE FUNCTIONS


// Create a folder hierarchy with a base folder for the semester, a section folder for shared documents, and one folder for each student for graded papers
function setupCreateFolderStructure() {
  try {
  if (alertYesNo(ALERT_SETUP_CREATE_FOLDER_STRUCTURE)) {
    var sheet = activeSheet();
    doCreateFolderStructure(sheet);
  }
  } catch(e) {
    ui.alert(ERROR_SETUP_CREATE_FOLDER_STRUCTURE);
    logError('setupCreateFolderStructure', e);
  }
}

FolderStructure.prototype.getSemesterFolder = function() {
    var folderIterator = this.rootFolder.getFoldersByName(this.semesterTitle);
    if (folderIterator.hasNext()) {
      var folder = folderIterator.next();
      if (folder.getName() === this.semesterTitle) {
        return folder;
      }
    }
    return null;
  };

FolderStructure.prototype.getCourseFolder = function() {
    if (this.semesterFolder !== null) {
      var folderIterator = this.semesterFolder.getFoldersByName(this.courseTitle);
      if (folderIterator.hasNext()) {
        var folder = folderIterator.next();
        if (folder.getName() === this.courseTitle) {
          return folder;
        }
      }
    }
    return null;
  };

FolderStructure.prototype.getGradedFolder = function() {
    if (this.semesterFolder !== null) {
      var folderIterator = this.semesterFolder.getFoldersByName(GRADED_PAPERS_FOLDER_NAME);
      if (folderIterator.hasNext()) {
        var folder = folderIterator.next();
        if (folder.getName() === GRADED_PAPERS_FOLDER_NAME) {
          return folder;
        }
      }
    }
    return null;
  };

FolderStructure.prototype.getStudentFolders = function() {
    if (this.gradedFolder !== null) {
      var studentFolders = [];
      var folderIterator = this.gradedFolder.getFolders();
      while (folderIterator.hasNext()) {
        var folder = folderIterator.next();
        studentFolders.push(folder);
      }
      return studentFolders;
    }
    return null;
  };

// Create a course folder hierarchy
// need to fix checking of Graded folder, because it asks to delete students from other sections
// this function is gargantuan at the moment, in desperate need of refactoring
function doCreateFolderStructure(sheet) {
    var semesterTitle = getSemesterTitle(sheet);
    var courseTitle = getCourseTitle(sheet);
    var folderStructure = new FolderStructure(semesterTitle, courseTitle);
    var root = folderStructure.rootFolder;
    var semesterFolder = folderStructure.getSemesterFolder();
    var courseFolder = folderStructure.getCourseFolder();
    var gradedFolder = folderStructure.getGradedFolder();
    var existingStudentFolders = folderStructure.getStudentFolders();
    var newStudents = getStudentNames(sheet);
    var createdFolders = [];
    var deletedFolders = [];
    var foldersNotCreated = [];
    var foldersToDelete = [];
    var error = 0;
    if (semesterFolder === null) { // I think I can refactor this bit into a separate function
      try {
        semesterFolder = root.createFolder(semesterTitle).setShareableByEditors(false).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
        var newSemesterFolder = new Folder(semesterTitle, root, 'My Drive/' + semesterTitle);
        createdFolders.push(newSemesterFolder);
      } catch(e) {
        if (!arrayContains(createdFolders, newSemesterFolder)) {
          foldersNotCreated.push(semesterTitle);
        }
        logError('doCreateFolderStructure', e);
        error = 1;
      }
    }
    if (courseFolder === null) {
      try {
        courseFolder = semesterFolder.createFolder(courseTitle).setShareableByEditors(false).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
        var newCourseFolder = new Folder(courseTitle, semesterFolder, 'My Drive/' + semesterTitle + '/' + courseTitle);
        createdFolders.push(newCourseFolder);
      } catch(e) {
        if (!arrayContains(createdFolders, newCourseFolder)) {
          foldersNotCreated.push(courseTitle);
        }
        logError('doCreateFolderStructure', e);
        error = 1;
      }
    }
    if (gradedFolder === null) {
      try {
        gradedFolder = semesterFolder.createFolder(GRADED_PAPERS_FOLDER_NAME).setShareableByEditors(false).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
        var newGradedFolder = new Folder(GRADED_PAPERS_FOLDER_NAME, semesterFolder, 'My Drive/' + semesterTitle + '/' + GRADED_PAPERS_FOLDER_NAME);
        createdFolders.push(newGradedFolder);
      } catch(e) {
        if (!arrayContains(createdFolders, newGradedFolder)) {
          foldersNotCreated.push(GRADED_PAPERS_FOLDER_NAME);
        }
        logError('doCreateFolderStructure', e);
        error = 1;
      }
    }
    if (existingStudentFolders !== null) {
      try {
        var updatedStudents = newStudents.slice();
        for (folder = 0; folder < existingStudentFolders.length; folder += 1) {
          var name = existingStudentFolders[folder].getName();
          if (arrayContains(newStudents, name)) {
              updatedStudents.splice(newStudents.indexOf(name), 1);
          } else {
            foldersToDelete.push(existingStudentFolders[folder]);
          }
        }
        newStudents = updatedStudents.slice();
//
//        var updatedStudents = newStudents;
//        existingStudentFolders.forEach( function(folder) {
//        var name = folder.getName();
//        if (this.contains(name)) {
//          this.splice(this.indexOf(name), 1);
//          } else if (!this.contains(name)) {
//            foldersToDelete.push(folder);
//          }
//        }, updatedStudents);
//        newStudents = updatedStudents;
      } catch(e) {
        logError('doCreateFolderStructure', e);
      }
    }
    newStudents.forEach( function(student) {
      try {
        var studentFolder = gradedFolder.createFolder(student).setShareableByEditors(false).setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
        var newStudentFolder = new Folder(student, GRADED_PAPERS_FOLDER_NAME, 'My Drive/' + semesterTitle + '/' + GRADED_PAPERS_FOLDER_NAME + '/' + student);
        createdFolders.push(newStudentFolder);
      } catch(e) {
        if (!arrayContains(createdFolders, newStudentFolder)) {
          foldersNotCreated.push(student);
        }
        logError('doCreateFolderStructure', e);
        error = 1;
      }
    });
    var deleteAlert = function() {
      var alert = 'Is it OK to delete the following student folders?\n\n';
      foldersToDelete.forEach( function(folder) { alert = alert.concat(folder.getName() + '\n'); });
      return alert;
    };
    var resultAlert = function() {
      var created = 'No new folders were created.\n';
      var deleted = '\nNo folders were trashed.\n';
      if (createdFolders.length !== 0) {
        created = 'These folders were created:\n\n';
        createdFolders.forEach( function(folder) { created = created.concat(folder.path + '\n'); });
      }
      if (deletedFolders.length !== 0) {
        deleted = '\nThese folders were trashed:\n\n';
        deletedFolders.forEach( function(folder) { deleted = deleted.concat(folder.getName() + '\n'); });
      }
      var alert = created + deleted;
      return alert;
    };
    var errorAlert = function() {
      var alert = 'There was a problem with some folders.\n\n';
      if (foldersNotCreated.length !== 0) {
        alert = alert.concat('The following folders could not be created:\n\n');
        foldersNotCreated.forEach( function(folder) { alert = alert.concat(folder + '\n'); });
      }
      if (foldersToDelete.length !== 0) {
        alert = alert.concat('The following folders could not be trashed:\n\n');
        foldersToDelete.forEach( function(folder) { alert = alert.concat(folder.getName() + '\n'); });
      }
      return alert;
    };
    if (foldersToDelete.length !== 0 && alertYesNo(deleteAlert())) {
      foldersToDelete.forEach( function(folder) {
        try {
          var parent = gradedFolder;
          parent.removeFolder(folder);
          folder.setTrashed(true);
          deletedFolders.push(folder);
          //foldersToDelete.splice(foldersToDelete.indexOf(folder), 1);
        } catch(e) {
          if (!arrayContains(deletedFolders, folder)) {
            error = 1;
          }
          logError('doCreateFolderStructure', e);
        }
      });
    }
    ui.alert(resultAlert());
    if (error === 1) {
      ui.alert(errorAlert());
    }
}

function setupShareFolders() {
  try {
  if (alertYesNo(ALERT_SETUP_SHARE_FOLDERS)) {
    var sheet = activeSheet();
    doSetupShareFolders(sheet);
  }
  } catch(e) {
    ui.alert(ERROR_SETUP_SHARE_FOLDERS);
    logError('setupShareFolders', e);
  }
}

function doSetupShareFolders(sheet) { // unshare needed for students who drop, also need to use addEditors instead of addEditor and maybe create an array of functions to call with student names
  var sheet = activeSheet();
  var students = getStudentsWithIds(sheet)
  var courseTitle = getCourseTitle(sheet);
  var folderIter = DriveApp.getFoldersByName(courseTitle);
  var courseFolder = folderIter.hasNext() ? folderIter.next() : null; // I would use Document Properties for this, but I can't be sure the folder ids won't change
  var studentsNullList = [];
  if (courseFolder === null) {
    ui.alert(ALERT_SETUP_SHARE_FOLDERS_MISSING_COURSE_FOLDER);
    return;
  }
  folderIter = DriveApp.getFoldersByName(GRADED_PAPERS_FOLDER_NAME);
  var gradedFolder = folderIter.hasNext() ? folderIter.next() : null;
  if (gradedFolder === null) {
    ui.alert(ALERT_SETUP_SHARE_FOLDERS_MISSING_GRADED_FOLDER);
    return;
  }
  students.forEach( function(student) {
    folderIter = DriveApp.getFoldersByName(student.name);
    var studentFolder = folderIter.hasNext() ? folderIter.next() : null;
    if (studentFolder === null) {
      studentsNullList.push(student.name);
    } else {
      courseFolder.addEditor(student.email);
      studentFolder.addEditor(student.email);
    }
  });
  var missingAlert = function() {
    var alert = 'The following students did not have folders:\n\n';
      studentsNullList.forEach( function(student) { alert = alert.concat(student.name + '\n'); });
      return alert;
    };
  missingAlert();
  spreadsheet.toast(ALERT_SETUP_SHARE_FOLDERS_SUCCESS, TOAST_TITLE, TOAST_DISPLAY_TIME);
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function diffPapers() {
  var id = '1YTThNXwde96tnRDTBUtEQZ18rOWzTEpDAc-Qecpd36w';
  var url = 'https://www.googleapis.com/drive/v2/files/' + id + '/revisions';
  var options = {

    'muteHttpExceptions' : true
  };
  var result = UrlFetchApp.fetch(url, options);
  Logger.log(result);
}

// Fetch revision history of document
function getRevisionHistory(id){
  var id = '1YTThNXwde96tnRDTBUtEQZ18rOWzTEpDAc-Qecpd36w';
  //var scope = 'https://www.googleapis.com/auth/drive';
  var scope = 'https://www.googleapis.com/drive/v2/files/';
  var fetchArgs = googleOAuth_('docs', scope);
  //fetchArgs.method = 'GET';
  var url = 'https://www.googleapis.com/drive/v2/files/' + id + '/revisions';
  var response = UrlFetchApp.fetch(url, fetchArgs);
  var json = JSON.parse(response);
  Logger.log(response);

  //var jsonFeed = Utilities.jsonParse(urlFetch.getContentText()).feed.entry;
  //return the revison history feed
  //return jsonFeed
}

function googleOAuth_(name, scope) {
  var oAuthConfig = UrlFetchApp.addOAuthService(name);
  oAuthConfig.setRequestTokenUrl('https://www.google.com/accounts/OAuthGetRequestToken?scope=' + scope);
  //oAuthConfig.setRequestTokenUrl('https://accounts.google.com/o/oauth2/');
  oAuthConfig.setAuthorizationUrl('https://www.google.com/accounts/OAuthAuthorizeToken');
  //oAuthConfig.setAuthorizationUrl('https://accounts.google.com/o/oauth2/auth');
  oAuthConfig.setAccessTokenUrl('https://www.google.com/accounts/OAuthGetAccessToken');
  //oAuthConfig.setAccessTokenUrl('https://www.googleapis.com/oauth2/v1/tokeninfo');
  oAuthConfig.setConsumerKey('anonymous');
  oAuthConfig.setConsumerSecret('anonymous');
  return {oAuthServiceName: name, oAuthUseToken: 'always', muteHttpExceptions : true};
}

function test1() {
  var type = DriveApp.getFileById('1b9fEFuDMXd8c4e1_AvBRY-055Z1uR0pvrOVwTgEm5eE').getMimeType();
  Logger.log(type);
}