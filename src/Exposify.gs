/**
 * Exposify
 * by Steven Syrek
 *
 * See LICENSE.txt for licensing information
 */

/**
 * @fileoverview Exposify is a Google Sheets add-on that automates a variety
 * of tasks related to the teaching of expository writing courses. Key features
 * include automatic setup of grade books, attendance records, and folder
 * hierarchies in Google Drive for organizing course sections and the return of
 * graded assignments; batch word counts of student assignments; and various
 * formatting and administrative tasks.
 * @author steven.syrek@gmail.com (Steven Syrek)
 */

/**
 * NOTE: Exposify will not work out-of-the-box if you just copy and paste the code
 * into a Google Scripts Editor. You will need to do the following:
 * 1. Create a project in your Google Developers Console and, from the Script Editor,
 * associate your script project with the project number from the console.
 * 2. Enable the Drive API and the Google Picker API in the Developers Console.
 * 3. Enable the Drive API from the "Resources : Advanced Google services" menu in the
 * Script editor.
 * 4. Create an API key and an OAuth client ID in the credentials section of your
 * project in the Developers console. Make a note of them.
 * 5. Select the "File : Project properties" menu in the Script editor and create four
 * script properties set to the following values:
 * DEVELOPER_KEY - API key from the console (API keys section).
 * CLIENT_ID - Client ID from the console (OAuth 2.0 section).
 * CLIENT_SECRET - Client secret from the console  (OAuth 2.0 section).
 * LOG_FILE_ID - The file ID of the spreadsheet you want to use for install and error logs
 * (you can skip this one and set ERROR_TRACKING and INSTALL_TRACKING to false, but if you
 * do want to track these things, make sure you create separate sheets called "Installs"
 * and "Errors" in your spreadsheet).
 */

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
var GRADED_PAPER_PREFIX = '(Graded) '; // text to prepend to the filenames of graded papers (note the trailing space!)
var ATTENDANCE_SHEET_COLUMN_WIDTH = 25; // width of columns in the attendance record part of the gradebook, 25 is the minimum recommended if you want all the dates to be visible
var COLOR_BLANK = '#ffffff'; // #ffffff is white
var COLOR_SHADED = '#ededed'; // #ededed is light grey, a nice color for contrast and also a pun on the purpose of this application
var FONT = ['verdana,sans,sans-serif']; // font for the gradebook, with fallbacks

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
var MIME_TYPE_GOOGLE_DOC = 'application/vnd.google-apps.document'

/**
 * Default text to display for various alert messages throughout the application.
 */
var ALERT_INSTALL_THANKS = 'Thanks for installing Exposify! To get started, add a section to your gradebook by selecting "Setup new Gradebook" in the Exposify "Setup" menu.';
var ALERT_ASSIGNMENTS_CREATE_TEMPLATES_SUCCESS = 'New document files successfully created for the students in section $.';
var ALERT_ASSIGNMENTS_COPY_SUCCESS = 'Successfully copied $ assignments for grading into the semester folder!';
var ALERT_ASSIGNMENTS_COPY_NOT_RETURNED = 'Some files could not be returned. Make sure all students have their own folders for graded papers and try again:';
var ALERT_ASSIGNMENTS_RETURN_SUCCESS = 'Successfully returned $ assignments!';
var ALERT_ASSIGNMENTS_NOTHING_FOUND = 'No assignments with the name "$" were found in the course folder.';
var ALERT_ASSIGNMENTS_NOTHING_RETURNED = 'No assignments were returned. Make sure you have individual folders for your students and try again.';
var ALERT_ADMIN_GENERATE_GRADEBOOK_SUCCESS = 'New gradebook file for $ successfully created in My Drive!';
var ALERT_ADMIN_GENERATE_WARNING_ROSTER_SUCCESS = 'Warning roster for section $ successfully created in My Drive!';
var ALERT_ADMIN_NO_WARNINGS = 'There are no warnings to issue for this section!';
var ALERT_NO_GRADEBOOK = 'You have not set up a gradebook yet for this sheet. You need to do that before Exposify can help with anything else.';
var ALERT_SETUP_ADD_STUDENTS_SUCCESS = '$ successfully imported! You might want to double-check the spreadsheet to make sure it is correct.';
var ALERT_SETUP_CREATE_CONTACTS_SUCCESS = 'New contact group successfully created for $!';
var ALERT_MISSING_COURSE_FOLDER = 'There is no course folder for this course. Use the Create Folder Structure command to create one before executing this command again.';
var ALERT_MISSING_GRADED_FOLDER = 'There is no graded papers folder for this course. Use the Create Folder Structure command to create one before executing this command again.';
var ALERT_MISSING_GRADED_PAPER_FOLDERS = 'There are no graded paper folders for individual students in the main graded papers folder. Use the Create Folder Structure command to create one before executing this command again.';
var ALERT_MISSING_SEMESTER_FOLDER = 'There is no semester folder for this section. Use the Create Folder Structure command to create one before executing this command again.';
var ALERT_SETUP_SHARE_FOLDERS_SUCCESS = 'The course folders for section $ were successfully shared!';
var ALERT_SETUP_NEW_GRADEBOOK_ALREADY_EXISTS = 'A gradebook for section $ already exists. If you want to overwrite it, make it the active spreadsheet and try again.';
var ALERT_SETUP_NEW_GRADEBOOK_SUCCESS = 'New gradebook created for $!';

var TOAST_DISPLAY_TIME = 10; // how long should the little toast window linger before disappearing
var TOAST_TITLE = 'Success!' // toast window title

var ERROR_INSTALL = 'There was a problem installing Exposify. I don\'t know why. I\'m sorry. Please try again.';
var ERROR_FORMAT_SET_SHADED_ROWS = 'There was a problem formatting the sheet. Probably somewhere in the "cloud." These things happen sometimes. Please try again.';
var ERROR_FORMAT_SWITCH_STUDENT_NAMES = 'There was a problem formatting the sheet. Sorry about that. Please try again. It will probably work eventually!';
var ERROR_SETUP_NEW_GRADEBOOK_FORMAT = 'There was a problem formatting the page. The cloud gods are against us today. Please try again with a supplicatory attitude.';
var ERROR_SETUP_ADD_STUDENTS = 'There was a problem reading the file. I\'m sure it was a fluke, and it\'ll work if you try a second (or third?) time.';
var ERROR_SETUP_ADD_STUDENTS_EMPTY = 'I could not find any students in the file "$" for this section. Make sure you didn\'t modify it after downloading it from Sakai and that it\'s the correct section.';
var ERROR_SETUP_ADD_STUDENTS_INVALID = '"$" is not a valid CSV or Google Sheets file. Please try again with a file that has one of those two formats.';

/**
 * The COURSE_FORMATS object literal contains the basic data used to format Exposify gradebooks,
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
      paperGrades: {
        requiredValues: ['A', 'B+', 'B', 'C+', 'C', 'NP'],
        helpText: 'Enter A, B+, B, C+, C, or NP',
        rangeToValidate: ['F4:F25', 'J4:J25', 'N4:N25', 'R4:R25', 'U4:V25']
      },
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
      /*
      numericGrades: {
        helpText: 'Enter a numeric grade from 0–100',
        rangeToValidate: ['F4:F25', 'J4:J25', 'N4:N25', 'R4:R25', 'U4:V25']
      },
      */
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
        var nonNumeric = [this.paperGrades, this.roughDraftStatus, this.lateFinalStatus, this.incompleteFinalStatus, this.proposalGrade, this.finalGrades];
        return {nonNumeric: nonNumeric}; // package and return validation data
        /*
        var nonNumeric = [this.roughDraftStatus, this.lateFinalStatus, this.incompleteFinalStatus, this.proposalGrade, this.finalGrades];
        var numeric = [this.numericGrades];
        return {nonNumeric: nonNumeric, numeric: numeric}; // package and return validation data
        */
      }
    }//,
    //finalGradeFormula: '((((F$ + J$ + N$) / 300) * .45) + ((R$ / 100) * .40) + ((U$ / 100) * .15)) * 100',
    //finalGradeFormulaRange: 'V4:V25'
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

/**
 * The SUMMER_SESSIONS object literal is a slightly obtuse way of storing information about the slightly obtuse summer session schedule.
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
    msg: 'This will replace all data on this sheet with a new gradebook. Are you sure you wish to proceed?',
    title: 'Setup New Gradebook'
  },
  dialog: {
    title: 'Setup New Gradebook',
    html: 'setupNewGradebookDialog.html',
    width: 525,
    height: 450
  },
  error_msg: 'There was a problem with the setup process. It could be a problem with the code, but it\'s probably a server issue. Please try again and see if it works eventually.'
};
var DIALOG_SETUP_ADD_STUDENTS = {
  alert: {
    alertType: YES_NO,
    msg: 'This will replace any students currently listed in this gradebook with a new student roster. Are you sure you wish to proceed?',
    title: 'Add Students to Section'
  },
  dialog: {
    title: 'Add Students to Section',
    html: 'addStudentsFilePickerDialog.html',
    width: 800,
    height: 600
  },
  error_msg: 'There was a problem accessing your Drive account. These things happen! Please try again.'
};
var DIALOG_SETUP_CREATE_CONTACTS = {
  alert: {
    alertType: YES_NO,
    msg: 'This will create or update contacts and create a contact group based on the students on this spreadsheet. Students no longer in this course will be removed from the contact group if it already exists, but the contacts themselves will not be deleted. Is this what you want to do?',
    title: 'Create Contact Group'
  },
  command: 'setupCreateContacts',
  error_msg: 'Unable to access contacts. For some reason. I don\'t know why. Please try again. It might work! Eventually!'
};
var DIALOG_SETUP_CREATE_FOLDER_STRUCTURE = {
  alert: {
    alertType: YES_NO,
    msg: 'This command will create or update the folder structure for your Expos section, including a shared coursework folder and individual folders for each student based on this sheet. Do you wish to proceed?',
    title: 'Create Folder Structure'
  },
  command: 'setupCreateFolderStructure',
  error_msg: 'There was a problem creating the folder structure. I know that\'s vague, but it might work if you try again. Sometimes the server is just slow. Please try again. Please.'
};
var DIALOG_SETUP_SHARE_FOLDERS = {
  alert: {
    alertType: YES_NO,
    msg: 'This command will share your course folder with all students in this section, and the student folders in your graded papers folder with each individual student, respectively. Do you wish to proceed?',
    title: 'Share Folders With Students'
  },
  command: 'setupShareFolders',
  error_msg: 'There was a problem sharing the folders. Sometimes the server doesn\'t like to share. Please try again.'
};
var DIALOG_ASSIGNMENTS_CREATE_TEMPLATES = {
  alert: {
    alertType: PROMPT,
    msg: 'This will create a blank document for each of the students in the current gradebook, based on a paper template designed to look nicer than the MLA format. If this is what you wish to do, enter the name of this assignment (each file will be named "Student Name [section] - Assignment Name", so if you enter "Assignment 1" and you have section AB then each filename will be "Student Name AB - Assignment 1):',
    title: 'Create Paper Templates'
  },
  command: 'assignmentsCreatePaperTemplates',
  error_msg: 'Unable to create new files. It could be the weather. Please try again.'
};
var DIALOG_ASSIGNMENTS_COPY = {
  alert: {
    alertType: PROMPT,
    msg: 'This will copy a set of assignments into the semester folder (not shared with students) for private commenting and grading. What is the name of the assignment I should look for? For example, if you enter "Assignment 1" then I will copy every document with a filename which contains that phrase, such as "Student Name AB - Assignment 1".',
    title: 'Copy Assignments for Grading'
  },
  command: 'assignmentsCopy',
  error_msg: 'Unable to copy assignments. It was probably the ionosphere. Please try again.'
};
var DIALOG_ASSIGNMENTS_RETURN = {
  alert: {
    alertType: PROMPT,
    msg: 'This will return a set of graded assignments to students in their private folders. What is the name of the assignment? For example, if you enter "Assignment 1" then I will move every document with a filename which contains that phrase, such as "Student Name AB - Assignment 1".',
    title: 'Return Graded Assignments'
  },
  command: 'assignmentsReturn',
  error_msg: 'Unable to return assignments. Are you sure you paid your taxes this year? Please try again.'
};
var DIALOG_ADMIN_WARNING_ROSTER = {
  alert: {
    alertType: YES_NO,
    msg: 'This will generate a warning roster for this section and place it in your root "My Drive" folder. Is that what you want to do?',
    title: 'Generate Warning Roster'
  },
  dialog: {
    title: 'Generate Warning Roster',
    html: 'adminWarningRoster.html',
    width: 600,
    height: 800
  },
  error_msg: 'There was a problem. I\'m sorry about that. It wasn\'t my fault. Or maybe it was. I can\'t tell from here. Please try again.'
};
var DIALOG_ADMIN_GRADEBOOK = {
  alert: {
    alertType: YES_NO,
    msg: 'This will create a separate gradebook file from the relevant data on this sheet as another Google Sheets file that you can download in Excel or another format for submission to the appropriate authorities. The file will be placed in your root "My Drive" folder. Is that what you want to do?',
    title: 'Generate Final Gradebook'
  },
  command: 'adminGenerateGradebook',
  error_msg: 'There was a problem. If I had more information, I would tell you, but it was probably the server. Or possibly God. Certainly one of those two. Please try again.'
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
  html: 'help.html'
};

/**
 * Toggles for whether or not I am tracking errors and installs and the names of the associated spreadsheets.
 */
var ERROR_TRACKING = true; // determines whether errors are sent to the error tracking spreadsheet (specified as the script property LOG_FILE_ID)
var ERROR_TRACKING_SHEET_NAME = 'Errors';
var INSTALL_TRACKING = true; // determine whether errors are sent to the install tracking spreadsheet (specified as the script property LOG_FILE_ID)
var INSTALL_TRACKING_SHEET_NAME = 'Installs';

/**
 * Text for the pre-formatted document templates used for student papers.
 */
var TEMPLATE_TITLE = 'This is the Thematic Title of My Paper: This is the Explanatory Subtitle of My Paper';
var TEMPLATE_PARAGRAPHS = [
  'This is the only document file you need for this assignment. Replace this text with your own writing, come up with your own title, and make sure the Works Cited is correct.',
  'Do not alter the formatting of this document. You do not need to create a heading, center your title, indent your paragraphs, double-space, or change the margins of your works cited—it’s all done for you already.',
  'You do not need to put your name, the date, etc. at the top of this document or add space between the end of your paper and your works cited. Seriously, I thought of everything!'
];
var TEMPLATE_WORKS_CITED = {
  author: 'Fredrickson, Barbara. ',
  title: '\'Selections from Love 2.0: How Our Supreme Emotion Affects Everything We Feel, Think, Do, and Become.\' ',
  volume: 'The New Humanities Reader. ',
  info: '5th ed. Stamford, CT: Cengage, 2015. 105–128. Print.'
};

// TRIGGER FUNCTIONS

/**
 * Execute as a trigger whenever the application is installed as an add-on to a Google Spreadsheet.
 */
function onInstall(e) {
  try {
    onOpen(e); // setup the custom menu, which is really the only important thing this function does
    expos.showHtmlSidebar(SIDEBAR_HELP);
    expos.alert({msg: ALERT_INSTALL_THANKS})();
    expos.logInstall(); // tell me when someone has installed the add-on, for my records
  } catch(e) {
    expos.alert({msg: ERROR_INSTALL})();
    expos.logError('onInstall', e); // tell me when something goes wrong, so I can fix things
  }
} // end onInstall

/**
 * Execute as a trigger whenever the attached Google Spreadsheet is opened. Add the custom
 * Exposify menu to the menu bar. Menu commands call the specified function, which passes control
 * to the command handler function, Exposify.prototype.executeMenuCommand.
 */
function onOpen() {
  var ui = expos.ui;
  var menu = expos.getMenu();
  try {
    menu
      .addSubMenu(ui.createMenu('Setup')
        .addItem('Setup new gradebook...', 'exposifySetupNewGradebook')
        .addItem('Add students to gradebook...', 'exposifySetupAddStudents')
        .addItem('Create contact group for students...', 'exposifySetupCreateContacts')
        .addItem('Create or update folder structure for this section...', 'exposifySetupCreateFolderStructure')
        .addItem('Share folders with students...', 'exposifySetupShareFolders'))
      .addSubMenu(ui.createMenu('Assignments')
        .addItem('Create paper templates for students...', 'exposifyCreatePaperTemplates')
        .addItem('Copy assignments for grading...', 'exposifyAssignmentsCopy')
        .addItem('Return graded assignments...', 'exposifyAssignmentsReturn')
        .addItem('Calculate word counts...', 'exposifyAssignmentsCalcWordCounts'))
      .addSubMenu(ui.createMenu('Administration')
        .addItem('Generate warning roster for this section...', 'exposifyAdminGenerateWarningRoster')
        .addItem('Generate final gradebook for this section...', 'exposifyAdminGenerateGradebook'))
      .addSubMenu(ui.createMenu('Format')
        .addItem('Switch order of student names', 'exposifyFormatSwitchStudentNames')
        .addItem('Refresh shading of alternating rows', 'exposifyFormatSetShadedRows'))
      .addSeparator()
      .addItem('Exposify help...', 'exposifyHelp')
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
 * A simple student record, containing the student's name and netid, which can be computed into
 * a valid email address. Assumes all emails have the same domain, but this can be modified for edge
 * cases using the {@code setEmail()} method.
 * @constructor
 */
function Student(name, netid) {
  this.name = name;
  this.netid = netid;
  this.email = (EMAIL_REGEX.test(netid + EMAIL_DOMAIN) === true ? netid + EMAIL_DOMAIN : ''); // make sure the email address is valid
}; // end Student

/**
 * Constructor for a course format, which describes how to reformat the spreadsheet to create a new, blank gradebook.
 * @constructor
 */
function Format(course) {
  this.course = course;
  this.courseNumber = course.number;
  this.courseTitle = course.nameSection;
  this.sheetName = course.numberSection;
  this.courseFormat = COURSE_FORMATS[this.courseNumber];
  this.section = course.section;
  this.semester = course.semester;
  this.rows = course.rows;
  this.lastRow = course.rows.length;
  this.columns = course.columns;
  this.lastColumn = course.columns.length;
  this.columnHeadings = course.columnHeadings;
  this.meetingDays = course.meetingDays;
  this.gradeValidations = course.gradeValidations === undefined ? false : true;
  this.setRowHeights = function(sheet) { this.rows.map(function(row, index) { sheet.setRowHeight(index + 1, this.rows[index]); }, this); } // set row heights
  this.setColumnWidths = function(sheet) {this.columns.map(function(column, index) { sheet.setColumnWidth(index + 1, this.columns[index]); }, this); } // set column widths
  this.shadedRows = false;
} // end Format

/**
 * The main Exposify constructor, a namespace for most (but not all) of the methods and properties of the add-on. This is probably overkill, but it seemed like a good idea at the time.
 * @constructor
 */
function Exposify() {
  // Private properties
  /**
   * Store a reference to the active Spreadsheet object, which shouldn't vary after the object is created.
   * @private {Spreadsheet}
   */
  var spreadsheet_ = SpreadsheetApp.getActiveSpreadsheet();
  /**
   * Store a reference to the Ui object for this spreadsheet, which shouldn't vary after the object is created.
   * @private {Ui}
   */
  var ui_ = SpreadsheetApp.getUi();
  /**
   * Store a reference to the Menu object for this spreadsheet.
   * @private {Menu}
   */
  var menu_ = ui_.createAddonMenu();
  /**
   * Store references to common UI button sets, so we don't have to look them up at runtime.
   * @private {ButtonSet}
   */
  var ok_ = ui_.ButtonSet.OK;
  var okCancel_ = ui_.ButtonSet.OK_CANCEL;
  var yes_ = ui_.Button.YES;
  var yesNo_ = ui_.ButtonSet.YES_NO;
  var prompt_ = ui_.Button.OK;
  // Protected methods
  /**
   * Return the active Spreadsheet object.
   * @protected
   * @return {Spreadsheet} spreadsheet_ - A Google Apps Spreadsheet object.
   */
  this.getSpreadsheet = function() { return spreadsheet_; };
  this.spreadsheet = this.getSpreadsheet();
  /**
   * Return the active Sheet object.
   * @protected
   * @return {Sheet} The Sheet object representing the active sheet.
   */
  this.getSheet = function() { return spreadsheet_.getActiveSheet(); }
  this.sheet = this.getSheet();
  /**
   * Return the Ui object for this spreadsheet.
   * @protected
   * @return {Ui} ui_ - The Ui object for the Spreadsheet object to which Exposify is attached.
   */
  this.getUi = function() { return ui_; };
  this.ui = this.getUi();
  /**
   * Return the Menu object for this spreadsheet.
   * @protected
   * @return {Menu} menu_ - The Menu object for the Spreadsheet object to which Exposify is attached.
   */
  this.getMenu = function() {return menu_; };
  this.menu = this.getMenu();
  /**
   * Set the default time zone for the spreadsheet. Return the spreadsheet for chaining.
   * @protected
   * @param {string} timezone - A string representing a timezone in "long" format, as listed by Joda.org
   * @return {Spreadsheet} spreadsheet_ - The Spreadsheet object to which Exposify is attached.
   */
  this.setTimezone = function(timezone) {
    spreadsheet_.setSpreadsheetTimeZone(timezone);
    return spreadsheet_;
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
   * This is a simple interface wrapper for accessing the built-in UI alert controls.
   * @protected {Object}
   */
  this.alertUi = {
    ok: ok_,
    okCancel: okCancel_,
    yes: yes_,
    yesNo: yesNo_,
    prompt: prompt_
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
function exposifySetupCreateFolderStructure() { expos.checkSheetStatus.call(expos, DIALOG_SETUP_CREATE_FOLDER_STRUCTURE); }
function exposifySetupShareFolders() { expos.checkSheetStatus.call(expos, DIALOG_SETUP_SHARE_FOLDERS); }
function exposifyCreatePaperTemplates() { expos.checkSheetStatus.call(expos, DIALOG_ASSIGNMENTS_CREATE_TEMPLATES); }
function exposifyAssignmentsCalcWordCounts() { expos.checkSheetStatus.call(expos, {command: 'assignmentsCalcWordCounts'}); }
function exposifyAssignmentsCopy() { expos.checkSheetStatus.call(expos, DIALOG_ASSIGNMENTS_COPY); }
function exposifyAssignmentsReturn() { expos.checkSheetStatus.call(expos, DIALOG_ASSIGNMENTS_RETURN); }
function exposifyAdminGenerateWarningRoster() { expos.checkSheetStatus.call(expos, DIALOG_ADMIN_WARNING_ROSTER); }
function exposifyAdminGenerateGradebook() { expos.checkSheetStatus.call(expos, DIALOG_ADMIN_GRADEBOOK); }
function exposifyFormatSwitchStudentNames() { expos.checkSheetStatus.call(expos, {command: 'formatSwitchStudentNames', error_msg: ERROR_FORMAT_SWITCH_STUDENT_NAMES}); }
function exposifyFormatSetShadedRows() { expos.checkSheetStatus.call(expos, {command: 'formatSetShadedRows', error_msg: ERROR_FORMAT_SET_SHADED_ROWS}); }
function exposifyHelp() { expos.executeMenuCommand.call(expos, {command: 'help'}); }

// CALLBACKS

/**
 * As with the menu commands, callbacks from user interfaces (from the client side) that use the
 * Google API must call functions in the global namespace, at least as far as I can tell. These
 * functions pass control to other functions that do the actual work. In the future, I may replace
 * these functions with a single callback handler, as with the menu functions, if it's possible.
 */
function assignmentsCalcWordCountsCallback(params) { return expos.assignmentsCalcWordCounts(expos.sheet, params); }
function assignmentsCalcWordCountsCallbackGetTitle() { return expos.assignmentsCalcWordCountsGetTitle(expos.sheet); }
function getOAuthToken() { return expos.getOAuthToken(); }
function setupNewGradebookCallback(courseInfo) { expos.setupNewGradebook(expos.sheet, courseInfo); }
function setupAddStudentsCallback(id) { expos.setupAddStudents(expos.sheet, id); }
function adminGenerateWarningRosterCallback(warnings) { expos.adminGenerateWarningRoster(warnings); }
function adminGenerateWarningRosterCallbackGetStudents() { return expos.adminGenerateWarningRosterGetStudents(expos.sheet); }

// EXPOSIFY FUNCTIONS

/**
 * Generate a warning roster based on information collected from the user about which students should receive which warning.
 * @param {Object} warnings - An object containing arrays of objects corresponding to the three warning codes.
 */
Exposify.prototype.adminGenerateWarningRoster = function(warnings) {
  try {
    var w1 = warnings['w1'];
    var w2 = warnings['w2'];
    var w3 = warnings['w3'];
    if (w1.length === 0 && w2.length === 0 && w3.length === 0) {
      var alert = this.alert({msg: ALERT_ADMIN_NO_WARNINGS, title: 'Generate Warning Roster'});
      alert();
      return;
    }
    var spreadsheet = this.spreadsheet;
    var sheet = this.sheet;
    var semester = this.getSemesterTitle(sheet);
    var section = this.getSectionTitle(sheet);
    var semesterFolder = this.getSemesterFolder(sheet);
    if (semesterFolder !== null) {
      var instructor = semesterFolder.getOwner().getName();
    }
    if (instructor === null) {
      var email = spreadsheet.getOwner().getEmail();
      var instructor = ContactsApp.getContact(email).getFullName();
    }
    if (instructor === null) {
      var instructor = '';
    }
    var course = '01:355:' + this.getCourseNumber(sheet) + ':' + section;
    var title = this.getCourseTitle(sheet) + ' - Warnings Roster - ' + semester;
    var styles = {};
    styles[DocumentApp.Attribute.BOLD] = false;
    styles[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
    styles[DocumentApp.Attribute.LINE_SPACING] = 1.3;
    styles[DocumentApp.Attribute.INDENT_FIRST_LINE] = 0;
    styles[DocumentApp.Attribute.FONT_FAMILY] = 'Georgia';
    styles[DocumentApp.Attribute.FONT_SIZE] = 10;
    styles[DocumentApp.Attribute.MARGIN_BOTTOM] = 72;
    styles[DocumentApp.Attribute.MARGIN_LEFT] = 72;
    styles[DocumentApp.Attribute.MARGIN_RIGHT] = 72;
    styles[DocumentApp.Attribute.MARGIN_TOP] = 72;
    var document = DocumentApp.create(title);
    var body = document.getBody();
    body.setAttributes(styles);
    var first = body.getParagraphs()[0];
    first.setText('Warnings Roster – ' + semester);
    body.appendParagraph('Instructor: ' + instructor);
    body.appendParagraph('Course and Section Number: ' + course);
    var tableCells = [
      ['Last Name', 'First Name', 'RUID', 'Warning']
    ];
    w1.forEach(function(student) { tableCells.push([student.last, student.first, student.id, 'W1']); });
    w2.forEach(function(student) { tableCells.push([student.last, student.first, student.id, 'W2']); });
    w3.forEach(function(student) { tableCells.push([student.last, student.first, student.id, 'W3']); });
    body.appendTable(tableCells);
    var pars = body.getParagraphs();
    var last = pars[pars.length - 1];
    last.setText('W1 Warning for poor performance (e.g. non-passing and/or missing work.)');
    body.appendParagraph('W2 Warning for poor attendance or "never attended."');
    body.appendParagraph('W3 Warning for poor performance and poor attendance.');
    spreadsheet.toast(ALERT_ADMIN_GENERATE_WARNING_ROSTER_SUCCESS.replace('$', section), TOAST_TITLE, TOAST_DISPLAY_TIME); // cute pop-up window
  } catch(e) { this.logError('Exposify.prototype.adminGenerateWarningRoster', e); }
} // end Exposify.prototype.adminGenerateWarningRoster

/**
 * Retrieve the names and ids of the students in this section, used as a callback to client-side code.
 * @param {Sheet} sheet - The active Sheet object.
 * @return {Array} students - An array of Student objects.
 */
Exposify.prototype.adminGenerateWarningRosterGetStudents = function(sheet) {
  try {
    var students = this.getStudents(sheet);
    var that = this;
    students.forEach(function(student) {
      student.name = that.getNameFirstLast(student.name);
    });
    students.sort(function (a, b) {
      if (a.name > b.name) {
        return 1;
      }
      return -1;
    });
    return students;
  } catch(e) { this.logError('Exposify.prototype.adminGenerateWarningRosterGetStudents', e); }
} // end Exposify.prototype.adminGenerateWarningRosterGetStudents

/**
 * Check that an incoming request to make an alert has the correct parameters and raise an
 * exception if it does not. The parameter is an object with two fields, one containing the type of alert
 * and one containing the message to be displayed to the user. The available alert types are OK, OK_CANCEL,
 * and YES_NO. These are defined as constant values. This function returns another function, which can be
 * executed to display the dialog box, i.e. with {@code alert(confirmation)();}.
 * @param {Object} confirmation - The object holding the arguments to the function.
 * @param {string} confirmation.string - Alert type of the alert dialog, which determines the buttons to display.
 * @param {string} confirmation.msg - The message to display in the alert dialog.
 * @return {Function} alertFunction - A function that can be executed to display the alert.
 */
Exposify.prototype.alert = function(confirmation) {
  try {
    if (!confirmation.hasOwnProperty('alertType')) {
      confirmation.alertType = OK;
      return this.doMakeAlert(confirmation); // A simple alert with an OK button is the default
    } else if (!this.alertUi.hasOwnProperty(confirmation.alertType)) {
      var e = 'Alert type ' + confirmation.alertType + ' is not defined on Exposify.';
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
  try {
    if (arr.length < 1) {
      return false;
    }
    var i = arr.length;
    while (i--) {
      if (arr[i] === item) {
        return true;
      }
    }
    return false;
  } catch(e) {this.logError('Exposify.prototype.arrayContains', e); }
} // end Exposify.prototype.arrayContains

/**
 * Return word counts for the set of files specified by the user.
 * @param {Sheet} sheet - The Sheet object.
 * @param {Object} params - An object containing the parameters passed to the callback function.
 * @param {string} params.students - Either "selected" or "all," the option to search all papers or a subset.
 * @param {string} params.filter - The filter for the files, optionally entered by the user.
 */
Exposify.prototype.assignmentsCalcWordCounts = function(sheet, params) {
  try {
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
 * @param {Sheet} sheet - The Sheet object.
 * @return {string} title - The course title and number of students enrolled.
 */
Exposify.prototype.assignmentsCalcWordCountsGetTitle = function(sheet) {
  try {
    var courseTitle = this.getCourseTitle(sheet);
    var enrollment = this.getStudentCount(sheet);
    var title = courseTitle + '<br />(' + enrollment + ' students)';
    return title;
  } catch(e) { this.logError('Exposify.prototype.assignmentsCalcWordCountsGetTitle', e); }
} // end Exposify.prototype.assignmentsCalcWordCountsGetTitle

/**
 * Copy student assignments from the course folder to the semester folder for private grading.
 * @param  {Sheet} sheet - The sheet object.
 * @param  {string} assignment - The name of the assignment to filter for.
 */
Exposify.prototype.assignmentsCopy = function(sheet, assignment) {
  try {
    var spreadsheet = this.spreadsheet;
    var courseFolder = this.getCourseFolder(sheet);
    var semesterFolder = this.getSemesterFolder(sheet);
    if (courseFolder === null) {
      var alert = this.alert({msg: ALERT_MISSING_COURSE_FOLDER, title: 'Copy Assignments for Grading'});
      alert();
      return;
    } else if (semesterFolder === null) {
      var alert = this.alert({msg: ALERT_MISSING_SEMESTER_FOLDER, title: 'Copy Assignments for Grading'});
      alert();
      return;
    }
    var regex = '.+' + assignment.trim() + '.*'; // the filename contains the name of the assignment somewhere
    var re = new RegExp(regex);
    var type = MIME_TYPE_GOOGLE_DOC;
    var papers = this.getMatchedFiles(courseFolder, re, type);
    if (papers.length === 0) {
      var alert = this.alert({msg: ALERT_ASSIGNMENTS_NOTHING_FOUND.replace('$', assignment), title: 'Copy Assignments for Grading'});
      alert();
      return;
    }
    var number = 0;
    papers.forEach(function(paper) {
      paper.makeCopy(semesterFolder);
      number += 1;
      });
    spreadsheet.toast(ALERT_ASSIGNMENTS_COPY_SUCCESS.replace('$', number), TOAST_TITLE, TOAST_DISPLAY_TIME);
  } catch(e) { this.logError('Exposify.prototype.assignmentsCopy', e); }
} // end Exposify.prototype.assignmentsCopy

/**
 * Create Google Docs files for students to use as templates for their assignments.
 * @param {Sheet} sheet - The Sheet object.
 * @param {string} assignment - The name of the assignment to use in the template filename.
 */
Exposify.prototype.assignmentsCreatePaperTemplates = function(sheet, assignment) {
  try {
    var spreadsheet = this.spreadsheet;
    var that = this;
    var students = this.getStudentNames(sheet);
    var folder = this.getCourseFolder(sheet);
    var section = this.getSectionTitle(sheet);
    var files = [];
    students.forEach(function(student) {
      var fileName = student + ' ' + section + ' - ' + assignment;
      var document = that.doMakeNewTemplate(fileName);
      Utilities.sleep(100); // just in case
      var id = document.getId();
      files.push(DriveApp.getFileById(id));
    });
    files.forEach(function(file) {
      folder.addFile(file);
      DriveApp.removeFile(file);
    });
    spreadsheet.toast(ALERT_ASSIGNMENTS_CREATE_TEMPLATES_SUCCESS.replace('$', section), TOAST_TITLE, TOAST_DISPLAY_TIME); // cute pop-up window
  } catch(e) { this.logError('Exposify.prototype.assignmentsCreatePaperTemplates', e); }
} // end Exposify.prototype.assignmentsCreatePaperTemplates

/**
 * Return student assignments from the semester folder to their individual, private folders for review.
 * @param {Sheet} sheet - The sheet object.
 * @param {string} assignment - The name of the assignment to filter for.
 */
Exposify.prototype.assignmentsReturn = function(sheet, assignment) {
  try {
    var spreadsheet = this.spreadsheet;
    var semesterFolder = this.getSemesterFolder(sheet);
    var gradedPapersFolder = this.getGradedPapersFolder(sheet);
    if (semesterFolder === null) {
      var alert = this.alert({msg: ALERT_MISSING_SEMESTER_FOLDER, title: 'Return Graded Assignments'});
      alert();
      return;
    } else if (gradedPapersFolder === null) {
      var alert = this.alert({msg: ALERT_MISSING_GRADED_FOLDER, title: 'Return Graded Assignments'});
      alert();
      return;
    }
    var section = this.getSectionTitle(sheet);
    var regex = '.+' + assignment.trim() + '.*'; // the filename contains the name of the assignment somewhere
    var re = new RegExp(regex);
    var type = MIME_TYPE_GOOGLE_DOC;
    var papers = this.getMatchedFiles(semesterFolder, re, type);
    if (papers.length === 0) {
      var alert = this.alert({msg: ALERT_ASSIGNMENTS_NOTHING_FOUND.replace('$', assignment), title: 'Return Graded Assignments'});
      alert();
      return;
    }
    notMoved = [];
    var that = this;
    papers.forEach(function(paper) {
      var re = new RegExp('Copy of');
      var fileName = paper.getName();
      if (fileName.match(re) !== null) {
        var fileName = fileName.slice(8);
      }
      paper.setName(GRADED_PAPER_PREFIX + fileName);
      var re = new RegExp(section);
      var folderName = fileName.slice(0, fileName.match(re).index).trim();
      var studentFolder = that.getFolder(gradedPapersFolder, folderName);
      if (studentFolder === null) {
        notMoved.push(paper);
      } else {
        studentFolder.addFile(paper);
        semesterFolder.removeFile(paper);
      }
    });
    if (notMoved.length > 0) {
      var list = notMoved.map(function(file) { return file.getName() + '\n'; });
      var text = ALERT_ASSIGNMENTS_COPY_NOT_RETURNED + '\n\n' + list.join('');
      var alert = this.alert({msg: text, title: 'Return Graded Assignments'});
      alert();
    } else if (notMoved.length === papers.length) {
      var alert = this.alert({msg: ALERT_ASSIGNMENTS_NOTHING_RETURNED, tite: 'Return Graded Assignments'});
      alert();
    } else if (notMoved.length < papers.length) {
      var number = papers.length;
      spreadsheet.toast(ALERT_ASSIGNMENTS_RETURN_SUCCESS.replace('$', number), TOAST_TITLE, TOAST_DISPLAY_TIME);
    }
  } catch(e) { this.logError('Exposify.prototype.assignmentsReturn', e); }
} // end Exposify.prototype.assignmentsReturn

/**
 * Check whether a gradebook has already been set up for this sheet. If so, pass
 * control to the {@code executeMenuCommand()} function. Return false otherwise.
 * @param {Object} params - The parameters passed to the original menu command.
 */
Exposify.prototype.checkSheetStatus = function(params) {
  try {
    var sheet = this.sheet;
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
 * @param {string} dialog.html - The HTML file to process and display as the dialog box content.
 * @param {number} dialog.width - The width of the dialog box.
 * @param {number} dialog.height - The height of the dialog box.
 * @return {HtmlOutput} htmlDialog - The sanitized html dialog box, ready to be displayed to the user.
 */
Exposify.prototype.createHtmlDialogFromFile = function(dialog) {
  try {
    var stylesheet = this.getHtmlOutputFromFile(STYLESHEET);
    var body = this.getHtmlOutputFromFile(dialog.html).getContent(); // Sanitize the HTML file
    var page = stylesheet.append(body).getContent(); // Combine the style sheet with the body
    var htmlDialog = this.getHtmlOutput(page)
      .setWidth(dialog.width)
      .setHeight(dialog.height);
    return htmlDialog;
  } catch(e) { this.logError('Exposify.prototype.createHtmlDialogFromFile', e); }
} // end Exposify.prototype.createHtmlDialogFromFile

/**
 * Create a dialog box to display to the user using information stored as a template an object literal
 * constant. The html field of the argument object should be raw HTML.
 * @param {Object} dialog - The template containing the specification of the dialog box.
 * @param {string} dialog.title - The title of the dialog box to display.
 * @param {string} dialog.html - The HTML content to process and display as the dialog box content.
 * @param {number} dialog.width - The width of the dialog box.
 * @param {number} dialog.height - The height of the dialog box.
 * @return {HtmlOutput} htmlDialog - The sanitized html dialog box, ready to be displayed to the user.
 */
Exposify.prototype.createHtmlDialogFromText = function(dialog) {
  try {
    var stylesheet = this.getHtmlOutputFromFile(STYLESHEET);
    var body = this.getHtmlOutput(dialog.html).getContent(); // Sanitize the HTML file
    var page = stylesheet.append(body).getContent(); // Combine the style sheet with the body
    var htmlDialog = this.getHtmlOutput(page)
      .setWidth(dialog.width)
      .setHeight(dialog.height);
    return htmlDialog;
  } catch(e) { this.logError('Exposify.prototype.createHtmlDialogFromText', e); }
} // end Exposify.prototype.createHtmlDialogFromText

/**
 * Insert student names into the spreadsheet.
 * @param {Object} params - Object containing the function parameters.
 * @param {Array} params.students - A list of Student objects containing the data to add to the spreadsheet.
 * @param {Sheet} params.sheet - A Google Apps Sheet object, the sheet to which names will be added.
 */
Exposify.prototype.doAddStudents = function (params) {
  try {
    var sheet = params.sheet;
    var students = params.students;
    var studentList = [];
    var fullRange = sheet.getRange(4, 1, MAX_STUDENTS, 2);
    fullRange.clearContent(); // erase whatever data is already on the sheet where we put student names
    var range = sheet.getRange(4, 1, students.length, 2); // get a range of two columns and a number of rows equal to the number of students to insert
    students.forEach(function(student) { studentList.push([student.name, student.netid]); } ); // add a row to the temporary studentList array for each student
    range.setValues(studentList); // set the value of the whole range at once, so I don't call the API more than necessary
  } catch(e) { this.logError('Exposify.prototype.doAddStudents', e); }
} // end Exposify.prototype.doAddStudents

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
 * Format a spreadsheet sheet for use as a gradebook for a specified course.
 * @param {Object} newCourse - Information about the new course on which to base the formatting.
 * @param {Course} newCourse.course - The course information requested from the user.
 * @param {Sheet} newCourse.sheet - The sheet to format.
 */
Exposify.prototype.doFormatSheet = function(newCourse) {
  try {
    var sheet = newCourse.sheet;
    var course = newCourse.course;
    var format = new Format(course);
    format.setShadedRows().apply(sheet);
    return true
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
 * Create a new Google Sheets file using only the gradebook portion of the active sheet,
 * for downloading in another format.
 * @param {Sheet} sheet - The Sheet object.
 */
Exposify.prototype.doGenerateGradebook = function(sheet) {
  try {
    var spreadsheet = this.spreadsheet;
    var title = this.getCourseTitle(sheet);
    var section = this.getSectionTitle(sheet);
    var semester = this.getSemesterTitle(sheet);
    var courseNumber = this.getCourseNumber(sheet);
    var rows = this.getStudentCount(sheet) + 3; // number of students plus the heading is the row delimiter for the spreadsheet to export
    var columns = COURSE_FORMATS[courseNumber].columns.length; // use the number of columns specified in COURSE_FORMATS as the column delimiter
    var exportRange = sheet.getRange(1, 1, rows, columns);
    var values = exportRange.getValues();
    var newSpreadsheet = SpreadsheetApp.create(title, MAX_STUDENTS + 3, columns);
    var importSheet = newSpreadsheet.getSheets()[0];
    var importRange = importSheet.getRange(1, 1, rows, columns);
    var courseInfo = {
      course: courseNumber,
      section: section,
      semester: semester,
      meetingDays: []
    };
    var course = new Course(courseInfo);
    var format = new Format(course);
    format.apply(importSheet); // make sure the new sheet looks pretty
    importRange.setValues(values); // copy the relevant data from this sheet to the new spreadsheet, starting at the top left corner
    spreadsheet.toast(ALERT_ADMIN_GENERATE_GRADEBOOK_SUCCESS.replace('$', title), TOAST_TITLE, TOAST_DISPLAY_TIME);
  } catch(e) { this.logError('Exposify.prototype.doGenerateGradebook', e); }
} // end Exposify.prototype.doGenerateGradebook

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
    var ui = this.ui;
    var title = (confirmation.hasOwnProperty('title') ? confirmation.title : ALERT_TITLE_DEFAULT);
    var alertUi = this.alertUi;
    var ok = alertUi.ok;
    var yes = alertUi.yes;
    var okCancel = alertUi.okCancel;
    var yesNo = alertUi.yesNo;
    var prompt = alertUi.prompt;
    var alerts = { // Map alert functions to different alert types
      ok: function() { return ui.alert(title, msg, ok); },
      okCancel: function() { return (ui.alert(title, msg, okCancel)) === ok ? true : false; },
      yesNo: function() { return (ui.alert(title, msg, yesNo)) === yes ? true : false; },
      prompt: function() {
        var response = ui.prompt(title, msg + '\n\n', okCancel);
        return response.getSelectedButton() === prompt ? response.getResponseText() : false;
      }
    };
    var dialog = alerts[alertType]; // Create a function using the closures stored in the {@code alerts} variable.
    return dialog; // Return the function without executing it.
  } catch(e) { this.logError('Exposify.prototype.doMakeAlert', e); }
} // end Exposify.prototype.doMakeAlert

/**
 * Make Google Apps Data Validation objects for applying grade validation to a new gradebook.
 * @param {number} courseNumber - The course number, used to pull data from the course template object literal.
 * @return {GradeValidationSet} gradeValidations - The grade validations to apply and the ranges to which to apply them.
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
 * Make a new Google Docs file for use as a paper template.
 * @param {string} title - The title of the Docs file.
 * @return {Document} document - The Google Apps Document object.
 */
Exposify.prototype.doMakeNewTemplate = function(title) {
  try {
    var styles = {};
    styles[DocumentApp.Attribute.BOLD] = false;
    styles[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
    styles[DocumentApp.Attribute.LINE_SPACING] = 1.29;
    styles[DocumentApp.Attribute.SPACING_AFTER] = 11.5;
    styles[DocumentApp.Attribute.INDENT_FIRST_LINE] = 0;
    var bodyStyles = {}
    bodyStyles[DocumentApp.Attribute.FONT_FAMILY] = 'Garamond';
    bodyStyles[DocumentApp.Attribute.FONT_SIZE] = 12;
    bodyStyles[DocumentApp.Attribute.MARGIN_BOTTOM] = 144;
    bodyStyles[DocumentApp.Attribute.MARGIN_LEFT] = 108;
    bodyStyles[DocumentApp.Attribute.MARGIN_RIGHT] = 144;
    bodyStyles[DocumentApp.Attribute.MARGIN_TOP] = 144;
    var indent = {}
    indent[DocumentApp.Attribute.INDENT_START] = 18;
    var that = this;
    var document = DocumentApp.create(title);
    var body = document.getBody();
    body.setAttributes(bodyStyles);
    var title = body.getParagraphs()[0].editAsText();
    title.setAttributes(styles);
    title.setText(TEMPLATE_TITLE).setBold(true);
    var bodyText = TEMPLATE_PARAGRAPHS;
    bodyText.forEach(function(paragraph) {
      var bodyParagraph = body.appendParagraph(paragraph).editAsText();
      bodyParagraph.setAttributes(styles);
    });
    body.appendParagraph('Works Cited').setAttributes(styles);
    var worksCited = body.appendParagraph(TEMPLATE_WORKS_CITED.author + TEMPLATE_WORKS_CITED.title);
    worksCited.appendText(TEMPLATE_WORKS_CITED.volume).setItalic(true);
    worksCited.appendText(TEMPLATE_WORKS_CITED.info).setItalic(false);
    worksCited.editAsText().setAttributes(styles).setAttributes(indent);
    body.appendParagraph('').setAttributes(styles).setIndentStart(0);
    return document;
  } catch(e) { this.logError('Exposify.prototype.doMakeNewTemplate', e); }
} // end Exposify.prototype.doMakeNewTemplate

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
 * @param {Object} params - Object containing the function parameters.
 * @param {Sheet} params.sheet - The Sheet object.
 * @param {string} params.id - The id of the Google Apps File object to open.
 * @param {string} params.mimeType - The mimeType of the Google Apps File object to open.
 * @return {Array} students - A list of Student objects, the students to add to the gradebook.
 */
Exposify.prototype.doParseSpreadsheet = function(params) {
  try {
    var sheet = params.sheet;
    var id = params.id;
    var mimeType = params.mimeType;
    var students = [];
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
 Exposify.prototype.doSetShadedRows = function(sheet) {
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
Exposify.prototype.doSwitchStudentNames = function(sheet) {
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
    var params = {sheet: sheet, students: students};
    this.doAddStudents(params); // repopulate the sheet with the student names
    sheet.sort(1); // sort sheet alphabetically
    this.doSetShadedRows(sheet); // because the sort will probably mess them up
  } catch(e) { this.logError('Exposify.prototype.doSwitchStudentNames', e); }
} // end Exposify.prototype.doSwitchStudentNames

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
      if (response === false) { return; }
    }
    if (params.hasOwnProperty('dialog') && response === true) { // show the dialog, if there is one, and if the alert response is true
      var dialog = params.dialog;
      var htmlDialog = this.createHtmlDialogFromFile(dialog); // "sanitize" the HTML so Google will display it
      this.showModalDialog(htmlDialog, dialog.title); // to limit the number of times I reference Ui
    }
    if (params.hasOwnProperty('command')) { // execute whatever other command the user is requesting, if no alert or dialog is needed
      var that = this; // have I told you lately that I love JavaScript?
      var commands = {
        setupCreateContacts: function() { that.setupCreateContacts(that.sheet); },
        setupCreateFolderStructure: function() { that.setupCreateFolderStructure(that.sheet); },
        setupShareFolders: function() { that.setupShareFolders(that.sheet); },
        assignmentsCreatePaperTemplates: function() { that.assignmentsCreatePaperTemplates(that.sheet, response); },
        assignmentsCalcWordCounts: function() { that.showHtmlSidebar(SIDEBAR_ASSIGNMENTS_CALC_WORD_COUNTS); },
        assignmentsCopy: function() { that.assignmentsCopy(that.sheet, response); },
        assignmentsReturn: function() { that.assignmentsReturn(that.sheet, response); },
        adminGenerateGradebook: function() { that.doGenerateGradebook(that.sheet); },
        formatSwitchStudentNames: function() { that.doSwitchStudentNames(that.sheet); },
        formatSetShadedRows: function() { that.doSetShadedRows(that.sheet); },
        help: function() { that.showHtmlSidebar(SIDEBAR_HELP); }
      };
      var command = commands[params.command];
      command();
    }
  } catch(e) {
    if (params.hasOwnProperty('error_msg')) {
      var alert = this.alert({msg: params.error_msg});
      alert(); // display alert if something goes wrong (this is the only error message a user should probably see)
    }
    this.logError('Exposify.prototype.executeMenuCommand', e);
  }
} // end Exposify.prototype.executeMenuCommand

/**
 * Determine whether this year uses a different schedule, only if September 1 falls on a Tuesday.
 * @param {number} year - The year to check.
 * @return {boolean}
 */
Exposify.prototype.getAlternateDesignationYearStatus = function(year) { // change in designation days are different if September 1 is a Tuesday (see http://senate.rutgers.edu/RLBAckS1003AAcademicCalendarPart2.pdf)
  try {
    var firstDayOfSeptember = (new Date(year, 8, 1)).getDay();
    return firstDayOfSeptember === 2 ? true : false; // return true if the first day of September of the year being checked is a Tuesday and false otherwise
  } catch(e) { this.logError('Exposify.prototype.getAlternateDesignationYearStatus', e); }
} // end Exposify.prototype.getAlternateDesignationYearStatus

/**
 * Get the client secret for this script for use in the OAuth2 authorization flow. The secret is stored
 * as a script property, because we don't want end users to be able to see it.
 * @return {string} secret - The client secret for this app.
 */
Exposify.prototype.getClientSecret = function() {
  try {
    var secret = PropertiesService.getScriptProperties().getProperty('CLIENT_SECRET');
    return secret;
  } catch(e) { this.logError('Exposify.prototype.getClientSecret', e); }
} // end Exposify.prototype.getClientSecret

/**
 * Get the client id for this script for use in the OAuth2 authorization flow. The id is stored
 * as a script property, because we don't want end users to be able to see it.
 * @return {string} id - The client id for this app.
 */
Exposify.prototype.getClientId = function() {
  try {
    var id = PropertiesService.getScriptProperties().getProperty('CLIENT_ID');
    return id;
  } catch(e) { this.logError('Exposify.prototype.getClientId', e); }
} // end Exposify.prototype.getClientId

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
 * @param {Sheet} sheet - A Google Apps Sheet object, the gradebook for which we want the associated course folder.
 */
Exposify.prototype.getCourseFolder = function(sheet) {
  try {
    var courseTitle = this.getCourseTitle(sheet);
    var semesterFolder = this.getSemesterFolder(sheet);
    var courseFolder = this.getFolder(semesterFolder, courseTitle);
    return courseFolder;
  } catch(e) { this.logError('Exposify.prototype.getCourseFolder', e); }
} // end Exposify.prototype.getCourseFolder

/**
 * Look up and return the course number for this course, based on its name in the spreadsheet.
 * @param {Sheet} sheet - The active Sheet object from which to retrieve the course number.
 * @return {string} courseNumber - The course number.
 */
Exposify.prototype.getCourseNumber = function(sheet) {
  try {
    var title = this.getCourseTitle(sheet);
    var section = this.getSectionTitle(sheet);
    var name = title.replace(section, '').trim();
    var numbers = Object.getOwnPropertyNames(COURSE_FORMATS);
    var courses = numbers.map(function(number) { return COURSE_FORMATS[number].name; } );
    var index = courses.indexOf(name);
    var courseNumber = numbers[index];
    return courseNumber;
  } catch(e) { this.logError('Exposify.prototype.getCourseNumber', e); }
} // end Exposify.prototype.getCourseNumber

/**
 * Return the name and section of the course for a given gradebook, e.g. "Expository Writing AB"
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
 * @return {number} - The day of the week on which spring break begins.
 */
Exposify.prototype.getFirstDayOfSpringBreak = function(year) {
  try {
    var firstDayOfMarch = new Date(year, 3, 1).getDay();
    return firstDayOfMarch + (6 - firstDayOfMarch) + 7; // Spring Break starts the second Saturday of March, so find out the first day of March, add days to get to Saturday, and add 7 to that
  } catch(e) { this.logError('Exposify.prototype.getFirstDayOfSpringBreak', e); }
} // end Exposify.prototype.getFirstDayOfSpringBreak

/**
 * Search for and return the subfolder, if one exists, if a given parent folder.
 * @param {Folder} folder - The parent folder to search in.
 * @param {string} name - The name of the folder we're looking for.
 */
Exposify.prototype.getFolder = function(folder, name) {
  try {
    if (folder !== null) {
      var folderIter = folder.getFoldersByName(name);
      return folderIter.hasNext() ? folderIter.next() : null; // return the first match found
    } else {
      return null;
    }
  } catch(e) { this.logError('Exposify.prototype.getFolder', e); }
} // end Exposify.prototype.getFolder

/**
 * Get the graded papers folder for the gradebook on the active spreadsheet.
 * @param {Sheet} sheet - A Google Apps Sheet object, the gradebook for which we want the associated graded papers folder.
 * @return {Folder} folder - The graded papers folder or null if it doesn't exist.
 */
Exposify.prototype.getGradedPapersFolder = function(sheet) {
  try {
    var courseFolder = this.getCourseFolder(sheet);
    return this.getFolder(courseFolder, GRADED_PAPERS_FOLDER_NAME);
  } catch(e) { this.logError('Exposify.prototype.getGradedPapersFolder', e); }
} // end Exposify.prototype.getGradedPapersFolder

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
 * @return {number} - The day of the week on which the last day of a month occurs.
 */
Exposify.prototype.getLastDayOfMonth = function(month, year) {
  try {
    month += 1;
    return month === 2 ? year & 3 || !(year % 25) && year & 15 ? 28 : 29 : 30 + (month + (month >> 3 ) & 1); // do some bit twiddling to figure out the last day of any given month, hard to read code courtesy of http://jsfiddle.net/TrueBlueAussie/H89X3/22/
  } catch (e) { this.logError('Exposify.prototype.getLastDayOfMonth', e); }
} // end Exposify.prototype.getLastDayOfMonth

/**
 * Search a given folder for a file matching the given regular expression and, optionally, with a specific MIME type.
 * @param  {Folder} folder - A Folder object.
 * @param  {RegExp} re - A RegExp object.
 * @param  {string=} type - A file MIME type to optionally select for.
 * @return  {Array} filtered - An array of matched files, or an empty array if none.
 */
Exposify.prototype.getMatchedFiles = function(folder, re, type) {
  try {
    var filesIter = folder.getFiles();
    var filtered = [];
    while (filesIter.hasNext()) {
      var file = filesIter.next();
      var match = file.getName().match(re);
      if (type === undefined && match !== null) {
        filtered.push(file);
      } else if (match !== null && file.getMimeType() === type) {
        filtered.push(file);
      }
    }
    return filtered;
  } catch(e) { this.logError('Exposify.prototype.getMatchedFiles', e); }
} // end Exposify.prototype.getMatchedFiles

/**
 * Switch a name from "last, first" to "first last" order.
 * @param {string} name - A name in last, first order.
 * @return {string} newName - The name in first last order.
 */
Exposify.prototype.getNameFirstLast = function(name) {
  try {
    var names = name.split(','); // if name string contains a comma, assume they are in last, first order and split them at the comma
    if (names.length > 1) { // if, for some reason, the name is already first name first
      var newName = names[1].trim() + ' ' + names[0].trim(); // remove leading and trailing whitespace but add a space between them
    } else {
      var newName = names[0];
    }
    return newName;
  } catch(e) { this.logError('Exposify.prototype.getNameFirstLast', e); }
} // end Exposify.prototype.getNameFirstLast

/**
 * Switch a name from "first last" to "last, first" order.
 * @param {string} name - A name in first last order.
 * @return {string} newName - The name in last, first order.
 */
Exposify.prototype.getNameLastFirst = function(name) {
  try {
    var names = name.split(' '); // if names are in first last order, split them at the space
    var newName = names.pop() + ', ' + names.join(' '); // insert commas between the names and add a space
    return newName;
  } catch(e) { this.logError('Exposify.prototype.getNameLastFirst', e); }
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
 * Get the root Google Drive folder.
 * @return {Folder} folder - The root "My Drive" folder.
 */
Exposify.prototype.getRootFolder = function() {
  try {
    return DriveApp.getRootFolder();
  } catch(e) { this.logError('Exposify.prototype.getRootFolder', e); }
} // end Exposify.prototype.getRootFolder

/**
 * Get the semester folder for the gradebook on the active spreadsheet.
 * @param {Sheet} sheet - A Google Apps Sheet object, the gradebook for which we want the associated semester folder.
 * @return {Folder} folder - The semester folder or null if it doesn't exist.
 */
Exposify.prototype.getSemesterFolder = function(sheet) {
  try {
    var semesterTitle = this.getSemesterTitle(sheet);
    var root = this.getRootFolder();
    var semesterFolder = this.getFolder(root, semesterTitle);
    return semesterFolder;
  } catch(e) { this.logError('Exposify.prototype.getSemesterFolder', e); }
} // end Exposify.prototype.getSemesterFolder

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
} // end Exposify.prototype.getSectionTitle

/**
 * Return the semester for which a given gradebook is used.
 * @param {Sheet} sheet - The Google Apps Sheet object from which to retrieve the semester.
 * @return {string} semesterTitle - The semester.
 */
Exposify.prototype.getSemesterTitle = function(sheet) {
  try {
    var semesterTitle = sheet.getRange('A2').getValue(); // the semester, from the gradebook
    return semesterTitle;
  } catch(e) { this.logError('Exposify.prototype.getSemesterTitle', e); }
} // end Exposify.prototype.getSemesterTitle

/**
 * Return a string that concatenates a given semester with the current year.
 * @param {string} semester - A semester.
 * @return {string} - The semester and year, e.g. "Fall 2015".
 */
Exposify.prototype.getSemesterYearString = function(semester) {
  try {
    var year = new Date().getFullYear(); // assume any given gradebook is being created for the current year (not sure if that's a good idea, but it seems likely in the vast majority of cases)
    return semester + ' ' + year; // create a string from the semester and the current year, i.e. 'Fall 2015'
  } catch(e) { this.logError('Exposify.prototype.getSemesterYearString', e); }
} // end Exposify.prototype.getSemesterYearString

/**
 * Check whether a gradebook has been set up for a Sheet. Return true if so, false otherwise.
 * @param {Sheet} sheet - The Google Apps Sheet object to check.
 * @return {boolean} - True if a gradebook has been set up for the active sheet and false otherwise.
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
    var rows = sheet.getRange(4, 1, MAX_STUDENTS, 1).getValues();
    var count = rows.filter(function(cell) { return cell.toString().length > 0 }).length; // have I told you lately how much I love JavaScript?
    return count;
  } catch (e) { this.logError('Exposify.prototype.getStudentCount', e); }
} // end Exposify.prototype.getStudentCount

/**
 * Search by name for the email address of a student
 * @param {Array} students - An array of student objects to search.
 * @param {string} name - The student's name to use in the search.
 * @return {string} email - The email address or null if none is found.
 */
Exposify.prototype.getStudentEmail = function(students, name) {
  try {
    var name = this.getNameFirstLast(name);
    var email = null;
    var that = this;
    students.forEach(function(student) {
      var studentName = that.getNameFirstLast(student.name);
      if (name === studentName) {
        email = student.email;
      }
    });
    return email;
    } catch(e) { this.logError('Exposify.prototype.getStudentEmail', e); }
} // end Exposify.prototype.getStudentEmail

/**
 * Get the student folders for the gradebook on the active spreadsheet.
 * @param {Sheet} sheet - A Google Apps Sheet object, the gradebook for which we want the associated student folders.
 * @return {Array} studentFolders - An array containing individual student folders or null if the graded papers folder doesn't exist.
 */
Exposify.prototype.getStudentFolders = function(sheet) {
  try {
    var gradedPapersFolder = this.getGradedPapersFolder(sheet);
    if (gradedPapersFolder !== null) {
      var studentFolders = [];
      var folderIterator = gradedPapersFolder.getFolders();
      while (folderIterator.hasNext()) {
        studentFolders.push(folderIterator.next());
      }
      return studentFolders;
    }
    else {
      return null;
    }
  } catch(e) { this.logError('Exposify.prototype.getStudentFolders', e); }
} // end Exposify.prototype.getStudentFolders

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
 * @return {number} - The day of the month on which Thanksgiving occurs.
 */
Exposify.prototype.getTuesdayOfThanksgivingWeek = function(year) {
  try {
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
  } catch(e) { this.logError('Exposify.prototype.getTuesdayOfThanksgivingWeek', e); }
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
    var type = MIME_TYPE_GOOGLE_DOC;
    var filtered = this.getMatchedFiles(courseFolder, re, type);
    var counts = [];
    filtered.forEach(function(file) {
      var doc = DocumentApp.openById(file.getId()).getBody().getText();
      var worksCitedRe = /Works Cited(\s|.)+/; // regular expression to find and remove Works Cited from word count
      var count = doc.replace(worksCitedRe, '').split(/\s+/g).length; // simple word count, but delete the Works Cited first
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
    var spreadsheet = this.spreadsheet;
    var logFileId = PropertiesService.getScriptProperties().getProperty('LOG_FILE_ID');
    var logs = SpreadsheetApp.openById(logFileId);
    var errorLogSheet = logs.getSheetByName(ERROR_TRACKING_SHEET_NAME);
    var date = new Date();
    var timestamp = date.toDateString() + ' ' + date.toTimeString();
    var email = spreadsheet.getOwner().getEmail();
    var id = spreadsheet.getId();
    var info = [timestamp, email, id, callingFunction, traceback];
    var lastRow = errorLogSheet.getLastRow();
    var pasteRange = errorLogSheet.getRange((lastRow + 1), 1, 1, 5);
    if (errorLogSheet.getMaxRows() === lastRow) {
      errorLogSheet.insertRowAfter(lastRow);
    }
    pasteRange.setValues([info]);
  }
  var msg = 'I\'m sorry, but there was a problem! Try again, because sometimes the problem is Google, not Exposify. But you can tell Steve you saw this error message, and maybe he can fix it:\n\n(' + errorLogSheet.getLastRow() + ') ' + traceback;
  var alert = this.alert({msg: msg});
  alert(); // this will be annoying if there are too many of them
} // end Exposify.prototype.logError

/**
 * Record the email address of someone who installs Exposify and the Google Docs
 * spreadsheet id number of the document to which it is attached. This is intended
 * for communication and updating purposes only. It can be turned off by setting the
 * INSTALL_TRACKING constant to false.
 */
Exposify.prototype.logInstall = function() {
  if (INSTALL_TRACKING === true) {
    var spreadsheet = this.spreadsheet;
    var logFileId = PropertiesService.getScriptProperties().getProperty('LOG_FILE_ID');
    var logs = SpreadsheetApp.openById(logFileId);
    var installLogSheet = logs.getSheetByName(INSTALL_TRACKING_SHEET_NAME);
    var date = new Date();
    var timestamp = date.toDateString() + ' ' + date.toTimeString();
    var email = spreadsheet.getOwner().getEmail();
    var id = spreadsheet.getId();
    var info = [timestamp, email, id];
    var lastRow = installLogSheet.getLastRow();
    var pasteRange = installLogSheet.getRange((lastRow + 1), 1, 1, 3);
    if (installLogSheet.getMaxRows() === lastRow) {
      errorLogSheet.insertRowAfter(lastRow);
    }
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
 * this sheet.
 * @param {Sheet} sheet - Google Apps Sheet object for which to set the property.
 * @return {boolean} - True if successful, false otherwise.
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
 * @param {Sheet} sheet - The Sheet object.
 * @param {string} id - The file id of the file from which to extract student names.
 */
Exposify.prototype.setupAddStudents = function(sheet, id) {
  try {
    var spreadsheet = this.spreadsheet;
    var file = DriveApp.getFileById(id);
    var mimeType = file.getMimeType(); // Google Sheets or csv format
    var filename = file.getName();
    var students = [];
    if (mimeType === MIME_TYPE_GOOGLE_SHEET) {
      var params = {sheet: sheet, id: id, mimeType: MIME_TYPE_GOOGLE_SHEET};
      students = this.doParseSpreadsheet(params);
    } else if (mimeType === MIME_TYPE_CSV) {
      var params = {sheet: sheet, id: id, mimeType: MIME_TYPE_CSV};
      students = this.doParseSpreadsheet(params);
    } else {
      var alert = this.alert({msg: ERROR_SETUP_ADD_STUDENTS_INVALID.replace('$', filename), title: 'Setup Add Students'}); // '$' is a wildcard value that is replaced with the filename
      alert();
      return;
    }
    if (students.length === 0) {
      var alert = this.alert({msg: ERROR_SETUP_ADD_STUDENTS_EMPTY.replace('$', filename), title: 'Setup Add Students'});
      alert();
    } else {
      var params = {sheet: sheet, students: students};
      this.doAddStudents(params);
      spreadsheet.toast(ALERT_SETUP_ADD_STUDENTS_SUCCESS.replace('$', filename), TOAST_TITLE, TOAST_DISPLAY_TIME);
    }
  } catch(e) {
    var alert = this.alert({msg: ERROR_SETUP_ADD_STUDENTS, title: 'Setup Add Students'});
    alert();
    this.logError('Exposify.prototype.setupAddStudents', e);
  }
} // end Exposify.prototype.setupAddStudents

/**
 * Create a Google Contacts contact group for the students listed in the active gradebook.
 * @param {Sheet} sheet - The Google Apps Sheet object from which to retrieve student names.
 */
Exposify.prototype.setupCreateContacts = function(sheet) {
  try {
    var spreadsheet = this.spreadsheet;
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
        var contact = ContactsApp.createContact(name[0], name[1], student.email); // if the student's email doesn't exist or is incorrectly formatted, this field will be blank
      } else {
        var contact = ContactsApp.getContact(student.email);
      }
      contactGroup.addContact(contact); // this is slow :(
      Utilities.sleep(100); // and I have to make it slower to avoid quota exceptions
    });
    spreadsheet.toast(ALERT_SETUP_CREATE_CONTACTS_SUCCESS.replace('$', contactGroupTitle), TOAST_TITLE, TOAST_DISPLAY_TIME); // '$' is replaced with the name of the contact group
  } catch(e) { this.logError('Exposify.prototype.setupCreateContacts', e); }
} // end Exposify.prototype.setupCreateContacts

/**
 * Create a folder structure in Google Drive for the gradebook on the active sheet.
 * @param {Sheet} sheet - The Google Apps Sheet object with the gradebook for which to create a folder structure in Drive.
 */
Exposify.prototype.setupCreateFolderStructure = function(sheet) {
  try {
    var rootFolder = DriveApp.getRootFolder();
    var semesterFolder = this.getSemesterFolder(sheet);
    var courseFolder = this.getCourseFolder(sheet);
    var gradedPapersFolder = this.getGradedPapersFolder(sheet);
    var studentFolders = this.getStudentFolders(sheet) || [];
    var createdFolders = [];
    var deletedFolders = [];
    if (semesterFolder === null) {
      var semesterTitle = this.getSemesterTitle(sheet);
      var semesterFolder = rootFolder.createFolder(semesterTitle);
      createdFolders.push(semesterFolder.getName());
    }
    if (courseFolder === null) {
      var courseTitle = this.getCourseTitle(sheet);
      var courseFolder = semesterFolder.createFolder(courseTitle);
      createdFolders.push(courseFolder.getName());
    }
    if (gradedPapersFolder === null) {
      var gradedPapersFolder = courseFolder.createFolder(GRADED_PAPERS_FOLDER_NAME);
      createdFolders.push(gradedPapersFolder.getName());
    }
    var studentNames = this.getStudentNames(sheet);
    var studentFolderNames = studentFolders.map(function(folder) { return folder.getName(); });
    var that = this;
    var foldersToCreate = studentNames.filter(function(studentName) { return that.arrayContains(studentFolderNames, studentName) ? false : true; });
    var foldersToDelete = studentFolderNames.filter(function(studentFolderName) { return that.arrayContains(studentNames, studentFolderName) ? false : true; });
    foldersToCreate.forEach(function(name) {
      var newFolder = gradedPapersFolder.createFolder(name);
      createdFolders.push(name);
    });
    var that = this;
    studentFolders.forEach(function(folder) {
      var name = folder.getName();
      if (that.arrayContains(foldersToDelete, name)) {
        folder.setTrashed(true);
        deletedFolders.push(name);
      }
    });
    var msg = 'Finished!\n';
    if (createdFolders.length > 0) {
      msg += '\nFolders created:\n\n' + createdFolders.join('\n');
      msg += '\n';
    }
    if (deletedFolders.length > 0) {
      msg += '\nFolders removed:\n\n' + deletedFolders.join('\n');
    }
    if (createdFolders.length === 0 && deletedFolders.length === 0) {
      msg += 'No folders have been created or destroyed.';
    }
    var alert = this.alert({msg: msg, title: 'Create Folder Structure'});
    alert();
  } catch(e) { this.logError('Exposify.prototype.setupCreateFolderStructure', e); }
} // end Exposify.prototype.setupCreateFolderStructure

/**
 * Convert user input, collected from a dialog box, into a newly formatted gradebook.
 * @param {Object} courseInfo - User input collected into an object.
 * @param {string} courseInfo.course - A course number.
 * @param {string} courseInfo.section - A section code.
 * @param {string} courseInfo.semester - A semester name.
 * @param {Array} courseInfo.meetingDays - A list of days of the week when the class meets.
 */
Exposify.prototype.setupNewGradebook = function(sheet, courseInfo) {
  try {
    var spreadsheet = this.spreadsheet;
    var newName = courseInfo.course === OTHER_COURSE_NUMBER ? courseInfo.section : courseInfo.course + ':' + courseInfo.section; // only show the course number if it's real
    var exists = spreadsheet.getSheetByName(newName);
    if (exists !== null && sheet.getName() === newName) {
      var alert = this.alert({msg: ALERT_SETUP_NEW_GRADEBOOK_ALREADY_EXISTS.replace('$', newName), title: 'Setup New Gradebook'}); // avoid creating a new sheet with the same name as an existing sheet
      alert();
      return;
    }
    var newCourse = new Course(courseInfo); // create new Course object with information collected from the user by the dialog box
    var check = this.doFormatSheet({course: newCourse, sheet: sheet}); // do the actual work, probably in a way that I should further refactor
    if (check === true) {
      var checkStatus = this.setSheetStatus(sheet);
    }
    if (checkStatus === true) {
      spreadsheet.toast(ALERT_SETUP_NEW_GRADEBOOK_SUCCESS.replace('$', newCourse.nameSection), TOAST_TITLE, TOAST_DISPLAY_TIME); // cute pop-up window
    } else {
      var alert = this.alert({msg: ERROR_SETUP_NEW_GRADEBOOK_FORMAT, title: 'Setup New Gradebook'});
      alert();
    }
  } catch(e) {
    var alert = this.alert({msg: ERROR_SETUP_NEW_GRADEBOOK_FORMAT, title: 'Setup New Gradebook'});
    alert();
    this.logError('Exposify.prototype.setupNewGradebook', e);
  }
} // end Exposify.prototype.setupNewGradebook

/**
 * Share the course folder with all students in the section, and share their graded papers folders with each
 * of them, respectively.
 * @param {Sheet} sheet - The Google Apps Sheet object with the gradebook containing the students with whom to share folders.
 */
Exposify.prototype.setupShareFolders = function(sheet) {
  try {
    var spreadsheet = this.spreadsheet;
    var section = this.getSectionTitle(sheet);
    var students = this.getStudents(sheet);
    var that = this;
    var emails = students.map(function(student) { return student.email; });
    var courseFolder = this.getCourseFolder(sheet);
    if (courseFolder === null) {
      var alert = this.alert({msg: ALERT_MISSING_COURSE_FOLDER, title: 'Share Folders With Students'});
      alert();
      return;
    }
    courseFolder.setShareableByEditors(false);
    var currentEditors = courseFolder.getEditors().map(function(editor) { return editor.getEmail(); });
    currentEditors.forEach(function(editor) {
      if (that.arrayContains(emails, editor) === false) { courseFolder.removeEditor(editor); }
    });
    emails.forEach(function(email) {
      if (that.arrayContains(currentEditors, email) === false) { courseFolder.addEditor(email); }
    });
    var gradedPapersFolder = this.getGradedPapersFolder(sheet);
    if (gradedPapersFolder === null) {
      var alert = this.alert({msg: ALERT_MISSING_GRADED_FOLDER, title: 'Share Folders With Students'});
      alert();
      return;
    }
    var subFolders = gradedPapersFolder.getFolders();
    if (subFolders === null) {
      var alert = this.alert({msg: ALERT_MISSING_GRADED_PAPER_FOLDERS, title: 'Share Folders With Students'});
      alert();
      return;
    }
    var studentFolders = [];
    //var studentNames = students.map(function(student) { return that.getNameFirstLast(student.name); });
    while (subFolders.hasNext()) { studentFolders.push(subFolders.next()); }
    studentFolders.forEach(function(folder) {
      folder.getEditors().forEach(function(editor) { folder.removeEditor(editor); });
      var folderName = folder.getName();
      //if (that.arrayContains(studentNames, folderName)) {
        var email = that.getStudentEmail(students, folderName);
        if (email !== null) { folder.addEditor(email); }
      //}
      Utilities.sleep(100); // to avoid quota exceptions
    });
    var gradedEditors = gradedPapersFolder.getEditors();
    gradedEditors.forEach(function(editor) { gradedPapersFolder.removeEditor(editor); });
    spreadsheet.toast(ALERT_SETUP_SHARE_FOLDERS_SUCCESS.replace('$', section), TOAST_TITLE, TOAST_DISPLAY_TIME);
  } catch(e) { this.logError('Exposify.prototype.setupShareFolders', e); }
} // end Exposify.prototype.setupShareFolders

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

// FORMAT FUNCTIONS

/**
 * Apply a set of formatting options to the active sheet object.
 * @param {Sheet} sheet - The Sheet object to which to apply the format.
 */
Format.prototype.apply = function(sheet) {
  try {
    var headingRange = sheet.getRange(3, 1, 1, this.columnHeadings.length); // cell range for gradebook column headings
    var centerRange = sheet.getRange(3, 3, this.lastRow, this.lastColumn); // cell range for central part of gradebook, where grade data is actually entered
    var topRowsRange = sheet.getRange(1, 1, 3, this.lastColumn); // rows to keep at the top of the spreadsheet view
    var titleRange = sheet.getRange('A1:A2'); // course name and semester titles
    var mergeTitleRange = sheet.getRange('A1:B2'); // we want to merge each of these with the following cell to create a bigger space for the titles
    var mergeRange = sheet.getRange(1, 3, 2, this.lastColumn - 2); // merge the empty columns in the top rows so it looks nicer
    var cornerRange = sheet.getRange('A3:B3'); // where the frozen rows and columns intersect
    var fullRange = sheet.getRange(1, 1, this.lastRow, this.lastColumn); // range of the entire gradebook
    var maxRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()); // range of the entire visible sheet
    sheet.clear();
    maxRange.clearDataValidations(); // this has to be done separately
    sheet.setFrozenRows(0); // make sure only the correct rows and columns are frozen when formatting is complete
    sheet.setFrozenColumns(0);
    fullRange.breakApart(); // break apart any merged cells
    this.setRowHeights(sheet);
    this.setColumnWidths(sheet);
    fullRange.setFontFamily(FONT); // set font
    fullRange.setFontSize(11); // student names and grades font size
    titleRange.setFontSizes([[16],[14]]); // titles font size
    headingRange.setFontSize(9); // headings font size
    cornerRange.setHorizontalAlignment('center'); // set text alignments
    cornerRange.setVerticalAlignment('middle');
    centerRange.setHorizontalAlignment('center');
    centerRange.setVerticalAlignment('middle');
    fullRange.setBorder(true, true, true, true, true, true); // set cell borders
    titleRange.setValues([[this.courseTitle], [this.semester]]); // set titles
    mergeTitleRange.mergeAcross(); // merge title cells
    mergeRange.mergeAcross(); // merge other cells in the first two rows
    headingRange.setValues([this.columnHeadings]); // set column headings
    headingRange.setWrap(true); // set word wrapping
    sheet.setFrozenRows(3); // freeze first three rows
    sheet.setFrozenColumns(2); // freeze first two columns
    if (this.gradeValidations === true) { expos.setGradeValidations(sheet, this.course.gradeValidations); } // set data validations for grades
    if (this.courseFormat.hasOwnProperty('finalGradeFormulaRange')) { expos.doSetFormulas(sheet, this.courseNumber); } // apply final grade formula to this range
    if (this.meetingDays.length !== 0) { expos.doFormatSheetAddAttendanceRecord(this.course, sheet); } // add an attendance sheet if the user asked for it
    if (this.shadedRows === true) {
      topRowsRange.setBackground(COLOR_SHADED);  // set background color of frozen rows
      expos.doSetShadedRows(sheet); // set alternating color of student rows
    }
    sheet.setName(this.sheetName); // name sheet with section number
  } catch(e) { expos.logError('Format.prototype.apply', e); }
} // end Format.prototype.apply

/**
 * Set the format to shade alternating rows.
 * @return {Format} this - Return this object for chaining.
 */
Format.prototype.setShadedRows = function() {
  try {
    this.shadedRows = true;
    return this;
  } catch(e) { expos.logError('Format.prototype.setShadedRows', e); }
} // end Format.prototype.setShadedRows