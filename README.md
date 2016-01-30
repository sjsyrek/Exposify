# Exposify

Exposify is a Google Sheets add-on written in [Google Apps Script](https://developers.google.com/apps-script/), HTML, CSS, and jQuery that automates a variety of tasks related to the teaching of expository writing courses. Key features include automatic setup of grade books, attendance records, and folder hierarchies in Google Drive for organizing course sections and the return of graded assignments; batch word counts of student assignments; and various formatting and administrative tasks. This add-on is based on the specific courses I have taught, but I am planning to create a generic gradebook application in the future. Consider this experimental software.

Some of the code is rough, as I more or less taught myself JavaScript while writing it. I have tried to go back and refactor as much as possible, as time has allowed, and I have also tried to follow best practices for both JavaScript and Google Apps Script, where applicable, but the coding style is not consistent and there is much I would change or do completely differently if I had to start all over again. Comments are in [JSDoc format](http://usejsdoc.org), and documentation can be automatically generated using the [jsdoc node package](https://www.npmjs.com/package/jsdoc).

NOTE: Exposify will not work out-of-the-box if you just copy and paste the code into a Google Scripts Editor. You will need to do the following:
1. Create a project in your Google Developers Console and, from the Script Editor.
2. Enable the Drive API and the Google Picker API in the Developers Console.
3. Enable the Drive API from the "Resources : Advanced Google services" menu in the Script editor.
4. Create an API key and an OAuth client ID in the credentials section of your project in the Developers console. Make a note of them.
5. Select the "File : Project properties" menu in the Script editor and create four script properties set to the following values:
- DEVELOPER_KEY - API key from the console (API keys section).
- CLIENT_ID - Client ID from the console (OAuth 2.0 section).
- CLIENT_SECRET - Client secret from the console  (OAuth 2.0 section).
- LOG_FILE_ID - The file ID of the spreadsheet you want to use for install and error logs
(you can skip this one and set ERROR_TRACKING and INSTALL_TRACKING to false, but if you do want to track these things, make sure you create separate sheets called "Installs" and "Errors" in your spreadsheet).

## Features:

- Automatically generate gradebooks, including attendance records
- Import student rosters directly into your gradebook, without having to type anything in
- Calculate final grades for courses with numeric grades
- Convenience features, like alternate row shading for easy legibility and the ability to swap student names around depending on your needs
- Create Google Contacts Groups for students, to facilitate autocompletion in Gmail
- Create a structure of folders in Google Drive for course materials and student papers
- Share folders with students and automatically collect and return papers for grading
- Generate warning rosters and final gradebooks for submission to higher powers
- Create easy-to-read, pre-formatted paper templates for students to write into directly, making file organization much easier and more consistent (goodbye MLA format)
- Calculate word counts for a given assignment for all or individual students
- Built-in grade validation, so your gradebook is always in the correct format