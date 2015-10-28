# Exposify

Exposify is a Google Sheets add-on that automates a variety of tasks related to the teaching of expository writing courses. Key features include automatic setup of grade books, attendance records, and folder hierarchies in Google Drive for organizing course sections and the return of graded assignments; batch word counts of student assignments; differential comparison of paper drafts (e.g. rough versus final); and various formatting and administrative tasks. This add-on is based on the specific courses I have taught, but I am planning to create a generic gradebook application in the future.

Some of the code is rough, as I more or less taught myself JavaScript while writing it. I have tried to go back and refactor as much as possible, as time has allowed, and I have also tried to follow best practices for both JS and GAS, where applicable.

NOTE: Exposify will not work out-of-the-box if you just copy and paste the code into a Google Scripts Editor. You will need to perform the following steps in order to run the code yourself:
- create a new project for Exposify in your Google Developers Console, and associate it with the script project (note your API key and Client Secret)
- activate the Drive API in both Resources : Advanced Google services in the Script Editor and in the Developers Console for the project
- activate the Google Picker API in the Developers Console
- also in the Developers Console, set the Redirect URI for your client id to https://script.google.com/macros/d/{PROJECT KEY}/getDriveServiceCallback (you can find your Project Key in the Info tab under the menu option File : Project properties)
- in the Script Editor, create a Script Property called DEVELOPER_KEY with your API key as the value and CLIENT_ID with your CLIENT_SECRET as the value
- add the [OAuth2 library](https://github.com/googlesamples/apps-script-oauth2) to your project

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
- Compare a rough and final draft and display the differences between them
- Built-in grade validation, so your gradebook is always in the correct format