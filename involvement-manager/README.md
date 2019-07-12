# Google CSSI Involvement Manager
##### Created by Andrew Dimmer
##### v0.1.1
This program is designed to help TAs automatically check in with their students each week, and track who is involved.

> <g-emoji class="g-emoji" alias="warning" fallback-src="https://github.githubassets.com/images/icons/emoji/unicode/26a0.png">⚠️</g-emoji> Warning! This was created in about 4 hours so that I could use it myself. No features claim to be stable, and bugs may exist. If you come across a bug or come up with a feature you would like to see, please log it [here](https://github.com/andrewdimmer/google-cssi/issues).

## Setup the Program
### 1. Clone the Spreadsheet
Go to the [template](https://drive.google.com/open?id=1aMgesjItl6rur0XAhSf4itqCfONWprCRxVMAUpxBHYM) and navigate to "File" -> "Make a copy". Create a copy in a folder on your Google Drive where you know where it will be.

This should create a copy of the spreadsheet, check-in form, and script file.

> Please then grant Morgan Lynch edit access to the spreadsheet so that she can view the check-in information as well.
### 2. Add Students
Next, you need to add your students to the spreadsheet. All of the sheets in the spreadsheet pull their user data from one master list, so you only need to add the names once.

To get your students' names populating the spreadsheet, simply fill out First Name, Last Name, and Email columns (A, B, and C) on the `Names and Info` sheet (color coded green).

> Note: you may want to add yourself you can test that everything is working

> <g-emoji class="g-emoji" alias="warning" fallback-src="https://github.githubassets.com/images/icons/emoji/unicode/26a0.png">⚠️</g-emoji> Warning! Do not add or remove rows from the spreadsheet. That may adversely impact formulas and scripts. The ability to add and remove rows will be something in a future update.

Once you have added students' names and emails, you can also fill out the rest of `Names and Info` sheet except for the "Course" column (D). Feel free to add or remove any fields you want with the exception of columns A-D.

> Note: On all sheets, columns that are populated automatically by formulas are italics so that you know not to fill them yourself, even if it looks like something you could.
### 3. Update the Pre-Filled Link for the Form
One of the benefits of using a script to send the check-in form to each person individually is that you can auto-fill parts of the form, in particular the name, email, and course.

To set this up, you need to generate a pre-filled link with placeholder values that the script will later replace with the actual values from the spreadsheet.

To create and add the pre-filled link to the script:
1. Open the form
2. Click the three dots in the upper right-hand corner
3. Click on "Get pre-filled link"

![Create A Link](https://github.com/andrewdimmer/google-cssi/blob/master/involvement-manager/images/CreateALink.png)

4. Enter "FIRST" in the `First Name` field
5. Enter "LAST" in the `Last Name` field
6. Enter "EMAIL" in the `Email Address` field
7. Select "Finished" from the drop down in the `Course` field
8. Scroll down to the bottom and click the "GET LINK" button
9. Click on "COPY LINK" in the black box that popped on in the lower left-hand corner
10. Open the spreadsheet
11. Click on the "Tools" menu
12. Click on "Script editor"
13. Paste the pre-filled link into the empty double quotes on line 6
14. Click the Save button (the red * next to the file name should disappear if the file saved successfully)

### 4. Test the Email System
The Google CSSI Involvement Manager parses the email from a Google Doc so that you can change the email without editing the code, once you set the source of the email.

To demonstrate this feature, there is a sample email already coded into the Involvement Manager. To test this, you can manually trigger the test function that sends an email to you just as if you were one of the students.

To run the test function:
1. Open the spreadsheet
2. Click on the "Tools" menu
3. Click on "Script editor"
4. Edit lines 13-15 to add your name and email
5. Select the `test` function from the runner drop down, then click run

![Run a Function](https://github.com/andrewdimmer/google-cssi/blob/master/involvement-manager/images/RunAFunction.PNG)

When you run a script the first time, it asks you to confirm the permission scope. When the "Authorization required" pop-up comes up, click "Review Permissions", then Login.

![Authorization Required](https://github.com/andrewdimmer/google-cssi/blob/master/involvement-manager/images/AuthorizationRequired.PNG)

Under some circumstances, there is an error saying the "The app isn't verified". Click "Advanced" then "Go to Google CSSI 2019 Involvement Sheet" to get past this error. After you do that, you can just click the "Allow" button, and the email should send.

![The app isn't verified](https://github.com/andrewdimmer/google-cssi/blob/master/involvement-manager/images/AppIsntVerified.PNG)

Check your inbox to see that the email delivered.
### 5. Edit the email content
One of the nice things about parsing the email from a Google Doc is that the creation of a new email template is as simple as making a new Google Doc and writing your email in it just like you would if you were writing a draft in your email client.

The first line of the document is always the subject line of the email. This line contains no formatting and uses no short codes. Given that the formatting doesn't show up in the email, it can be sometimes helpful to make the subject line larger than the rest of the body so you remember it is in fact the subject. In the example, the subject line was denoted by Heading 1.

Once you add the subject line, just click enter to and start typing your body on the next line. The whole rest of the document, however long of short it may be, will be parsed as the email body.

Currently, the parser supports the following formatting features:
* bold
* italics
* underline
* hyperlinks
* unordered lists (bullet points: circles, discs, and squares)
* ordered lists (all standard glyphs: 1., A., a., I., i.)

The parser also supports the following short codes:
* `[FIRST_NAME_SHORTCODE]` : replaces the short code with the student's first name
* `[LAST_NAME_SHORTCODE]` : replaces the short code with the student's last name
* `[LINK_SHORTCODE]` : replaces the short code with the student's pre-filled form link
  - HTML Body: displays the text "Click here to check in."
  - Plain Text Body: displays the raw hyperlink (ex. https://form-link-here)

If you would like an example of what to say for the email [here](https://docs.google.com/document/d/1o_YfPdB8OinPgO_C0AqCorRt_3Zmmldh74l2a-zPdII/edit?usp=sharing) is the email I will be using.

Once you have written the email, you just need to add the link to the script so it knows to pull the new email body.

To update the email link:
1. Open the spreadsheet.
2. Click on the "Tools" menu
3. Click on "Script editor"
4. Replace the link that is already on line 5 with the share link of your new Google Doc
5. Click the Save button (the red * next to the file name should disappear if the file saved successfully)
### 6. Create Triggers to Send Emails
To send an email at a specific date and time, you can create project triggers that will execute the code for you.

To do this:
1. Open the spreadsheet.
2. Click on the "Tools" menu
3. Click on "Script editor"
4. From the script editor, click on the "Edit" menu
5. Click on "Current project's triggers"
6. Click "+ Add Trigger" in the bottom right hand corner for each trigger you wish to add. Follow the image below for the settings to create a date-time trigger.

![Date Time Trigger](https://github.com/andrewdimmer/google-cssi/blob/master/involvement-manager/images/CreateATrigger.png)

Recommended dates for setting up the triggers based on when the spreadsheet things the week starts are on Thursday nights so that students can report what they have done for the week anytime between Friday and Monday.
* 2019-07-11 20:00
* 2019-07-18 20:00
* 2019-07-25 20:00
* 2019-08-01 20:00
* 2019-08-08 20:00
* 2019-08-15 20:00
* 2019-08-22 20:00
* 2019-08-29 20:00

> Note: if you have already missed the execution time of any trigger, you can always send the email yourself by following the same steps as 4. Test the Email System, just select the `sendEmails` function instead of `test`.
## Using the Involvement Manager
### Keep attendance at different events
You might be interested in tracking who and how many people are coming to each of your events (office hours, tech talks, and cohort meetings).

To log attendance, simply open the spreadsheet and go to the correct sheet for each event:
* Office Hours: `Office Hours Attendance` (color coded blue)
* Tech Talks: `Tech Talk Attendance` (color coded blue)
* Cohort Meetings: `Cohort Meeting Attendance` (color coded blue)

Then add an identifier of the event in the first row of the first open column (the topic discussed, the date of the event, "Week 1 Cohort Meeting", etc.).

Finally, check off the checkbox for each person who attended.

The spreadsheet will automatically count the number of people at the event based on the number of boxes checked, as well as increase the total number of events that each checked off user attended (stored on the `Names and Info` sheet).

> Note: If you do more events than there are columns, you can safely add columns anywhere to the right of column C with no problem with the programming or formulas. Just be sure to copy and paste in the formula to get the total number of attendees for that event from another column.
### Viewing Weekly Check-Ins
You may wish to view all Weekly Check-Ins at the same time. This can be done on the `Check-In Data` sheet (color coded red), although this view cramped and somewhat unhelpful.

To instead view all of the check-ins for a specific week, use the `Check-In Summary` sheet (color coded yellow). This provides a cleaner view, as well as conditional formatting to alert you to things you may want to follow up with your students about.

To select the week for which the check-in data is displayed, select the week number from the drop down in cell `B1`. The rest of the data will automatically populate.

Finally, there is some useful conditional formatting to help you identify possible issues to discuss with your students:
* "What did I work on this week?" (column D) highlights red if the latest week they are working on is more than two weeks behind the current week (they are working on week 2 work during week 5 of the program)
  - By the time that this highlights red, they will need to do more than one week of the Coursera class per week of the summer to finish before the school year starts
* "What did I work on this week?" (column D) highlights yellow if it is blank. This is something you would likely want to follow up with them about as this could mean:
  - They have not yet submitted a weekly check in form
  - They didn't work on the Coursera class this week (perhaps they were on vacation)
  - They are stuck on something for Coursera, and did not report they worked on anything because they haven't made it any further
* "Overall, how do you feel about this week's material?" (column E) highlights red if it is either a 1 or 2
* "What learning zone were you in? Please, check all that apply." (column I) highlights red if they were in the strain zone
* "How would you rate the social community?" (column K) highlights red if it is either a 1 or 2
## Issue Tracking
If you come across a bug or come up with a feature you would like to see, please log it [here](https://github.com/andrewdimmer/google-cssi/issues).

I'll get to fixing these as soon as I can!
## Changelog
* v0.1.0
  - Created the basic feature set including keeping attendance, sending weekly check-in forms, and basic check-in analysis
