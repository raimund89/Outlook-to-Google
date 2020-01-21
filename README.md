# Outlook-to-Google
Standalone Windows Application to export an ICS-file that can be added to Google Calendar or any other calendar app.

![Screenshot of settings](https://github.com/raimund89/Outlook-to-Google/blob/master/OutlookToGoogle.png)

## Introduction
Outlook has several ways to send a calendar to others, either via E-mail or by publishing it to a directory. The problem is, this functionality leaves very little control over what to export and how. On top of that, the exported ICS is incompatible with most calendar applications (like Google Calendar or Lightning) because it doesn't comply with the ICS file specification ([RFC 5545](https://tools.ietf.org/html/rfc5545)).

This standalone application allows exporting the calendar of the *current* Outlook user to be exported to an \*.ics file. If you place this file in for example Dropbox, it can then be imported into Google Calendar through it's Sharing link.

## Todo list
The application is functional, however I would like to add several things to it:
- [ ] Specify a start and end date. At the moment it's fixed at 30 days before and 90 days after today
- [ ] Expand the amount of ICS tags exported, especially extended (non-RFC5545) tags used by Outlook
- [ ] Implement calendar functions like CANCEL and UPDATE
- [ ] In general, use full functionality of RFC5545 specs, instead of only the basics
- [ ] And at any point, make the code a bit more consistent :)
- [ ] Create installer
- [ ] Maybe switch from System.Threading.Timer to a scheduled Windows Service.

## Known issues
- [ ] Not nicely cleaning up, every calendar update the RAM-usage increases with 2-3 MB.
- [ ] ICS validators say Europe/Amsterdam is not a valid timezone. Calendar programs don't have a problem though.
- [ ] Summary doesn't have a language tag. Not required, but recommended
- [ ] Recurring items are converted to multiple single items
- [ ] olResponseOrganized doesn't have the right partstat
- [ ] The sensitivity 'private' doesn't have the right ICS classification
- [ ] Reminders are not included
- [ ] Outlook itself exports a 'TRANSP' tag, no idea what to do with that...
