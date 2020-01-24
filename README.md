# Outlook-to-Google
Standalone Windows Application to export an ICS-file that can be added to Google Calendar or any other calendar app.

## Introduction
Outlook has several ways to send a calendar to others, either via E-mail or by publishing it to a directory. The problem is, this functionality leaves very little control over what to export and how. On top of that, the exported ICS is incompatible with most calendar applications (like Google Calendar or Lightning) because it doesn't comply with the ICS file specification ([RFC 5545](https://tools.ietf.org/html/rfc5545)).

This standalone application allows exporting the calendar of the *current* Outlook user to be exported to an \*.ics file. If you place this file in for example Dropbox, it can then be imported into Google Calendar through it's Sharing link. The update frequency can be set to anything from every 5 minutes to every day.

![Screenshot of settings](https://github.com/raimund89/Outlook-to-Google/blob/e5435ff5527049ae1a0120beda813edc4504e393/OutlookToGoogle.png)

## Notes on usage
Currently, the ICS-file needs to be exported to a cloud-platform and then imported into Google Calendar by using the shared link. I've only tested this with Dropbox, but other platforms should work as well. I've found in Dropbox, that you need to change a small thing in the sharing link. The link ends with this sequence: ``"?dl=0"``. This indicates to Dropbox that it should show a preview-webpage. You don't want that, you want Dropbox to directly serve the file. Change this sequence to ```"?dl=1"``` does that, and this works fine.

## Acknowledgements
The application uses several libraries, all available through NuGet in Visual Studio. The main library is the NetOffice library ( specifically the Outlook and Office APIs) for communication with Outlook. Using this library makes the application independent of the version of Outlook. It also uses the TimeZoneConverter library to convert between Windows and IANA-compliant timezone designations. Last, the stdole.dll file is added to make sure Interop-functionality works, but it might not be necessary to include it.

## Todo list
The application is functional, however I would like to add several things to it:
- [ ] Expand the amount of ICS tags exported, especially extended (non-RFC5545) tags used by Outlook
- [ ] Implement calendar functions like CANCEL and UPDATE
- [ ] In general, use full functionality of RFC5545 specs, instead of only the basics
- [ ] And at any point, make the code a bit more consistent :)
- [ ] Maybe switch from System.Threading.Timer to a scheduled Windows Service.
- [ ] Directly interface with Google to upload calendar changes.

## Known issues
- [ ] Cancelled events don't show as cancelled in Google Calendar
- [ ] Any exception to a recurring event is now seen as a cancellation, but this doesn't have to be the case of course!
- [ ] Not nicely cleaning up, every calendar update the RAM-usage increases with 2-3 MB. But automatic cleanup does kick in at some point.
- [ ] Summary doesn't have a language tag. Not required, but recommended
- [ ] olResponseOrganized doesn't have the right partstat
- [ ] The sensitivity 'private' doesn't have the right ICS classification
- [ ] Only basic reminders are supported
- [ ] Outlook itself exports a 'TRANSP' tag, no idea what to do with that...
- [ ] Recurring Yearly, YearNth and Monthly events are not completely correct in the ICS
