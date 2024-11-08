# reporter-handy-scripts

The `/scripts` folder contains some handy PowerShell scripts that use Fastvue Reporter's APIs to achieve functionality not available out of the box.

You can schedule these scripts using Windows Task Scheduler to run at desired times. Just make sure the user the task runs as is an Administrator for the Fastvue Reporter website.

## CreateAndEmailReport.ps1

This script generates a report for a specified time range in the day, filtered by a list of users and sends an email with the private link to the report, as well as PDF and CSV attachments.

> Example: Generate a report on a class of students, as soon as the class finishes. Use this script to generate a report between 11:00am and 11:45 am, for a specific list of users and schedule it on week days at 11:46am.

Here's a [video](https://www.loom.com/share/21cfc712542d434d803c0034f6accad3?sid=0937bcde-9ce0-48a0-9669-45a66601cef7) on how to use it.

## AddStaffNamesToKeywordGroup.ps1

This script queries Active Directory for a list of staff names and adds them to a keyword group. This is useful for schools that want to know when students are searching for staff online to find their social media profiles.

Edit the "User-Defined Variables" section to enter your AD server and optionally specify a LDAP query or list of Security Groups to search. You can change the AD attribute use as the keywords (displayName is used by default), as well as the name of the Keyword Group to import them into.

## CloseOpenIndexes

This script closes all open indexes (dates) in Fastvue Reporter's Elasticsearch database, excluding today’s and yesterday’s indexes. Closing indexes frees up memory resources, and Fastvue Reporter will automatically reopen these indexes when needed (e.g., when running reports on older dates).

Instructions:

1. Update the `baseURL` variable to match your Fastvue Reporter URL.
2. In Chrome, go to your Fastvue Reporter's web interface, and navigate to **Settings > Diagnostics > Database**.
   - If the database status is "Bad" or "Unknown," restart the Fastvue Reporter service in **services.msc**
   - Wait for the status to show either "Connected", "Waiting for index recovery", "Preparing Elasticsearch" or 'Operational"
3. Open Chrome DevTools:
   - Press `F12` or `Ctrl+Shift+I` (Windows/Linux) or `Cmd+Option+I` (Mac).
4. Go to the **Console** tab, paste the script, and press `Enter`.
   \*/
