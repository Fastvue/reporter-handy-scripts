# reporter-handy-scripts

The `/scripts` folder contains some handy scripts that use Fastvue Reporter's APIs to achieve functionality not available out of the box.

The PowerShell scripts can be secheduled using Windows Task Scheduler to run at desired times. Just make sure the user the task runs as is an Administrator for the Fastvue Reporter website.

## CreateAndEmailReport.ps1

This script generates a report for a specified time range in the day, filtered by a list of users and sends an email with the private link to the report, as well as PDF and CSV attachments.

> Example: Generate a report on a class of students, as soon as the class finishes. Use this script to generate a report between 11:00am and 11:45 am, for a specific list of users and schedule it on week days at 11:46am.

Here's a [video](https://www.loom.com/share/21cfc712542d434d803c0034f6accad3?sid=0937bcde-9ce0-48a0-9669-45a66601cef7) on how to use it.

## AddStaffNamesToKeywordGroup.ps1

This script queries Active Directory for a list of staff names and adds them to a keyword group. This is useful for schools that want to know when students are searching for staff online to find their social media profiles.

Edit the "User-Defined Variables" section to enter your AD server and optionally specify a LDAP query or list of Security Groups to search. You can change the AD attribute use as the keywords (displayName is used by default), as well as the name of the Keyword Group to import them into.

## CloseOpenIndexes.js

This script closes all open indexes (dates) in Fastvue Reporter's Elasticsearch database. Closing indexes frees up memory resources, and Fastvue Reporter will automatically reopen indexes when needed. You'll notice today's and yesterday's indexes will automatically re-open after running this script, as Fastvue Reporter opens these to import logs and run queries for the live dashboards.

Instructions:

1. In Fastvue Reporter, go to Settings > Diagnostic > Database and note the Cluster URI
2. Remote onto the Fastvue Server, open Chrome and enter the Cluster URI into the address bar.
3. Open Chrome's developer tools by pressing F12 or right-clicking and selecting "Inspect"
4. Navigate to the "Console" tab
5. Copy and paste the entire script into the console and press Enter
6. A success message will be displayed in the console if the indices were closed successfully. An error message will be displayed if there was an issue closing the indices.
