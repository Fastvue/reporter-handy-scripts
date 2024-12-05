# reporter-handy-scripts

The `/scripts` folder contains some handy scripts that use Fastvue Reporter's APIs to achieve functionality not available out of the box.

The PowerShell scripts can be scheduled using Windows Task Scheduler to run at desired times. Just make sure the user the task runs as is an Administrator for the Fastvue Reporter website.

## CreateAndEmailReport.ps1

Although you can easily schedule daily, weekly or monthly reports with Fastvue Reporter, they all run just after midnight, ready in your inbox at the beginning of the day. But what if you want to run a report after each shift, class, or at 5pm each day?

This script generates a report for a specified time range in the day, filtered by a list of users and sends an email with the private link to the report, as well as PDF and CSV attachments.

> Example: Generate a report on a class of students, as soon as the class finishes. Use this script to generate a report between 11:00am and 11:45 am, for a specific list of users and schedule it on weekdays at 11:46am.

Here's a [video](https://www.loom.com/share/21cfc712542d434d803c0034f6accad3?sid=0937bcde-9ce0-48a0-9669-45a66601cef7) on how to use it.

## AddStaffNamesToKeywordGroup.ps1

Want to know when someone is searching for, or watching videos about your staff members? For example, schools may want to know when students are searching for their teachers to find their social media profiles.

You can easily create a 'Staff Names' Keyword Group in Fastvue Reporter and paste in a list of names. However, with regular staff changes, it makes sense to automate this.

This script queries Active Directory for a list of staff names and adds them to a keyword group.

Edit the "User-Defined Variables" section to enter your AD server and optionally specify an LDAP query or list of Security Groups to search. You can change the AD attribute used as the keywords (displayName is used by default), as well as the name of the Keyword Group to import them into.

## CloseOpenIndexes.ps1

This script closes all open indexes (dates) in Fastvue Reporter's Elasticsearch database.

Closing indexes frees up memory resources, and Fastvue Reporter will automatically reopen indexes when needed.
You'll notice today's and yesterday's indexes will automatically re-open after running this script,
as Fastvue Reporter opens these to import logs and run queries for the live dashboards.

Instructions:

1. In Fastvue Reporter, go to Settings > Diagnostic > Database and note the Cluster URI
2. Enter the cluster URI in the $ElasticsearchUrl variable below
3. Log into the Fastvue server and run this script in PowerShell
