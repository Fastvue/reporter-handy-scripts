# reporter-handy-scripts

The `/scripts` folder contains some handy PowerShell scripts that use Fastvue Reporter's APIs to achieve functionality not available out of the box.

You can schedule these scripts using Windows Task Scheduler to run at desired times. Just make sure the user the task runs as is an Administrator for the Fastvue Reporter website.

## CreateAndEmailReport.ps1

This script generates a report for a specified time range in the day, filtered by a list of users and sends an email with the private link to the report, as well as PDF and CSV attachments.

Here's a video on how to use it

<div style="position: relative; padding-bottom: 56.25%; height: 0;"><iframe src="https://www.loom.com/embed/21cfc712542d434d803c0034f6accad3?sid=0937bcde-9ce0-48a0-9669-45a66601cef7" frameborder="0" webkitallowfullscreen mozallowfullscreen allowfullscreen style="position: absolute; top: 0; left: 0; width: 100%; height: 100%;"></iframe></div>
