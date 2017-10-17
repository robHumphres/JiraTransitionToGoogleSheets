# JiraTransitionToGoogleSheets
Transitioning Jira CSV Export to Google Sheets

This application was built using MAVEN, and implements the use of jexcelapi.

Jexcelapi = http://jexcelapi.sourceforge.net/

In order for this to work you need to download a csv from JIRA containing the following filters in this order.

- Issue Type
- Issue key
- Issue id
- Parent id
- Summary,Assignee
- Priority
- Status
- Resolution
- Created
- Updated
- Due Date
- Custom field (Story Points)

One you have that you'll need the arguments to run.
- [0] = location of the jira csv
- [1] = Sprint Name

The only stuff that will get added are stories and bugs, which is inforced in the readCSVFile function. Once the file has been read it it'll use an override comparable function to 
alphabetically give you an order ranging from the categories from JIRA (Ex. ACA = Android, etc.). 

# The Output
The output will give you all the tickets on an overview page, and individually give you each ticket with a templated test plan
- Test Case Name (Auto generated)
- Description (Auto generated)
- Test Case Completed Date
- Run By
- Start Date
- Finish Date
- Jira ticket (Auto generated)
- Time
- Environment
- Build #
- Prerequisite
- Os / Browser
- Assumptions
- Overall Pass or Fail
