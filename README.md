Display Registration — Data Loader Flow
Date: 2025-10-31
Scope: Belgium, France, Netherlands, Italy

This is the description for the whole process not the code 

1. Purpose
This document outlines the end-to-end process for managing weekly Display Registration data for the countries BE, FR, NL, and IT. The goal is to automate the process of downloading, formatting, and importing display data into Salesforce while maintaining accuracy, oversight, and traceability.
Key Objectives:
•	Ensure CSV files are consistently formatted for Salesforce Data Loader.
•	Reduce operational errors and save time through automation.
•	Enable reloading of files by country and filename with duplicate checks.

2. High-Level Process Overview
1.	Email Processing
•	Incoming emails are filtered in Outlook based on sender and subject keywords.
•	Attachments (.xlsx) are saved to SharePoint under /Display Load/<Country>/<YYYY-MM-DD>/.
2.	File Formatting
•	Remove unnecessary columns.
•	Rename columns to Salesforce names.
•	Correct values as per country-specific rules.
•	Standardize date formats to yyyy-mm-dd.
•	Check for duplicates and other validation rules.
•	Export formatted CSV for Salesforce import.
 
3.	Notification
o	Once a file is formatted, COE specialists receive an email with the file path and row count.
4.	Salesforce Data Loader Import
o	Power Automate Desktop (PAD) or manual Data Loader imports the CSV into Salesforce.
o	Releases the display registration queue via browser.
5.	Verification
o	Check the Salesforce report using the load date and user filter.
o	Compare row counts with notification emails.
o	Run a dummy check (D-0000000000) to ensure the queue is cleared.
6.	Confirmation
o	Reply to the original email confirming successful data load.

3. Folder Structure & Naming
SharePoint Root:
C:\Users\********\Anheuser-Busch InBev\SFDC COE - EUR - General\Display Load\{country}Display Load\YYYY-MM-DD
File Naming:
•	Raw files: Raw_<Country>_Data_<YYYY-MM-DD>.xlsx
•	Processed files: Processed_<Country>_<YYYY-MM-DD>.csv

 
4. Automation Components
Tools Used:
•	Python (pandas + win32) — email processing, file transformation, notifications.
•	Power Automate Desktop — GUI-based Data Loader automation.
•	Salesforce Data Loader — importing data into Salesforce.
•	Excel — file review/editing.
•	Outlook — receiving/sending emails.
Python Automation Covers:
•	Reading Outlook inbox and filtering emails.
•	Saving attachments to SharePoint folders.
•	Standardizing, renaming, and cleaning data for Salesforce.
•	Sending email notifications with file path and row count.
5. Python Script — Email to File Processor
Notes:
•	Handles country-specific adjustments for FR, IT, BE, NL.
•	Ensures date columns are standardized.
•	Skips files if already processed.
•	Sends email notifications only for newly processed files.

6. Data Loader Flow & Reload
•	Allows reloading by specifying file path and country.
•	Duplicate checks ensure no redundant imports.
•	Recommended: PAD flow or command-line wrapper for Salesforce Data Loader:

7. Verification Steps
•	Open Salesforce > Reports > Display Registration Report.
•	Filter by Load Date and User.
•	Compare row count to email notification.
•	Run dummy check (D-0000000000) to clear the queue.
•	If mismatch, review error.csv from Data Loader.

8. Operational Checks
•	Verify Outlook connection.
•	Ensure country filters applied (BE, FR, NL, IT).
•	Confirm all emails in processing string match criteria.
•	Duplicate table auto-refresh completed.
•	Week number validation passed.

9. Alerts & Error Handling
•	Move failed files to /Failed/ and notify owner.
•	Archive processed files in /Archive/<Country>/<YYYY-MM-DD>/.
•	Log all runs (timestamp, user, file, rows_in, rows_loaded, errors).
 
10. Next Steps / Recommendations
•	Review and update Salesforce API mappings per country.
•	Implement central duplicate table and connect Power Query.
•	Test Python formatter with historical files.
•	Build PAD flow with FilePath and Country inputs.
•	Schedule weekly audits for week-number trends and missing files.

11. Special note Italy
Italy does not deliver their displays weekly like the other countries and the template is also different therefore we do the saving and formatting of this manually.
