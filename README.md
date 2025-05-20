Add Emails Based on Codes – Excel VBA Script

This VBA macro automates the process of scanning a reference list in the "ITR" sheet and generating a unique list of email addresses based on embedded letter codes. The compiled list is then displayed in the "EMAIL" sheet.

What It Does:
1) Scans Column B in the "ITR" worksheet for predefined codes (e.g., "AB", "QA", "PM", etc.).
2) Matches each code to a set of predefined email addresses using a dictionary.
3) Compiles all matched email addresses into a single string.
4) Removes duplicate emails and provides a clean, unique list.

Outputs:
Full email list (with duplicates) in EMAIL!A1
Unique email list (no duplicates) in EMAIL!A2

Example of Code Matching:
If B2 contains the text "Inspection - QA-001" and B5 contains "AB-Form":

1) QA maps to qa9.team@checkmail.com
2) AB maps to test1.email@example.com
3) These emails will be added to the result list.

How It Works
Uses a dictionary object to store and retrieve emails for each code.
Uses InStr() to perform partial string matching on each cell in Column B.
Compiles results into a ; -separated string.
Splits the string and stores each address in another dictionary to ensure uniqueness.

Benefits
No manual lookup or copying of emails.

Ensures accuracy by avoiding duplicate emails.

Makes it easy to send group emails or build dynamic distribution lists based on ITR content.

Sample Output: Cell	Content
A1	Full email list including duplicates
A2	Clean list with only unique email addresses ready for copy/paste
