# email-mapi-automation (EXAMPLE, not all the data)
A small excel table in html format is received in the body of the email.

We extract the data from the email, modify it and create new columns in the excel table so that the database entry is correct.

We generate the creation of a new fully automated email that indicates the data that has been transmitted and updated to the database. We indicate who would be the people to send this email to based on the data and generate an .xlsx file and add it to the body of the email for transmission.

As soon as we finish sending the generated email through Outlook, we select the excel tables that have been scraped and update the data in the database.

This program will capture all incoming emails from the specified address every 30 minutes from 7.30am to 17.00pm.

-----------------------------------------------------------------------------------------------------------

LIBRARIES:
- pandas.
- numpy.
- shutil.
- os.
- Beautifulsoup
- Regex
- xlsxwriter
- timestamp
- win32com.client
