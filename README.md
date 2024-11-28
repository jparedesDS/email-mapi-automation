# Email MAPI Automation

## Project Description
**Email MAPI Automation** is a fully automated solution designed to process HTML tables embedded in the body of emails, modify the data, and update a database. Additionally, it generates and sends a new email with the updated data and includes a dynamically generated Excel file as an attachment.

### Project Workflow
1. **Email Reception:**
   - The program captures emails from a specified address every 30 minutes, between 7:30 a.m. and 5:00 p.m.
   - Emails include an HTML table embedded in the body.

2. **Data Extraction and Modification:**
   - Data from the HTML table is extracted and converted into a manipulable format (Pandas DataFrame).
   - Modifications are applied to the data, and new columns are created to ensure compatibility with the database entry.

3. **Automated Email Generation:**
   - A new automated email is created, which includes:
     - Information about the data transmitted and updated in the database.
     - An `.xlsx` file containing the processed data attached to the email.
     - Recipients are automatically selected based on the processed data.

4. **Database Update:**
   - After sending the email via Outlook, the processed tables are selected, and the database is updated accordingly.

## Libraries Used
The project uses the following Python libraries:
- **pandas**: For data manipulation and analysis.
- **numpy**: For advanced numerical operations.
- **shutil**: For file and directory management.
- **os**: For operating system interactions.
- **BeautifulSoup**: For HTML parsing and data extraction.
- **re (Regex)**: For pattern matching and text manipulation.
- **xlsxwriter**: For creating Excel files.
- **timestamp**: For handling date and time.
- **win32com.client**: For Outlook automation to manage emails.

## Key Features
- **Fully Automated:** Automatic email and data processing.
- **Email Management:** Automated generation and sending of emails with updated data and attachments.
- **Database Integration:** Prepares data for seamless integration into an existing database.
- **Dynamic Processing:** Handles HTML tables of various formats dynamically.

## System Requirements
- **Operating System:** Windows (required for Outlook automation).
- **Python 3.x**
- Additional dependencies installed via `pip`.

## How to Run
1. Clone this repository:
   ```bash
   git clone https://github.com/your-username/email-mapi-automation.git
2. Install the required dependencies:
```
pip install -r requirements.txt
```
3. Configure the email address and time intervals in the configuration file.
4. Run the main script:
```
python main.py
```
## Contributions
Contributions are welcome! If you'd like to improve this project, please:

1. Fork the repository.
2. Create a new branch (git checkout -b feature/new-feature).
3. Make your changes and commit them (git commit -am 'Add new feature').
4. Push your branch (git push origin feature/new-feature).
5. Open a Pull Request.

## License
This project is licensed under the MIT License. See the LICENSE file for details.
