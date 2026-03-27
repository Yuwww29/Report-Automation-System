# Report-Automation-System
A fully functional system was developed during my internship at SATS Ltd. to automate the Excel report generation process. The system replaces manual steps involved in creating the report with automation, producing fully formatted reports in seconds.

The primary goal of the project is to automate the creation of Excel reports from both daily and weekly raw data, reducing the time required from approximately 2 hours to under 1 minute. Data from each automated report is uploaded to a Supabase database, enabling Streamlit web interface to access the data for performance monitoring and staff tracking through data visualisations.

The system is designed to be user-friendly and future-proof, allowing other staff members to use it seamlessly should it be deployed beyond the prototype stage. It can be run locally or hosted on a web server for collaborative use, though web deployment requires appropriate data protection measures to prevent data leaks and unathorised usage.

# Backend System
The backend of the system is coded with Python using libraries such as Pandas, OpenPyXL, and Flask.
- Pandas -> Cleaning and manipulation of raw data, creating pivot tables, and storing relevant data in various data structures to be inserted into the Excel report.
- OpenPyXL -> Inserting of pivot tables and stored data into respective Excel sheets, styling, formatting, and reorganising of Excel sheets.
- Flask -> Serves as the communication between frontend system (user interface), backend system, and database.


The structure of the system is as follows:
- User upload raw data file to the system, with a maximum file size of 1.5MB
- The system will compare existing data (in the form of dates) in the raw data file with the database to check if duplicate data is being uploaded. This ensures that there is no conflicting data in the database, which will affect the data visualisations on Streamlit. If existing data is found by the system, user can either overwrite the data (upsert data in Supabase) or exit the process.
- The system will then check for any staff members not assigned to the two respective teams in the raw data. This is achieved by iterating through every row and comparing names in the raw data with names stored in the database. The system will prompt user to assign the unrecognised staff members to their respective teams.
- The automation process will then take place. For daily data, the process will take less than 30 seconds. For weekly data, the process will take less than 1 minute.
- After the automation process is complete, the automated Excel report will be downloadable and will also be stored in the database.

The workflow of the system as described is shown below:

<img width="600" height="600" alt="Screenshot 2026-03-28 011314" src="https://github.com/user-attachments/assets/4a4725e8-5f79-4461-b567-9581c072117f" />

# Frontend System
The frontend of the system is coded with CSS, HTML and JavaScript.

The interface of the system consists of 5 pages.
- An upload page where user can upload the raw data to be processed
- A page to manage names of the Commandos Team
- A page to manage names of the Support Team
- A page to view and download past reports that were generated
- A manage historic data page to view and delete historic data stored in the database. This directly affects graphs on streamlit.
