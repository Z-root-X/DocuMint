=================================================
 Admit Card Automator - Production Package
=================================================

Thank you for using the Admit Card Automator!

This folder contains everything you need to run the application and generate admit cards.

----------------
1. What's Included
----------------

*   `AdmitCardAutomator.exe`: The main application file. Double-click this to run the program.
*   `example_data.xlsx`: An example Excel file with dummy data to show you the required format.
*   `example_template.docx`: An example Word document template with all the available placeholders.
*   `README.txt`: This file.

----------------
2. How to Use the Application
----------------

1.  **Run the Application**: Double-click the `AdmitCardAutomator.exe` file to start the application.

2.  **Follow the Wizard**:
    *   **Step 1: Setup Your Files**: Use the "Browse" buttons to select your Excel data file, your Word template file, and the folders where you want to save the generated PDFs and the log files. You can use the included example files to get started.
    *   **Step 2: Customize Your Email**: Edit the email subject and body as needed.
    *   **Step 3: Review & Run**: Review all your settings in the summary panel. When you are ready, click "Start Process" to begin.

3.  **In-App Help**: For more detailed instructions, click the **❓** icon in the top-right corner of the application.

----------------
3. How to Prepare Your Files
----------------

**Excel Data Sheet:**

Your Excel file must contain the following columns. The names must be an *exact match*.

*   `ID`
*   `Email`
*   `Name`
*   `Phone`
*   `DOB`
*   `Age`
*   `Father`
*   `District`
*   `Category`
*   `Last_Academic_Degree`
*   `Last_Academic_Institute`
*   `Current_Academic_Institute`
*   `Online_Exam_Score`

**Word DOCX Template:**

The Word document acts as the template for your admit cards. Insert placeholders into the document where you want candidate-specific information to appear.

The placeholders must correspond to the column names in your Excel file, enclosed in angle brackets. For example, to display a candidate's name, you would use `<Name>` in your template.

=================================================
