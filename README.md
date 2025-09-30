# DocuMint - Personalized Document Generation and Email Automation

## Project Overview

DocuMint is a robust and user-friendly desktop application designed to automate the process of generating personalized documents and distributing them via email. It streamlines workflows that involve merging data from an Excel spreadsheet into a Word document template, converting the resulting documents to PDF, and then sending these PDFs as email attachments through Microsoft Outlook.

This application is ideal for tasks such as:
*   Generating personalized admit cards for students.
*   Creating custom certificates or letters.
*   Distributing individualized reports or statements.

## Features

*   **Dynamic Placeholders**: Utilize any column header from your Excel data file as a placeholder (e.g., `<Name>`, `<ID>`, `<Course>`) directly within your Word document template. The application intelligently replaces these placeholders with corresponding data for each recipient.
*   **Intuitive Graphical User Interface (GUI)**: A modern, wizard-style interface built with `tkinter` guides users through the entire process, from file selection to email configuration and execution.
*   **Automated PDF Conversion**: Seamlessly converts the generated Word documents (.docx) into universally viewable PDF files.
*   **Outlook Email Integration**: Leverages Microsoft Outlook to send personalized emails, attaching the generated PDF documents to the respective recipients.
*   **Configurable Settings**: Advanced settings allow users to customize:
    *   **PDF Filename Format**: Define dynamic filenames for generated PDFs using placeholders (e.g., `AdmitCard_{<StudentID>}_{<LastName>}.pdf`).
    *   **Email Retries**: Specify the number of attempts to resend an email if the initial attempt fails.
    *   **Email Delay**: Set a delay (in seconds) between sending emails to prevent overwhelming mail servers or hitting rate limits.
*   **Configuration Persistence**: User settings are automatically saved and reloaded upon application restart, ensuring a consistent experience.
*   **Comprehensive Logging**: Provides real-time feedback and logs the status of each document generation and email sending operation.
*   **Standalone Executable**: Includes a build script (`build.bat`) to create a single, portable executable file for easy distribution without requiring Python installation on target machines.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Prerequisites

Before you begin, ensure you have the following installed:

*   **Python 3.x**: Download and install from [python.org](https://www.python.org/downloads/).
*   **Microsoft Outlook**: Required for sending emails.
*   **Microsoft Word**: Required for document template processing and PDF conversion.
*   **Git**: For cloning the repository (optional, you can also download the ZIP).

### Installation

1.  **Clone the repository** (or download the ZIP and extract it):
    ```bash
    git clone https://github.com/Z-root-X/DocuMint.git
    cd DocuMint
    ```

2.  **Navigate to the project directory**:
    ```bash
    cd DocuMint
    ```
    *(Ensure you are in the root directory of the cloned project.)*

3.  **Install the required Python libraries**:
    It's highly recommended to use a virtual environment to manage dependencies.

    ```bash
    # Create a virtual environment
    python -m venv venv

    # Activate the virtual environment
    # On Windows:
    .\venv\Scripts\activate
    # On macOS/Linux:
    # source venv/bin/activate

    # Install dependencies
    pip install -r requirements.txt
    ```

## Usage

Once installed, you can run the application directly from the source code or use the standalone executable if you've built it.

### Running from Source

To launch the application, execute the `gui.py` script:

```bash
# Ensure your virtual environment is activated
python gui.py
```

The application's GUI will appear, guiding you through the following steps:

1.  **File Setup**: Select your Excel data file, Word template file, and specify output folders for PDFs and logs.
2.  **Email Customization**: Define the email subject and body (HTML supported) for the personalized emails.
3.  **Review & Run**: Review your settings, perform a dry run, send a test email, or start the full process.

### Example Files

The `examples` directory contains:
*   `example_data.csv`: A sample Excel-compatible CSV file with dummy data to help you get started.
*   `example_template.txt`: A sample Word template (you'll need to save this as a `.docx` file) demonstrating placeholder usage.

## Building a Standalone Executable

The project includes a batch script (`build.bat`) to create a standalone executable using `PyInstaller`. This allows you to distribute the application without requiring the end-user to install Python or its dependencies.

1.  **Install PyInstaller** (if you haven't already):
    ```bash
    pip install pyinstaller
    ```

2.  **Run the build script**:
    ```bash
    build.bat
    ```

    This script will generate the executable and all necessary dependencies in the `dist/DocuMint_Package` folder within your project directory.

## Project Structure (Detailed)

*   `.gitignore`: Specifies intentionally untracked files to ignore.
*   `build.bat`: Batch script for building the standalone executable.
*   `core.py`: Core logic for document processing, PDF conversion, and email sending.
*   `gui.py`: Main application file, implementing the `tkinter` GUI.
*   `LICENSE`: Contains the licensing information for the project (e.g., MIT License).
*   `README.md`: This comprehensive guide to the project.
*   `requirements.txt`: Lists all Python package dependencies.
*   `example_template.txt`: A text file demonstrating the structure of a Word template with placeholders.
*   `examples/`: Directory containing example data and templates.
    *   `example_data.csv`: Sample data for testing the application.
*   `venv/`: (Ignored by Git) Python virtual environment for dependency management.
*   `__pycache__/`: (Ignored by Git) Python compiled bytecode cache.
*   `build/`: (Ignored by Git) Temporary directory used by PyInstaller during the build process.
*   `dist/`: (Ignored by Git) Output directory for the standalone executable.

## Contributing

Contributions are welcome! If you have suggestions for improvements, bug reports, or want to add new features, please feel free to:

1.  Fork the repository.
2.  Create a new branch (`git checkout -b feature/YourFeature`).
3.  Make your changes.
4.  Commit your changes (`git commit -m 'Add some feature'`).
5.  Push to the branch (`git push origin feature/YourFeature`).
6.  Open a Pull Request.

## License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.

## Contact

For any questions or inquiries, please open an issue on the GitHub repository.