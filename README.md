# DocuMint

DocuMint is a professional, production-ready desktop application for automating the generation and emailing of personalized documents. It reads data from an Excel file, merges it into a Word template, and sends the resulting PDF via Outlook.

![DocuMint Screenshot](screenshot.png)

## Features

*   **Dynamic Placeholders**: Use any column from your Excel file as a placeholder in your Word template.
*   **Modern UI**: A sleek, professional dark theme inspired by the Microsoft Store UI.
*   **Wizard-Style Workflow**: An intuitive, step-by-step guide to set up and run the process.
*   **Responsive Design**: The application window is fully responsive and adapts to any screen size.
*   **Advanced Settings**: Configure the PDF filename format, email retries, and delay.
*   **Configuration Persistence**: Your settings are saved when you close the application and reloaded when you open it again.
*   **In-App Instructions**: A comprehensive user guide is available within the application.
*   **Standalone Executable**: A build script is included to create a single executable file for easy distribution.

## Example Files

The `examples` directory contains:

*   `example_data.csv`: An example Excel file with dummy data.
*   `example_template.txt`: An example Word template with all the available placeholders.

## Getting Started

### Prerequisites

*   Microsoft Outlook
*   Microsoft Word

### Installation

1.  Clone the repository:
    ```
    git clone https://github.com/Z-root-X/DocuMint.git
    ```
2.  Navigate to the project directory:
    ```
    cd documint
    ```
3.  Install the required Python libraries:
    ```
    pip install -r requirements.txt
    ```

## Usage

To run the application, execute the following command from the root directory:

```
python -m src.documint.gui
```

Alternatively, you can run the `DocuMint.exe` file from the `dist/DocuMint_Package` folder after building the application.

## How to Build

To build the standalone executable, you need to have `PyInstaller` installed (`pip install pyinstaller`).

Then, simply run the `build.bat` script:

```
build.bat
```

The final package will be created in the `dist/DocuMint_Package` folder.
