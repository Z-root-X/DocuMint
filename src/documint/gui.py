import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox, simpledialog
import threading
import json
import os
from .core import process_emails

class SettingsWindow(tk.Toplevel):
    """A Toplevel window for configuring advanced settings."""
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Advanced Settings")
        self.geometry("500x300")
        self.configure(bg="#4B4B4B")

        self.pdf_filename_format_var = tk.StringVar(value=self.parent.pdf_filename_format_var.get())
        self.retries_var = tk.IntVar(value=self.parent.retries_var.get())
        self.delay_var = tk.IntVar(value=self.parent.delay_var.get())

        main_frame = ttk.Frame(self, padding="20", style='Card.TFrame')
        main_frame.pack(expand=True, fill="both")

        ttk.Label(main_frame, text="PDF Filename Format:", style='Card.TLabel').grid(row=0, column=0, sticky="w", padx=10, pady=10)
        ttk.Entry(main_frame, textvariable=self.pdf_filename_format_var, style='Dark.TEntry').grid(row=0, column=1, sticky="ew", padx=10, pady=10)

        ttk.Label(main_frame, text="Email Retries:", style='Card.TLabel').grid(row=1, column=0, sticky="w", padx=10, pady=10)
        ttk.Spinbox(main_frame, from_=0, to=10, textvariable=self.retries_var).grid(row=1, column=1, sticky="ew", padx=10, pady=10)

        ttk.Label(main_frame, text="Email Delay (seconds):", style='Card.TLabel').grid(row=2, column=0, sticky="w", padx=10, pady=10)
        ttk.Spinbox(main_frame, from_=0, to=60, textvariable=self.delay_var).grid(row=2, column=1, sticky="ew", padx=10, pady=10)

        button_frame = ttk.Frame(main_frame, style='Card.TFrame')
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)

        ttk.Button(button_frame, text="Save", command=self.save_and_close, style='Primary.TButton').pack(side="left", padx=10)
        ttk.Button(button_frame, text="Cancel", command=self.destroy, style='Secondary.TButton').pack(side="left", padx=10)

    def save_and_close(self):
        """Saves the settings and closes the window."""
        self.parent.pdf_filename_format_var.set(self.pdf_filename_format_var.get())
        self.parent.retries_var.set(self.retries_var.get())
        self.parent.delay_var.set(self.delay_var.get())
        self.destroy()

class DocuMint(tk.Tk):
    """The main application window for DocuMint."""
    def __init__(self):
        super().__init__()

        self.title("DocuMint")
        self.geometry("1024x768")
        self.minsize(900, 600)
        self.configure(bg="#4B4B4B")

        self.config_file = "config.json"

        # Configuration variables
        self.data_file_var = tk.StringVar()
        self.template_file_var = tk.StringVar()
        self.pdf_folder_var = tk.StringVar()
        self.logs_folder_var = tk.StringVar()
        self.pdf_filename_format_var = tk.StringVar(value="Admit_{<ID>}")
        self.retries_var = tk.IntVar(value=2)
        self.delay_var = tk.IntVar(value=2)

        # Style configuration
        self.setup_styles()

        # Header for settings and help buttons
        header = ttk.Frame(self, style='App.TFrame')
        header.place(relx=1.0, rely=0, anchor='ne')
        self.settings_button = ttk.Button(header, text="⚙️", style='Icon.TButton', command=self.open_settings)
        self.settings_button.pack(side="right", padx=5, pady=5)
        self.help_button = ttk.Button(header, text="❓", style='Icon.TButton', command=self.show_instructions)
        self.help_button.pack(side="right", padx=5, pady=5)
        self.about_button = ttk.Button(header, text="ℹ️", style='Icon.TButton', command=self.show_about)
        self.about_button.pack(side="right", padx=5, pady=5)

        # Main container
        container = ttk.Frame(self, style='App.TFrame')
        container.pack(expand=True, fill="both")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (WelcomePage, FileSetupPage, EmailPage, RunPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.load_config()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.show_frame("WelcomePage")

    def setup_styles(self):
        """Sets up the styling for the application's widgets."""
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('App.TFrame', background='#4B4B4B')
        style.configure('Page.TFrame', background='#4B4B4B')
        style.configure('Card.TFrame', background='#5c5c5c')
        style.configure('H1.TLabel', background='#4B4B4B', foreground='#D3D3D3', font=("Segoe UI", 24, "bold"))
        style.configure('H2.TLabel', background='#4B4B4B', foreground='#D3D3D3', font=("Segoe UI", 14))
        style.configure('Card.TLabel', background='#5c5c5c', foreground='#D3D3D3', font=("Segoe UI", 11))
        style.configure('Review.TLabel', background='#5c5c5c', foreground='#D3D3D3', font=("Segoe UI", 12, "bold"))
        style.configure('Primary.TButton', background='#6A5ACD', foreground='#ffffff', font=("Segoe UI", 12, "bold"), padding=(20, 10), borderwidth=0)
        style.map('Primary.TButton', background=[('active', '#836FFF')])
        style.configure('Secondary.TButton', background='#5c5c5c', foreground='#ffffff', font=("Segoe UI", 10), padding=(10, 5), borderwidth=1, relief="solid")
        style.map('Secondary.TButton', background=[('active', '#6c6c6c')])
        style.configure('Icon.TButton', background='#4B4B4B', foreground='#D3D3D3', font=("Segoe UI", 12), padding=5, borderwidth=0)
        style.map('Icon.TButton', background=[('active', '#5c5c5c')])

    def open_settings(self):
        """Opens the advanced settings window."""
        SettingsWindow(self)

    def show_instructions(self):
        """Shows the instructions window."""
        instruction_window = tk.Toplevel(self)
        instruction_window.title("Instructions")
        instruction_window.geometry("900x700")
        instruction_window.configure(bg="#4B4B4B")

        instruction_text = scrolledtext.ScrolledText(instruction_window, wrap=tk.WORD, relief="flat", padx=20, pady=20, bg="#4B4B4B", fg="#D3D3D3")
        instruction_text.pack(expand=True, fill="both")
        
        instructions = self.get_instructions()
        
        # Add tags for styling
        instruction_text.tag_configure('h1', font=('Segoe UI', 20, 'bold'), foreground='#FF7F50', spacing3=15)
        instruction_text.tag_configure('h2', font=('Segoe UI', 14, 'bold'), foreground='#D3D3D3', spacing3=10)
        instruction_text.tag_configure('bold', font=('Segoe UI', 11, 'bold'), foreground='#FF7F50')
        instruction_text.tag_configure('italic', font=('Segoe UI', 11, 'italic'))
        instruction_text.tag_configure('normal', font=('Segoe UI', 11))
        instruction_text.tag_configure('code', font=('Courier New', 11), background='#5c5c5c')

        for line in instructions.split('\n'):
            if line.startswith('==='):
                instruction_text.insert(tk.END, line.strip('=').strip() + '\n\n', 'h1')
            elif line.startswith('---'):
                instruction_text.insert(tk.END, line.strip('-').strip() + '\n', 'h2')
            elif line.startswith('*'):
                instruction_text.insert(tk.END, '  •  ', 'bold')
                instruction_text.insert(tk.END, line[1:].strip() + '\n', 'normal')
            elif '`' in line:
                parts = line.split('`')
                for i, part in enumerate(parts):
                    if i % 2 == 1:
                        instruction_text.insert(tk.END, part, 'code')
                    else:
                        instruction_text.insert(tk.END, part, 'normal')
                instruction_text.insert(tk.END, '\n')
            else:
                instruction_text.insert(tk.END, line + '\n', 'normal')

        instruction_text.config(state="disabled")

    def get_instructions(self):
        """Returns the instruction text."""
        return '''=================================================
 DocuMint - User Guide
=================================================

Welcome to DocuMint! This guide will help you get started.

----------------
1. Overview
----------------

DocuMint is a powerful tool designed to automate the entire process of generating and emailing personalized documents. It reads data from an Excel file, merges it into a Word template, and sends the resulting PDF via Outlook.

----------------
2. Dynamic Placeholders
----------------

DocuMint now supports dynamic placeholders! This means you can use any column from your Excel file as a placeholder in your Word template. For example, if you have a column named `CourseName` in your Excel file, you can use `<CourseName>` in your Word template, and DocuMint will automatically replace it with the correct data.

----------------
3. Required Libraries
----------------

To run this application, you need to install the following Python libraries:

*   `pandas`
*   `python-docx`
*   `pywin32`

You can install them using pip:
`pip install pandas python-docx pywin32`

----------------
4. Advanced Settings
----------------

Click the ⚙️ icon to configure:

*   **PDF Filename Format**: Define a custom format for the generated PDF files. You can use any column from your Excel file as a placeholder, e.g., `Admit_{<ID>}_{<Name>}`.
*   **Email Retries**: Set the number of times the application should attempt to resend a failed email.
*   **Email Delay**: Specify a delay (in seconds) between sending each email to avoid potential issues with email servers.

================================================='''

    def show_about(self):
        """Shows the about window."""
        messagebox.showinfo("About DocuMint", "DocuMint\nVersion: 10.0\n\nDeveloped by Gemini")

    def show_frame(self, page_name):
        """Shows a frame for the given page name."""
        frame = self.frames[page_name]
        if hasattr(frame, 'on_show'):
            frame.on_show()
        frame.tkraise()

    def on_closing(self):
        """Called when the application window is closed."""
        self.save_config()
        self.destroy()

    def save_config(self):
        """Saves the current configuration to a file."""
        email_page = self.frames["EmailPage"]
        config = {
            "data_file": self.data_file_var.get(),
            "template_file": self.template_file_var.get(),
            "pdf_folder": self.pdf_folder_var.get(),
            "logs_folder": self.logs_folder_var.get(),
            "email_subject": email_page.email_subject_entry.get(),
            "email_body": email_page.email_body_text.get("1.0", tk.END),
            "pdf_filename_format": self.pdf_filename_format_var.get(),
            "retries": self.retries_var.get(),
            "delay": self.delay_var.get()
        }
        with open(self.config_file, 'w') as f:
            json.dump(config, f, indent=4)

    def load_config(self):
        """Loads the configuration from a file."""
        try:
            with open(self.config_file, 'r') as f:
                config = json.load(f)
                self.data_file_var.set(config.get("data_file", ""))
                self.template_file_var.set(config.get("template_file", ""))
                self.pdf_folder_var.set(config.get("pdf_folder", ""))
                self.logs_folder_var.set(config.get("logs_folder", ""))
                self.pdf_filename_format_var.set(config.get("pdf_filename_format", "Admit_{<ID>}"))
                self.retries_var.set(config.get("retries", 2))
                self.delay_var.set(config.get("delay", 2))

                email_page = self.frames["EmailPage"]
                email_page.email_subject_entry.delete(0, tk.END)
                email_page.email_subject_entry.insert(0, config.get("email_subject", "Your Admit Card and Instructions"))
                email_page.email_body_text.delete("1.0", tk.END)
                email_page.email_body_text.insert(tk.END, config.get("email_body", email_page.get_default_email_body()))
        except FileNotFoundError:
            pass # No config file yet

class WelcomePage(ttk.Frame):
    """The welcome page of the application."""
    def __init__(self, parent, controller):
        super().__init__(parent, style='Page.TFrame')
        self.controller = controller

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main_content = ttk.Frame(self, style='Page.TFrame')
        main_content.grid(row=0, column=0)

        ttk.Label(main_content, text="Welcome to DocuMint", style='H1.TLabel').pack(pady=(10, 10))
        ttk.Label(main_content, text="The professional solution for automating your document workflow.", style='H2.TLabel').pack(pady=10)
        ttk.Button(main_content, text="Get Started", style='Primary.TButton', command=lambda: controller.show_frame("FileSetupPage")).pack(pady=40)

class FileSetupPage(ttk.Frame):
    """The page for setting up the required files."""
    def __init__(self, parent, controller):
        super().__init__(parent, style='Page.TFrame')
        self.controller = controller

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main_content = ttk.Frame(self, style='Page.TFrame')
        main_content.grid(row=0, column=0)

        ttk.Label(main_content, text="Step 1: Setup Your Files", style='H1.TLabel').pack(pady=(20, 20))

        card = ttk.Frame(main_content, style='Card.TFrame', padding=40)
        card.pack()

        self.create_browse_row(card, "Data File (Excel):", self.controller.data_file_var, self.browse_data_file, 0)
        self.create_browse_row(card, "Template File (DOCX):", self.controller.template_file_var, self.browse_template_file, 1)
        self.create_browse_row(card, "PDF Output Folder:", self.controller.pdf_folder_var, self.browse_pdf_folder, 2)
        self.create_browse_row(card, "Logs Folder:", self.controller.logs_folder_var, self.browse_logs_folder, 3)

        nav_frame = ttk.Frame(main_content, style='Page.TFrame')
        nav_frame.pack(pady=40)
        ttk.Button(nav_frame, text="Back", style='Secondary.TButton', command=lambda: controller.show_frame("WelcomePage")).pack(side="left", padx=10)
        ttk.Button(nav_frame, text="Next", style='Primary.TButton', command=lambda: controller.show_frame("EmailPage")).pack(side="left", padx=10)

    def create_browse_row(self, parent, label_text, var, command, row):
        ttk.Label(parent, text=label_text, style='Card.TLabel').grid(row=row, column=0, sticky="w", padx=10, pady=10)
        ttk.Entry(parent, textvariable=var, width=60).grid(row=row, column=1, padx=10, pady=10)
        ttk.Button(parent, text="Browse...", style='Secondary.TButton', command=command).grid(row=row, column=2, padx=10, pady=10)

    def browse_data_file(self):
        file = filedialog.askopenfilename(title="Select Excel Data File", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file:
            self.controller.data_file_var.set(file)

    def browse_template_file(self):
        file = filedialog.askopenfilename(title="Select DOCX Template", filetypes=[("Word Documents", "*.docx")])
        if file:
            self.controller.template_file_var.set(file)

    def browse_pdf_folder(self):
        folder = filedialog.askdirectory(title="Select PDF Output Folder")
        if folder:
            self.controller.pdf_folder_var.set(folder)

    def browse_logs_folder(self):
        folder = filedialog.askdirectory(title="Select Logs Folder")
        if folder:
            self.controller.logs_folder_var.set(folder)

class EmailPage(ttk.Frame):
    """The page for customizing the email content."""
    def __init__(self, parent, controller):
        super().__init__(parent, style='Page.TFrame')
        self.controller = controller

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        main_content = ttk.Frame(self, style='Page.TFrame')
        main_content.grid(row=0, column=0)

        ttk.Label(main_content, text="Step 2: Customize Your Email", style='H1.TLabel').pack(pady=(20, 20))

        card = ttk.Frame(main_content, style='Card.TFrame', padding=40)
        card.pack()

        ttk.Label(card, text="Subject:", style='Card.TLabel').grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.email_subject_entry = ttk.Entry(card, width=80)
        self.email_subject_entry.grid(row=0, column=1, padx=10, pady=10)

        ttk.Label(card, text="Body (HTML):", style='Card.TLabel').grid(row=1, column=0, sticky="nw", padx=10, pady=10)
        self.email_body_text = scrolledtext.ScrolledText(card, wrap=tk.WORD, width=80, height=15, relief="flat")
        self.email_body_text.grid(row=1, column=1, padx=10, pady=10)

        nav_frame = ttk.Frame(main_content, style='Page.TFrame')
        nav_frame.pack(pady=40)
        ttk.Button(nav_frame, text="Back", style='Secondary.TButton', command=lambda: controller.show_frame("FileSetupPage")).pack(side="left", padx=10)
        ttk.Button(nav_frame, text="Next", style='Primary.TButton', command=lambda: controller.show_frame("RunPage")).pack(side="left", padx=10)

    def get_default_email_body(self):
        """Returns the default email body."""
        return '''<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Your Document</title>
</head>
<body style="margin:0; padding:0; background-color:#f4f4f4; font-family: sans-serif;">
  <p>Dear {<Name>},</p>
  <p>Your document is attached.</p>
</body>
</html>'''

class RunPage(ttk.Frame):
    """The page for running the process and viewing logs."""
    def __init__(self, parent, controller):
        super().__init__(parent, style='Page.TFrame')
        self.controller = controller

        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        ttk.Label(self, text="Step 3: Review & Run", style='H1.TLabel').grid(row=0, column=0, columnspan=2, pady=(20, 10))

        # Review Frame
        review_card = ttk.Frame(self, style='Card.TFrame', padding=20)
        review_card.grid(row=1, column=0, sticky="nsew", padx=(20, 10), pady=10)
        ttk.Label(review_card, text="Review Settings", style='Review.TLabel').pack(pady=10)
        self.review_text = tk.Text(review_card, wrap=tk.WORD, relief="flat", bg="#5c5c5c", fg="#D3D3D3", font=("Segoe UI", 10))
        self.review_text.pack(expand=True, fill="both", padx=10, pady=10)

        # Log Frame
        log_card = ttk.Frame(self, style='Card.TFrame', padding=20)
        log_card.grid(row=1, column=1, sticky="nsew", padx=(10, 20), pady=10)
        ttk.Label(log_card, text="Log", style='Review.TLabel').pack(pady=10)
        self.log_text = scrolledtext.ScrolledText(log_card, wrap=tk.WORD, relief="flat")
        self.log_text.pack(expand=True, fill="both", padx=10, pady=10)
        self.log_text.configure(bg="#1c2833", fg="#D3D3D3")

        # Control Frame
        control_frame = ttk.Frame(self, style='Page.TFrame')
        control_frame.grid(row=2, column=0, columnspan=2, pady=20)

        self.start_button = ttk.Button(control_frame, text="▶ Start Process", style='Primary.TButton', command=self.start_process)
        self.start_button.pack(side="left", padx=10)

        self.dry_run_button = ttk.Button(control_frame, text="■ Dry Run", style='Secondary.TButton', command=self.dry_run)
        self.dry_run_button.pack(side="left", padx=10)

        self.test_email_button = ttk.Button(control_frame, text="✉ Send Test Email", style='Secondary.TButton', command=self.send_test_email)
        self.test_email_button.pack(side="left", padx=10)

        # Navigation Frame
        nav_frame = ttk.Frame(self, style='Page.TFrame')
        nav_frame.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(nav_frame, text="Back", style='Secondary.TButton', command=lambda: controller.show_frame("EmailPage")).pack()

    def on_show(self):
        """Called when the frame is shown."""
        self.review_text.config(state="normal")
        self.review_text.delete("1.0", tk.END)
        review_content = f"Data File: {self.controller.data_file_var.get()}\n"
        review_content += f"Template File: {self.controller.template_file_var.get()}\n"
        review_content += f"PDF Output Folder: {self.controller.pdf_folder_var.get()}\n"
        review_content += f"Logs Folder: {self.controller.logs_folder_var.get()}\n\n"
        review_content += f"Email Subject: {self.controller.frames['EmailPage'].email_subject_entry.get()}\n\n"
        review_content += f"PDF Filename Format: {self.controller.pdf_filename_format_var.get()}\n"
        review_content += f"Email Retries: {self.controller.retries_var.get()}\n"
        review_content += f"Email Delay: {self.controller.delay_var.get()}s"
        self.review_text.insert(tk.END, review_content)
        self.review_text.config(state="disabled")

    def start_process(self, dry_run=False, test_email=None):
        """Starts the admit card generation and email sending process."""
        if not self.validate_inputs():
            return

        if not dry_run and not test_email:
            if not messagebox.askyesno("Confirmation", "Are you sure you want to start the email sending process?"):
                return

        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.set_buttons_state("disabled")
        self.append_log("🚀 Process started...")

        threading.Thread(target=self.run_processing, args=(dry_run, test_email)).start()

    def validate_inputs(self):
        """Validates the user's inputs."""
        paths = {
            "Data File": self.controller.data_file_var.get(),
            "Template File": self.controller.template_file_var.get(),
            "PDF Output Folder": self.controller.pdf_folder_var.get(),
            "Logs Folder": self.controller.logs_folder_var.get()
        }
        for name, path in paths.items():
            if not path:
                messagebox.showerror("Input Error", f"{name} is not set.")
                return False
            if not os.path.exists(path):
                messagebox.showerror("Input Error", f"{name} not found at: {path}")
                return False
        return True

    def dry_run(self):
        """Starts the process in dry run mode."""
        self.start_process(dry_run=True)

    def send_test_email(self):
        """Sends a single test email."""
        email = simpledialog.askstring("Test Email", "Enter the email address to send a test email to:")
        if email:
            self.start_process(test_email=email)

    def run_processing(self, dry_run, test_email):
        """Runs the core processing logic in a separate thread."""
        email_page = self.controller.frames["EmailPage"]
        process_emails(
            self.controller.data_file_var.get(),
            self.controller.template_file_var.get(),
            self.controller.pdf_folder_var.get(),
            self.controller.logs_folder_var.get(),
            self.append_log,
            email_page.email_subject_entry.get(),
            email_page.email_body_text.get("1.0", tk.END),
            self.controller.pdf_filename_format_var.get(),
            self.controller.retries_var.get(),
            self.controller.delay_var.get(),
            dry_run=dry_run,
            test_email=test_email
        )
        self.append_log("🏁 Process completed.")
        self.set_buttons_state("normal")

    def set_buttons_state(self, state):
        """Sets the state of the control buttons."""
        self.start_button.config(state=state)
        self.dry_run_button.config(state=state)
        self.test_email_button.config(state=state)

    def append_log(self, message):
        """Appends a message to the log text widget."""
        self.log_text.config(state="normal")
        tag = "normal"
        if "SUCCESS" in message:
            tag = "success"
        elif "FAILED" in message or "Error" in message:
            tag = "error"
        
        self.log_text.insert(tk.END, message + "\n", tag)
        self.log_text.tag_config("success", foreground="#2ecc71", font=("Courier New", 11, "bold"))
        self.log_text.tag_config("error", foreground="#e74c3c", font=("Courier New", 11, "bold"))
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

if __name__ == "__main__":
    app = DocuMint()
    app.mainloop()