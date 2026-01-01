import os
import sys
import threading
import pythoncom
from flask import Flask, render_template, request, jsonify, send_from_directory

# Add the src directory to path so we can import documint.core
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.dirname(current_dir)
sys.path.append(src_dir)

from documint.core import process_emails, validate_placeholders

app = Flask(__name__)
app.secret_key = "documint_secure_key"

# Global Status for Progress Tracking
job_status = {
    "is_running": False,
    "logs": []
}

def log_callback(message):
    """Callback to capture logs from core.py"""
    job_status["logs"].append(message)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/docs')
def docs():
    return render_template('docs.html')

@app.route('/download/<filename>')
def download_file(filename):
    # Examples are in project_root/examples
    # src_dir is project_root/src
    examples_dir = os.path.abspath(os.path.join(src_dir, '..', 'examples'))
    return send_from_directory(examples_dir, filename, as_attachment=True)

@app.route('/api/validate', methods=['POST'])
def validate():
    data = request.json
    data_path = data.get('data_path')
    template_path = data.get('template_path')
    
    if not os.path.exists(data_path) or not os.path.exists(template_path):
        return jsonify({"valid": False, "error": "Files not found on server."})

    is_valid, missing = validate_placeholders(data_path, template_path)
    return jsonify({"valid": is_valid, "missing": missing})

@app.route('/api/run', methods=['POST'])
def run_job():
    if job_status["is_running"]:
        return jsonify({"status": "error", "message": "Job already running"})

    data = request.json
    job_status["is_running"] = True
    job_status["logs"] = ["🚀 Job Started..."]
    
    def background_task():
        # Initialize COM for Outlook on this thread
        pythoncom.CoInitialize()
        try:
            process_emails(
                data_file=data['data_path'],
                template_file=data['template_path'],
                pdf_folder=data['pdf_folder'],
                logs_folder=data['logs_folder'],
                log_callback=log_callback,
                email_subject=data['subject'],
                email_body=data['body'],
                pdf_filename_format=data['filename_format'],
                retries=int(data['retries']),
                delay=int(data['delay']),
                email_config=data['email_config'],
                dry_run=False
            )
        except Exception as e:
            log_callback(f"❌ Critical Error: {str(e)}")
        finally:
            job_status["is_running"] = False
            log_callback("🏁 Job Finished.")
            pythoncom.CoUninitialize()

    thread = threading.Thread(target=background_task)
    thread.start()
    
    return jsonify({"status": "success", "message": "Job started in background"})

@app.route('/api/status')
def status():
    return jsonify(job_status)

if __name__ == '__main__':
    print("🌍 DocuMint Web Server Running at http://localhost:5000")
    app.run(debug=True, port=5000)
