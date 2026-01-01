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
from documint.database import DatabaseManager

app = Flask(__name__)
app.secret_key = "documint_secure_key"
db = DatabaseManager(os.path.join(src_dir, '..', 'history.db'))

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

# --- Profile Management ---
PROFILES_DIR = os.path.join(src_dir, '..', 'profiles')
if not os.path.exists(PROFILES_DIR):
    os.makedirs(PROFILES_DIR)

@app.route('/api/profiles', methods=['GET'])
def list_profiles():
    files = [f for f in os.listdir(PROFILES_DIR) if f.endswith('.json')]
    return jsonify({"profiles": files})

@app.route('/api/profiles', methods=['POST'])
def save_profile():
    data = request.json
    name = data.get('name')
    config = data.get('config')
    if not name or not config:
        return jsonify({"error": "Missing name or config"}), 400
    
    # Sanitize name
    safe_name = "".join([c for c in name if c.isalnum() or c in (' ', '_', '-')]).strip()
    file_path = os.path.join(PROFILES_DIR, f"{safe_name}.json")
    
    import json
    with open(file_path, 'w') as f:
        json.dump(config, f, indent=4)
        
    return jsonify({"status": "success", "message": f"Profile '{safe_name}' saved."})

@app.route('/api/profiles/<filename>', methods=['GET'])
def load_profile(filename):
    file_path = os.path.join(PROFILES_DIR, filename)
    if not os.path.exists(file_path):
        return jsonify({"error": "Profile not found"}), 404
        
    import json
    with open(file_path, 'r') as f:
        config = json.load(f)
        
    return jsonify(config)
# --------------------------

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
            
            # Calculate Stats
            success = sum(1 for l in job_status["logs"] if "✅" in l)
            fail = sum(1 for l in job_status["logs"] if "❌" in l or "⚠️" in l)
            total = success + fail # Approx
            
            # Log to DB
            mode = "SMTP (Parallel)" if data.get('email_config', {}).get('provider') == 'smtp' else "Outlook (Serial)"
            db.log_job(total, success, fail, mode)
            
            log_callback("🏁 Job Finished. Stats saved to History.")
            pythoncom.CoUninitialize()

    thread = threading.Thread(target=background_task)
    thread.start()
    
    return jsonify({"status": "success", "message": "Job started in background"})

@app.route('/api/status')
def status():
    return jsonify(job_status)

@app.route('/api/stats')
def get_stats():
    return jsonify(db.get_stats())

if __name__ == '__main__':
    print("🌍 DocuMint Web Server Running at http://localhost:5000")
    app.run(debug=True, port=5000)
