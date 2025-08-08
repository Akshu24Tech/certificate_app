import os
import pandas as pd
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from flask import Flask, render_template, request, jsonify

# --- Flask App Initialization ---
app = Flask(__name__)

# --- Configuration ---
# Define paths relative to the app's location
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, 'assets')
UPLOADS_DIR = os.path.join(BASE_DIR, 'uploads')
CERTIFICATES_DIR = os.path.join(BASE_DIR, 'certificates')
OUTPUT_CSV_PATH = os.path.join(CERTIFICATES_DIR, 'output.csv')

# Create necessary directories if they don't exist
os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(CERTIFICATES_DIR, exist_ok=True)


# --- PART 1: CSV Compression Logic ---
def compress_csv_logic(input_path):
    """Reads the input CSV, processes it, and saves the output."""
    try:
        df = pd.read_csv(input_path)
        df["Program Name"] = df["Program Name"].astype(str)
        df["Semester"] = df["Semester"].astype(str)
        df["Program & Semester"] = df["Program Name"] + " (" + df["Semester"] + ")"
        
        df_grouped = (
            df.groupby(["Name", "Designation", "Email"])["Program & Semester"]
            .apply(lambda x: "; ".join(sorted(set(x))))
            .reset_index()
        )
        
        df_grouped.to_csv(OUTPUT_CSV_PATH, index=False)
        return {"status": "success", "message": f"CSV compressed successfully. Output saved to {OUTPUT_CSV_PATH}"}
    except Exception as e:
        return {"status": "error", "message": str(e)}


# --- PART 2: PDF Generation and Emailing Logic ---

# Custom PDF class from your script, with updated paths for assets
class CertificateGenerator(FPDF):
    def __init__(self):
        super().__init__()
        # Correctly path the fonts from the /assets directory
        self.add_font('OpenSans', '', os.path.join(ASSETS_DIR, 'OpenSans-Regular.ttf'), uni=True)
        self.add_font('OpenSans', 'B', os.path.join(ASSETS_DIR, 'OpenSans-Bold.ttf'), uni=True)
    
    def header(self):
        # Correctly path the certificate template from the /assets directory
        self.image(os.path.join(ASSETS_DIR, "certificate_template.png"), x=0, y=0, w=210, h=297)

    def create_program_table_with_semesters(self, programs, semesters, start_x=15, start_y=115):
        self.set_xy(start_x, start_y)
        self.set_font("OpenSans", "B", 13)
        col1_width, col2_width, col3_width, row_height = 25, 110, 50, 8
        
        self.cell(col1_width, row_height, "Sr.No.", border=1, align='C')
        self.cell(col2_width, row_height, "Program Name", border=1, align='C')
        self.cell(col3_width, row_height, "Semester(s)", border=1, align='C', ln=1)
        
        self.set_font("OpenSans", "", 13)
        for i, (program, semester) in enumerate(zip(programs, semesters), 1):
            self.set_x(start_x)
            self.cell(col1_width, row_height, str(i), border=1, align='C')
            self.cell(col2_width, row_height, program, border=1, align='L')
            self.cell(col3_width, row_height, semester, border=1, align='C', ln=1)
            
        table_end_y = start_y + (len(programs) + 1) * row_height
        signature_x = start_x + col1_width + col2_width - 15
        signature_y = table_end_y + 10
        
        try:
            self.image(os.path.join(ASSETS_DIR, "Signature.png"), x=signature_x, y=signature_y, w=50, h=35)
        except Exception:
            print("Warning: Signature.png not found")
            
        self.set_font("OpenSans", "B", 13)
        controller_x = start_x + col1_width + col2_width - 18
        self.set_xy(controller_x, signature_y + 35)
        self.cell(50, 9, "Controller of Examinations", align='C')

    def add_text(self, name, designation, sr, program_semester_data):
        self.set_font("OpenSans", "", 13)
        self.set_x(10)
        self.cell(10, 70, f"COE/GUG/2025/209 (1-680)/{sr}", ln=1)
        self.set_xy(17, 68)
        self.multi_cell(0, 9, f"This is to certify that {name}, {designation} has set the Question Paper for below mentioned Program during May 2025 examinations. The Question Paper was meticulously prepared, keeping in mind the curriculum requirements and other instructions of the University.", align='L')
        
        all_programs, all_semesters = [], []
        if pd.notna(program_semester_data) and str(program_semester_data).strip():
            programs_list = [p.strip() for p in str(program_semester_data).split(';') if p.strip()]
            for program_item in programs_list:
                if '(' in program_item and ')' in program_item:
                    last_open = program_item.rfind('(')
                    program_name = program_item[:last_open].strip()
                    semester_num = program_item[last_open+1:program_item.rfind(')')].strip()
                    all_programs.append(program_name)
                    all_semesters.append(semester_num)
                else:
                    all_programs.append(program_item)
                    all_semesters.append("N/A")
        
        self.create_program_table_with_semesters(all_programs, all_semesters)

def send_certificate_email(recipient_email, certificate_file, logs):
    """Sends a single certificate email."""
    # IMPORTANT: Update with your email credentials
    sender_email = os.environ.get("SENDER_EMAIL")  # <-- REPLACE WITH YOUR EMAIL
    sender_password = os.environ.get("SENDER_PASSWORD") # <-- REPLACE WITH YOUR GMAIL APP PASSWORD
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "Certificate of Question Paper Setting"
    body = "Dear Sir/Madam,\n\nPlease find attached your certificate for setting the Question Paper during May-2025 examinations.\n\nBest regards,\nController of Examinations"
    msg.attach(MIMEText(body, 'plain'))
    
    with open(certificate_file, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(certificate_file)}')
    msg.attach(part)
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        logs.append(f"✅ Certificate sent successfully to {recipient_email}")
    except Exception as e:
        logs.append(f"❌ FAILED to send email to {recipient_email}: {str(e)}")

def generate_and_send_logic():
    """Reads the compressed CSV, generates all PDFs, and sends emails."""
    logs = []
    try:
        data = pd.read_csv(OUTPUT_CSV_PATH)
        for idx, row in data.iterrows():
            pdf = CertificateGenerator()
            pdf.add_page()
            sr_number = idx + 1
            
            pdf.add_text(row['Name'], row['Designation'], sr_number, row['Program & Semester'])
            
            person_name = row['Name'].replace(' ', '_').replace('.', '').replace(',', '')
            filename = f"certificate_{person_name}.pdf"
            filepath = os.path.join(CERTIFICATES_DIR, filename)
            pdf.output(filepath)
            
            recipient_email = row.get('Email', '')
            if pd.notna(recipient_email) and '@' in recipient_email:
                logs.append(f"📧 Sending certificate for {row['Name']} to {recipient_email}...")
                send_certificate_email(recipient_email, filepath, logs)
            else:
                logs.append(f"⚠️ No valid email for {row['Name']}. PDF saved at {filepath}")
        
        logs.append("✅ All certificates processed successfully!")
        return {"status": "success", "logs": logs}

    except FileNotFoundError:
        logs.append("❌ Error: 'output.csv' not found. Please compress a CSV file first.")
        return {"status": "error", "logs": logs}
    except Exception as e:
        logs.append(f"❌ An unexpected error occurred: {str(e)}")
        return {"status": "error", "logs": logs}

# --- Flask Routes (The Web Server Endpoints) ---

@app.route('/')
def index():
    """Serves the main HTML page."""
    return render_template('index.html')

@app.route('/compress', methods=['POST'])
def compress_route():
    """Handles the file upload and triggers the CSV compression."""
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "No file part in the request."}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"status": "error", "message": "No file selected."}), 400
    if file and file.filename.endswith('.csv'):
        input_path = os.path.join(UPLOADS_DIR, 'input.csv')
        file.save(input_path)
        result = compress_csv_logic(input_path)
        return jsonify(result)
    return jsonify({"status": "error", "message": "Invalid file type. Please upload a .csv file."}), 400

@app.route('/send', methods=['POST'])
def send_route():
    """Triggers the PDF generation and sending process."""
    result = generate_and_send_logic()
    return jsonify(result)


if __name__ == '__main__':
    # Runs the Flask app
    app.run(debug=True)
