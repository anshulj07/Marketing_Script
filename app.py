import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

app = Flask(__name__)
app.secret_key = os.getenv('SECRECT_KEY')

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

SENDGRID_API_KEY = os.getenv('SENDGRID_KEY')
EMAIL = os.getenv('EMAIL_ID')
# Function to send emails
def send_email(to_email, to_name):
    message = Mail(
        from_email= EMAIL,
        to_emails=to_email,
        subject=f'Hello {to_name}, Here’s Your Mail!',
        html_content=f'<strong>Dear {to_name},</strong><br><br>This is a test mail sent via SendGrid!'
    )
    
    try:
        sg = SendGridAPIClient(SENDGRID_API_KEY)
        response = sg.send(message)
        print(f'Email sent to {to_email} | Status: {response.status_code}')
    except Exception as e:
        print(f'Error sending email to {to_email}: {e}')

# Route for uploading file and sending emails
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part in the request')
            return redirect(request.url)
        
        file = request.files['file']
        
        # Check if a file was selected
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        # Ensure the file is an Excel file
        if not file.filename.endswith('.xlsx'):
            flash('Only Excel files with .xlsx extension are allowed!')
            return redirect(request.url)
        
        # Validate the content without saving the file first
        try:
            # Use the file stream directly to avoid saving it before validation
            df = pd.read_excel(file)

            # Check for required columns
            if 'Email' not in df.columns or 'Name' not in df.columns:
                flash('Excel file must contain "Email" and "Name" columns')
                return redirect(request.url)

        except Exception as e:
            flash(f'Error processing file: {e}')
            return redirect(request.url)
        
        # If validation passed, save the file
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        # Send the emails after the file has been saved
        for index, row in df.iterrows():
            send_email(row['Email'], row['Name'])
        
        flash('Emails sent successfully!')
        return redirect(url_for('upload_file'))
    
    return render_template('upload.html')

if __name__ == "__main__":
    app.run(debug=True)
