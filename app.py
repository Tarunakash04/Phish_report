from flask import Flask, render_template, request, redirect, url_for, send_file, session
import pandas as pd
import os
from io import BytesIO
from werkzeug.utils import secure_filename
from flask_session import Session

app = Flask(__name__)
app.secret_key = 'phishsecret'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['SESSION_TYPE'] = 'filesystem'
Session(app)

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def parse_file(file_storage):
    filename = secure_filename(file_storage.filename)
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file_storage.save(save_path)

    if filename.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(save_path, engine='openpyxl')
    elif filename.endswith('.csv'):
        df = pd.read_csv(save_path)
    else:
        raise ValueError("Unsupported file format.")
    return df

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        master_file = request.files['master_file']
        master_df = parse_file(master_file)
        master_df.columns = [str(c).strip() for c in master_df.columns]

        session['master_df'] = master_df.to_json()
        return redirect(url_for('upload_phish'))
    return render_template('index.html')

@app.route('/upload_phish', methods=['GET', 'POST'])
def upload_phish():
    if request.method == 'POST':
        if 'master_df' not in session:
            return redirect(url_for('index'))

        master_df = pd.read_json(session['master_df'])
        phish_file = request.files['phish_file']
        phish_df = parse_file(phish_file)
        phish_df.columns = [str(c).strip() for c in phish_df.columns]

        # Use exact column names
        email_col_master = 'OFFICE_EMAIL_ADDRESS'
        email_col_phish = 'email'

        if email_col_master not in master_df.columns or email_col_phish not in phish_df.columns:
            return "❌ Email column not found in one or both files.", 400

        # Merge on email
        merged_df = pd.merge(phish_df, master_df, left_on=email_col_phish, right_on=email_col_master, how='left')

        # Columns to keep
        required_cols = [
            'EMPLOYEE_CODE', 
            'Full Name', 
            'OFFICE_EMAIL_ADDRESS', 
            'status', 
            'L1_MANAGER', 
            'L2_MANAGER', 
            'SBU', 
            'DEPARTMENT', 
            'ZONE', 
            'LOCATION'
        ]

        # Filter merged columns
        merged_df = merged_df[required_cols]
        merged_df = merged_df.rename(columns={'Full Name': 'Name'})

        # Create summary sheet
        summary_df = phish_df['status'].value_counts().reset_index()
        summary_df.columns = ['Status', 'Count']

        # Store in session
        session['merged_df'] = merged_df.to_json()
        session['phish_df'] = phish_df.to_json()
        session['summary_df'] = summary_df.to_json()

        # Render table
        return render_template("summary.html",
                               summary_table_html=summary_df.to_html(index=False, classes="summary-table"))
    return render_template('upload_phish.html')

@app.route('/download')
def download():
    try:
        merged_df = pd.read_json(session['merged_df'])
        phish_df = pd.read_json(session['phish_df'])
        summary_df = pd.read_json(session['summary_df'])

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Summary Stats")
            merged_df.to_excel(writer, index=False, sheet_name="Mapped Data")
            phish_df.to_excel(writer, index=False, sheet_name="Raw Phish Data")
        output.seek(0)

        return send_file(output,
                         download_name="Phisherman_Report.xlsx",
                         as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return f"Download failed: {str(e)}"

@app.route('/reset')
def reset():
    session.clear()
    for f in os.listdir(app.config['UPLOAD_FOLDER']):
        os.remove(os.path.join(app.config['UPLOAD_FOLDER'], f))
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)