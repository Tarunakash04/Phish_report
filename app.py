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
        raise ValueError("Unsupported file type")
    return df

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        master_file = request.files['master_file']
        if not master_file:
            return "Master file missing.", 400

        master_df = parse_file(master_file)
        master_df.columns = [c.strip() for c in master_df.columns]

        session['master_df'] = master_df.to_json()
        return redirect(url_for('upload_phish'))

    return render_template('index.html')

@app.route('/upload_phish', methods=['GET', 'POST'])
def upload_phish():
    if 'master_df' not in session:
        return redirect(url_for('index'))

    if request.method == 'POST':
        master_df = pd.read_json(session['master_df'])
        phish_file = request.files['phish_file']
        if not phish_file:
            return "Phish file missing.", 400

        phish_df = parse_file(phish_file)
        phish_df.columns = [c.strip() for c in phish_df.columns]

        # Merge on OFFICE_EMAIL_ADDRESS
        merged_df = phish_df.merge(master_df, on='OFFICE_EMAIL_ADDRESS', how='left')

        # Columns to include in mapped data
        final_columns = [
            'EMPLOYEE_CODE',
            'Name',
            'OFFICE_EMAIL_ADDRESS',
            'Status',
            'L1 manager',
            'L2 manager',
            'SBU',
            'DEPARTMENT',
            'Zone',
            'Location'
        ]

        # Ensure all expected columns exist
        for col in final_columns:
            if col not in merged_df.columns:
                merged_df[col] = pd.NA

        # Reorder
        mapped_df = merged_df[final_columns]

        # Create summary count by Status
        summary_df = phish_df['Status'].value_counts().reset_index()
        summary_df.columns = ['Status', 'Count']

        # Save in session
        session['mapped_df'] = mapped_df.to_json()
        session['summary_df'] = summary_df.to_json()
        session['raw_df'] = phish_df.to_json()

        return render_template("summary.html",
                               supporting_table_html=summary_df.to_html(index=False, classes="summary-table"),
                               target='OFFICE_EMAIL_ADDRESS')

    return render_template('upload_phish.html')

@app.route('/download')
def download():
    try:
        mapped_df = pd.read_json(session['mapped_df'])
        summary_df = pd.read_json(session['summary_df'])
        raw_df = pd.read_json(session['raw_df'])

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            summary_df.to_excel(writer, sheet_name="Status Count", index=False)
            mapped_df.to_excel(writer, sheet_name="Mapped Data", index=False)
            raw_df.to_excel(writer, sheet_name="Raw Data", index=False)
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
