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
        session['phish_reports'] = []
        return redirect(url_for('upload_phish'))
    return render_template('index.html')

@app.route('/upload_phish', methods=['GET', 'POST'])
def upload_phish():
    if 'master_df' not in session:
        return redirect(url_for('index'))

    if request.method == 'POST':
        master_df = pd.read_json(session['master_df'])
        uploaded_files = request.files.getlist('phish_file')

        for phish_file in uploaded_files:
            phish_df = parse_file(phish_file)
            phish_df.columns = [str(c).strip() for c in phish_df.columns]

            email_col_master = 'OFFICE_EMAIL_ADDRESS'
            email_col_phish = 'email'

            if email_col_master not in master_df.columns or email_col_phish not in phish_df.columns:
                return "❌ Email column not found in one or both files.", 400

            # Merge on email
            merged_df = pd.merge(phish_df, master_df, left_on=email_col_phish, right_on=email_col_master, how='left', indicator=True)

            # Identify unmatched
            unmatched_df = merged_df[merged_df['_merge'] == 'left_only'][[email_col_phish, 'status']]
            matched_df = merged_df[merged_df['_merge'] == 'both']

            # Filter final columns for mapped data
            required_cols = ['EMPLOYEE_CODE', 'Full Name', 'OFFICE_EMAIL_ADDRESS', 'status',
                             'L1_MANAGER', 'L2_MANAGER', 'SBU', 'DEPARTMENT', 'ZONE', 'LOCATION']
            final_mapped_df = matched_df[required_cols].rename(columns={'Full Name': 'Name'})

            # Save report info
            phish_reports = session.get('phish_reports', [])
            phish_reports.append({
                'phish_df': phish_df.to_json(),
                'merged_df': final_mapped_df.to_json(),
                'unmatched_df': unmatched_df.to_json(),
                'filename': phish_file.filename
            })
            session['phish_reports'] = phish_reports

        return redirect(url_for('summary'))

    return render_template('upload_phish.html')

@app.route('/summary')
def summary():
    if 'phish_reports' not in session or not session['phish_reports']:
        return redirect(url_for('upload_phish'))

    master_df = pd.read_json(session['master_df'])
    phish_reports = session['phish_reports']

    # Prepare consolidated employee report
    consolidated_df = master_df[['Full Name', 'OFFICE_EMAIL_ADDRESS']].rename(columns={'Full Name': 'Name', 'OFFICE_EMAIL_ADDRESS': 'email'})

    all_status = []
    for report in phish_reports:
        phish_df = pd.read_json(report['phish_df'])
        month = "Unknown"
        if 'send_date' in phish_df.columns:
            first_date = pd.to_datetime(phish_df['send_date'].iloc[0], errors='coerce')
            if not pd.isna(first_date):
                month = first_date.strftime('%b')
        status_df = phish_df[['email', 'status']].copy()
        status_df = status_df.rename(columns={'status': month})
        all_status.append(status_df)

    for status_df in all_status:
        consolidated_df = pd.merge(consolidated_df, status_df, on='email', how='left')

    # Count clicks or submissions
    status_cols = consolidated_df.columns[2:]
    consolidated_df['Count'] = consolidated_df[status_cols].apply(lambda row: sum(row.fillna('').str.lower().isin(['clicked', 'submitted data'])), axis=1)

    # Create global summary stats
    all_status_concat = pd.concat([pd.read_json(r['phish_df'])['status'] for r in phish_reports], ignore_index=True)
    summary_df = all_status_concat.value_counts().reset_index()
    summary_df.columns = ['Status', 'Count']

    session['consolidated_df'] = consolidated_df.to_json()
    session['summary_df'] = summary_df.to_json()

    return render_template("summary.html",
                           summary_table_html=summary_df.to_html(index=False, classes="summary-table"))

@app.route('/download')
def download():
    try:
        summary_df = pd.read_json(session['summary_df'])
        consolidated_df = pd.read_json(session['consolidated_df'])
        phish_reports = session['phish_reports']

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Summary Stats")
            consolidated_df.to_excel(writer, index=False, sheet_name="Season Consolidation")

            for i, report in enumerate(phish_reports, 1):
                merged_df = pd.read_json(report['merged_df'])
                unmatched_df = pd.read_json(report['unmatched_df'])
                sheet_name = f"Mapped_{i}"
                unmatched_sheet = f"Unmatched_{i}"

                merged_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
                unmatched_df.to_excel(writer, index=False, sheet_name=unmatched_sheet[:31])

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
