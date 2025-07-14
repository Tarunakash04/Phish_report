from flask import Flask, render_template, request, redirect, url_for, send_file, session
import pandas as pd
import os
from io import BytesIO, StringIO
from werkzeug.utils import secure_filename
from flask_session import Session
from collections import Counter

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
        master_df = pd.read_json(StringIO(session['master_df']))
        uploaded_files = request.files.getlist('phish_file')

        phish_reports = []

        for phish_file in uploaded_files:
            phish_df = parse_file(phish_file)
            phish_df.columns = [str(c).strip() for c in phish_df.columns]
            phish_reports.append(phish_df.to_json())

        session['phish_reports'] = phish_reports
        return redirect(url_for('summary'))

    return render_template('upload_phish.html')

@app.route('/summary')
def summary():
    if 'phish_reports' not in session or not session['phish_reports']:
        return redirect(url_for('upload_phish'))

    master_df = pd.read_json(StringIO(session['master_df']))
    phish_reports = session['phish_reports']

    consolidated_df = master_df[['EMPLOYEE_CODE', 'Full Name', 'OFFICE_EMAIL_ADDRESS',
                                 'L1_MANAGER', 'L2_MANAGER', 'SBU', 'DEPARTMENT', 'ZONE', 'LOCATION']]
    consolidated_df = consolidated_df.rename(columns={'Full Name': 'Name', 'OFFICE_EMAIL_ADDRESS': 'email'})

    all_status_dfs = []
    month_status_map = {}

    for report_json in phish_reports:
        phish_df = pd.read_json(StringIO(report_json))
        phish_df.columns = [str(c).strip() for c in phish_df.columns]

        if 'send_date' in phish_df.columns:
            dates = pd.to_datetime(phish_df['send_date'], errors='coerce')
        elif 'modified_date' in phish_df.columns:
            dates = pd.to_datetime(phish_df['modified_date'], errors='coerce')
        else:
            continue

        months = dates.dt.month.dropna()
        if not months.empty:
            most_common_month = Counter(months).most_common(1)[0][0]
            month_name = pd.to_datetime(f"2023-{most_common_month:02d}-01").strftime('%b')
            month_status_map[month_name] = phish_df[['email', 'status']].rename(columns={'status': month_name})
    
    # Deduplicate months, only one column per month
    for month, df_status in month_status_map.items():
        all_status_dfs.append(df_status)

    for status_df in all_status_dfs:
        consolidated_df = pd.merge(consolidated_df, status_df, on='email', how='left')

    # Count clicks or submissions
    status_cols = consolidated_df.columns[10:]  # from first month column onward
    consolidated_df['Count'] = consolidated_df[status_cols].apply(
        lambda row: sum(row.fillna('').str.lower().isin(['clicked', 'submitted data'])),
        axis=1
    )

    # Global summary count
    all_status_values = []
    for df in all_status_dfs:
        all_status_values.extend(df.iloc[:, 1].dropna().tolist())

    summary_df = pd.Series(all_status_values).value_counts().reset_index()
    summary_df.columns = ['Status', 'Count']

    # Prepare unmapped data
    all_uploaded_emails = set()
    for df in all_status_dfs:
        all_uploaded_emails.update(df['email'].dropna().unique())
    mapped_emails = set(consolidated_df['email'].dropna().unique())
    unmapped_emails = all_uploaded_emails - mapped_emails
    unmapped_df = pd.DataFrame({'email': list(unmapped_emails)})

    # Save
    session['consolidated_df'] = consolidated_df.to_json()
    session['summary_df'] = summary_df.to_json()
    session['unmapped_df'] = unmapped_df.to_json()

    return render_template("summary.html",
                           summary_table_html=summary_df.to_html(index=False, classes="summary-table"))

@app.route('/download')
def download():
    try:
        summary_df = pd.read_json(StringIO(session['summary_df']))
        consolidated_df = pd.read_json(StringIO(session['consolidated_df']))
        unmapped_df = pd.read_json(StringIO(session['unmapped_df']))

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Summary Stats")
            consolidated_df.to_excel(writer, index=False, sheet_name="Mapped Data")
            unmapped_df.to_excel(writer, index=False, sheet_name="Unmapped Data")

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
