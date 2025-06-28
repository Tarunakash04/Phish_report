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

def read_file(file_path):
    if file_path.endswith(('.xlsx', '.xls')):
        return pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file format.")

def find_real_header(df):
    HEADER_KEYWORDS = ['email', 'event', 'team', 'manager', 'designation']
    for i, row in df.head(10).iterrows():
        if any(any(k.lower() in str(cell).lower() for k in HEADER_KEYWORDS) for cell in row):
            return i
    return 0

def parse_file(file_storage):
    filename = secure_filename(file_storage.filename)
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file_storage.save(save_path)

    if filename.endswith(('.xlsx', '.xls')):
        temp_df = pd.read_excel(save_path, header=None)
    elif filename.endswith('.csv'):
        temp_df = pd.read_csv(save_path, header=None)
    else:
        raise ValueError("Unsupported file type")

    header_row = find_real_header(temp_df)
    df = pd.read_excel(save_path, header=header_row) if filename.endswith(('.xlsx', '.xls')) else pd.read_csv(save_path, header=header_row)
    return df

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        employee_file = request.files['employee_file']
        emp_df = parse_file(employee_file)
        emp_df.columns = [str(c).strip().lower() for c in emp_df.columns]
        session['employee_df'] = emp_df.to_json()
        return redirect(url_for('upload_phish'))
    return render_template('index.html')

@app.route('/upload_phish', methods=['GET', 'POST'])
def upload_phish():
    if request.method == 'POST':
        if 'employee_df' not in session:
            return redirect(url_for('index'))

        emp_df = pd.read_json(session['employee_df'])
        phishing_file = request.files['phishing_file']
        phish_df = parse_file(phishing_file)
        phish_df.columns = [str(c).strip().lower() for c in phish_df.columns]

        # Use 'name' as the common column
        name_col_phish = next((col for col in phish_df.columns if 'name' in col), None)
        name_col_emp = next((col for col in emp_df.columns if 'name' in col), None)

        if not name_col_phish or not name_col_emp:
            return "❌ Name column not found in one or both files.", 400

        merged_df = phish_df.merge(emp_df, left_on=name_col_phish, right_on=name_col_emp, how='left')

        unmatched_df = merged_df[merged_df[emp_df.columns[0]].isna()]

        summary_cols = [col for col in ['team', 'designation', 'l1 manager'] if col in merged_df.columns]
        action_col = next((col for col in phish_df.columns if 'clicked' in col or 'action' in col or 'status' in col), None)

        if action_col:
            summary = merged_df.groupby(summary_cols + [action_col]).size().reset_index(name='Count')
        else:
            summary = pd.DataFrame([{'Note': 'Action column not found in phishing file'}])

        session['merged'] = merged_df.to_json()
        session['unmatched'] = unmatched_df.to_json()
        session['summary'] = summary.to_json()

        return render_template("summary.html",
                               target=name_col_phish,
                               supporting_table_html=summary.to_html(index=False, classes="summary-table"),
                               comparison_table_html=None,
                               model_summary_text=None)
    return render_template('upload_phish.html')

@app.route('/download')
def download():
    try:
        merged_df = pd.read_json(session['merged'])
        unmatched_df = pd.read_json(session['unmatched'])
        summary_df = pd.read_json(session['summary'])

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Enriched Data")
            unmatched_df.to_excel(writer, index=False, sheet_name="Unmatched Emails")
            summary_df.to_excel(writer, index=False, sheet_name="Summary Stats")
        output.seek(0)

        return send_file(output,
                         download_name="PhishMapper_Report.xlsx",
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
