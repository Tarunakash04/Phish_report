# 📄 Excel Summarizer & PhishMapper

Welcome to **Excel Summarizer**, now enhanced as a phishing campaign analysis and mapping tool ("PhishMapper")!  
This Flask-based app automates your manual process of merging phishing simulation results with employee master data, and generates clean Excel reports.

---

## 🚀 Features

✅ Upload employee master sheet (Name, Email, Team, Designation, etc.)  
✅ Upload phishing result sheet (Name, Action columns — Clicked, Opened, Ignored, etc.)  
✅ Auto-map by Name and merge extra details  
✅ Generate a fully formatted Excel summary report  
✅ Download final merged file in seconds  
✅ Clean Flask web UI with step-wise uploads

---

## 🏗️ How it works

### 1️⃣ Upload Master Sheet

- Contains official employee data.
- Supported formats: `.xlsx`, `.xls`, `.csv`.

---

### 2️⃣ Upload Phishing Result Sheet

- Contains campaign results.
- Must include at least **Name** and **Action** columns.
- Supported formats: `.xlsx`, `.xls`, `.csv`.

---

### 3️⃣ Auto Mapping

- The system matches **Name** between the two sheets.
- Appends columns like Email, Team, Designation to each phishing record.

---

### 4️⃣ Download

- Generates a new Excel file with all merged columns.
- Ready for further analysis or direct reporting.

---

## 🌐 Usage

1. Open your browser and go to: `http://127.0.0.1:5000`
2. Upload your **Employee Master Sheet** first.
3. Upload your **Phishing Result Sheet** next.
4. View merged summary instantly.
5. Click **Download Excel** to save your report.

---

## 🛡️ Security

- All processing runs locally.
- No data leaves your system.
- Uploaded files and sessions cleared after reset.

---

## 💡 Example use case

> You split 24,000 employees into 3 waves of 8,000 each. You send phishing emails monthly to different groups, rotating every quarter. After each campaign, upload the result sheet, and get an auto-generated merged report — with details like Designation, Team, Manager — all in one file. Quick, consistent, and no manual VLOOKUP nightmares.

---

## 📝 License

MIT License.  
See `License` file for full details.
