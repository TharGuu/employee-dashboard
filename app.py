import pandas as pd
from flask import Flask, send_file, render_template_string
import tempfile
import os

app = Flask(__name__)
TEMP_DIR = tempfile.gettempdir()

def generate_sample_data():
    """Creates sample Excel files in temporary storage"""
    # Sample data for Pattama's interviews
    pattama_data = {
        "Date": ["15-Jan-2025", "15-Jan-2025"],
        "Candidate Name": ["Mr.A", "Mr.B"],
        "Role": ["Data Analyst", "Web Developer"],
        "Interview": ["Yes", "Yes"],
        "Status": ["Pass", "Fail"],
        "Remark": ["", ""]
    }
    pd.DataFrame(pattama_data).to_excel(f"{TEMP_DIR}/Daily_report_Pattama.xlsx", index=False)

    # Sample data for Raewwadee's interviews
    raewwadee_data = {
        "Date": ["15-Jan-2025", "15-Jan-2025"],
        "Candidate Name": ["Mr.C", "Mr.D"],
        "Role": ["Software Tester", "Project Coordinator"],
        "Interview": ["Yes", "Yes"],
        "Status": ["Fail", "Pass"],
        "Remark": ["", ""]
    }
    pd.DataFrame(raewwadee_data).to_excel(f"{TEMP_DIR}/Daily_report_Raewwadee.xlsx", index=False)

    # Sample data for new employees
    new_employee_data = {
        "Employee Name": ["Mr.A", "Mr.D"],
        "Join Date": ["3-Feb-2025", "17-Feb-2025"],
        "Role": ["Data Analyst", "Project Coordinator"],
        "DOB": ["01-01-2000", "01-12-2001"],
        "ID Card": ["1-1111-11111-11-1", "2-2222-2222-22-2"],
        "Remark": ["", ""]
    }
    pd.DataFrame(new_employee_data).to_excel(f"{TEMP_DIR}/New_Employees.xlsx", index=False)

def create_dashboard():
    """Processes data and creates final dashboard"""
    generate_sample_data()
    
    # Load all Excel files
    pattama = pd.read_excel(f"{TEMP_DIR}/Daily_report_Pattama.xlsx")
    raewwadee = pd.read_excel(f"{TEMP_DIR}/Daily_report_Raewwadee.xlsx")
    new_employees = pd.read_excel(f"{TEMP_DIR}/New_Employees.xlsx")
    
    # Add interviewer names
    pattama["Interviewer"] = "Pattama Sooksan"
    raewwadee["Interviewer"] = "Raewwadee Jaidee"
    
    # Combine and merge data
    all_interviews = pd.concat([pattama, raewwadee])
    result = pd.merge(
        new_employees,
        all_interviews[["Candidate Name", "Role", "Interviewer"]],
        left_on=["Employee Name", "Role"],
        right_on=["Candidate Name", "Role"],
        how="left"
    )
    final = result[["Employee Name", "Join Date", "Role", "Interviewer"]].dropna()
    final.to_excel(f"{TEMP_DIR}/Employee_Dashboard.xlsx", index=False)
    return final

@app.route("/")
def home():
    """Main endpoint that shows data in browser and offers download"""
    df = create_dashboard()
    
    # Create HTML table
    html_table = df.to_html(classes='table table-striped', index=False)
    
    # Simple HTML page with download button
    html = f"""
    <html>
        <head>
            <title>Employee Dashboard</title>
            <style>
                table {{ width: 80%; margin: 20px auto; border-collapse: collapse; }}
                th, td {{ padding: 10px; border: 1px solid #ddd; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                .download-btn {{ 
                    display: block; 
                    width: 200px; 
                    margin: 20px auto; 
                    padding: 10px; 
                    text-align: center; 
                    background: #4CAF50; 
                    color: white; 
                    text-decoration: none;
                }}
            </style>
        </head>
        <body>
            <h1 style="text-align: center;">Employee Dashboard</h1>
            {html_table}
            <a href="/download" class="download-btn">Download Excel</a>
        </body>
    </html>
    """
    return html

@app.route("/download")
def download():
    """Endpoint to download the Excel file"""
    return send_file(
        f"{TEMP_DIR}/Employee_Dashboard.xlsx",
        as_attachment=True,
        download_name="Employee_Dashboard.xlsx"
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))