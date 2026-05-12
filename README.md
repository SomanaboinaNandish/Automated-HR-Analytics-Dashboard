Overview

This project is an automated HR Analytics Dashboard built using Python and Microsoft Excel. It consolidates employee data across multiple departments and geographies, generates real-time HR insights, automates KPI reporting, and creates a professional dashboard for workforce analytics.

The system processes employee, finance, productivity, risk, and attrition datasets to generate a centralized HR reporting dashboard with automated alerts and visual analytics.

Features
HR Automation
Automated dashboard generation
Multi-sheet Excel integration
Live HR KPI calculations
Automated HR alerts
Quarterly attrition analytics
Workforce risk monitoring
Dashboard Analytics
Total employee headcount
India vs US workforce distribution
Probation tracking
Intern LWD alerts
Productivity analytics
Department-wise finance analytics
Risk distribution charts
Attrition monitoring
Automated Alerts
Intern Exit Alerts

Detects interns whose Last Working Day (LWD) is due within the next 45 days.

Probation Alerts

Identifies employees whose probation confirmation is due within the next 30 days.

Technologies Used
Technology	Purpose
Python	Automation & data processing
pandas	Data analysis
openpyxl	Excel automation
xlsxwriter	Dashboard formatting & charts
Microsoft Excel	Reporting dashboard
Workbook Structure

The dashboard processes the following Excel sheets:

Sheet Name	Purpose
India Employee Database	India employee records
US Employee Database	US employee records
RM Data	Resource allocation data
Finance	Employee salary & budget data
Productivity	Productivity tracking
Risk Report	HR risk monitoring
Offboarded Resources	Attrition & exits
Dashboard	Final analytics dashboard
KPI Metrics

The dashboard automatically calculates:

Total Headcount
Active Employees
Employees Under Probation
Active Interns
Attrition Rate
Average Productivity
Department Budget Allocation
Risk Distribution
Project Structure
HR-Automation-Dashboard/
│
├── generate_dashboard.py
├── HR_Dashboard_Data.xlsx
├── HR_Automation_Dashboard.xlsx
├── requirements.txt
└── README.md
Installation
Clone Repository
git clone <your-repo-link>
cd HR-Automation-Dashboard
Install Dependencies
py -m pip install -r requirements.txt

OR

py -m pip install pandas openpyxl xlsxwriter
Running the Project
Step 1

Place the source dataset:

HR_Dashboard_Data.xlsx

inside the project folder.

Step 2

Run the automation script:

py generate_dashboard.py
Step 3

Generated dashboard:

HR_Automation_Dashboard.xlsx

will be created automatically.

Dashboard Workflow
Source Excel Data
        ↓
Python Processing
        ↓
KPI Calculations
        ↓
Alert Generation
        ↓
Excel Dashboard Creation
Business Logic Implemented
Probation Monitoring

Employees under probation are automatically tracked using Date of Joining calculations.

Intern Tracking

Intern LWD dates are monitored for upcoming exits.

Attrition Analytics

Quarterly attrition is calculated from Offboarded Resources data.

Risk Analytics

Risk Report updates are reflected dynamically in dashboard analytics.

Sample Dashboard Components
KPI Cards
Finance Charts
Risk Pie Charts
Alert Tables
Department Analytics
Productivity Metrics
Key Learnings

This project demonstrates:

Business analytics
Workflow automation
HR analytics
Excel automation
Data engineering
Dashboard development
Reporting automation
Future Improvements
Streamlit web dashboard
SQL database integration
Power BI integration
Email alert automation
Predictive attrition analytics
Real-time dashboard refresh
Employee search functionality
Resume Description

Developed an automated HR analytics dashboard using Python, Pandas, OpenPyXL, and Excel integrating employee, finance, productivity, and risk datasets with automated KPI reporting, attrition analytics, workflow automation, and dynamic dashboard generation.

Author

Nandish Somanaboina
