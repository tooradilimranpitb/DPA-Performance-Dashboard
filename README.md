DPA Performance & Target Management Dashboard
A comprehensive, interactive web application built with Streamlit for tracking, analyzing, and managing Daily Performance Analytics (DPA) and scanning targets across districts and supervisors.

🚀 Key Features
Executive Performance Analytics: Instant visibility into total scanned pages, active DPAs, averages, and working days tracked.

Interactive Dashboards & Leaderboards: View granular performance ratings (🌟 Exceptional, ✅ On Track, ⚠️ Needs Support), daily/fortnightly trends, and comparative bar charts.

Projections & Automated Target Setting: Project month-end outputs based on current run rates and automatically compute upcoming monthly targets using adjustable growth multipliers.

Supervisor Roll-Up View: Roll up district outputs by assigned supervisor with automated color-coded performance indicators and detailed drill-down expanders.

Secure Admin Access Control: Role-based security model. Data uploads, live spreadsheet edits, month/year sheet management, and supervisor re-mappings are strictly locked behind an administrator authentication wall (adminpass123).

Persistent Excel Backend: All edits, data updates, and mappings are saved directly back to the core workbook (dpa_data.xlsx).

🛠️ Project Structure
Plaintext
├── app.py                # Main Streamlit application script
├── dpa_data.xlsx         # Core Excel workbook containing performance data & supervisor mappings
├── appLogoIcon.png       # Application header logo icon
└── requirements.txt      # Required Python packages for deployment
📦 Installation & Local Setup
Clone or download this repository to your local machine.

Install the required dependencies using pip:

Bash
pip install -r requirements.txt
Run the Streamlit application:

Bash
streamlit run app.py
🔒 Administrator Authentication
By default, data modifications and administrative controls are locked.

To unlock data editing and management tools, enter the password adminpass123 in the Admin Authentication panel located in the sidebar.

🌐 Online Deployment
This application can be deployed for free on Streamlit Community Cloud:

Push all project files (app.py, dpa_data.xlsx, appLogoIcon.png, and requirements.txt) to a GitHub repository.

Go to Streamlit Community Cloud and log in with GitHub.

Click New app, choose your repository, branch (main), and set the main file path to app.py.

Click Deploy!
