![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
# 📊 Ally Accessibility Report Generator

This project provides a Python script (`ally_report.py`) that pulls accessibility data from the Blackboard Ally API, compares two academic terms, and generates school-level accessibility insights.

## 🚀 Features
- Connects to the Ally API using token authentication
- Compares two academic terms
- Aggregates results by School
- Calculates accessibility scores and growth

## 📁 Project Structure
```
.
├── ally_report.py
├── dept_mapping.csv
├── .env
└── README.md
```

## 🔑 Prerequisites
- Python 3.8+
- pip

Install dependencies:
```
pip install pandas requests python-dotenv
```

## ⚙️ Setup

Clone the repository:
```
git clone https://github.com/nmastroianni/accessibility-reports
cd accessibility-reports
```

Create a `.env` file:
```
ALLY_CLIENT_ID=your_client_id
ALLY_API_TOKEN=your_api_token
```

## 🗂️ Required File

Create `dept_mapping.csv` with columns:
- departmentId
- departmentName
- School

Example:

164,ACCT,Business & Social Sciences

## ▶️ Usage
Run:
```
python ally_report.py
```

Enter terms when prompted.

## 📊 Output
A timestamped folder is created with reports comparing both terms.

## ⚠️ Notes
- Courses with zero files are excluded
- Departments ending in `-DL` are treated as distance learning

## 📄 License
This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.
