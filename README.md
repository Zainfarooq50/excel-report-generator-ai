# Excel Report Generator with AI Insights

This Python project generates detailed Excel reports from sales data, complete with interactive charts and AI-powered insights using OpenAI's GPT model. Ideal for businesses looking to automate sales reporting and add professional, AI-driven analysis.

---

## ⚡ Features
✅ Loads and cleans sales data from CSV files  
✅ Automatically classifies high and low performers based on sales  
✅ Generates an Excel report with:  
- Full sales dataset  
- City-wise sales summary  
- Performance breakdown  
✅ Interactive bar and pie charts inside the report  
✅ AI-generated business insights using GPT  
✅ Saves a ready-to-share `.xlsx` report in seconds  

---

## 🛠️ Requirements
- Python 3.x  
- `pandas`, `numpy`, `xlsxwriter`, `openpyxl`, `requests`  
- `GPT_utils_excel.py` with AI integration (requires OpenAI API key)  

---

## 🚀 How to Use
1. Ensure your OpenAI API key is configured inside `GPT_utils_excel.py`  
2. Prepare your sales CSV file with columns: `Name`, `City`, `Sales`  
3. Run the script:  
   ```bash
   python Excel_Report_Generator_AI.py
Enter your CSV file name when prompted

The AI-enhanced Excel report will be saved in the outputs folder

📁 Example Output Structure
Copy
Edit
outputs/
├── AI_Sales_Report_salesdata.xlsx
The report includes:

All_Sales_Data — Full cleaned dataset

City_Summary — Total sales by city with bar chart

Performance_Summary — High/Low performer breakdown with pie chart

AI_Insights — Text summary and GPT-generated business insights

🤖 AI Integration
AI Model: GPT-3.5-Turbo

Provides professional, human-friendly insights based on summary data

Fully automated AI analysis embedded within the Excel report

💼 Ideal Use Cases
✔ Automated sales reporting
✔ Quick performance analysis for businesses
✔ AI-enhanced Excel reports for management
✔ Visual dashboards with actionable insights