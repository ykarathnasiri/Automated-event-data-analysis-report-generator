CeylonEvent Analysis Report Generator
This Flask application generates and provides downloadable event analysis reports based on CeylonEvent data.
Features

Generate comprehensive event analysis reports
Download generated reports in PDF format
Simple and intuitive web interface

Installation

Clone this repository:
Copygit clone https://github.com/yourusername/ceylonevent-analysis-report-generator.git
cd ceylonevent-analysis-report-generator

Install the required dependencies:
Copypip install flask pandas matplotlib seaborn python-docx docx2pdf


Usage

Run the Flask application:
Copypython app.py

Open a web browser and navigate to http://localhost:5000
Click the "Generate Report" button to create a new report
Once the report is generated, click the "Download Report" button to download the PDF

Project Structure

app.py: Main Flask application file
report_generator.py: Contains the generate_report() function for creating the analysis report
templates/index.html: HTML template for the web interface
Data/event.csv: Input data file (included dummy datset in this repository)
Report/: Directory where generated reports are saved

Dependencies

Flask
pandas
matplotlib
seaborn
python-docx
docx2pdf

Note
This application uses a dummy dataset for demonstration purposes. The analysis is based on cleaned records from August and September 2024, comprising approximately 4,045 entries.
Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.