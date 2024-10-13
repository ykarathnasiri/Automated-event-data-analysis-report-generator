from flask import Flask, render_template, send_file, jsonify
from report_generator import generate_report
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_report', methods=['POST'])
def create_report():
    try:
        pdf_path = generate_report()
        return jsonify({"status": "success", "message": "Report generated successfully"})
    except Exception as e:
        app.logger.error(f"Error generating report: {str(e)}")
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/download_report')
def download_report():
    pdf_path = 'Report/event_data_analysis_report.pdf'
    try:
        return send_file(pdf_path, as_attachment=True)
    except FileNotFoundError:
        return 'Report file not found', 404

if __name__ == '__main__':
    app.run(debug=True)