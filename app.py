from flask import Flask, jsonify
from flask_cors import CORS
import os
from odoo_reporter_local import OdooSubscriptionReporter
from excel_exporter import create_excel_report_base64

app = Flask(__name__)
CORS(app)

@app.route('/')
def health_check():
    return jsonify({"status": "Backend is running", "message": "Odoo Reporter API"})

@app.route('/api/reports')
def get_reports():
    try:
        reporter = OdooSubscriptionReporter()
        reports_data = reporter.generate_structured_reports()
        return jsonify(reports_data)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"An unexpected error occurred: {str(e)}"}), 500

@app.route('/api/reports/excel')
def get_excel_report():
    try:
        reporter = OdooSubscriptionReporter()
        reports_data = reporter.generate_structured_reports()
        
        if not reports_data:
            return jsonify({"error": "No data available to generate Excel report."}), 400
        
        base64_excel = create_excel_report_base64(reports_data)
        return jsonify({"fileContent": base64_excel})
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"An unexpected error occurred: {str(e)}"}), 500
