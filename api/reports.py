from flask import Flask, jsonify
from flask_cors import CORS
import sys
import os

# Add parent directory to path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from odoo_reporter_local import OdooSubscriptionReporter

app = Flask(__name__)
CORS(app)

@app.route('/')
def get_reports():
    try:
        reporter = OdooSubscriptionReporter()
        reports_data = reporter.generate_structured_reports()
        return jsonify(reports_data)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"An unexpected error occurred: {str(e)}"}), 500

# Vercel handler
def handler(event, context):
    return app(event, context)
