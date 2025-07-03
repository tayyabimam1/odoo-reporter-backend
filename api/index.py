from flask import Flask, jsonify
from flask_cors import CORS
import sys
import os

# Add parent directory to path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

app = Flask(__name__)
CORS(app)

@app.route('/')
def health_check():
    return jsonify({"status": "Backend is running", "message": "Odoo Reporter API"})

# Vercel handler
def handler(event, context):
    return app(event, context)
