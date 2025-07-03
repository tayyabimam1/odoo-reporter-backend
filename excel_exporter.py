import base64
import io
from typing import Dict, List

import openpyxl

def create_excel_report_base64(data: List[Dict]) -> str:
    """
    Creates an Excel report from subscription data and returns it as a base64 encoded string.
    """
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Subscription Report"

    # Define headers based on the desired fields
    headers = [
        "Name", "Status", "Plan", "Start Date", "End Date",
        "Customer Name", "Customer Address", "Customer Phone",
        "Delivery Name", "Delivery Status", "Delivery Date",
        "Product", "Quantity", "Unit Price", "Subtotal",
        "Payment Terms", "Untaxed Amount", "Total Amount"
    ]
    sheet.append(headers)

    # Populate data, creating a new row for each product in a subscription
    for report in data:
        if report.get("products"):
            for product in report["products"]:
                row = [
                    report.get("name"),
                    report.get("status"),
                    report.get("plan"),
                    report.get("start_date"),
                    report.get("end_date"),
                    report.get("customer", {}).get("name"),
                    report.get("customer", {}).get("address"),
                    report.get("customer", {}).get("phone"),
                    report.get("delivery", {}).get("name"),
                    report.get("delivery", {}).get("status"),
                    report.get("delivery", {}).get("date"),
                    product.get("name"),
                    product.get("quantity"),
                    product.get("unit_price"),
                    product.get("subtotal"),
                    report.get("payment_terms"),
                    report.get("untaxed_amount"),
                    report.get("total_amount"),
                ]
                sheet.append(row)
        else:
            # If no products, still add a row for the subscription
            row = [
                report.get("name"),
                report.get("status"),
                report.get("plan"),
                report.get("start_date"),
                report.get("end_date"),
                report.get("customer", {}).get("name"),
                report.get("customer", {}).get("address"),
                report.get("customer", {}).get("phone"),
                report.get("delivery", {}).get("name"),
                report.get("delivery", {}).get("status"),
                report.get("delivery", {}).get("date"),
                "N/A", 0, 0, 0,  # Placeholder for product info
                report.get("payment_terms"),
                report.get("untaxed_amount"),
                report.get("total_amount"),
            ]
            sheet.append(row)

    # Save to an in-memory buffer
    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    # Encode as base64 and return as a string
    base64_encoded = base64.b64encode(buffer.read()).decode('utf-8')
    return base64_encoded
