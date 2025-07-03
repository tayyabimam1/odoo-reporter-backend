import requests
import json
import logging
import os
from datetime import datetime
from typing import Dict, List, Optional, Union
import argparse
from excel_exporter import create_excel_report_base64

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

class OdooSubscriptionReporter:
    def __init__(self):
        """Initializes the reporter, loading configuration from environment variables."""
        logger.info("Odoo reporter initialized")
        self.url = os.environ.get("ODOO_URL")
        self.db = os.environ.get("ODOO_DB")
        self.uid = int(os.environ.get("ODOO_UID"))
        self.password = os.environ.get("ODOO_PASSWORD")

        if not all([self.url, self.db, self.uid, self.password]):
            msg = "Missing Odoo configuration. Ensure ODOO_URL, ODOO_DB, ODOO_UID, and ODOO_PASSWORD are set."
            logger.error(msg)
            raise ValueError(msg)

    def _make_request(self, model: str, method: str, args: List, kwargs: Dict = None) -> Dict:
        payload = {
            "jsonrpc": "2.0",
            "method": "call",
            "params": {
                "service": "object",
                "method": "execute_kw",
                "args": [self.db, self.uid, self.password, model, method, args, kwargs or {}]
            }
        }
        try:
            response = requests.post(self.url, json=payload, timeout=30)
            response.raise_for_status()
            result = response.json()
            if 'error' in result:
                logger.error(f"Odoo API Error: {result['error']}")
            return result.get('result', [])
        except Exception as e:
            logger.error(f"Request failed: {str(e)}")
            return []

    def get_all_subscriptions(self) -> List[Dict]:
        """Fetch all subscriptions from Odoo"""
        logger.info("Fetching subscriptions...")
        # NOTE: The domain filter `["subscription_state", "!=", False]` was removed because it
        # requires the 'sale_subscription' module in Odoo, which may not be installed
        # (as suggested by the "Object sale.subscription doesn't exist" error).
        # This script will now fetch ALL sale orders. For accurate subscription reporting,
        # please install the "Subscriptions" application in your Odoo database.
        domain = []
        fields = [
            "id", "name", "subscription_state", "plan_id", "date_order", 
            "partner_id", "order_line", "payment_term_id",
            "amount_untaxed", "amount_total"
        ]
        result = self._make_request(
            "sale.order",
            "search_read",
            [domain],
            {"fields": fields}
        )
        
        # If 'subscription_state' is missing, add a placeholder to avoid crashes
        if result and 'subscription_state' not in result[0]:
            logger.warning("'subscription_state' field not found. The 'sale_subscription' module is likely missing in Odoo.")
            for r in result:
                r['subscription_state'] = 'n/a'

        logger.info(f"Found {len(result)} sale orders")
        return result

    def get_customer_details(self, partner_id: int) -> Dict:
        if not partner_id:
            return {}
            
        fields = ["name", "street", "street2", "city", "state_id", "country_id", "phone", "email"]
        # The result of a read is a list, so we take the first element.
        details = self._make_request("res.partner", "read", [[partner_id]], {"fields": fields})
        return details[0] if details else {}

    def get_order_lines(self, line_ids: List[int]) -> List[Dict]:
        if not line_ids:
            return []
            
        fields = ["product_id", "name", "product_uom_qty", "price_unit", "price_subtotal"]
        return self._make_request("sale.order.line", "read", [line_ids], {"fields": fields})

    def get_delivery_orders(self, origin: str) -> List[Dict]:
        if not origin:
            return []
            
        domain = [("origin", "=", origin), ("picking_type_id.code", "=", "outgoing")]
        fields = ["name", "state", "scheduled_date"]
        return self._make_request(
            "stock.picking", 
            "search_read", 
            [domain], 
            {"fields": fields, "order": "scheduled_date desc", "limit": 1}
        )

    def get_many2one_value(self, field_value: Union[bool, list], default: str = "Not Available") -> str:
        """Safely extract value from Many2one field"""
        if isinstance(field_value, list) and len(field_value) > 1:
            return field_value[1]
        return default

    def get_partner_id(self, partner_field: Union[bool, list]) -> int:
        """Safely extract partner ID"""
        if isinstance(partner_field, list) and partner_field:
            return partner_field[0]
        return 0

    @staticmethod
    def format_date(date_str: str) -> str:
        if not date_str:
            return "Not Available"
        try:
            # Handle both date and datetime formats
            if ' ' in date_str:
                dt_obj = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
            else:
                dt_obj = datetime.strptime(date_str, '%Y-%m-%d')
            return dt_obj.strftime('%m/%d/%Y')
        except (ValueError, TypeError):
            return "Invalid Date"

    @staticmethod
    def format_address(customer: Dict) -> str:
        state = customer.get('state_id', [False, ''])[1]
        country = customer.get('country_id', [False, ''])[1]
        address_parts = [
            customer.get('street', ''),
            customer.get('street2', ''),
            customer.get('city', ''),
            f"{state} ({country})" if state and country else ''
        ]
        return ", ".join(filter(None, address_parts))

    @staticmethod
    def map_status(status: str) -> str:
        status_mapping = {
            '4_close': 'Closed',
            '6_churn': 'Churned',
            '3_pending': 'Pending',
            '2_open': 'Active',
            '1_draft': 'Draft'
        }
        return status_mapping.get(status, status.capitalize())

    @staticmethod
    def map_delivery_status(status: str) -> str:
        status_mapping = {
            'draft': 'Draft',
            'waiting': 'Waiting',
            'confirmed': 'Confirmed',
            'assigned': 'Preparation',
            'done': 'Delivered',
            'cancel': 'Cancelled'
        }
        return status_mapping.get(status, status.capitalize())

    def generate_structured_reports(self) -> List[Dict]:
        """Fetches subscriptions and formats them into a list of structured dictionaries."""
        subscriptions = self.get_all_subscriptions()
        if not subscriptions:
            return []
        
        reports_data = []
        for sub in subscriptions:
            try:
                partner_id = self.get_partner_id(sub.get('partner_id', False))
                customer = self.get_customer_details(partner_id)
                origin = sub.get('name', '')
                deliveries = self.get_delivery_orders(origin)
                delivery = deliveries[0] if deliveries else {}
                products = self.get_order_lines(sub.get('order_line', []))

                reports_data.append({
                    "name": sub.get('name', 'N/A'),
                    "status": self.map_status(sub.get('subscription_state', '')),
                    "plan": self.get_many2one_value(sub.get('plan_id'), 'Not Available'),
                    "start_date": self.format_date(sub.get('date_order')),
                    "end_date": "Not Available",
                    "customer": {
                        "name": customer.get('name', 'N/A'),
                        "address": self.format_address(customer),
                        "phone": customer.get('phone', 'N/A'),
                    },
                    "delivery": {
                        "name": delivery.get('name', 'N/A'),
                        "status": self.map_delivery_status(delivery.get('state', 'N/A')),
                        "date": self.format_date(delivery.get('scheduled_date', 'N/A')),
                    },
                    "products": [
                        {
                            "name": p.get('name', 'N/A').split('\n')[0],
                            "quantity": p.get('product_uom_qty', 0),
                            "unit_price": p.get('price_unit', 0),
                            "subtotal": p.get('price_subtotal', 0),
                        }
                        for p in products
                    ],
                    "payment_terms": self.get_many2one_value(sub.get('payment_term_id', [False, 'N/A']), 'N/A'),
                    "untaxed_amount": sub.get('amount_untaxed', 0),
                    "total_amount": sub.get('amount_total', 0),
                })
            except Exception as e:
                logger.error(f"Error processing subscription {sub.get('name', 'N/A')}: {e}")
        
        logger.info(f"Successfully processed {len(reports_data)} subscriptions")
        return reports_data

# Main execution
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate reports from Odoo.")
    parser.add_argument("--format", type=str, default="json", choices=["json", "excel"],
                        help="Output format for the report.")
    args = parser.parse_args()

    try:
        reporter = OdooSubscriptionReporter()
        logger.info("Starting report generation")
        reports_data = reporter.generate_structured_reports()

        if args.format == "excel":
            if not reports_data:
                print(json.dumps({"error": "No data available to generate Excel report."}))
            else:
                base64_excel = create_excel_report_base64(reports_data)
                print(json.dumps({"fileContent": base64_excel}))
        else:  # Default to json
            print(json.dumps(reports_data))

    except ValueError as e:
        print(json.dumps({"error": str(e)}))
    except Exception as e:
        logger.error(f"An unexpected error occurred: {str(e)}")
        print(json.dumps({"error": f"An unexpected error occurred: {str(e)}"}))

    logger.info("Process completed successfully")