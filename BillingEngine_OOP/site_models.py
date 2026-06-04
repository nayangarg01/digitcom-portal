import pandas as pd
import os

class Site:
    """
    Base class representing a single Digitcom Site.
    Holds all data from Master Tracker and Min Dump mapped to standardized attributes.
    """
    def __init__(self, site_id, pmp_id=None):
        self.site_id = str(site_id).strip()
        self.pmp_id = str(pmp_id).strip() if pmp_id else None
        
        # 1. Identity & Location
        self.sector_id = "N/A"
        self.hop_id = "N/A"
        self.latitude = None
        self.longitude = None
        
        # 2. Logistics & Infrastructure
        self.tower_type = "N/A"
        self.jc = "N/A"
        self.wh = "N/A"
        self.vehicle_no = "N/A"
        self.km_actual = 0.0
        self.km_wo = 0.0
        self.km_threshold = 0.0
        
        # 3. Financial & Billing
        self.wo = "N/A"
        self.dc_no = "N/A"
        self.performa_no = "N/A"
        self.wbs_id = "N/A"
        self.po_no = "N/A"
        
        # 4. Progress & Documentation
        self.activity_type = "N/A"
        self.min_no = "N/A"
        self.min_date = None
        self.completion_date = None
        self.remarks = ""
        
        # 5. Service & Material Quantities
        self.items = {} # Generic mapping for material codes
        self.dispatches = [] # List of MIN Dump material deliveries
        
        # Specific billable attributes (for quick access)
        self.no_of_sectors = 0.0
        self.clubbing = "N/A"
        self.charge_radio = 0.0
        self.qty_a6 = 0.0
        self.charge_atp = 0.0
        self.qty_cpri = 0.0
        self.qty_power = 0.0
        self.charge_sealant = 0.0
        self.qty_extra_visit = 0
        self.charge_gbm = 0.0
        
    def add_item(self, sap_code, quantity):
        self.items[str(sap_code).strip()] = float(quantity)

    def add_dispatch(self, sap_code, description, quantity, min_number, date, remarks="", pmp_id=None, activity="A6"):
        self.dispatches.append({
            "sap_code": str(sap_code).strip(),
            "description": str(description).strip(),
            "quantity": float(quantity) if quantity is not None else 0.0,
            "min_number": str(min_number).strip() if min_number else "N/A",
            "date": date,
            "remarks": str(remarks).strip() if remarks else "",
            "pmp_id": str(pmp_id).strip() if pmp_id else None,
            "activity": str(activity).strip()
        })

    def get_consumed_quantity(self, sap_code):
        return float(self.items.get(str(sap_code).strip(), 0.0))

    def get_dispatched_quantity(self, sap_code):
        sap_code_str = str(sap_code).strip()
        return sum(d["quantity"] for d in self.dispatches if d["sap_code"] == sap_code_str)

    def get_material_variance(self, sap_code):
        return self.get_dispatched_quantity(sap_code) - self.get_consumed_quantity(sap_code)

    def __repr__(self):
        return f"<Site {self.site_id} | WO: {self.wo} | Activity: {self.activity_type} | Dispatches: {len(self.dispatches)}>"

class A6Site(Site):
    """Specific logic for A6 billing sites."""
    def __init__(self, site_id, pmp_id=None):
        super().__init__(site_id, pmp_id)
        self.activity_type = "A6"
        
    def calculate_km_billing(self):
        # A6 logic: use km_actual but caps or thresholds might apply
        return min(self.km_actual, self.km_wo)

class A6B6Site(Site):
    """Specific logic for A6+B6 billing sites."""
    def __init__(self, site_id, pmp_id=None):
        super().__init__(site_id, pmp_id)
        self.activity_type = "A6+B6"
        
    def calculate_km_billing(self):
        # A6+B6 specific logic
        return self.km_threshold
