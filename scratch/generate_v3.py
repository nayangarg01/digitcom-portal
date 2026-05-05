import openpyxl
import os
from openpyxl.drawing.image import Image as OpenpyxlImage

def generate_v3_alt_logo(output_path, ref_path, logo_path):
    wb_new = openpyxl.load_workbook(output_path)
    ws = wb_new['Main WCC']
    # Clear images if any
    ws._images = []
    
    img = OpenpyxlImage(logo_path)
    img.width, img.height = 120, 40 # Standard Jio logo size
    ws.add_image(img, 'B2')
    wb_new.save(output_path)
    print(f"- Saved V3 with alternative logo to {output_path}")

if __name__ == "__main__":
    v2_path = "/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/DC0105_Clean_Billing_V2.xlsx"
    v3_path = "/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/DC0105_Clean_Billing_V3_ALT_LOGO.xlsx"
    alt_logo = "/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/logos/reliance-jio.png"
    
    import shutil
    shutil.copy(v2_path, v3_path)
    generate_v3_alt_logo(v3_path, v2_path, alt_logo)
