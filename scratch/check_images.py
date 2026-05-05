import openpyxl

def check_images():
    ref_path = '/Users/nayangarg/Desktop/DigitcomWebsiteRenovation/Old_Codebase_renovated_v6.1/FinaliseBillingFormat/DIGITCOM_ AIRFIBER_DC0105_ JDPR_29-JAN-26_A6 (REJECT) 2.xlsx'
    wb = openpyxl.load_workbook(ref_path)
    ws = wb['Main WCC']
    print(f"Number of images in Main WCC: {len(ws._images)}")
    for i, img in enumerate(ws._images):
        print(f"Image {i}: Anchor {img.anchor}")

if __name__ == "__main__":
    check_images()
