import os
import glob
from parse_work_order import parse_work_order

def main():
    input_dir = 'WorkOrders'
    output_dir = 'Output'

    if not os.path.exists(input_dir):
        print(f"Directory '{input_dir}' does not exist.")
        return

    os.makedirs(output_dir, exist_ok=True)

    pdf_files = glob.glob(os.path.join(input_dir, '**', '*.pdf'), recursive=True)
    if not pdf_files:
        print(f"No PDF files found in '{input_dir}'.")
        return

    for pdf_file in pdf_files:
        base_name = os.path.basename(pdf_file)
        name_without_ext = os.path.splitext(base_name)[0]
        output_excel = os.path.join(output_dir, f"{name_without_ext}_Summary.xlsx")
        
        print(f"--------------------------------------------------")
        parse_work_order(pdf_file, output_excel)

    print(f"--------------------------------------------------")
    print(f"Batch processing complete! All files saved to '{output_dir}'.")

if __name__ == '__main__':
    main()
