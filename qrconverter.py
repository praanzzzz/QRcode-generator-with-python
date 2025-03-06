


import pandas as pd
import qrcode
import os

def generate_qr_code_vcard(first_name, last_name, position, contact, email, business, address, qr_filename):
    vcard_data = f"""BEGIN:VCARD
VERSION:2.1
FN:{first_name} {last_name}
N:{last_name};{first_name};;;
TITLE:{position}
TEL:{contact}
EMAIL:{email}
ORG:{business}
ADR:;;{address}
END:VCARD"""
    
    qr = qrcode.QRCode(
        version=5,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    
    qr.add_data(vcard_data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill="black", back_color="white")
    qr_img.save(qr_filename, "JPEG")

def process_excel(file_path, output_dir):
    df = pd.read_excel(file_path)
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for index, row in df.iterrows():
        first_name = row["first name"]
        last_name = row["last name"]
        position = row["position"]
        contact = row["contact number"]
        email = row["email"]
        business = row["business name"]
        address = row["address"]
        
        qr_filename = os.path.join(output_dir, f"{first_name}_{last_name}_qr.jpg")
        generate_qr_code_vcard(first_name, last_name, position, contact, email, business, address, qr_filename)
        print(f"QR code saved: {qr_filename}")

if __name__ == "__main__":
    excel_file = "personal_info.xlsx"  # Change this to your actual file path
    output_folder = "qr_codes"  # Folder where QR codes will be saved
    process_excel(excel_file, output_folder)
