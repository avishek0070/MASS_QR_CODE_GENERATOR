import pandas as pd
import qrcode
import os
import re

def sanitize_for_path(name):
    name = name.strip()  # Trim whitespace from the start and end
    return re.sub(r'[<>:"/\\|?*]', '_', name)

def extract_srns_and_names(row):
    srns_and_names = [(row['SRN'], row['Name'])]
    for i in range(1, 4):
        srn_col = f'SRN.{i}'
        name_col = f'Name.{i}'
        if pd.notna(row[srn_col]) and pd.notna(row[name_col]):
            srns_and_names.append((row[srn_col], row[name_col]))
    return srns_and_names

def generate_qr_codes_from_excel(file_path):
    df = pd.read_excel(file_path)
    
    for index, row in df.iterrows():
        team_name = sanitize_for_path(row['Team Name'])
        srns_and_names = extract_srns_and_names(row)
        
        team_dir = fr"C:\Users\agarw\OneDrive\Desktop\qrcode_terrathon\{team_name}"
        os.makedirs(team_dir, exist_ok=True)
        
        for srn, name in srns_and_names:
            qr_content = f"{srn.upper()}"
            qr = qrcode.make(qr_content)
            
            sanitized_srn = sanitize_for_path(srn)
            qr_file_path = os.path.join(team_dir, f"{sanitized_srn}.png")
            qr.save(qr_file_path)

# Ensure the path to your Excel file is correct and uncomment to execute the script
generate_qr_codes_from_excel(r"C:\Users\agarw\Downloads\Terrathon 3.0 Registration Form (Responses).xlsx")
