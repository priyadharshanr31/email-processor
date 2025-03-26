import pandas as pd
import os
from config import TEMPLATE_FILE, OUTPUT_FOLDER

def process_latest_file(file_path, start_row=10, preview_only=False):
    """Processes the latest statement file and processes it into the template.
    
    If preview_only=True, it returns only the extracted rows for preview.
    """
    try:
        print(f"[INFO] Using latest downloaded file: {file_path}")
        
        # Read the statement file
        statement_df = pd.read_excel(file_path)
        print("[INFO] Reading statement file...")

        # Filter rows based on the 'Status' column
        filter_values = ["Y", "YES", "Yes", "yes", "y", "Done", "done"]
        extracted_df = statement_df[statement_df["Status"].astype(str).str.strip().isin(filter_values)]
        print("[INFO] Filtering data where Status = 'Y', 'YES', 'Done'...")

        if preview_only:
            return None, extracted_df  # Return only extracted rows for preview

        # Read the template file
        template_df = pd.read_excel(TEMPLATE_FILE, sheet_name=None)  # Read all sheets
        sheet_name = list(template_df.keys())[0]  # Assuming first sheet is the target
        template_df = template_df[sheet_name]
        print("[INFO] Reading template file...")

        # Ensure the start_row is within range
        if start_row > len(template_df):
            print(f"⚠️ [WARNING] Start row {start_row} is beyond template file range. Appending at the end.")
            start_row = len(template_df)

        # Split the template into two parts
        upper_part = template_df.iloc[:start_row]  # Data above the insertion point
        lower_part = template_df.iloc[start_row:]  # Data below the insertion point

        # Merge extracted data into the template
        processed_df = pd.concat([upper_part, extracted_df, lower_part], ignore_index=True)
        print("[INFO] Processing files...")

        # Save the processed file
        output_file = os.path.join(OUTPUT_FOLDER, f"Processed_Statement.xlsx")
        processed_df.to_excel(output_file, index=False)
        print(f"[SUCCESS] Processed file created successfully: {output_file}")

        return output_file, extracted_df  # Returning processed file and extracted data preview

    except Exception as e:
        print(f"❌ [ERROR] {e}")
        return None, None
