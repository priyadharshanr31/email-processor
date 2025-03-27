import pandas as pd
import asyncio
from openpyxl import load_workbook
from config import TEMPLATE_FILE, OUTPUT_FOLDER
import re


async def process_latest_file(file_path, start_row=10):
    """Asynchronously process the latest statement file and update the template with extracted data."""
   
    try:
        print(f"[INFO] Processing file: {file_path}")

        # Read statement file asynchronously
        df = await asyncio.to_thread(pd.read_excel, file_path, engine="openpyxl")

       

        if "Status" not in df.columns:
            print("❌ [ERROR] 'Status' column not found in the file.")
            return None, None
        
        df_filtered = df[df["Status"].astype(str).str.upper().isin(["Y", "YES", "DONE"])].copy()


        if df_filtered.empty:
            print("[INFO] No matching data found, skipping processing.")
            return None, None



        # Extract currency values from the "Description" column
        if "Description" in df_filtered.columns:
            df_filtered["Extracted_Desc"] = df_filtered["Description"].apply(
                lambda x: re.search(r"([A-Z]{3} \d+)", str(x)).group(1) if pd.notna(x) and re.search(r"([A-Z]{3} \d+)", str(x)) else None
            )
            # Drop the "Description" column for final output
            df_filtered_final = df_filtered.drop(columns=["Description"], errors="ignore")  # ✅ Final output without "Description"

 
        else:
            df_filtered_final = df_filtered.copy()  # No change if "Description" is missing

        

        wb = await asyncio.to_thread(load_workbook, TEMPLATE_FILE, keep_vba=True)
        ws = wb["CS"]

        

        existing_data = list(ws.values)

        

        header_row = existing_data[0] if existing_data else []
        extracted_desc_index = None


        if "Extracted_Desc" in header_row:
            extracted_desc_index = header_row.index("Extracted_Desc")
        else:
            
            header_row = list(header_row) + ["Extracted_Desc"]
            extracted_desc_index = len(header_row) - 1

        
        for col in header_row:
            if col not in df_filtered_final.columns:
                df_filtered_final[col] = None  
        
        # Reorder DataFrame to match template headers
        df_filtered_final = df_filtered_final[header_row]

        # Split the template data at the start row
        before_insert = existing_data[:start_row - 1]  
        after_insert = existing_data[start_row - 1:]  


        # Convert DataFrame to list of lists
        extracted_data_list = [list(df_filtered.columns)] + df_filtered.values.tolist()  # ✅ Preview 1 with "Description"
        final_output_list = df_filtered_final.values.tolist()  # ✅ Final output without "Description"


        # Clear the sheet and rewrite data with inserted extracted data
        ws.delete_rows(1, ws.max_row)  # Clear existing data
        ws.append(header_row)  # Ensure headers are written first


        for row in before_insert[1:]:  # Skip first row since we wrote headers already
            ws.append(row)


        # Append extracted data
        for row in final_output_list:
            ws.append(row)

        for row in after_insert:
            ws.append(row)


        # Save the processed file asynchronously
        output_file = f"{OUTPUT_FOLDER}/Processed_Statement.xlsm"
        await asyncio.to_thread(wb.save, output_file)


        print(f"[✅ SUCCESS] Processed file saved: {output_file}")
        return output_file, extracted_data_list  


    except Exception as e:
        print(f"❌ [ERROR] {e}")
        return None, None
