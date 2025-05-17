import pandas as pd
import os
import re
from datetime import datetime

def safe_convert_currency(value):
    """Safely convert currency strings to float"""
    if pd.isna(value) or value == '' or str(value).strip() == 'Total Orders (Delivered + Cancelled)':
        return 0.0
    try:
        return float(re.sub(r"[^\d.]", "", str(value)))
    except:
        return 0.0

def safe_convert_int(value):
    """Safely convert strings to integers"""
    if pd.isna(value) or value == '' or str(value).strip() == 'Total Orders (Delivered + Cancelled)':
        return 0
    try:
        return int(re.sub(r"[^\d]", "", str(value)))
    except:
        return 0

def extract_summary_data(file_path):
    """Extracts data from Summary tab with robust error handling"""
    try:
        df = pd.read_excel(file_path, sheet_name="Summary", header=None)
        
        data = {
            "Brand": None,
            "Location": None,
            "City": None,
            "Res-Id": None,
            "GSTIN": None,
            "Payout Period": None,
            "Payout Settlement Date": None,
            "Total Payout": 0.0,
            "Total Orders": 0,
            "Bank UTR": None,
            "File Name": os.path.basename(file_path)
        }
        
        for i in range(len(df)):
            if pd.notna(df.iloc[i, 1]):
                cell = str(df.iloc[i, 1]).strip()
                
                if any(x in cell for x in ["Biryani", "Restaurant"]):
                    data["Brand"] = cell
                elif "Whitefield" in cell:
                    data["Location"] = cell
                elif "Bangalore" in cell:
                    data["City"] = cell
                elif "Rest. ID" in cell:
                    data["Res-Id"] = cell.split("Rest. ID - ")[-1]
                elif "GSTIN" in cell:
                    data["GSTIN"] = cell.split("GSTIN  - ")[-1]
                elif "Payout Period" in cell:
                    data["Payout Period"] = df.iloc[i+1, 1] if i+1 < len(df) else None
                elif "Payout Settlement Date" in cell:
                    data["Payout Settlement Date"] = df.iloc[i+1, 1] if i+1 < len(df) else None
                elif "Total Payout" in cell:
                    val = df.iloc[i+1, 1] if i+1 < len(df) else "0"
                    data["Total Payout"] = safe_convert_currency(val)
                elif "Total Orders" in cell:
                    val = df.iloc[i+1, 1] if i+1 < len(df) else "0"
                    data["Total Orders"] = safe_convert_int(val)
                elif "Bank UTR" in cell:
                    data["Bank UTR"] = df.iloc[i+1, 1] if i+1 < len(df) else None
        
        return pd.DataFrame([data])
    
    except Exception as e:
        print(f"⚠️ Error processing Summary in {os.path.basename(file_path)}: {str(e)}")
        return None

def extract_payout_breakup(file_path):
    """Extracts data from Payout Breakup tab"""
    try:
        df = pd.read_excel(file_path, sheet_name="Payout Breakup", header=None)
        
        # Get summary data for reference
        summary = extract_summary_data(file_path)
        if summary is None or summary.empty:
            return None
            
        # Find table start
        start_row = None
        for i in range(len(df)):
            if pd.notna(df.iloc[i, 2]) and "Particulars" in str(df.iloc[i, 2]):
                start_row = i + 1
                break
        
        if start_row is None:
            return None
            
        # Extract data rows
        data = []
        for i in range(start_row, len(df)):
            if pd.notna(df.iloc[i, 2]):
                try:
                    data.append({
                        "SR.No": df.iloc[i, 0],
                        "Particulars": df.iloc[i, 2],
                        "Delivered Orders": safe_convert_int(df.iloc[i, 3]),
                        "Cancelled Orders": safe_convert_int(df.iloc[i, 4]),
                        "Total": safe_convert_currency(df.iloc[i, 5]),
                        "Brand": summary.iloc[0]["Brand"],
                        "Res-Id": summary.iloc[0]["Res-Id"],
                        "Payout Period": summary.iloc[0]["Payout Period"],
                        "File Name": os.path.basename(file_path)
                    })
                except:
                    continue
        
        return pd.DataFrame(data) if data else None
    
    except Exception as e:
        print(f"⚠️ Error processing Payout Breakup in {os.path.basename(file_path)}: {str(e)}")
        return None

def extract_order_level(file_path):
    """Extracts data from Order Level tab"""
    try:
        df = pd.read_excel(file_path, sheet_name="Order Level")
        
        # Find header row
        header_row = None
        for i in range(len(df)):
            if "Order ID" in str(df.iloc[i, 0]):
                header_row = i
                break
        
        if header_row is None:
            return None
            
        # Process data
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row+1:].reset_index(drop=True)
        
        # Add metadata
        summary = extract_summary_data(file_path)
        if summary is not None and not summary.empty:
            df["Brand"] = summary.iloc[0]["Brand"]
            df["Res-Id"] = summary.iloc[0]["Res-Id"]
            df["Payout Period"] = summary.iloc[0]["Payout Period"]
        df["File Name"] = os.path.basename(file_path)
        
        return df
    
    except Exception as e:
        print(f"⚠️ Error processing Order Level in {os.path.basename(file_path)}: {str(e)}")
        return None

def process_files(folder_path):
    """Process all files in the folder"""
    summary_data = []
    payout_data = []
    order_data = []
    
    processed_files = 0
    total_files = 0
    
    for file in os.listdir(folder_path):
        if file.startswith("invoice_Annexure_") and file.endswith(".xlsx"):
            total_files += 1
            file_path = os.path.join(folder_path, file)
            print(f"Processing: {file}")
            
            # Extract data from all sheets
            summary = extract_summary_data(file_path)
            payout = extract_payout_breakup(file_path)
            orders = extract_order_level(file_path)
            
            if summary is not None and not summary.empty:
                summary_data.append(summary)
                processed_files += 1
            if payout is not None and not payout.empty:
                payout_data.append(payout)
            if orders is not None and not orders.empty:
                order_data.append(orders)
    
    # Combine all data
    final_summary = pd.concat(summary_data, ignore_index=True) if summary_data else None
    final_payout = pd.concat(payout_data, ignore_index=True) if payout_data else None
    final_orders = pd.concat(order_data, ignore_index=True) if order_data else None
    
    # Save to Excel
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"Consolidated_Annexure_Data_{timestamp}.xlsx"
    
    with pd.ExcelWriter(output_file) as writer:
        if final_summary is not None:
            final_summary.to_excel(writer, sheet_name="Summary", index=False)
        if final_payout is not None:
            final_payout.to_excel(writer, sheet_name="Payout Breakup", index=False)
        if final_orders is not None:
            final_orders.to_excel(writer, sheet_name="Order Level", index=False)
    
    print(f"\n✅ Successfully processed {processed_files}/{total_files} files")
    print(f"Results saved to: {output_file}")
    return output_file

if __name__ == "__main__":
    folder_path = r"D:\Downloads\Cleaning of Data & Merging into single excel\Cleaning of Data & Merging into single excel\Payout Summary & Order Level Sales"
    process_files(folder_path)