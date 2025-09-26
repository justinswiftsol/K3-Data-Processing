import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import subprocess
import os
import pandas as pd


def automate_k3p_thickness_workflow():
    """
    Automate the K3P thickness data workflow:
    1. Extract data from Google Sheets (WDXRF and Absorbance page)
    2. Create JMP table with the data
    3. Save as "WDXRF and absorption.jmp" in the specified folder
    """
    
    print("📏 Starting K3P Thickness Workflow Automation...")
    
    # Step 1: Extract data from Google Sheets
    print("\n📊 Extracting data from Google Sheets...")
    sheet_data = extract_wdxrf_absorbance_data()
    
    if sheet_data is None:
        print("❌ Failed to extract data from Google Sheets")
        return
    
    print(f"✅ Extracted {len(sheet_data)} rows of data from Google Sheets")
    
    # Step 2: Create JMP script to import the data
    print("\n📝 Generating JMP script for thickness data...")
    jmp_script = generate_thickness_jmp_script(sheet_data)
    
    # Step 3: Save JMP script
    script_filename = f"automate_k3p_thickness_{datetime.now().strftime('%Y%m%d_%H%M')}.jsl"
    script_path = os.path.join(os.getcwd(), script_filename)
    
    with open(script_path, 'w') as f:
        f.write(jmp_script)
    
    print(f"✅ JMP script saved: {script_filename}")
    
    # Step 4: Execute JMP script
    print("\n🚀 Executing JMP script...")
    execute_jsl_script(script_path)
    
    print("\n🎉 K3P Thickness Workflow Step 1 Complete!")
    print("\n📋 What was accomplished:")
    print("1. ✅ Extracted data from Google Sheets (WDXRF and Absorbance page)")
    print("2. ✅ Created JMP table with column names")
    print("3. ✅ Saved as 'WDXRF and absorption.jmp' in the specified folder")
    print("\n📋 Next Steps (manual):")
    print("1. Review the created JMP table")
    print("2. Continue with additional workflow steps as needed")


def extract_wdxrf_absorbance_data():
    """Extract data from the WDXRF and Absorbance Google Sheet"""
    try:
        # Set up Google Sheets access
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        # Get the directory where this script is located
        script_dir = os.path.dirname(os.path.abspath(__file__))
        service_account_path = os.path.join(script_dir, "service_account.json")
        creds = ServiceAccountCredentials.from_json_keyfile_name(service_account_path, scope)
        client = gspread.authorize(creds)

        # Open the Witness Sample Tracker sheet
        print("📊 Opening Witness Sample Tracker Google Sheet...")
        sheet = client.open_by_url(
            "https://docs.google.com/spreadsheets/d/14SGfNV-zyRL8lLgCGM9DKC0J8I0yz2xK6YIifx8hVVs")
        
        # List available worksheets for debugging
        print("📋 Available worksheets:")
        worksheets = sheet.worksheets()
        for i, ws in enumerate(worksheets):
            print(f"   {i+1}. '{ws.title}'")
        
        # Try to access the "WDXRF and Absorbance" worksheet
        print("📋 Accessing WDXRF and Absorbance worksheet...")
        try:
            worksheet = sheet.worksheet("WDXRF and Absorbance")
        except gspread.exceptions.WorksheetNotFound:
            print("❌ 'WDXRF and Absorbance' worksheet not found")
            # Try alternative names
            alternative_names = ["WDXRF and absorbance", "WDXRF", "Absorbance", "WDXRF and Absorption"]
            worksheet = None
            for alt_name in alternative_names:
                try:
                    print(f"🔍 Trying alternative name: '{alt_name}'")
                    worksheet = sheet.worksheet(alt_name)
                    print(f"✅ Found worksheet: '{alt_name}'")
                    break
                except gspread.exceptions.WorksheetNotFound:
                    continue
            
            if worksheet is None:
                print("❌ Could not find the target worksheet with any alternative names")
                return None
        
        # Get all values from the sheet
        all_values = worksheet.get_all_values()
        
        if not all_values:
            print("❌ No data found in the worksheet")
            return None
        
        print(f"📊 Found {len(all_values)} rows in the worksheet")
        
        # Find the last row with data (check column A for non-empty values)
        last_row_with_data = 0
        for i, row in enumerate(all_values):
            if len(row) > 0 and row[0].strip():  # Column A has data
                last_row_with_data = i
        
        print(f"📊 Last row with data: {last_row_with_data + 1}")
        
        # Extract data from A1 to column Q (index 16) for all rows with data
        extracted_data = []
        for i in range(last_row_with_data + 1):
            row = all_values[i]
            # Ensure we get up to column Q (17 columns: A=0 to Q=16)
            row_data = row[:17] if len(row) >= 17 else row + [''] * (17 - len(row))
            extracted_data.append(row_data)
        
        print(f"✅ Extracted data: {len(extracted_data)} rows × {len(extracted_data[0]) if extracted_data else 0} columns")
        
        return extracted_data
        
    except gspread.exceptions.SpreadsheetNotFound:
        print("❌ Spreadsheet not found - please check:")
        print("   1. The spreadsheet URL is correct")
        print("   2. The service account has access to this spreadsheet")
        print("   3. The spreadsheet is shared with the service account email")
        return None
    except gspread.exceptions.APIError as e:
        print(f"❌ Google Sheets API Error: {e}")
        print("   This might be a permissions or authentication issue")
        return None
    except FileNotFoundError:
        print("❌ service_account.json file not found")
        print("   Please ensure the service account credentials file is in the current directory")
        return None
    except Exception as e:
        print(f"❌ Unexpected error extracting data from Google Sheets: {e}")
        print(f"   Error type: {type(e).__name__}")
        import traceback
        print("   Full traceback:")
        traceback.print_exc()
        return None


def generate_thickness_jmp_script(sheet_data):
    """Generate JSL script to create JMP table with the extracted data"""
    
    if not sheet_data or len(sheet_data) < 2:
        print("❌ Insufficient data to create JMP script")
        return ""
    
    # Get column headers (first row)
    headers = sheet_data[0]
    data_rows = sheet_data[1:]  # All rows except header
    
    print(f"📊 Column headers: {headers}")
    print(f"📊 Data rows: {len(data_rows)}")
    
    # Create JSL script to build the data table
    jsl_script = f'''// Auto-generated JSL script for K3P Thickness workflow
// Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
// Source: WDXRF and Absorbance Google Sheet
// Rows: {len(data_rows)} data rows + header

Print("📏 Starting K3P Thickness Data Import...");

// Create new data table
dt = New Table( "WDXRF and absorption",
    Add Rows( {len(data_rows)} ),
'''

    # Add column definitions
    for i, header in enumerate(headers):
        if header.strip():  # Only add columns with non-empty headers
            # Determine if column should be numeric or character based on header name
            if any(keyword in header.lower() for keyword in ['nm', 'absorbance', 'speed', 'percentage', 'offset', 'thickness']):
                col_type = 'Numeric, "Continuous", Format( "Best", 12 )'
            else:
                col_type = 'Character, "Nominal"'
            
            jsl_script += f'    New Column( "{header}", {col_type} ),\n'
    
    jsl_script += ');\n\n'
    
    # Add data to the table
    jsl_script += '// Populate data\n'
    for col_idx, header in enumerate(headers):
        if header.strip():  # Only process columns with headers
            jsl_script += f'// Column {col_idx + 1}: {header}\n'
            jsl_script += f'dt:Name("{header}") << Set Values({{\n'
            
            # Add all values for this column
            values_list = []
            for row_idx, row in enumerate(data_rows):
                value = row[col_idx] if col_idx < len(row) else ""
                
                # Handle numeric vs text values
                if any(keyword in header.lower() for keyword in ['nm', 'absorbance', 'speed', 'percentage', 'offset', 'thickness']):
                    # Numeric column - convert empty strings to missing values
                    if value.strip() == "":
                        values_list.append('    .')
                    else:
                        try:
                            # Try to convert to number
                            float(value)
                            values_list.append(f'    {value}')
                        except:
                            values_list.append('    .')  # Missing value if can't convert
                else:
                    # Character column - wrap in quotes
                    escaped_value = str(value).replace('"', '\\"')  # Escape quotes
                    values_list.append(f'    "{escaped_value}"')
            
            # Join values with commas, but no trailing comma
            jsl_script += ',\n'.join(values_list) + '\n'
            jsl_script += '});\n\n'
    
    # Save the file
    jsl_script += f'''// Save the data table
save_path = "$HOME/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/R&D - Passivation/Passivation Projects 2024+/RD.E5: Process window specs and drift for inline ETL1   vendor demo/2025Q3 Data/WDXRF and absorption.jmp";

Try(
    dt << Save( save_path );
    Print("✅ Successfully saved: WDXRF and absorption.jmp");
    Print("📁 Location: " || save_path);
,
    Print("❌ Error saving file - please check the folder path exists");
    Print("📁 Attempted path: " || save_path);
);

Print("🎉 K3P Thickness data import complete!");
Print("📊 Created table with {len(data_rows)} rows and {len([h for h in headers if h.strip()])} columns");
'''

    return jsl_script


def execute_jsl_script(script_path):
    """Execute the JSL script in JMP"""
    try:
        print("🚀 Opening JSL script in JMP...")
        subprocess.run([
            "open",
            "-a", "JMP 18",
            script_path
        ], check=True)
        print("✅ JSL script opened in JMP successfully!")
        print("👆 The script should run automatically, or press Cmd+R in JMP to execute")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error opening JSL script in JMP: {e}")
        print(f"📁 You can manually open the file: {os.path.abspath(script_path)}")


if __name__ == "__main__":
    automate_k3p_thickness_workflow()
