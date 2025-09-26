import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import subprocess
import os
import glob
import re


def automate_temp_workflow():
    """
    Automate the complete temperature data workflow:
    1. Get reference lines from Google Sheets
    2. Get date inputs from user (YYMMDD format)
    3. Find the appropriate CSV file and JMP file
    4. Generate comprehensive JSL script for temperature data
    5. Execute the JSL script in JMP
    """
    
    print("🌡️ Starting Temperature Workflow Automation...")
    
    # Get date inputs from user (YYMMDD format)
    previous_date = input("Enter the previous experiment date (YYMMDD format, e.g., 250821): ")
    current_date = input("Enter the current experiment date (YYMMDD format, e.g., 250827): ")
    current_exp = input("Enter the current experiment number (#### format, e.g., 0407): ")
    
    print(f"\n📋 Processing temperature data: {previous_date} → {current_date}")
    print(f"🧪 Using experiment number: K3P{current_exp}")
    
    # Step 1: Get reference lines from Google Sheets
    print("\n📊 Fetching reference lines from Google Sheets...")
    sheet_data = get_reference_lines_from_sheets()

    if not sheet_data:
        print("❌ Failed to get reference lines from Google Sheets")
        return

    print(f"✅ Retrieved reference lines and time ranges from Google Sheets")

    # Step 2: Verify file paths
    print("\n📁 Verifying temperature data files...")
    previous_jmp_file, current_csv_file = verify_temp_files(previous_date, current_date)

    if not previous_jmp_file or not current_csv_file:
        print("❌ Required files not found")
        return

    print(f"✅ Found JMP file: {os.path.basename(previous_jmp_file)}")
    print(f"✅ Found CSV file: {os.path.basename(current_csv_file)}")

    # Step 3: Generate JSL script
    print("\n📝 Generating comprehensive temperature JSL script...")
    jsl_script = generate_temp_jsl(previous_date, current_date, current_exp, previous_jmp_file, current_csv_file, sheet_data)

    # Step 4: Save JSL script
    script_filename = f"automate_temp_{previous_date}_to_{current_date}.jsl"
    script_path = os.path.join(os.getcwd(), script_filename)

    with open(script_path, 'w') as f:
        f.write(jsl_script)

    print(f"✅ JSL script saved: {script_filename}")

    # Step 5: Execute JSL script in JMP
    print("\n🚀 Executing JSL script in JMP...")
    execute_jsl_script(script_path)

    print("\n🎉 Temperature Workflow Automation Complete!")
    print("\n📋 Next Steps:")
    print("1. Manually concatenate the CSV data with the temperature data table")
    print("2. Save the updated file with the new date")
    print("3. The automation has completed all JMP table script updates")


def get_reference_lines_from_sheets():
    """Get pre-formatted JSL strings from Google Sheets using both worksheets"""
    try:
        # Set up Google Sheets access
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
        client = gspread.authorize(creds)

        # Open the reference times sheet
        sheet = client.open_by_url(
            "https://docs.google.com/spreadsheets/d/1g9Vb8OJJ67kEaGhXoxyBRHD-3oxouyLQr66b5V6JnVE")
        
        # Step 1: Get the target date from "K3 PC3 Runs" worksheet
        print("📊 Getting target date from K3 PC3 Runs...")
        runs_worksheet = sheet.worksheet("K3 PC3 Runs")
        runs_data = runs_worksheet.get_all_records()
        
        # Find the latest date
        latest_date = None
        
        for row in runs_data:
            if 'Date' in row and row['Date']:
                try:
                    date_str = str(row['Date'])
                    if '/' in date_str:
                        if date_str.count('/') == 2:
                            if len(date_str.split('/')[0]) == 4:  # YYYY/MM/DD
                                row_date = datetime.strptime(date_str, "%Y/%m/%d").date()
                            else:  # MM/DD/YYYY or M/D/YYYY
                                row_date = datetime.strptime(date_str, "%m/%d/%Y").date()
                        else:
                            continue
                    else:
                        continue
                    
                    if latest_date is None or row_date > latest_date:
                        latest_date = row_date
                        
                except:
                    continue
    
        print(f"📅 Target date from K3 PC3 Runs: {latest_date}")
        
        # Step 2: Get the JSL strings from "K3 Timestamps to JMP Time" worksheet  
        print("📊 Getting JSL strings from K3 Timestamps to JMP Time...")
        jsl_worksheet = sheet.worksheet("K3 Timestamps to JMP Time")
        all_values = jsl_worksheet.get_all_values()
        
        if not all_values or len(all_values) < 2:
            print("❌ No data found in the K3 Timestamps to JMP Time worksheet")
            return {}
        
        print(f"🔍 Found {len(all_values)} rows of JSL data (header + data rows)")
        
        # Get data from ALL rows (each row contains one reference line)
        result = {
            'date': latest_date,
            'time_range_string': '',     # Column P (index 15) from first data row
            'ref_lines_h': [],           # Column H (index 7) from all data rows
            'ref_lines_i': []            # Column I (index 8) from all data rows
        }
        
        # Get time range string from Column P (index 15) of first data row
        if len(all_values) > 1 and len(all_values[1]) > 15:
            result['time_range_string'] = str(all_values[1][15])  # Column P from first row
            print(f"📊 Time range string (P): {result['time_range_string']}")
        else:
            print("❌ Column P (index 15) not available!")
        
        # Collect ALL reference lines from Column H (index 7) from all data rows
        ref_lines_h = []
        for i, row in enumerate(all_values[1:], 1):  # Skip header row
            if len(row) > 7 and row[7].strip():  # Column H
                ref_line = str(row[7]).strip()
                if ref_line and "Add Ref Line" in ref_line:
                    ref_lines_h.append(ref_line)
        
        result['ref_lines_h'] = ref_lines_h
        print(f"📏 Reference lines H: Found {len(ref_lines_h)} lines")
        
        # Collect ALL reference lines from Column I (index 8) from all data rows  
        ref_lines_i = []
        for i, row in enumerate(all_values[1:], 1):  # Skip header row
            if len(row) > 8 and row[8].strip():  # Column I
                ref_line = str(row[8]).strip()
                if ref_line and "Add Ref Line" in ref_line:
                    ref_lines_i.append(ref_line)
        
        result['ref_lines_i'] = ref_lines_i
        print(f"📏 Reference lines I: Found {len(ref_lines_i)} lines")

        return result
        
    except Exception as e:
        print(f"❌ Error accessing Google Sheets: {e}")
        return {}


def verify_temp_files(previous_date, current_date):
    """Verify that the required temperature files exist"""
    base_path = "/Users/justinparayno/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/Shared with Swift computers (INSECURE)/K3/K3 PC3 source temperature"
    
    # Previous JMP file path
    previous_jmp_file = os.path.join(base_path, previous_date, f"{previous_date} SL4824-LR.jmp")
    
    # Current CSV file path
    current_csv_file = os.path.join(base_path, current_date, "SL4824-Combined-LR.csv")
    
    # Check if files exist
    if not os.path.exists(previous_jmp_file):
        print(f"❌ Previous JMP file not found: {previous_jmp_file}")
        return None, None
        
    if not os.path.exists(current_csv_file):
        print(f"❌ Current CSV file not found: {current_csv_file}")
        return None, None
        
    return previous_jmp_file, current_csv_file


def generate_temp_jsl(previous_date, current_date, current_exp, previous_jmp_file, current_csv_file, sheet_data):
    """Generate the complete JSL script for temperature data workflow"""

    # Extract pre-formatted JSL strings from Google Sheets
    time_range_string = sheet_data.get('time_range_string', '')
    ref_lines_h = sheet_data.get('ref_lines_h', [])
    ref_lines_i = sheet_data.get('ref_lines_i', [])
    experiment_number = f'K3P{current_exp}'  # Use user-provided experiment number
    
    # Parse the time range string to extract Min and Max values
    min_time = 3839122800  # Default fallback
    max_time = 3839180400  # Default fallback
    
    if time_range_string:
        # Extract Min and Max values from the time_range_string
        min_match = re.search(r'Min\(\s*(\d+)\s*\)', time_range_string)
        max_match = re.search(r'Max\(\s*(\d+)\s*\)', time_range_string)
        
        if min_match:
            min_time = int(min_match.group(1))
        if max_match:
            max_time = int(max_match.group(1))
    
    print(f"🕐 Using time range: Min({min_time}), Max({max_time}) - covers full day")
    print(f"🧪 Using experiment number: {experiment_number}")
    
    # Combine reference lines from columns H and I  
    all_ref_lines = ref_lines_h + ref_lines_i
    
    # Remove trailing commas from each line and then join properly
    clean_ref_lines = []
    for line in all_ref_lines:
        clean_line = line.rstrip(',').strip()
        if clean_line:
            clean_ref_lines.append(clean_line)
    
    combined_ref_lines = ",\n\t\t\t\t".join(clean_ref_lines) + "," if clean_ref_lines else ""

    # Generate the comprehensive JSL script for temperature data
    jsl_script = f'''// Auto-generated JSL script for Temperature workflow automation
// Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
// Date transition: {previous_date} → {current_date}
// CSV file: {os.path.basename(current_csv_file)}
// Using pre-formatted JSL strings from Google Sheets

Print("🌡️ Starting Temperature Workflow Automation...");
Print("📋 Processing: {previous_date} → {current_date}");

// Step 1: Open previous temperature data file
Print("📂 Opening previous temperature data file...");

Try(
    // Open previous experiment temperature data
    temp_dt = Open("$HOME/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/Shared with Swift computers (INSECURE)/K3/K3 PC3 source temperature/{previous_date}/{previous_date} SL4824-LR.jmp");
    Print("✅ Opened {previous_date} SL4824-LR.jmp");
,
    Print("❌ Error opening {previous_date} SL4824-LR.jmp");
    Throw("Failed to open previous temperature data file");
);

// Step 2: Clean up data (delete all rows except first row - keeps only header/first data point)
Print("🧹 Cleaning up previous data (keeping only first row)...");

Try(
    // Select and delete rows with the specific timestamp (this deletes all rows except the first row)
    temp_dt << Select Where( :Date and time == 3822799504 );
    temp_dt << Invert Row Selection;
    temp_dt << Delete Rows;
    Print("✅ Cleaned up previous data - kept only first row");
,
    Print("⚠️ Could not clean up data - continuing anyway");
);

// Step 3: Import new CSV data
Print("📥 Importing new temperature CSV data...");

Try(
    csv_dt = Open("$HOME/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/Shared with Swift computers (INSECURE)/K3/K3 PC3 source temperature/{current_date}/SL4824-Combined-LR.csv",
        columns(
            New Column( "Date/Time", Character, "Nominal" ),
            New Column( "SL4824-LR-1.PV", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-1.SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-1.CVH", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-1.CVC", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-1.RS_SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-2.PV", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-2.SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-2.CVH", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-2.CVC", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-2.RS_SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-3.PV", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-3.SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-3.CVH", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-3.CVC", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-3.RS_SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-4.PV", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-4.SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-4.CVH", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-4.CVC", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-4.RS_SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-5.PV", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-5.SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-5.CVH", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-5.CVC", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-5.RS_SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-6.PV", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-6.SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-6.CVH", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-6.CVC", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "SL4824-LR-6.RS_SP", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Version", Character, "Nominal" )
        ),
        Import Settings(
            End Of Line( CRLF, CR, LF ),
            End Of Field( Comma, CSV( 1 ) ),
            Treat Leading Zeros as Character( 1 ),
            Strip Quotes( 0 ),
            Use Apostrophe as Quotation Mark( 0 ),
            Use Regional Settings( 0 ),
            Scan Whole File( 1 ),
            Treat empty columns as numeric( 0 ),
            CompressNumericColumns( 0 ),
            CompressCharacterColumns( 0 ),
            CompressAllowListCheck( 0 ),
            Labels( 1 ),
            Column Names Start( 1 ),
            First Named Column( 1 ),
            Data Starts( 2 ),
            Lines To Read( "All" ),
            Year Rule( "20xx" )
        )
    );
    Print("✅ Temperature CSV data imported successfully");
,
    Print("❌ Error importing temperature CSV data");
    Throw("Failed to import temperature CSV data");
);

// Step 4: Create new table scripts (copies of previous experiment scripts)
Print("🔄 Creating new temperature table scripts...");

// Create new scripts for current experiment based on the JMP log pattern
// From the log, we see the pattern: rename "K3P#### XXX T 2" to "K3P{current_date} XXX T"
script_types = {{"C60 T", "EDAI T", "MgF2 T"}};

For( i = 1, i <= N Items( script_types ), i++,
    script_type = script_types[i];
    
    Try(
        // Get all table script names to find the right one to copy
        script_names = temp_dt << Get Property Names;
        found_script = "";
        
        // Look for scripts ending with " 2" to rename
        For( j = 1, j <= N Items( script_names ), j++,
            script_name = script_names[j];
            If( Contains( script_name, script_type ) & Contains( script_name, " 2" ),
                found_script = script_name;
                Break();
            );
        );
        
        If( found_script != "",
            new_name = "{experiment_number} " || script_type;
            temp_dt << Rename Table Script( found_script, new_name );
            Print("✅ Created new script: " || found_script || " → " || new_name);
        ,
            Print("⚠️ Could not find script ending with '" || script_type || " 2'");
            
            // Alternative: try to find any script with this type and copy it
            For( j = 1, j <= N Items( script_names ), j++,
                script_name = script_names[j];
                If( Contains( script_name, script_type ) & !Contains( script_name, "{experiment_number}" ),
                    // Found a script to copy - create new name
                    new_name = "{experiment_number} " || script_type;
                    
                    // Get the script content and create a new one
                    script_content = temp_dt << Get Property( script_name );
                    temp_dt << Set Property( new_name, script_content );
                    Print("✅ Copied script: " || script_name || " → " || new_name);
                    Break();
                );
            );
        );
    ,
        Print("❌ Error creating script for: " || script_type);
    );
);

// Step 5: Update table scripts with new reference lines and time ranges
Print("📊 Updating temperature table scripts with new reference lines...");

// C60 Temperature script (SL4824-LR-3.PV and SL4824-LR-4.PV)
Try(
    temp_dt << Set Property( "{experiment_number} C60 T",
        Graph Builder(
            Size( 921, 333 ),
            Show Control Panel( 0 ),
            Fit to Window( "Off" ),
            Variables(
                X( :Date and time ),
                Y( :"SL4824-LR-3.PV"n ),
                Y( :"SL4824-LR-4.PV"n, Position( 1 ) )
            ),
            Elements( Line( X, Y( 1 ), Y( 2 ), Legend( 6 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Date and time", ScaleBox,
                    {{Min( {min_time} ), Max( {max_time} ), Interval( "Minute" ), Inc( 60 ), Minor Ticks( 5 ),
                    {combined_ref_lines}
                    Label Row(
                        {{Label Orientation( "Perpendicular" ), Show Major Grid( 1 ),
                        Show Minor Grid( 1 )}}
                    )}}
                ),
                Dispatch( {{}}, "SL4824-LR-3.PV", ScaleBox,
                    {{Min( 450 ), Max( 600 ), Inc( 50 ), Minor Ticks( 4 ),
                    Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
                )
            )
        )
    );
    Print("✅ Updated {experiment_number} C60 T script");
,
    Print("❌ Error updating C60 T script");
);

// EDAI Temperature script (SL4824-LR-1.PV and SL4824-LR-2.PV)
Try(
    temp_dt << Set Property( "{experiment_number} EDAI T",
        Graph Builder(
            Size( 921, 333 ),
            Show Control Panel( 0 ),
            Fit to Window( "Off" ),
            Variables(
                X( :Date and time ),
                Y( :"SL4824-LR-1.PV"n ),
                Y( :"SL4824-LR-2.PV"n, Position( 1 ) )
            ),
            Elements( Line( X, Y( 1 ), Y( 2 ), Legend( 6 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Date and time", ScaleBox,
                    {{Min( {min_time} ), Max( {max_time} ), Interval( "Minute" ), Inc( 60 ), Minor Ticks( 5 ),
                    {combined_ref_lines}
                    Label Row(
                        {{Label Orientation( "Perpendicular" ), Show Major Grid( 1 ),
                        Show Minor Grid( 1 )}}
                    )}}
                ),
                Dispatch( {{}}, "SL4824-LR-1.PV", ScaleBox,
                    {{Min( 145 ), Max( 170 ), Inc( 5 ), Minor Ticks( 4 ),
                    Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
                )
            )
        )
    );
    Print("✅ Updated {experiment_number} EDAI T script");
,
    Print("❌ Error updating EDAI T script");
);

// MgF2 Temperature script (SL4824-LR-5.PV and SL4824-LR-6.PV)
Try(
    temp_dt << Set Property( "{experiment_number} MgF2 T",
        Graph Builder(
            Size( 921, 333 ),
            Show Control Panel( 0 ),
            Fit to Window( "Off" ),
            Variables(
                X( :Date and time ),
                Y( :"SL4824-LR-5.PV"n ),
                Y( :"SL4824-LR-6.PV"n, Position( 1 ) )
            ),
            Elements( Line( X, Y( 1 ), Y( 2 ), Legend( 6 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Date and time", ScaleBox,
                    {{Min( {min_time} ), Max( {max_time} ), Interval( "Minute" ), Inc( 60 ), Minor Ticks( 5 ),
                    {combined_ref_lines}
                    Label Row(
                        {{Label Orientation( "Perpendicular" ), Show Major Grid( 1 ),
                        Show Minor Grid( 1 )}}
                    )}}
                ),
                Dispatch( {{}}, "SL4824-LR-5.PV", ScaleBox,
                    {{Min( 1175 ), Max( 1300 ), Inc( 5 ), Minor Ticks( 4 ),
                    Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
                )
            )
        )
    );
    Print("✅ Updated {experiment_number} MgF2 T script");
,
    Print("❌ Error updating MgF2 T script");
);

// Step 6: Script automation complete - ready for manual concatenation
Print("✅ Temperature script setup complete!");
Print("🌡️ Temperature table scripts have been updated with correct reference lines and time ranges");
Print("📋 Next manual steps:");
Print("   1. Manually concatenate the CSV data with the temperature data table");
Print("   2. Save the concatenated file as {current_date} SL4824-LR.jmp");
Print("   3. The temperature data workflow is ready");

Print("🎉 Temperature automation phase complete!");
Print("📋 Manual steps remaining:");
Print("   1. Concatenate the CSV data with the temperature data table");
Print("   2. Save as {current_date} SL4824-LR.jmp in the {current_date} folder");
Print("   3. Temperature plots are ready with updated reference lines");
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
    automate_temp_workflow()
