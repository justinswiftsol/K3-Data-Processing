import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import subprocess
import os
import glob


def automate_qcm_workflow():
    """
    Automate the complete QCM workflow:
    1. Get reference lines from Google Sheets
    2. Find the most recent CSV file
    3. Get experiment numbers from user
    4. Generate comprehensive JSL script
    5. Execute the JSL script in JMP
    """
    
    print("🚀 Starting QCM Workflow Automation...")
    
    # Get experiment numbers from user
    previous_exp = input("Enter the previous experiment number (e.g., 0395): K3P")
    current_exp = input("Enter the current experiment number (e.g., 0398): K3P")
    
    print(f"\n📋 Processing experiments: K3P{previous_exp} → K3P{current_exp}")
    
    # Step 1: Get reference lines from Google Sheets
    print("\n📊 Fetching reference lines from Google Sheets...")
    sheet_data = get_reference_lines_from_sheets()

    if not sheet_data:
        print("❌ Failed to get reference lines from Google Sheets")
        return

    print(f"✅ Retrieved reference lines and time ranges from Google Sheets")

    # Step 2: Find the most recent CSV file
    print("\n📁 Finding most recent CSV file...")
    csv_file = find_most_recent_csv()

    if not csv_file:
        print("❌ No CSV file found")
        return

    print(f"✅ Found CSV file: {os.path.basename(csv_file)}")

    # Step 3: Generate JSL script
    print("\n📝 Generating comprehensive JSL script...")
    jsl_script = generate_comprehensive_jsl(previous_exp, current_exp, csv_file, sheet_data)

    # Step 4: Save JSL script
    script_filename = f"automate_qcm_K3P{previous_exp}_to_K3P{current_exp}.jsl"
    script_path = os.path.join(os.getcwd(), script_filename)

    with open(script_path, 'w') as f:
        f.write(jsl_script)

    print(f"✅ JSL script saved: {script_filename}")

    # Step 5: Execute JSL script in JMP
    print("\n🚀 Executing JSL script in JMP...")
    execute_jsl_script(script_path)

    print("\n🎉 QCM Workflow Automation Complete!")
    print("\n📋 Next Steps:")
    print("1. Check that the 'QCM_log_data K3P#### By (Run)' table opened correctly")
    print("2. Manually copy Rate and Frequency data to Google Sheets results")
    print("3. The automation has completed all JMP operations")


def get_reference_lines_from_sheets():
    """Get pre-formatted JSL strings from Google Sheets using column indices"""
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
        worksheet = sheet.worksheet("K3 PC3 Runs")
        
        # Get all data as values to access by column index
        all_values = worksheet.get_all_values()
        
        if not all_values or len(all_values) < 2:
            print("❌ No data found in the worksheet")
            return {}
        
        print(f"🔍 Found {len(all_values)} rows of data (header + data rows)")
        
        # Since you've cleaned the sheet to only have recent experiment data,
        # we can just use the first data row (row 1, after header row 0)
        latest_row = all_values[1]  # First data row after header
        
        # Get the date from column A for reference
        date_str = str(latest_row[0]) if len(latest_row) > 0 else ""
        latest_date = None
        
        if date_str:
            try:
                latest_date = datetime.strptime(date_str, "%Y/%m/%d").date()
            except:
                try:
                    latest_date = datetime.strptime(date_str, "%m/%d/%Y").date()
                except:
                    try:
                        latest_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                    except:
                        latest_date = "Unknown"
                        
        print(f"📅 Using experiment date: {latest_date} from first data row")
        
        result = {
            'date': latest_date,
            'time_range_string': '',     # Column P (index 15)
            'ref_lines_h': '',           # Column H (index 7)  
            'ref_lines_i': ''            # Column I (index 8)
        }
        
        print(f"🔍 Row has {len(latest_row)} columns")
        
        # Get time range string from Column P (index 15)
        if len(latest_row) > 15:
            result['time_range_string'] = str(latest_row[15])  # Column P
            print(f"📊 Time range string (P): {result['time_range_string']}")
        else:
            print("❌ Column P (index 15) not available!")
        
        # Get reference lines from Column H (index 7)
        if len(latest_row) > 7:
            result['ref_lines_h'] = str(latest_row[7])  # Column H
            print(f"📏 Reference lines H: Found {len(result['ref_lines_h'])} characters")
            if len(result['ref_lines_h']) > 0:
                print(f"📏 Reference lines H content: {result['ref_lines_h'][:100]}...")
        else:
            print("❌ Column H (index 7) not available!")
        
        # Get reference lines from Column I (index 8)
        if len(latest_row) > 8:
            result['ref_lines_i'] = str(latest_row[8])  # Column I
            print(f"📏 Reference lines I: Found {len(result['ref_lines_i'])} characters")
            if len(result['ref_lines_i']) > 0:
                print(f"📏 Reference lines I content: {result['ref_lines_i'][:100]}...")
        else:
            print("❌ Column I (index 8) not available!")

        return result
        
    except Exception as e:
        print(f"❌ Error accessing Google Sheets: {e}")
        return {}


def find_most_recent_csv():
    """Find the most recent QCM_log_data CSV file"""
    csv_folder = "/Users/justinparayno/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/Shared with Swift computers (INSECURE)/K3/K3 PC3 QCM Colnatec"

    # Look for files matching the pattern QCM_log_data YYYYMMDD HHMM.csv
    pattern = os.path.join(csv_folder, "QCM_log_data ???????? ????.csv")
    csv_files = glob.glob(pattern)

    if not csv_files:
        return None

    # Sort by modification time and return the most recent
    csv_files.sort(key=os.path.getmtime, reverse=True)
    return csv_files[0]


def generate_comprehensive_jsl(previous_exp, current_exp, csv_file, sheet_data):
    """Generate the complete JSL script based on the log analysis"""

    # Extract pre-formatted JSL strings from Google Sheets
    time_range_string = sheet_data.get('time_range_string', '')
    ref_lines_h = sheet_data.get('ref_lines_h', '')
    ref_lines_i = sheet_data.get('ref_lines_i', '')
    
    # Combine reference lines from columns H and I
    combined_ref_lines = ref_lines_h
    if ref_lines_i:
        combined_ref_lines += "\n\t\t\t\t" + ref_lines_i if combined_ref_lines else ref_lines_i

    # Generate the comprehensive JSL script
    jsl_script = f'''// Auto-generated JSL script for QCM workflow automation
// Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
// Experiment transition: K3P{previous_exp} → K3P{current_exp}
// CSV file: {os.path.basename(csv_file)}
// Using pre-formatted JSL strings from Google Sheets

Print("🚀 Starting QCM Workflow Automation...");
Print("📋 Processing: K3P{previous_exp} → K3P{current_exp}");

// Step 1: Open required data tables
Print("📂 Opening required data tables...");

Try(
    // Open PC3 Time Windows (required for scripts to run)
    pc3_dt = Open("$HOME/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/Shared with Swift computers (INSECURE)/K3/K3 Tool Data/2025/PC3 Time Windows.jmp");
    Print("✅ Opened PC3 Time Windows.jmp");
,
    Print("❌ Error opening PC3 Time Windows.jmp");
    Throw("Failed to open PC3 Time Windows.jmp");
);

Try(
    // Open previous experiment data
    qcm_dt = Open("$HOME/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/Shared with Swift computers (INSECURE)/K3/K3 PC3 QCM Colnatec/QCM_log_data K3P{previous_exp}.jmp");
    Print("✅ Opened QCM_log_data K3P{previous_exp}.jmp");
,
    Print("❌ Error opening QCM_log_data K3P{previous_exp}.jmp");
    Throw("Failed to open QCM_log_data K3P{previous_exp}.jmp");
);

// Step 2: Import new CSV data
Print("📥 Importing new CSV data...");

Try(
    csv_dt = Open("{csv_file}",
        columns(
            New Column( "Date", Numeric, "Continuous", Format( "yyyy-mm-dd", 10 ), Input Format( "yyyy-mm-dd" ) ),
            New Column( "Time", Numeric, "Continuous", Format( "h:m:s", 15, 3 ), Input Format( "h:m:s", 3 ) ),
            New Column( "Unix Timestamp", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Rate QCM1 EDAI GB 160S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Rate QCM2 EDAI NGB 160S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Rate QCM3 C60 GB C7S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Rate QCM4 C60 NGB C4S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Rate QCM5 MgF2 GB C7S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Rate QCM6 MgF2 NGB C4S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "MA QCM1 EDAI GB 160S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "MA QCM2 EDAI NGB 160S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "MA QCM3 C60 GB C7S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "MA QCM4 C60 NGB C4S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "MA QCM5 MgF2 GB C7S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "MA QCM6 MgF2 NGB C4S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Frequency QCM1 EDAI GB 160S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Frequency QCM2 EDAI NGB 160S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Frequency QCM3 C60 GB C7S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Frequency QCM4 C60 NGB C4S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Frequency QCM5 MgF2 GB C7S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Frequency QCM6 MgF2 NGB C4S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Thickness QCM1 EDAI GB 160S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Thickness QCM2 EDAI NGB 160S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Thickness QCM3 C60 GB C7S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Thickness QCM4 C60 NGB C4S2", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Thickness QCM5 MgF2 GB C7S1", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "Thickness QCM6 MgF2 NGB C4S1", Numeric, "Continuous", Format( "Best", 12 ) ),
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
    Print("✅ CSV data imported successfully");
,
    Print("❌ Error importing CSV data");
    Throw("Failed to import CSV data");
);

// Step 3: Rename table scripts (prepare for new experiment)
Print("🔄 Renaming table scripts...");

script_types = {{"C60 early", "C60 late", "EDAI early", "EDAI late", "MgF2 early", "MgF2 late"}};

For( i = 1, i <= N Items( script_types ), i++,
    old_name = "K3P{previous_exp} " || script_types[i] || " 2";
    new_name = "K3P{current_exp} " || script_types[i];

    Try(
        qcm_dt << Rename Table Script( old_name, new_name );
        Print("✅ Renamed: " || old_name || " → " || new_name);
    ,
        Print("⚠️ Could not rename: " || old_name);
    );
);

// Step 4: Update table scripts with new reference lines
Print("📊 Updating table scripts with new reference lines...");

// C60 early script
Try(
    qcm_dt << Set Property( "K3P{current_exp} C60 early",
        Graph Builder(
            Size( 900, 250 ),
            Show Control Panel( 0 ),
            Fit to Window( "Off" ),
            Variables(
                X( :Date and Time ),
                Y( :Rate QCM3 C60 GB C7S2 ),
                Y( :Rate QCM4 C60 NGB C4S2, Position( 1 ) )
            ),
            Elements( Line( X, Y( 1 ), Y( 2 ), Legend( 6 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Date and Time", ScaleBox,
                    {{{time_range_string}
                    {combined_ref_lines}
                    Label Row(
                        {{Label Orientation( "Vertical" ), Show Major Grid( 1 ),
                        Show Minor Grid( 1 )}}
                    )}}
                ),
                Dispatch( {{}}, "Rate QCM3 C60 GB C7S2", ScaleBox,
                    {{Min( 0 ), Max( 4 ), Inc( 1 ), Minor Ticks( 4 ),
                    Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
                )
            )
        )
    );
    Print("✅ Updated K3P{current_exp} C60 early script");
,
    Print("❌ Error updating C60 early script");
);

// C60 late script
Try(
    qcm_dt << Set Property( "K3P{current_exp} C60 late",
        Graph Builder(
            Size( 900, 250 ),
            Show Control Panel( 0 ),
            Fit to Window( "Off" ),
            Variables(
                X( :Date and Time ),
                Y( :Rate QCM3 C60 GB C7S2 ),
                Y( :Rate QCM4 C60 NGB C4S2, Position( 1 ) )
            ),
            Elements( Line( X, Y( 1 ), Y( 2 ), Legend( 6 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Date and Time", ScaleBox,
                    {{{time_range_string}
                    {combined_ref_lines}
                    Label Row(
                        {{Label Orientation( "Vertical" ), Show Major Grid( 1 ),
                        Show Minor Grid( 1 )}}
                    )}}
                ),
                Dispatch( {{}}, "Rate QCM3 C60 GB C7S2", ScaleBox,
                    {{Min( 0 ), Max( 4 ), Inc( 1 ), Minor Ticks( 4 ),
                    Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
                )
            )
        )
    );
    Print("✅ Updated K3P{current_exp} C60 late script");
,
    Print("❌ Error updating C60 late script");
);

// EDAI early script
Try(
    qcm_dt << Set Property( "K3P{current_exp} EDAI early",
        Graph Builder(
            Size( 957, 250 ),
            Fit to Window( "Off" ),
            Variables(
                X( :Date and Time ),
                Y( :Rate QCM1 EDAI GB 160S2 ),
                Y( :Rate QCM2 EDAI NGB 160S1, Position( 1 ) )
            ),
            Elements( Line( X, Y( 1 ), Y( 2 ), Legend( 6 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Date and Time", ScaleBox,
                    {{{time_range_string}
                    {combined_ref_lines}
                    Label Row(
                        {{Label Orientation( "Vertical" ), Show Major Grid( 1 ),
                        Show Minor Grid( 1 )}}
                    )}}
                ),
                Dispatch( {{}}, "Rate QCM1 EDAI GB 160S2", ScaleBox,
                    {{Min( 0 ), Max( 10 ), Inc( 2 ), Minor Ticks( 1 ),
                    Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
                )
            )
        )
    );
    Print("✅ Updated K3P{current_exp} EDAI early script");
,
    Print("❌ Error updating EDAI early script");
);

// EDAI late script
Try(
    qcm_dt << Set Property( "K3P{current_exp} EDAI late",
        Graph Builder(
            Size( 957, 250 ),
            Fit to Window( "Off" ),
            Variables(
                X( :Date and Time ),
                Y( :Rate QCM1 EDAI GB 160S2 ),
                Y( :Rate QCM2 EDAI NGB 160S1, Position( 1 ) )
            ),
            Elements( Line( X, Y( 1 ), Y( 2 ), Legend( 6 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Date and Time", ScaleBox,
                    {{{time_range_string}
                    {combined_ref_lines}
                    Label Row(
                        {{Label Orientation( "Vertical" ), Show Major Grid( 1 ),
                        Show Minor Grid( 1 )}}
                    )}}
                ),
                Dispatch( {{}}, "Rate QCM1 EDAI GB 160S2", ScaleBox,
                    {{Min( 0 ), Max( 10 ), Inc( 2 ), Minor Ticks( 1 ),
                    Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
                )
            )
        )
    );
    Print("✅ Updated K3P{current_exp} EDAI late script");
,
    Print("❌ Error updating EDAI late script");
);

// MgF2 early script
Try(
    qcm_dt << Set Property( "K3P{current_exp} MgF2 early",
        Graph Builder(
            Size( 921, 352 ),
            Show Control Panel( 0 ),
            Fit to Window( "Off" ),
            Variables(
                X( :Date and Time ),
                Y( :Rate QCM5 MgF2 GB C7S1 ),
                Y( :Rate QCM6 MgF2 NGB C4S1, Position( 1 ) ),
                Y( :MA QCM5 MgF2 GB C7S1, Position( 1 ) ),
                Y( :MA QCM6 MgF2 NGB C4S1, Position( 1 ) )
            ),
            Elements( Line( X, Y( 1 ), Y( 2 ), Y( 3 ), Y( 4 ), Legend( 6 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Date and Time", ScaleBox,
                    {{{time_range_string}
                    {combined_ref_lines}
                    Label Row(
                        {{Label Orientation( "Vertical" ), Show Major Grid( 1 ),
                        Show Minor Grid( 1 )}}
                    )}}
                ),
                Dispatch( {{}}, "Rate QCM5 MgF2 GB C7S1", ScaleBox,
                    {{Format( "Fixed Dec", 12, 2 ), Min( 0 ), Max( 8 ), Inc( 1 ),
                    Minor Ticks( 2 ), Label Row(
                        {{Show Major Grid( 1 ), Show Minor Grid( 1 )}}
                    )}}
                )
            )
        )
    );
    Print("✅ Updated K3P{current_exp} MgF2 early script");
,
    Print("❌ Error updating MgF2 early script");
);

// MgF2 late script
Try(
    qcm_dt << Set Property( "K3P{current_exp} MgF2 late",
        Graph Builder(
            Size( 921, 352 ),
            Show Control Panel( 0 ),
            Fit to Window( "Off" ),
            Variables(
                X( :Date and Time ),
                Y( :Rate QCM5 MgF2 GB C7S1 ),
                Y( :Rate QCM6 MgF2 NGB C4S1, Position( 1 ) ),
                Y( :MA QCM5 MgF2 GB C7S1, Position( 1 ) ),
                Y( :MA QCM6 MgF2 NGB C4S1, Position( 1 ) )
            ),
            Elements( Line( X, Y( 1 ), Y( 2 ), Y( 3 ), Y( 4 ), Legend( 6 ) ) ),
            SendToReport(
                Dispatch( {{}}, "Date and Time", ScaleBox,
                    {{{time_range_string}
                    {combined_ref_lines}
                    Label Row(
                        {{Label Orientation( "Vertical" ), Show Major Grid( 1 ),
                        Show Minor Grid( 1 )}}
                    )}}
                ),
                Dispatch( {{}}, "Rate QCM5 MgF2 GB C7S1", ScaleBox,
                    {{Format( "Fixed Dec", 12, 2 ), Min( 0 ), Max( 8 ), Inc( 1 ),
                    Minor Ticks( 2 ), Label Row(
                        {{Show Major Grid( 1 ), Show Minor Grid( 1 )}}
                    )}}
                )
            )
        )
    );
    Print("✅ Updated K3P{current_exp} MgF2 late script");
,
    Print("❌ Error updating MgF2 late script");
);

// Step 5: Script automation complete - ready for manual concatenation
Print("✅ Script setup complete!");
Print("📊 Table scripts have been updated with correct reference lines and time ranges");
Print("📋 Next manual steps:");
Print("   1. Manually concatenate the CSV data with the QCM data table");
Print("   2. Save the concatenated file as QCM_log_data K3P{current_exp}.jmp");
Print("   3. Generate the By(Run) analysis table");
Print("   4. Copy Rate and Frequency data to Google Sheets");

Print("🎉 Automation phase complete!");
Print("📋 Manual steps remaining:");
Print("   1. Concatenate the CSV data with the QCM data table");
Print("   2. Save as QCM_log_data K3P{current_exp}.jmp");
Print("   3. Generate By(Run) analysis and copy data to Google Sheets");
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
    automate_qcm_workflow()
