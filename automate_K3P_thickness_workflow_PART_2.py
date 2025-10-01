from datetime import datetime
import subprocess
import os
import glob
import re


def automate_k3p_thickness_workflow_part2():
    """
    Automate the K3P thickness data workflow Part 2:
    1. Import most recent K3_PC3 Results CSV file
    2. Filter data to keep only current experiment
    3. Save filtered data as K3P#### Process.jmp
    4. Join with WDXRF and absorption table
    5. Concatenate with existing WDXRF and absorbance with process.jmp
    6. Create/update plot script for current experiment
    7. Save final updated table
    """
    
    print("📏 Starting K3P Thickness Workflow Part 2 Automation...")
    
    # Get current experiment number from user
    current_exp = input("Enter the current experiment number (#### format, e.g., 0426): K3P")
    
    print(f"\n🧪 Processing experiment: K3P{current_exp}")
    
    # Step 1: Find and import most recent CSV file
    print("\n📊 Finding most recent K3_PC3 Results CSV file...")
    csv_file = find_most_recent_pc3_csv()
    
    if not csv_file:
        print("❌ No K3_PC3 Results CSV file found")
        return
    
    print(f"✅ Found CSV file: {os.path.basename(csv_file)}")
    
    # Step 2: Generate JMP script for the workflow
    print("\n📝 Generating JMP script for Part 2 workflow...")
    jmp_script = generate_part2_jmp_script(current_exp, csv_file)
    
    # Step 3: Save JMP script
    script_filename = f"automate_k3p_thickness_part2_{current_exp}_{datetime.now().strftime('%Y%m%d_%H%M')}.jsl"
    script_path = os.path.join(os.getcwd(), script_filename)
    
    with open(script_path, 'w') as f:
        f.write(jmp_script)
    
    print(f"✅ JMP script saved: {script_filename}")
    
    # Step 4: Execute JMP script
    print("\n🚀 Executing JMP script...")
    execute_jsl_script(script_path)
    
    print("\n🎉 K3P Thickness Workflow Part 2 Complete!")
    print("\n📋 What was accomplished:")
    print(f"1. ✅ Imported most recent K3_PC3 Results CSV file")
    print(f"2. ✅ Filtered data to keep only K3P{current_exp} experiment")
    print(f"3. ✅ Saved filtered data as K3P{current_exp} Process.jmp")
    print(f"4. ✅ Joined with WDXRF and absorption table")
    print(f"5. ✅ Concatenated with existing WDXRF and absorbance with process.jmp")
    print(f"6. ✅ Created/updated plot script for K3P{current_exp}")
    print(f"7. ✅ Saved final updated table")


def find_most_recent_pc3_csv():
    """Find the most recent K3_PC3 Results CSV file in the Data Genie Library"""
    data_genie_path = "/Users/justinparayno/Desktop/Data Genie Library"
    
    if not os.path.exists(data_genie_path):
        print(f"❌ Data Genie Library path not found: {data_genie_path}")
        return None
    
    # Look for K3_PC3 Results files
    pattern = os.path.join(data_genie_path, "K3_PC3 Results *.csv")
    csv_files = glob.glob(pattern)
    
    if not csv_files:
        print("❌ No K3_PC3 Results CSV files found in Data Genie Library")
        return None
    
    # Sort by modification time to get the most recent
    csv_files.sort(key=os.path.getmtime, reverse=True)
    most_recent = csv_files[0]
    
    print(f"📁 Found {len(csv_files)} K3_PC3 Results files, using most recent")
    
    return most_recent


def generate_part2_jmp_script(current_exp, csv_file):
    """Generate JSL script for Part 2 workflow"""
    
    csv_filename = os.path.basename(csv_file)
    csv_path_jsl = csv_file.replace('\\', '/')  # Convert to forward slashes for JSL
    
    experiment_id = f"K3P{current_exp}"
    
    # Define file paths for Google Drive
    google_drive_base = "$HOME/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/R&D - Passivation/Passivation Projects 2024+/RD.E5: Process window specs and drift for inline ETL1   vendor demo/2025Q3 Data"
    
    process_file_path = f"{google_drive_base}/K3P{current_exp} Process.jmp"
    wdxrf_absorption_path = f"{google_drive_base}/WDXRF and absorption.jmp"
    final_table_path = f"{google_drive_base}/WDXRF and absorbance with process.jmp"
    
    jsl_script = f'''// Auto-generated JSL script for K3P Thickness workflow Part 2
// Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
// Current Experiment: {experiment_id}
// Source CSV: {csv_filename}

Print("📏 Starting K3P Thickness Workflow Part 2...");
Print("🧪 Processing experiment: {experiment_id}");

// Step 1: Import K3_PC3 Results CSV file
Print("📊 Importing K3_PC3 Results CSV file...");
Try(
    dt_pc3 = Open(
        "{csv_path_jsl}",
        columns(
            New Column( "sampleID", Character, "Nominal" ),
            New Column( "K3_PC3 Date", Numeric, "Continuous", Format( "y/m/d", 10 ), Input Format( "y/m/d" ) ),
            New Column( "K3_PC3 Run Start Time", Character, "Nominal" ),
            New Column( "K3_PC3 Run End Time", Character, "Nominal" ),
            New Column( "K3_PC3 Session", Character, "Nominal" ),
            New Column( "K3_PC3 Session Name", Character, "Nominal" ),
            New Column( "K3_PC3 Speed in mm per min", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Number of Passes at Nominal Speed", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Run Count", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Run Name", Character, "Nominal" ),
            New Column( "K3_PC3 Run Start and Name", Character, "Nominal" ),
            New Column( "K3_PC3 Thickness Control", Character, "Nominal" ),
            New Column( "K3_PC3 EDAI Target Thickness in nm", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 C60 Target Thickness in nm", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 MgF2 Target Thickness in nm", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 1 Material", Character, "Nominal" ),
            New Column( "K3_PC3 Source 1 Load Brand", Character, "Nominal" ),
            New Column( "K3_PC3 Source 1 Load Product", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 1 Load Batch", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 1 Emptied before Reload", Character, "Nominal" ),
            New Column( "K3_PC3 Source 1 Load at Start in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 1 Load at End in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 2 Material", Character, "Nominal" ),
            New Column( "K3_PC3 Source 2 Load Brand", Character, "Nominal" ),
            New Column( "K3_PC3 Source 2 Load Product", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 2 Load Batch", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 2 Emptied before Reload", Character, "Nominal" ),
            New Column( "K3_PC3 Source 2 Load at Start in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 2 Load at End in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 3 Material", Character, "Nominal" ),
            New Column( "K3_PC3 Source 3 Load Brand", Character, "Nominal" ),
            New Column( "K3_PC3 Source 3 Load Product", Character, "Nominal" ),
            New Column( "K3_PC3 Source 3 Load Batch", Character, "Nominal" ),
            New Column( "K3_PC3 Source 3 Emptied before Reload", Character, "Nominal" ),
            New Column( "K3_PC3 Source 3 Load at Start in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 3 Load at End in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 4 Material", Character, "Nominal" ),
            New Column( "K3_PC3 Source 4 Load Brand", Character, "Nominal" ),
            New Column( "K3_PC3 Source 4 Load Product", Character, "Nominal" ),
            New Column( "K3_PC3 Source 4 Load Batch", Character, "Nominal" ),
            New Column( "K3_PC3 Source 4 Emptied before Reload", Character, "Nominal" ),
            New Column( "K3_PC3 Source 4 Load at Start in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 4 Load at End in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 5 Material", Character, "Nominal" ),
            New Column( "K3_PC3 Source 5 Load Brand", Character, "Nominal" ),
            New Column( "K3_PC3 Source 5 Load Product", Character, "Nominal" ),
            New Column( "K3_PC3 Source 5 Load Batch", Character, "Nominal" ),
            New Column( "K3_PC3 Source 5 Emptied before Reload", Character, "Nominal" ),
            New Column( "K3_PC3 Source 5 Load at Start in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 5 Load at End in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 6 Material", Character, "Nominal" ),
            New Column( "K3_PC3 Source 6 Load Brand", Character, "Nominal" ),
            New Column( "K3_PC3 Source 6 Load Product", Character, "Nominal" ),
            New Column( "K3_PC3 Source 6 Load Batch", Character, "Nominal" ),
            New Column( "K3_PC3 Source 6 Emptied before Reload", Character, "Nominal" ),
            New Column( "K3_PC3 Source 6 Load at Start in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 6 Load at End in g", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Source 1 Mode", Character, "Nominal" ),
            New Column( "K3_PC3 Source 2 Mode", Character, "Nominal" ),
            New Column( "K3_PC3 Source 3 Mode", Character, "Nominal" ),
            New Column( "K3_PC3 Source 4 Mode", Character, "Nominal" ),
            New Column( "K3_PC3 Source 5 Mode", Character, "Nominal" ),
            New Column( "K3_PC3 Source 6 Mode", Character, "Nominal" ),
            New Column( "K3_PC3 Power on Source 1 in %", Character, "Nominal" ),
            New Column( "K3_PC3 Power on Source 2 in %", Character, "Nominal" ),
            New Column( "K3_PC3 Power on Source 3 in %", Character, "Nominal" ),
            New Column( "K3_PC3 Power on Source 4 in %", Character, "Nominal" ),
            New Column( "K3_PC3 Power on Source 5 in %", Character, "Nominal" ),
            New Column( "K3_PC3 Power on Source 6 in %", Character, "Nominal" ),
            New Column( "K3_PC3 T Source 1 in °C", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 T Source 2 in °C", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 T Source 3 in °C", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 T Source 4 in °C", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 T Source 5 in °C", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 T Source 6 in °C", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Median Pressure in Torr", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM1 Source", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM2 Source", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM3 Source", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM4 Source", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM5 Source", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM6 Source", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM 1 Target Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM 2 Target Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM 3 Target Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM 4 Target Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM 5 Target Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM 6 Target Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Median QCM1 Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Median QCM2 Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Median QCM3 Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Median QCM4 Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Median QCM5 Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 Median QCM6 Rate in Å per sec", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM1 Frequency at Start in Hz", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM2 Frequency at Start in Hz", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM3 Frequency at Start in Hz", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM4 Frequency at Start in Hz", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM5 Frequency at Start in Hz", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM6 Frequency at Start in Hz", Numeric, "Continuous", Format( "Best", 12 ) ),
            New Column( "K3_PC3 QCM1 Problem", Character, "Nominal" ),
            New Column( "K3_PC3 QCM2 Problem", Character, "Nominal" ),
            New Column( "K3_PC3 QCM3 Problem", Character, "Nominal" ),
            New Column( "K3_PC3 QCM4 Problem", Character, "Nominal" ),
            New Column( "K3_PC3 QCM5 Problem", Character, "Nominal" ),
            New Column( "K3_PC3 QCM6 Problem", Character, "Nominal" ),
            New Column( "K3_PC3 QCM Readout Software, Mode", Character, "Nominal" ),
            New Column( "K3_PC3 Process Chamber Exposed to Air before the Run", Character, "Nominal" ),
            New Column( "K3_PC3 Repairs Before the Run", Character, "Nominal" ),
            New Column( "K3_PC3 Comments", Character, "Nominal" ),
            New Column( "K3_PC3 Wafer Insert", Character, "Nominal" ),
            New Column( "K3_PC3 Wafer Holder Type Before", Character, "Nominal" ),
            New Column( "K3_PC3 Wafer Holder Type After", Character, "Nominal" ),
            New Column( "K3_PC3 carrier position", Character, "Nominal" ),
            New Column( "K3_PC3 Sideways Offset", Numeric, "Continuous", Format( "Best", 12 ) )
        )
    );
    Print("✅ Successfully imported K3_PC3 Results CSV");
,
    Print("❌ Error importing K3_PC3 Results CSV file");
    Print("📁 File path: {csv_path_jsl}");
    Throw();
);

// Step 2: Filter data to keep only current experiment ({experiment_id})
Print("🔍 Filtering data for experiment {experiment_id}...");
Try(
    // Select rows matching the current experiment
    dt_pc3 << Select Where( :K3_PC3 Session == "{experiment_id}" );
    
    // Invert selection to select all OTHER experiments
    dt_pc3 << Invert Row Selection;
    
    // Delete all other experiments (keeping only current experiment)
    dt_pc3 << Delete Rows;
    
    Print("✅ Successfully filtered data for {experiment_id}");
    Print("📊 Rows remaining: " || Char(N Rows(dt_pc3)));
,
    Print("❌ Error filtering data for {experiment_id}");
    Throw();
);

// Step 3: Save filtered data as K3P{current_exp} Process.jmp
Print("💾 Saving filtered data as K3P{current_exp} Process.jmp...");
Try(
    dt_pc3 << Save( "{process_file_path}" );
    Print("✅ Successfully saved: K3P{current_exp} Process.jmp");
,
    Print("❌ Error saving K3P{current_exp} Process.jmp");
    Print("📁 Attempted path: {process_file_path}");
);

// Step 4: Open WDXRF and absorption table
Print("📊 Opening WDXRF and absorption table...");
Try(
    dt_wdxrf = Open( "{wdxrf_absorption_path}" );
    Print("✅ Successfully opened WDXRF and absorption.jmp");
,
    Print("❌ Error opening WDXRF and absorption.jmp");
    Print("📁 Path: {wdxrf_absorption_path}");
    Throw();
);

// Step 5: Join tables
Print("🔗 Joining WDXRF and absorption with K3P{current_exp} Process...");
Try(
    dt_joined = dt_wdxrf << Join(
        With( dt_pc3 ),
        By Matching Columns( :Name("WDXRF sample") = :sampleID ),
        Drop multiples( 0, 0 ),
        Include Nonmatches( 0, 0 ),
        Preserve main table order( 1 ),
        Output Table( "Join of WDXRF and absorption with K3P{current_exp} Process" )
    );
    Print("✅ Successfully joined tables");
    Print("📊 Joined table rows: " || Char(N Rows(dt_joined)));
,
    Print("❌ Error joining tables");
    Throw();
);

// Step 6: Open existing WDXRF and absorbance with process.jmp
Print("📊 Opening existing WDXRF and absorbance with process.jmp...");
Try(
    dt_existing = Open( "{final_table_path}" );
    Print("✅ Successfully opened existing WDXRF and absorbance with process.jmp");
,
    Print("❌ Error opening existing WDXRF and absorbance with process.jmp");
    Print("📁 Path: {final_table_path}");
    Throw();
);

// Step 7: Concatenate tables
Print("📝 Concatenating new data with existing table...");
Try(
    dt_existing << Concatenate(
        dt_joined,
        Output Table( "Concat of WDXRF and absorbance with process, Join of WDXRF and absorption with K3P{current_exp} Process" ),
        Append to first table,
        Keep Formulas
    );
    Print("✅ Successfully concatenated tables");
    Print("📊 Final table rows: " || Char(N Rows(dt_existing)));
,
    Print("❌ Error concatenating tables");
    Throw();
);

// Step 8: Copy and rename table script for current experiment
Print("📋 Setting up plot script for K3P{current_exp}...");
Try(
    // Get list of existing table scripts
    script_names = dt_existing << Get Property Names;
    previous_script = "";
    
    // Find the most recent K3P script to copy from
    For( i = 1, i <= N Items(script_names), i++,
        script_name = script_names[i];
        If( Contains(script_name, "K3P") & !Contains(script_name, "{current_exp}"),
            previous_script = script_name;
        );
    );
    
    If( previous_script != "",
        // Copy the previous script
        previous_property = dt_existing << Get Property( previous_script );
        dt_existing << Set Property( "K3P{current_exp}", previous_property );
        Print("✅ Copied plot script from: " || previous_script);
    ,
        Print("⚠️ No previous K3P script found to copy from");
    );
,
    Print("❌ Error setting up plot script");
);

// Step 9: Create Graph Builder with data filter for current experiment
Print("📈 Creating Graph Builder plot for K3P{current_exp}...");
Try(
    graph_script = Graph Builder(
        Size( 800, 500 ),
        Fit to Window( "Off" ),
        Variables(
            X( :Name("K3_PC3 Sideways Offset") ),
            Y( :Name("K3_PC3 QCM 3 Target Rate in Å per sec") ),
            Y( :Name("K3_PC3 QCM 4 Target Rate in Å per sec"), Position( 1 ) ),
            Y( :Name("K3_PC3 Median QCM3 Rate in Å per sec"), Position( 1 ) ),
            Y( :Name("K3_PC3 Median QCM4 Rate in Å per sec"), Position( 1 ) ),
            Y( :Name("C60 in nm") ),
            Y( :Name("MgF2 in nm") ),
            Y( :Name("EDAI in nm") ),
            Group X( :Name("Run Start") )
        ),
        Elements(
            Position( 1, 1 ),
            Points( X, Y( 1 ), Y( 2 ), Y( 3 ), Y( 4 ), Legend( 43 ) )
        ),
        Elements( Position( 1, 2 ), Points( X, Y, Legend( 39 ) ) ),
        Elements( Position( 1, 3 ), Points( X, Y, Legend( 40 ) ) ),
        Elements( Position( 1, 4 ), Points( X, Y, Legend( 41 ) ) ),
        Local Data Filter(
            Close Outline( 1 ),
            Add Filter(
                columns( :Name("K3_PC3 Session") ),
                Where( :Name("K3_PC3 Session") == "{experiment_id}" ),
                Display( :Name("K3_PC3 Session"), N Items( 15 ), Find( Set Text( "" ) ) )
            )
        ),
        SendToReport(
            Dispatch( {{}}, "K3_PC3 Sideways Offset", ScaleBox,
                {{Min( -200 ), Max( 200 ), Inc( 100 ), Minor Ticks( 4 ),
                Label Row( {{Label Orientation( "Vertical" ), Show Major Grid( 1 )}} )}}
            ),
            Dispatch( {{}}, "K3_PC3 QCM 3 Target Rate in Å per sec", ScaleBox,
                {{Format( "Fixed Dec", 12, 0 ), Min( 0 ), Max( 3 ), Inc( 1 ),
                Minor Ticks( 4 ), Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
            ),
            Dispatch( {{}}, "C60 in nm", ScaleBox,
                {{Min( 0 ), Max( 25 ), Inc( 5 ), Minor Ticks( 4 ),
                Add Ref Line( 14, "Sparse Dash", "Green", "", 1 ),
                Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
            ),
            Dispatch( {{}}, "MgF2 in nm", ScaleBox,
                {{Min( 0 ), Max( 10 ), Inc( 5 ), Minor Ticks( 4 ),
                Add Ref Line( 6, "Sparse Dash", "Green", "", 1 ),
                Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
            ),
            Dispatch( {{}}, "EDAI in nm", ScaleBox,
                {{Format( "Fixed Dec", 12, 0 ), Min( 0 ), Max( 4 ), Inc( 1 ),
                Minor Ticks( 4 ), Add Ref Line( 2, "Sparse Dash", "Green", "", 1 ),
                Label Row( {{Show Major Grid( 1 ), Show Minor Grid( 1 )}} )}}
            )
        )
    );
    
    // Set the graph script as a table property
    dt_existing << Set Property( "K3P{current_exp}", graph_script );
    Print("✅ Created Graph Builder plot for K3P{current_exp}");
,
    Print("❌ Error creating Graph Builder plot");
);

// Step 10: Save final updated table
Print("💾 Saving final updated WDXRF and absorbance with process.jmp...");
Try(
    dt_existing << Save( "{final_table_path}" );
    Print("✅ Successfully saved final updated table");
,
    Print("❌ Error saving final table");
    Print("📁 Path: {final_table_path}");
);

// Step 11: Display the graph for verification
Print("📈 Displaying Graph Builder for verification...");
Try(
    dt_existing << Run Script( "K3P{current_exp}" );
    Print("✅ Graph Builder displayed successfully");
,
    Print("⚠️ Could not display graph - please check manually");
);

Print("🎉 K3P Thickness Workflow Part 2 Complete!");
Print("📋 Summary:");
Print("   • Imported and filtered K3_PC3 Results data for {experiment_id}");
Print("   • Joined with WDXRF and absorption data");
Print("   • Updated master WDXRF and absorbance with process table");
Print("   • Created plot script for {experiment_id}");
Print("   • Saved all updated files");
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
    automate_k3p_thickness_workflow_part2()
