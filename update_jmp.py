import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
import pandas as pd

def convert_pc3_runs():
    # Set up access
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scope)
    client = gspread.authorize(creds)

    # Open the target sheet and worksheet
    sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1g9Vb8OJJ67kEaGhXoxyBRHD-3oxouyLQr66b5V6JnVE")
    worksheet = sheet.worksheet("K3 PC3 Runs")

    # Read full range of Date, Start Time, End Time
    dates = worksheet.col_values(1)[1:]  # Skip header
    starts = worksheet.col_values(2)[1:]
    ends = worksheet.col_values(3)[1:]

    rows = []
    for date_str, start_str, end_str in zip(dates, starts, ends):
        try:
            # Parse full datetime objects
            start_dt = datetime.strptime(f"{date_str} {start_str}", "%Y/%m/%d %H:%M")
            end_dt = datetime.strptime(f"{date_str} {end_str}", "%Y/%m/%d %H:%M")

            # Store formatted strings
            rows.append((
                start_dt.date(),  # Keep just date for filtering
                start_dt.strftime("%y%m%d%H%M"),
                end_dt.strftime("%y%m%d%H%M")
            ))
        except Exception:
            continue  # Skip rows with formatting issues

    if not rows:
        print("No valid rows found.")
        return

    # Determine most recent date
    latest_date = max(row[0] for row in rows)
    print(f"\nMost recent run date: {latest_date}")

    # Filter by most recent date
    recent_rows = [(a, b) for d, a, b in rows if d == latest_date]

    # Output final result
    print("\nFormatted timestamps:")
    for start, end in recent_rows:
        print(f"{start}   {end}")

    # Generate JMP script to add the data
    generate_jmp_script(recent_rows)

    return recent_rows


def generate_jmp_script(timestamp_pairs):
    """
    Generate a JMP script (JSL) that can be run in JMP to add timestamp pairs to the existing JMP file
    
    Args:
        timestamp_pairs: List of tuples (start_time, end_time) in YYMMDDHHMMM format
    """
    if not timestamp_pairs:
        print("No timestamp pairs to add to JMP file.")
        return
    
    # JMP file path (escaped for JMP script)
    jmp_file_path = '/Users/justinparayno/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/Shared with Swift computers (INSECURE)/K3/K3 Tool Data/2025/PC3 Time Windows.jmp'
    
    # Generate JSL script content
    jsl_script = f'''// JMP Script to add new PC3 timestamp data
// Generated automatically by update_jmp.py on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

// Open the existing JMP file
dt = Open("{jmp_file_path}");

// Get current number of rows
current_rows = N Rows(dt);

// Add new rows for the timestamp data
dt << Add Rows({len(timestamp_pairs)});

// Get column references (assuming first two columns are for timestamps)
col_names = dt << Get Column Names();
start_col = Column(dt, col_names[1]);
end_col = Column(dt, col_names[2]);

// Add the new timestamp data
'''
    
    # Add data insertion code for each timestamp pair
    for i, (start_time, end_time) in enumerate(timestamp_pairs):
        row_num = f"current_rows + {i + 1}"
        jsl_script += f'''start_col[{row_num}] = Num("{start_time}");
end_col[{row_num}] = Num("{end_time}");
'''
    
    # Add script footer
    jsl_script += f'''
// Save the updated file
dt << Save();

// Display confirmation message
Print("Successfully added {len(timestamp_pairs)} new timestamp pairs to JMP file");
Print("JMP file now contains " || Char(N Rows(dt)) || " total rows");

// Show what was added
Print("\\nAdded to JMP file:");
'''
    
    # Add confirmation display code
    for i, (start_time, end_time) in enumerate(timestamp_pairs):
        row_num = f"current_rows + {i + 1}"
        jsl_script += f'''Print("  Row " || Char({row_num}) || ": " || col_names[1] || "={start_time}, " || col_names[2] || "={end_time}");
'''
    
    # Write the JSL script to a file
    script_filename = "add_pc3_timestamps.jsl"
    script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), script_filename)
    
    try:
        with open(script_path, 'w') as f:
            f.write(jsl_script)
        
        print(f"\\n✅ JMP script generated successfully!")
        print(f"📄 Script saved as: {script_path}")
        print(f"📊 Will add {len(timestamp_pairs)} timestamp pairs to JMP file")
        
        print("\\n🔧 How to use:")
        print("1. Open JMP software")
        print("2. Go to File → Open → Script...")
        print(f"3. Select the file: {script_filename}")
        print("4. Click 'Run Script' or press Ctrl+R (Cmd+R on Mac)")
        print("5. The script will automatically add the new data to your JMP file")
        
        # Display what was added
        print("\nTimestamp pairs that will be added to JMP file:")
        for i, (start, end) in enumerate(timestamp_pairs):
            print(f"  Pair {i + 1}: Start={start}, End={end}")
            
    except Exception as e:
        print(f"Error generating JMP script: {e}")


# Run it
if __name__ == "__main__":
    convert_pc3_runs()