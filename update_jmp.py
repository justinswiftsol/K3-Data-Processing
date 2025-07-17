import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import subprocess
import os

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

    # Generate JSL script
    generate_jsl_script(recent_rows)
    
    return recent_rows

def generate_jsl_script(timestamp_pairs):
    """Generate JSL script with timestamp pairs"""
    
    # Create the JSL script content
    jsl_content = f"""// Auto-generated JSL script for K3 PC3 QCM temperature logger
// Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
// Number of timestamp pairs: {len(timestamp_pairs)}

// Open the JMP file
dt = Open("/Users/justinparayno/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/Shared with Swift computers (INSECURE)/K3/K3 PC3 Post Run Data/K3 PC3 Post Run Data.jmp");

// Add timestamp pairs to the "Run Start" and "Run End" columns
"""

    # Add each timestamp pair
    for i, (start, end) in enumerate(timestamp_pairs, 1):
        jsl_content += f"""
// Timestamp pair {i}
dt << Add Rows(1);
current_row = N Rows(dt);
dt:"Run Start"[current_row] = {start};
dt:"Run End"[current_row] = {end};"""

    # Add save command
    jsl_content += """

// Save the file
dt << Save();

// Display completion message
Print("✅ Successfully added """ + str(len(timestamp_pairs)) + """ timestamp pairs to the JMP file!");
Print("💾 File has been saved.");
"""

    # Write the JSL script to file
    jsl_filename = "update_timestamps.jsl"
    with open(jsl_filename, 'w') as f:
        f.write(jsl_content)
    
    print(f"\n✅ JSL script generated: {jsl_filename}")
    print(f"📝 Added {len(timestamp_pairs)} timestamp pairs to the script")
    
    # Open the JSL file in JMP
    try:
        print("🚀 Opening JSL script in JMP...")
        subprocess.run([
            "open", 
            "-a", "JMP Pro", 
            jsl_filename
        ], check=True)
        print("✅ JSL script opened in JMP successfully!")
        print("👆 Click 'Run' or press Cmd+R in JMP to execute the script")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error opening JSL script in JMP: {e}")
        print(f"📁 You can manually open the file: {os.path.abspath(jsl_filename)}")

# Run it
if __name__ == "__main__":
    convert_pc3_runs()