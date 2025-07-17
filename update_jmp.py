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

    # Update JMP file with the new data
    update_jmp_file(recent_rows)

    return recent_rows


def update_jmp_file(timestamp_pairs):
    """
    Update the JMP file with new timestamp pairs
    
    Args:
        timestamp_pairs: List of tuples (start_time, end_time) in YYMMDDHHMMM format
    """
    jmp_file_path = '/Users/justinparayno/Library/CloudStorage/GoogleDrive-justin.parayno@swiftsolar.com/Shared drives/Shared with Swift computers (INSECURE)/K3/K3 Tool Data/2025/PC3 Time Windows.jmp'
    
    if not timestamp_pairs:
        print("No timestamp pairs to add to JMP file.")
        return
    
    try:
        # Check if file exists
        if not os.path.exists(jmp_file_path):
            print(f"Warning: JMP file not found at {jmp_file_path}")
            print("Please verify the file path is correct.")
            return
        
        # Read existing JMP file
        # JMP files are typically tab-separated text files
        try:
            existing_data = pd.read_csv(jmp_file_path, sep='\t')
            print(f"Successfully read existing JMP file with {len(existing_data)} rows")
        except Exception as e:
            print(f"Error reading JMP file: {e}")
            print("Attempting to read as CSV...")
            try:
                existing_data = pd.read_csv(jmp_file_path)
                print(f"Successfully read JMP file as CSV with {len(existing_data)} rows")
            except Exception as e2:
                print(f"Error reading JMP file as CSV: {e2}")
                return
        
        # Get column names (assuming first two columns are for start and end times)
        columns = existing_data.columns.tolist()
        if len(columns) < 2:
            print("Error: JMP file must have at least 2 columns for start and end times")
            return
        
        start_col = columns[0]
        end_col = columns[1]
        print(f"Using columns: '{start_col}' and '{end_col}'")
        
        # Create new data to append
        new_data = []
        for start_time, end_time in timestamp_pairs:
            new_row = {col: None for col in columns}  # Initialize all columns with None
            new_row[start_col] = start_time
            new_row[end_col] = end_time
            new_data.append(new_row)
        
        # Create DataFrame for new data
        new_df = pd.DataFrame(new_data)
        
        # Append new data to existing data
        updated_data = pd.concat([existing_data, new_df], ignore_index=True)
        
        # Write back to JMP file (preserve original format)
        if jmp_file_path.endswith('.jmp'):
            # Save as tab-separated for JMP compatibility
            updated_data.to_csv(jmp_file_path, sep='\t', index=False)
        else:
            updated_data.to_csv(jmp_file_path, index=False)
        
        print(f"Successfully added {len(timestamp_pairs)} new timestamp pairs to JMP file")
        print(f"JMP file now contains {len(updated_data)} total rows")
        
        # Display what was added
        print("\nAdded to JMP file:")
        for i, (start, end) in enumerate(timestamp_pairs):
            print(f"  Row {len(existing_data) + i + 1}: {start_col}={start}, {end_col}={end}")
            
    except Exception as e:
        print(f"Error updating JMP file: {e}")
        print("Please check the file path and permissions.")


# Run it
if __name__ == "__main__":
    convert_pc3_runs()