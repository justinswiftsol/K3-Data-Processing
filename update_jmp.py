import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

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

    return recent_rows


# Run it
if __name__ == "__main__":
    convert_pc3_runs()