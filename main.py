from flask import Flask, render_template, request, redirect
import gspread
from datetime import datetime
import string
from oauth2client.service_account import ServiceAccountCredentials
import base64
import json
import os
from dotenv import load_dotenv

import os
from dotenv import load_dotenv

# ✅ Load the .env file
load_dotenv()

# ✅ Retrieve the Spreadsheet ID
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME")

# Debugging: Print the loaded value
print(f"🔹 Loaded Spreadsheet ID: {SPREADSHEET_ID}")

# ✅ Check if the variable is being loaded properly
if not SPREADSHEET_ID:
    raise ValueError("❌ ERROR: SPREADSHEET_ID is not found in the .env file")


# 🔹 Google Sheets API Setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]


# json_keyfile = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
# creds = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile, scope)
# client = gspread.authorize(creds)
google_creds_b64 = os.getenv("GOOGLE_CREDENTIALS_B64")

if not google_creds_b64:
    raise ValueError("🚨 GOOGLE_CREDENTIALS_B64 is missing. Set it in environment variables.")

google_creds_json = base64.b64decode(google_creds_b64).decode('utf-8')

# ✅ Convert JSON String to a Temporary File
with open("temp_google_creds.json", "w") as temp_file:
    temp_file.write(google_creds_json)

creds = ServiceAccountCredentials.from_json_keyfile_name("temp_google_creds.json", scope)

# ✅ Authorize gspread
client = gspread.authorize(creds)

worksheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Define fields that require special formatting
    number_fields = {"actual_weight", "post_fat_removal_weight", "pieces_per_kg","ingredient", "actual_weight", "post_fat_removal_weight",
        "defective_material", "bong_1", "bong_so", "bong_cat", "bong_8_10", "bong_less_8", "bong_b",
        "duoi", "vun", "mo", "dat", "bong_1_2", "bong_so_2", "bong_la", "bong_la_b", "dat_2", "vun_2",
        "la_total", "la_1", "la_1_vang", "la_2", "la_2_vang", "la_3", "la_3_vang", "la_4", "la_4_vang",
        "la_5", "la_5_6", "la_5_vang", "la_6", "la_6_vang", "la_7", "la_7_vang", "la_8", "la_8_vang",
        "la_8_9", "la_9", "la_9_vang", "la_10", "la_10_vang", "la_11", "la_11_vang", "la_12", "la_12_vang",
        "la_sample", "la_yellow", "la_green", "la_yellow_no_pass", "la_yellow_bad_smell_pass",
        "la_yellow_bad_smell", "la_white", "la_puff", "la_no_pass", "lb_total", "lb_1", "lb_1_vang",
        "lb_60_80", "lb_80_120", "lb_120_160", "lb_2", "lb_2_vang", "lb_3", "lb_3_vang",
        "lb_200_280", "lb_4", "lb_4_vang", "lb_5", "lb_5_6", "lb_5_vang", "lb_6", "lb_6_vang",
        "lb_7", "lb_7_vang", "lb_8", "lb_8_vang", "lb_8_9", "lb_9", "lb_480_760", "lb_9_vang",
        "lb_10", "lb_10_vang", "lb_10_vang_1", "lb_760_1000", "lb_sample", "lb_piece", "lb_yellow",
        "lb_green", "lb_yellow_sample", "lb_yellow_no_pass_bad_smell", "lb_yellow_bad_smell_pass",
        "lb_white", "lb_white_bad_smell", "lb_yellow_no_pass", "lb_yellow_bad_smell", "lb_no_pass",
        "lc_total", "lc_1", "lc_1_vang", "lc_1_good", "lc_0", "lc_0_vang", "lc_2", "lc_2_vang",
        "lc_3", "lc_3_vang", "lc_4", "lc_4_vang", "lc_5", "lc_5_vang", "lc_6", "lc_7", "lc_8",
        "lc_9", "lc_10", "lc_10_vang", "lc_yellow", "lc_sample", "lc_good", "lc_good_yellow",
        "lc_white", "lc_bad_smell_white", "lc_yellow_hole", "lc_yellow_no_pass", "lc_green",
        "lc_yellow_bad_smell_pass", "lc_black_no_pass", "lc_black_bad_smell_no_pass", "lc_no_pass"}
    date_fields = {"entry_date", "production_date", "fat_removal_date"}

    # List of all form fields
    form_fields = [
        "entry_date", "supplier", "production_date", "pieces_per_kg", "fat_removal_date",
        "raw_material_code", "remarks", "ingredient", "actual_weight", "post_fat_removal_weight",
        "defective_material", "bong_1", "bong_so", "bong_cat", "bong_8_10", "bong_less_8", "bong_b",
        "duoi", "vun", "mo", "dat", "bong_1_2", "bong_so_2", "bong_la", "bong_la_b", "dat_2", "vun_2",
        "la_total", "la_1", "la_1_vang", "la_2", "la_2_vang", "la_3", "la_3_vang", "la_4", "la_4_vang",
        "la_5", "la_5_6", "la_5_vang", "la_6", "la_6_vang", "la_7", "la_7_vang", "la_8", "la_8_vang",
        "la_8_9", "la_9", "la_9_vang", "la_10", "la_10_vang", "la_11", "la_11_vang", "la_12", "la_12_vang",
        "la_sample", "la_yellow", "la_green", "la_yellow_no_pass", "la_yellow_bad_smell_pass",
        "la_yellow_bad_smell", "la_white", "la_puff", "la_no_pass", "lb_total", "lb_1", "lb_1_vang",
        "lb_60_80", "lb_80_120", "lb_120_160", "lb_2", "lb_2_vang", "lb_3", "lb_3_vang",
        "lb_200_280", "lb_4", "lb_4_vang", "lb_5", "lb_5_6", "lb_5_vang", "lb_6", "lb_6_vang",
        "lb_7", "lb_7_vang", "lb_8", "lb_8_vang", "lb_8_9", "lb_9", "lb_480_760", "lb_9_vang",
        "lb_10", "lb_10_vang", "lb_10_vang_1", "lb_760_1000", "lb_sample", "lb_piece", "lb_yellow",
        "lb_green", "lb_yellow_sample", "lb_yellow_no_pass_bad_smell", "lb_yellow_bad_smell_pass",
        "lb_white", "lb_white_bad_smell", "lb_yellow_no_pass", "lb_yellow_bad_smell", "lb_no_pass",
        "lc_total", "lc_1", "lc_1_vang", "lc_1_good", "lc_0", "lc_0_vang", "lc_2", "lc_2_vang",
        "lc_3", "lc_3_vang", "lc_4", "lc_4_vang", "lc_5", "lc_5_vang", "lc_6", "lc_7", "lc_8",
        "lc_9", "lc_10", "lc_10_vang", "lc_yellow", "lc_sample", "lc_good", "lc_good_yellow",
        "lc_white", "lc_bad_smell_white", "lc_yellow_hole", "lc_yellow_no_pass", "lc_green",
        "lc_yellow_bad_smell_pass", "lc_black_no_pass", "lc_black_bad_smell_no_pass", "lc_no_pass"

    ]

    # Extract form data
    form_data = {field: request.form.get(field, "") for field in form_fields}

    # ✅ **Find Next Available Column (Always Start from Column C)**
    sheet_data = worksheet.get_all_values()

    # ✅ **Check if Column C is empty first**
    if not sheet_data or all(
            not cell for cell in [row[2] for row in sheet_data if len(row) > 2]):  # Check if C is empty
        next_column = 3  # Start at Column C
    else:
        # ✅ **Find last filled column that follows the pattern C → E → G → I**
        last_filled_col = 3
        while True:
            col_letter = string.ascii_uppercase[last_filled_col - 1]
            col_values = worksheet.col_values(last_filled_col)
            if all(cell == '' for cell in col_values[1:]):  # Check if column is empty
                break
            last_filled_col += 2  # Skip one column for spacing

        next_column = last_filled_col

    # ✅ **Ensure Column is Within Google Sheets Limit**
    if next_column > 98:
        next_column = 3  # Reset to column C if limit is exceeded

    # ✅ **Rows to be Skipped While Keeping Existing Values**
    skipped_rows = {8, 29, 30, 67, 110, 144}
    existing_values = {row: worksheet.cell(row, 2).value for row in skipped_rows}  # Read values from Column B

    formatted_data = []
    row_counter = 1
    form_index = 0  # Track index of form_fields

    while form_index < len(form_fields):
        if row_counter in skipped_rows:
            formatted_data.append([existing_values.get(row_counter, "")])
        else:
            field = form_fields[form_index]
            value = form_data.get(field, "")
            try:
                if field in number_fields:
                    formatted_value = float(value)
                elif field in date_fields:
                    formatted_value = datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")
                else:
                    formatted_value = value
            except ValueError:
                formatted_value = value

            formatted_data.append([formatted_value])
            form_index += 1  # Only increase form index when data is added

        row_counter += 1  # Always increase row counter to track correct row

    # ✅ Convert next_column to Google Sheets Column Letter
    def get_column_letter(n):
        result = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            result = string.ascii_uppercase[remainder] + result
        return result

    column_letter = get_column_letter(next_column)
    cell_range = f"{column_letter}1:{column_letter}{len(formatted_data)}"

    print(f"Updating range: {cell_range}")  # Debugging line
    try:
        worksheet.update(cell_range, formatted_data)
        print(f"✅ Data saved to Column {column_letter}, skipped rows: {skipped_rows}")
    except gspread.exceptions.APIError as e:
        print(f"❌ Google Sheets API Error: {e}")

    return redirect('/')


if __name__ == '__main__':
    app.run(debug=True, port=5002)