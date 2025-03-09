from flask import Flask, render_template, request, redirect
import gspread
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# 🔹 Google Sheets API Setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
JSON_KEYFILE = "deep-ground-385217-93ab33993add.json"

creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEYFILE, scope)
client = gspread.authorize(creds)

SPREADSHEET_ID = "18W_YejaqG03ozfAAsehCfgkW4V1UMNax-BnAGyEg3sA"
worksheet = client.open_by_key(SPREADSHEET_ID).worksheet("Sheet1")

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Define fields that require special formatting
    number_fields = {"actual_weight", "post_fat_removal_weight", "pieces_per_kg", "ingredient"}
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
        "la_yellow_bad_smell", "la_white", "la_puff", "la_no_pass"
    ]

    # Extract form data
    form_data = {field: request.form.get(field, "") for field in form_fields}

    # ✅ **Find the Next Available Column**
    sheet_data = worksheet.get_all_values()
    last_filled_col = len(sheet_data[0]) if sheet_data else 1  # Get last used column (1-based index)
    next_column = last_filled_col + 1  # Move to the next column

    # ✅ **Ensure Next Column is Within Google Sheets Limit**
    if next_column > 98:
        next_column = 3  # Reset to column C if limit is exceeded

    # ✅ **Rows to be Skipped in Google Sheets**
    skipped_rows = {8, 29, 30, 67, 110, 144}

    # ✅ **Prepare Data for Update**
    formatted_data = []
    row_counter = 1

    for field in form_fields:
        if row_counter in skipped_rows:
            # ✅ Keep existing value for skipped rows
            existing_value = worksheet.cell(row_counter, next_column).value
            formatted_data.append(existing_value)
        else:
            value = form_data.get(field, "")
            if value:
                try:
                    if field in number_fields:
                        formatted_value = float(value)
                    elif field in date_fields:
                        formatted_value = datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")
                    else:
                        formatted_value = value
                except ValueError:
                    formatted_value = value
            else:
                formatted_value = ""

            formatted_data.append(formatted_value)

        row_counter += 1

    # ✅ **Calculate Column Letter for Update**
    column_letter = chr(64 + next_column)  # Convert column number to letter (A, B, C, ...)
    cell_range = f"{column_letter}1:{column_letter}{len(form_fields)}"

    try:
        # Update the worksheet with the new data
        worksheet.update(cell_range, [[value] for value in formatted_data])
        print(f"✅ Data saved to Column {column_letter}, skipped rows: {skipped_rows}")
    except gspread.exceptions.APIError as e:
        print(f"❌ Google Sheets API Error: {e}")

    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True, port=5002)