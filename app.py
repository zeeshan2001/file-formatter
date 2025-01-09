from flask import Flask, render_template, request, redirect, url_for, send_file
import os
from openpyxl import load_workbook

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    # Check for source file
    if 'source_file' not in request.files or request.files['source_file'].filename == '':
        return "Source file is required"
    
    # Check for target file
    if 'target_file' not in request.files or request.files['target_file'].filename == '':
        return "Target file is required"
    
    # Save source file
    source_file = request.files['source_file']
    source_filepath = os.path.join(app.config['UPLOAD_FOLDER'], source_file.filename)
    source_file.save(source_filepath)
    
    # Save target file
    target_file = request.files['target_file']
    target_filepath = os.path.join(app.config['UPLOAD_FOLDER'], target_file.filename)
    target_file.save(target_filepath)
    
    return redirect(url_for('process_data', source_file=source_file.filename, target_file=target_file.filename))

@app.route('/process/<source_file>/<target_file>')
def process_data(source_file, target_file):
    # Paths to uploaded files
    source_file_path = os.path.join(app.config['UPLOAD_FOLDER'], source_file)
    target_file_path = os.path.join(app.config['UPLOAD_FOLDER'], target_file)

    # Load target and source workbooks
    target_workbook = load_workbook(target_file_path, data_only=True)
    source_workbook = load_workbook(source_file_path, data_only=True)

    # Select the "Financial Statements" sheet from the target workbook
    target_sheet = target_workbook["Financial Statements"]

    # Define starting row and columns for "Year" (F), "Tab" (N), and "Row" (M)
    start_row = 6  # Start from row 6 to skip the header
    year_col = 'F'
    tab_col = 'N'
    row_col = 'M'

    # Iterate through the rows in the target sheet
    for row in range(start_row, target_sheet.max_row + 1):
        tab_value = target_sheet[f"{tab_col}{row}"].value  # Get Tab value (sheet name)
        row_value = target_sheet[f"{row_col}{row}"].value  # Get Row value

        if tab_value and row_value:
            try:
                # Access the source sheet based on the Tab value
                source_sheet = source_workbook[tab_value.strip()]

                # Find the row in source sheet where column D matches the Row value
                found_value = None
                for source_row in range(1, source_sheet.max_row + 1):
                    source_row_value = source_sheet.cell(row=source_row, column=4).value  # Column D
                    if source_row_value == int(row_value):  # Match Row value
                        found_value = source_sheet.cell(row=source_row, column=7).value  # Column G
                        break

                # Update the Year column in the target sheet
                target_sheet[f"{year_col}{row}"] = found_value

                print(f"Updated Row {row}: F{row} with value '{found_value}' from Sheet '{tab_value}', Row {row_value}")

            except KeyError:
                print(f"Sheet '{tab_value}' not found in source file.")
            except ValueError:
                print(f"Invalid Row value '{row_value}' in Row {row}.")
            except Exception as e:
                print(f"Error processing Tab '{tab_value}', Row '{row_value}': {e}")
        else:
            print(f"Skipping Row {row} due to missing Tab or Row value.")

    # Save the updated target workbook
    updated_target_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'updated_target_sheet.xlsx')
    target_workbook.save(updated_target_file_path)
    print(f"Updated target sheet saved as '{updated_target_file_path}'")

    return send_file(updated_target_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
