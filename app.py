from flask import Flask, render_template, request, redirect, url_for, send_file
import os
import pandas as pd

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    if file:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        return redirect(url_for('process_data', filename=file.filename))


@app.route('/process/<filename>')
def process_data(filename):
    # Paths to target and source sheets
    target_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'target sheet.xlsx')
    source_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'sample source sheet.xlsx')
    
    # Load Target Sheet with correct header row (row 5, zero-indexed as 4)
    target_data = pd.read_excel(target_file_path, sheet_name='Financial Statements', header=4)
    target_data.columns = target_data.columns.str.strip()  # Remove any extra spaces from column names

    # Iterate over the rows of the Target Sheet
    for index, row in target_data.iterrows():
        tab = row['Tab']  # Extract the Tab (sheet name)
        source_row = row['Row']  # Extract the Row number
        
        # Skip rows with missing Tab or Row
        if pd.isna(tab) or pd.isna(source_row):
            continue

        # Load the specific sheet from the Source File
        try:
            source_data = pd.read_excel(source_file_path, sheet_name=tab)
        except Exception as e:
            print(f"Error loading sheet {tab}: {e}")
            continue

        # Extract value from the specified row for Year 1 (2023)
        try:
            value = source_data.loc[int(source_row) - 1, '2023']  # Adjust column name if necessary
            # Update column F (6th column) in the Target Sheet
            target_data.iloc[index, 5] = value  # Column F is the 6th column (index 5)
        except KeyError:
            print(f"Column '2023' not found in sheet {tab}")
        except IndexError:
            print(f"Row {source_row} not found in sheet {tab}")

    # Save the updated Target Sheet
    updated_target_path = os.path.join(app.config['UPLOAD_FOLDER'], 'updated_target_sheet.xlsx')
    target_data.to_excel(updated_target_path, index=False)

    return send_file(updated_target_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
