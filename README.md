# Flask File Processing App

This project is a Flask web application that processes Excel files. It uploads a source file and a target file, extracts data from the source, and updates the target file. The processed target file can be previewed and downloaded.

---

## Features

1. **Upload Files**:
   - Upload a source Excel file and a target Excel file via the web interface.

2. **Data Extraction and Update**:
   - Extracts data from specific columns in the source file and updates corresponding columns in the target file.
   - Data alignment is based on values in the target sheet's `M` and `N` columns.

3. **File Preview**:
   - Displays a preview of the updated target file.

4. **Download Processed File**:
   - Provides a download link for the processed target file.

---

## Installation

### Prerequisites

- Python 3.8+
- `pip` for managing Python packages
- Flask
- `openpyxl` for working with Excel files

### Steps

1. Clone the repository:
   ```bash
   git clone <repository_url>
   cd <repository_directory>
   ```

2. Create a virtual environment and activate it:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run the application:
   ```bash
   python app.py
   ```

5. Open your browser and go to:
   ```
   http://127.0.0.1:5000/
   ```

---

## File Structure

```
project_directory/
|├── app.py                 # Main Flask application
|├── templates/           # HTML templates
|   |├── index.html     # Upload form
|   |├── preview.html   # Preview processed file
|├── static/            # Static files (CSS, JS)
|├── uploads/           # Uploaded files
|├── output/            # Processed files
|└── requirements.txt   # Dependencies
```

---

## Usage

1. **Upload Files**:
   - Navigate to the homepage.
   - Select and upload both the source and target files.

2. **Processing**:
   - The app processes the files and updates the target sheet based on data from the source file.

3. **Preview and Download**:
   - Preview the updated target file in the browser.
   - Download the processed file via the provided link.

---

## Configuration

- **Upload Folder**:
  - Configured in `app.py`:
    ```python
    app.config['UPLOAD_FOLDER'] = 'uploads/'
    app.config['OUTPUT_FOLDER'] = 'output/'
    ```
- Ensure the `uploads` and `output` directories exist or are created at runtime.

---

## Dependencies

- Flask
- openpyxl
- Tailwind CSS (via CDN)
- Flowbite (via CDN)

---

## Known Issues

1. **Row Alignment**:
   - Ensure that data alignment between source and target files is correct.

2. **Missing Sheets**:
   - Verify that the target file contains a sheet named `Financial Statements`.




