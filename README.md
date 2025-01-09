
# File Formatter Web Application

This is a Flask-based web application that processes and updates a target Excel sheet based on data from a source Excel sheet. It allows users to upload Excel files and automatically updates the target sheet with specified data.

## Features

- Upload Excel files for processing.
- Extract specific data from a source Excel sheet and update the target sheet.
- Download the updated target sheet.

---

## Installation

1. **Clone the Repository**
   ```bash
   git clone https://github.com/yourusername/file-formatter.git
   cd file-formatter
   ```

2. **Set Up a Virtual Environment** (optional, but recommended)
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

1. **Run the Flask App**
   ```bash
   python app.py
   ```

2. **Access the App**
   Open your browser and go to:  
   `http://127.0.0.1:5000/`

3. **Upload Files**
   - Upload your target and source Excel files.
   - The app will process the data and provide the updated target sheet for download.

---

## Folder Structure

```
file-formatter/
├── app.py                # Main application file
├── templates/
│   └── index.html        # HTML template for the web interface
├── uploads/              # Folder to store uploaded and processed files
├── requirements.txt      # List of dependencies
└── README.md             # Project documentation
```

---

## Dependencies

- Flask
- pandas
- openpyxl

---

## How It Works

1. The user uploads an Excel file through the web interface.
2. The app reads a target and a source Excel sheet.
3. Data from the source sheet is mapped to the target sheet based on specified rows and columns.
4. The updated target sheet is saved and provided for download.
