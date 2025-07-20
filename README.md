# CardScannerApp

A Python application designed for scanning and extracting data from Magic: The Gathering (MTG) card images using OCR (Optical Character Recognition). The app processes images of MTG cards, extracts key details such as card title, collector number, and set/expansion code, and organizes the data for export to a CSV file. It features a user-friendly Tkinter-based GUI to facilitate reviewing, editing, and managing MTG card data, making it ideal for collectors, traders, and inventory managers.

## Description
The CardScannerApp is specifically built to streamline the process of filtering and cataloging Magic: The Gathering card data. It automates the extraction of critical information from card images, such as the card title, collector number (e.g., "123/456" or "123"), and set/expansion code (e.g., "DMU" for Dominaria United). The app uses the `rapidocr-onnxruntime` library for OCR to read text from card images and applies custom parsing logic to identify and clean MTG-specific data formats (e.g., removing dots from set codes like "DMU.EN" to "DMU" and extracting the first part of collector numbers like "123/456" to "123"). Users can review and manually correct extracted data through the GUI, ensuring accuracy before exporting to a CSV file for use in collection management, trading, or inventory tracking.

## Features
- Select a folder containing MTG card images (.jpg, .jpeg, .png, .bmp, .tiff, .webp).
- Perform OCR to extract card details like title, collector number, and set/expansion code (optional, using `rapidocr-onnxruntime`).
- Display card images alongside extracted text in a split-pane GUI.
- Edit extracted data to correct OCR errors or update card information.
- Export processed MTG card data to a CSV file for easy integration with collection tools.
- Progress tracking and detailed processing log for monitoring operations.

## Prerequisites
- Python 3.6+
- Required Python packages (see `requirements.txt`)

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/<your-username>/CardScannerApp.git
   cd CardScannerApp
   ```
2. Create a virtual environment (optional but recommended):
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
   python card_scanner_app.py
   ```

## Usage
1. Launch the app using `python card_scanner_app.py`.
2. Select an input folder containing MTG card images.
3. Specify an output CSV file (default: `card_data.csv`).
4. Click "Process Images" to start scanning and extracting card data.
5. Review and edit extracted data (e.g., card title, collector number, set code) in the GUI.
6. Click "Export to CSV" to save the filtered MTG card data.

## Dependencies
Listed in `requirements.txt`. Note that `rapidocr-onnxruntime` is optional for OCR functionality. Without it, the app will still allow manual data entry and image viewing but won't perform text recognition.

## Notes
- If `rapidocr-onnxruntime` is not installed, the app will display a warning, and OCR functionality will be disabled.
- For best OCR results, ensure MTG card images are clear, well-lit, and oriented correctly.
- The app includes specific parsing logic for MTG card data, such as handling set codes (e.g., "DMU.EN" → "DMU") and collector numbers (e.g., "123/456C" → "123").
- The exported CSV includes columns for filename, title, collector number, set/expansion, and processing status.

## License
MIT License (or specify your preferred license).

## Contributing
Contributions are welcome! Please submit a pull request or open an issue for suggestions, bug reports, or enhancements, especially for improving MTG-specific data parsing or adding support for additional card formats.