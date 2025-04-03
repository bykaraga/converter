# File Converter

A modern desktop application that allows you to convert files between various formats with a user-friendly interface.

## Features

- Excel → PDF conversion
- Word (.docx) → PDF conversion
- PDF → Word conversion
- Image to PDF conversion
- Drag & drop support
- Modern and intuitive user interface
- Batch file conversion support
- File preview capability
- Automatic save options
- CLI support

## Installation

1. Install the required Python packages:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python main.py
```

## Usage

### GUI Mode
1. Launch the application
2. Select a file by clicking the "Select File" button or by dragging and dropping
3. Choose your desired conversion type
4. Select the output location
5. You'll be notified when the conversion is complete

### CLI Mode
```bash
python main.py --input input_file.xlsx --output output_file.pdf --type excel_to_pdf
```

Available conversion types:
- `excel_to_pdf`
- `word_to_pdf`
- `pdf_to_word`
- `image_to_pdf`

## Requirements

- Python 3.8 or higher
- PyQt6
- pandas
- openpyxl
- reportlab
- python-docx
- pdfplumber
- pdf2docx
- Pillow

## Development

This project uses:
- PyQt6 for the modern GUI
- pandas for Excel file handling
- python-docx for Word document processing
- reportlab for PDF generation
- pdfplumber and pdf2docx for PDF conversion
- Pillow for image processing

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 
