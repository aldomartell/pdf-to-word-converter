# PDF to Word Converter

Convert PDF documents to Word (.docx) format with custom logo and description text. This repository includes two versions: one for simple folder structures and one for nested patient/subfolder structures.

## Features

- ✅ Batch convert multiple PDFs to Word format
- ✅ Add custom logo image at the top of each document (centered)
- ✅ Add description text at the bottom (bold, formatted)
- ✅ Error handling and progress reporting
- ✅ Simple and nested folder structure versions

## Requirements

- Python 3.7+
- pdf2docx
- python-docx

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/pdf-to-word-converter.git
cd pdf-to-word-converter
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

### Version 1: Simple (Flat Folder Structure)

Use `pdf_to_word_simple.py` when all your PDFs are in a single folder.

1. Create an `input_pdfs` folder and place your PDF files in it
2. Place your logo image (e.g., `logo.png`) in the project root
3. Edit `pdf_to_word_simple.py` and update:
   - `PDF_FOLDER` - path to your input PDFs folder
   - `OUTPUT_FOLDER` - where converted files should be saved
   - `LOGO_PATH` - path to your logo image
   - `DESCRIPTION_TEXT` - text to appear at the bottom of each document

4. Run the script:
```bash
python pdf_to_word_simple.py
```

### Version 2: Nested Folders (Patient/Subfolder Structure)

Use `pdf_to_word_nested.py` when PDFs are organized in patient folders or other subfolders.

**Key difference:** This version **preserves folder structure** in the output. If you have:
```
patient_data/
├── Patient_A/
│   ├── Report1.pdf
│   └── Report2.pdf
└── Patient_B/
    └── Report3.pdf
```

The output will be:
```
converted_documents/
├── Patient_A/
│   ├── Report1.docx
│   └── Report2.docx
└── Patient_B/
    └── Report3.docx
```

1. Create a main folder with patient subfolders containing PDFs
2. Place your logo image in the project root
3. Edit `pdf_to_word_nested.py` and update:
   - `PDF_FOLDER` - path to your main folder with subfolders
   - `OUTPUT_FOLDER` - where converted files should be saved
   - `LOGO_PATH` - path to your logo image
   - `DESCRIPTION_TEXT` - text to appear at the bottom

4. Run the script:
```bash
python pdf_to_word_nested.py
```

## Configuration

Both scripts have the same configuration variables at the top:

```python
PDF_FOLDER = 'input_pdfs'              # Input folder path
OUTPUT_FOLDER = 'output_documents'     # Output folder path
LOGO_PATH = 'logo.png'                 # Logo image file
DESCRIPTION_TEXT = 'Your text here'    # Description text
```

## File Formats Supported

- **Input:** PDF files (.pdf)
- **Output:** Word documents (.docx)
- **Logo:** PNG, JPG, GIF, BMP, etc. (any image format supported by python-docx)

## Troubleshooting

**Issue:** "No such file or directory" error
- Make sure your PDF_FOLDER path exists
- Check that LOGO_PATH points to an existing image file
- Verify folder names are spelled correctly

**Issue:** Logo not appearing
- Ensure the LOGO_PATH file exists and is a valid image
- Check that the file path is correct (relative or absolute)

**Issue:** PDF conversion fails
- Some complex PDFs may not convert perfectly
- The script will skip problematic PDFs and continue with others
- Check console output for specific error messages

## Output

The script will:
1. Create the output folder if it doesn't exist
2. Convert each PDF to Word format
3. Add your logo image centered at the top
4. Add your description text in bold at the bottom
5. Print progress updates to the console

Example output:
```
Found 5 PDF files in all subfolders.
Processing patient_data/Patient_A/Report1.pdf...
✓ Successfully converted patient_data/Patient_A/Report1.pdf
Processing patient_data/Patient_A/Report2.pdf...
✓ Successfully converted patient_data/Patient_A/Report2.pdf
...
All PDFs converted successfully with logo centered at the top!
```

