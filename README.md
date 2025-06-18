# Work Order PDF Extractor

A Python GUI application that processes work order PDFs, extracts work order numbers and equipment numbers using OpenAI's GPT-4 Vision API, and organizes files based on CSV reference data.

## Features

- **PDF Processing**: Converts PDF files to images and crops specific regions
- **AI Text Extraction**: Uses OpenAI GPT-4o-mini to extract work order numbers and equipment numbers
- **Intelligent File Organization**: Renames matching files and moves non-matching files to a separate folder
- **Configurable Crop Area**: User-defined cropping coordinates (default: top-left 1/4 of PDF)
- **GUI Interface**: Easy-to-use tkinter-based interface with tabs for settings, processing, and logs
- **Progress Tracking**: Real-time progress updates and detailed logging

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

2. For PDF processing, you may also need to install poppler-utils:
   - **Windows**: Download from [poppler for Windows](https://blog.alivate.com.au/poppler-windows/)
   - **macOS**: `brew install poppler`
   - **Linux**: `sudo apt-get install poppler-utils` or `sudo yum install poppler-utils`

## Usage

1. Run the application:
```bash
python work_order_extractor.py
```

2. Configure settings in the **Settings** tab:
   - Enter your OpenAI API key
   - Adjust crop coordinates if needed (default is top-left 1/4 of PDF)
   - Set file paths for PDF folder and CSV reference file

3. Test your crop settings using the **Test Crop** button in the **Process PDFs** tab

4. Click **Start Processing** to begin processing all PDFs

## File Structure

```
WOExtractor/
├── work_order_extractor.py     # Main application
├── requirements.txt            # Python dependencies
├── workOrderPDF/              # Folder containing PDF files to process
├── workOrderRef/              # Folder containing CSV reference file
│   └── MCAN_work_inprogress.csv
├── not_match/                 # Folder for non-matching PDFs
├── exampleCroppedImage/       # Example cropped images
└── config.json               # Saved configuration (created automatically)
```

## How It Works

1. **PDF to Image**: Each PDF is converted to an image (first page only)
2. **Cropping**: The specified region is cropped from the image
3. **AI Extraction**: The cropped image is sent to OpenAI GPT-4o-mini with a prompt to extract:
   - Work Order Number (8 digits after "Work Order No.")
   - Equipment Number
4. **Matching**: Extracted work order number is compared against the CSV reference file
5. **File Organization**:
   - **Match found**: File renamed to `CS-{WorkOrderNumber}-{EquipmentNumber}.pdf`
   - **No match**: File moved to `not_match/` folder

## Configuration

The application saves settings in `config.json`. Key settings include:

- `openai_api_key`: Your OpenAI API key
- `crop_x1`, `crop_y1`, `crop_x2`, `crop_y2`: Crop coordinates (0-1 relative to image size)
- `pdf_folder`: Path to folder containing PDFs to process
- `ref_csv_file`: Path to CSV reference file
- `not_match_folder`: Path to folder for non-matching files

## CSV Reference Format

The CSV file should have work order numbers listed one per line under an "Order" column header.

## API Requirements

- OpenAI API key with access to GPT-4o-mini model
- Sufficient API credits for image processing

## Error Handling

- Comprehensive logging to both file and GUI
- Graceful handling of PDF conversion failures
- API error recovery
- File operation error handling

## Troubleshooting

1. **PDF conversion fails**: Ensure poppler-utils is installed
2. **API errors**: Check your OpenAI API key and account credits
3. **File permission errors**: Ensure the application has write permissions to the target folders
4. **Missing dependencies**: Run `pip install -r requirements.txt`

## License

This project is provided as-is for internal use.