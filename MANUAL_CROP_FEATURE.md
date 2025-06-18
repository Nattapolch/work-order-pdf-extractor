# Manual Crop Feature

## Overview

The Manual Crop feature provides an intuitive drag-and-drop interface for selecting crop regions on PDF documents. Instead of manually entering coordinates, users can visually select the exact area they want to extract from their PDFs.

## Features

### 🎯 Visual Crop Selection
- **Drag-to-Select**: Click and drag on the PDF preview to select crop region
- **Real-time Selection**: Red dashed rectangle shows current selection
- **Scrollable Canvas**: Navigate large PDFs with scrollbars

### 👁️ Live Preview
- **Instant Feedback**: See cropped result immediately as you select
- **Maintains Aspect Ratio**: Preview scales correctly
- **Error Handling**: Shows "Selection too small" for invalid regions

### 📊 Coordinate Display
- **Absolute Coordinates**: Pixel positions on original image
- **Relative Coordinates**: 0.0-1.0 values used for processing
- **Real-time Updates**: Coordinates update as you drag

### ⚙️ Easy Application
- **Apply to All PDFs**: Selected region applies to entire batch
- **Config Integration**: Saves settings to config.json
- **UI Synchronization**: Updates other tabs automatically

## How to Use

### 1. Access the Feature
- Open the application
- Navigate to the **"Manual Crop"** tab

### 2. Load a PDF
- Click **"Refresh List"** to scan for PDF files
- Select a PDF from the dropdown menu
- Click **"Load PDF"** to display it

### 3. Select Crop Region
- **Click and drag** on the PDF preview to select area
- Watch the **red dashed rectangle** show your selection
- See **real-time preview** on the right side
- Check **coordinates** in the text box below

### 4. Apply Settings
- Click **"Apply Crop Settings"** to save the region
- Settings are automatically saved to config.json
- All future PDF processing will use this crop region

### 5. Reset if Needed
- Click **"Reset Selection"** to clear current selection
- Start over with a new drag selection

## Technical Details

### Coordinate System
- **Display Coordinates**: Canvas pixels (scaled for display)
- **Original Coordinates**: Actual image pixels
- **Relative Coordinates**: 0.0-1.0 range (saved to config)

### Image Processing
- **PDF Conversion**: Uses pdf2image with PyMuPDF fallback
- **Scaling**: Maintains aspect ratio for preview
- **Quality**: DPI optimized for performance vs quality

### Integration
- **Config Sync**: Updates crop_x1, crop_y1, crop_x2, crop_y2 in config
- **UI Updates**: Synchronizes with existing crop coordinate inputs
- **Batch Processing**: All PDFs use the same crop region

## Error Handling

### Common Issues
- **"No PDF files found"**: Check PDF folder path in settings
- **"Failed to load PDF"**: Ensure PDF is valid and not corrupted
- **"Selection too small"**: Make larger selection (minimum 10x10 pixels)
- **"Preview error"**: May occur with very large or small selections

### Troubleshooting
1. **PDF won't load**: Verify file exists and is accessible
2. **No preview**: Check if PDF has content on first page
3. **Coordinates wrong**: Ensure proper drag completion (release mouse)
4. **Settings not saved**: Check write permissions for config.json

## Workflow Integration

### Before Manual Crop
1. Set up API keys in Settings tab
2. Configure folder paths
3. Test with a sample PDF

### After Manual Crop
1. Crop settings automatically apply to all processing
2. Run batch processing as normal
3. All PDFs will use the selected crop region

## Benefits

### ✅ Improved Accuracy
- Visual selection eliminates guesswork
- Precise coordinate specification
- Real-time feedback prevents errors

### ✅ User Experience
- Intuitive drag-and-drop interface
- No need to calculate coordinates manually
- Immediate visual confirmation

### ✅ Efficiency
- Fast crop region selection
- One-time setup for entire batch
- Integrated with existing workflow

## Example Workflow

```
1. Open Manual Crop tab
2. Refresh PDF list
3. Select "sample-workorder.pdf"
4. Load PDF - see full page preview
5. Drag rectangle around work order number area
6. See cropped preview appear instantly
7. Check coordinates: (0.125, 0.050) to (0.375, 0.200)
8. Apply settings - saved to config
9. Go to Processing tab
10. Run batch processing - all PDFs use new crop region
```

## Technical Architecture

### UI Components
- **Canvas**: Scrollable PDF display with mouse event binding
- **Selection Rectangle**: Canvas rectangle with dashed red outline
- **Preview Label**: Image widget showing cropped result
- **Coordinates Text**: Read-only text display for coordinate info

### Event Handling
- **Button-1**: Start selection (mouse down)
- **B1-Motion**: Update selection (mouse drag)
- **ButtonRelease-1**: Finalize selection (mouse up)

### Scaling System
- **Canvas Scale**: PDF→Display scaling factor
- **Preview Scale**: Crop→Preview scaling factor
- **Coordinate Conversion**: Display→Original→Relative coordinates

This feature transforms the manual coordinate entry process into an intuitive visual experience, making PDF crop region selection accessible to all users regardless of technical expertise.