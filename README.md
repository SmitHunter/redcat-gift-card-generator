# Gift Card Generator (GCG)

A powerful, configurable GUI application for generating professional gift cards with barcodes and custom text layouts. Supports batch processing from CSV/Excel files with live preview functionality.

## Features

- **Batch Processing**: Generate hundreds of gift cards from CSV/Excel data files
- **Live Preview**: Real-time preview of gift card design with drag-and-drop positioning
- **Flexible Barcode Generation**: Code128 barcodes with multiple size options
- **Advanced Layout Designer**: Visual positioning controls with live preview
- **Custom Text Layouts**: Configurable text positioning with background options
- **Professional Output**: High-quality PNG images optimized for printing
- **User-Friendly Interface**: Modern dark theme with intuitive controls
- **Configurable Settings**: Customize fonts, sizes, colors, and default positions

## How It Works

1. **Select Background**: Choose your gift card background image (PNG, JPG, etc.)
2. **Import Data**: Load CSV or Excel file with gift card information
3. **Configure Layout**: Position barcodes and text using visual controls
4. **Preview Design**: See live preview of your gift card design
5. **Batch Generate**: Create all gift cards with one click

## Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   python main.py
   ```

## Configuration

Customize the application by editing `config.json`:

```json
{
  "business": {
    "name": "Your Business Name",
    "default_font": "Arial",
    "default_font_size": 12
  },
  "barcode": {
    "format": "Code128",
    "default_size": "Medium",
    "sizes": {
      "Small": [120, 30],
      "Medium": [160, 40],
      "Large": [200, 50],
      "XL": [240, 60]
    }
  },
  "ui": {
    "window_width": 1200,
    "window_height": 800,
    "canvas_width": 400,
    "canvas_height": 250
  },
  "defaults": {
    "barcode_position": "Bottom-Right",
    "text_position": "Bottom-Left",
    "text_background": "White Box",
    "text_alignment": "Left"
  }
}
```

### Configuration Options

#### Business Settings
- `name`: Your business name (displayed in title bar)
- `default_font`: Default font family for text rendering
- `default_font_size`: Base font size for text

#### Barcode Settings
- `format`: Barcode format (currently supports Code128)
- `default_size`: Default barcode size selection
- `sizes`: Custom size definitions for different barcode sizes

#### UI Settings
- `window_width/height`: Application window dimensions
- `canvas_width/height`: Preview canvas dimensions

#### Default Settings
- `barcode_position`: Default barcode placement
- `text_position`: Default text placement
- `text_background`: Default text background style
- `text_alignment`: Default text alignment

## Data File Format

Your CSV or Excel file should contain these columns (names can be customized in the app):

| Column | Description | Example |
|--------|-------------|---------|
| barcode | Barcode data | 1234567890123 |
| member_number | Member/Customer ID | 12345 |
| verification_code | Security/PIN code | ABCD1234 |

### Example CSV:
```csv
barcode,member_number,verification_code
1234567890123,12345,ABCD1234
1234567890124,12346,EFGH5678
1234567890125,12347,IJKL9012
```

## Usage Guide

### 1. File Selection
- **Background Image**: Select your gift card template (PNG, JPG, etc.)
- **Data File**: Choose CSV or Excel file with gift card data

### 2. Column Configuration
Map your data file columns to the required fields:
- **Barcode Column**: Column containing barcode data
- **Member Number**: Customer/member identification
- **Verification Code**: Security or PIN code

### 3. Barcode Configuration
- **Position**: Choose from presets or use custom positioning
- **Size**: Select from Small, Medium, Large, or XL
- **Custom Position**: Fine-tune X/Y coordinates (percentage-based)

### 4. Text Configuration
- **Position**: Choose text placement on the card
- **Background**: Select text background style:
  - None: No background
  - White/Black Box: Solid color backgrounds
  - Semi-Transparent: Adjustable transparency overlays
- **Alignment**: Left, Center, or Right text alignment
- **Scale**: Adjust text size (50% - 200%)

### 5. Output Configuration
- **Output Folder**: Choose where to save generated gift cards
- **File Format**: PNG files with high quality (95% compression)

### 6. Generation Controls
- **Preview Sample**: Generate a preview with sample data
- **Generate All Cards**: Process entire data file

## Output Files

Generated gift cards are saved as:
```
gift_card_{card_serial}_{sequential_number}.png
```

Example: `gift_card_GC001_0001.png`

## Building Executable

Create a standalone executable for distribution:

```bash
pip install pyinstaller
pyinstaller build.spec
```

The executable will be created in the `dist/` folder.

## Advanced Features

### Live Preview
- Real-time preview updates as you adjust settings
- Visual positioning with percentage-based coordinates
- Sample data preview before batch generation

### Flexible Positioning
- Preset positions: Top-Left, Top-Right, Bottom-Left, Bottom-Right, Center
- Custom positioning with precise X/Y coordinates
- Percentage-based positioning for scalability

### Text Backgrounds
- Solid color backgrounds (White/Black boxes)
- Semi-transparent overlays with adjustable opacity
- Automatic text color adjustment for readability

### Batch Processing
- Progress tracking during generation
- Error handling for invalid data
- Detailed status messages and logging

## Troubleshooting

### Common Issues

**"No file selected" errors:**
- Ensure background image and data file are selected
- Check file permissions and accessibility

**"Missing columns" errors:**
- Verify column names match your configuration
- Check CSV/Excel file structure

**Font errors:**
- Application falls back to default font if custom font unavailable
- Install required fonts system-wide for best results

**Memory issues with large batches:**
- Process in smaller batches if memory errors occur
- Close other applications to free up RAM

### Performance Tips

- Use optimized background images (reasonable resolution)
- Process large datasets in batches of 100-500 cards
- Close preview updates during batch generation for faster processing

## Requirements

- Python 3.7+
- CustomTkinter 5.2.0+
- Pillow 10.0.0+
- pandas 2.0.0+
- python-barcode 0.15.1+
- openpyxl 3.1.0+ (for Excel support)

## License

This project is designed for business use in gift card and loyalty program management.

## Support

For issues or feature requests, please check the documentation or create an issue in the repository.