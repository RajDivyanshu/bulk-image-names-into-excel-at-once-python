# Image Names to Excel Generator

A professional Python script to scan image files from folders and automatically generate a clean, styled Excel report containing all image names and folder details.

This tool is perfect for managing large image collections, organizing app assets, preparing datasets, tracking design libraries, and maintaining structured records of media files.

---

## Features

* Scans all folders and subfolders recursively
* Detects multiple image formats automatically
* Supports:

  * JPG / JPEG
  * PNG
  * WEBP
  * BMP
  * GIF
  * TIFF / TIF
* Generates a professional Excel (.xlsx) file
* Includes serial number for each image
* Captures image name
* Captures folder/category name
* Captures file extension
* Styled Excel sheet with:

  * Colored header row
  * Alternate row background
  * Borders for all cells
  * Auto column width setup
  * Freeze header row
  * Final summary row with total image count

---

## Example Output

### Excel Columns

```text id="m3i9ga"
Sr. No. | Image Name | Folder | Extension
```

### Example Data

```text id="2e1hwp"
1 | bridal_design_01.jpg | Bridal Mehndi | JPG
2 | flower_pattern.png   | Arabic Mehndi | PNG
3 | hand_art.webp        | Festival Design | WEBP
```

### Final Summary Row

```text id="r0c2pn"
Total | 1500
```

---

## Configuration

Edit these values at the top of the script:

```python id="q4f8sm"
ROOT_FOLDER = "Your Image Folder Path"
OUTPUT_FILE = "Your Output Excel File Path"
```

---

## Variable Description

| Variable    | Purpose                                |
| ----------- | -------------------------------------- |
| ROOT_FOLDER | Main folder containing all image files |
| OUTPUT_FILE | Path where Excel file will be saved    |

---

## Excel Styling Included

The generated Excel file includes:

* Professional blue header design
* White bold header text
* Center and left aligned columns
* Light alternate row colors
* Borders on all cells
* Fixed row height
* Frozen top header row
* Total count summary row

This makes the file ready for professional use without manual formatting.

---

## Best Use Cases

Perfect for:

* Mehndi design collections
* Android app asset management
* Wallpaper app resources
* Dataset preparation
* Client delivery reports
* Bulk image documentation
* Design library organization
* Media inventory management

---

## How to Run

```bash id="1jw8gx"
python image_names_to_excel.py
```

---

## Requirements

* Python 3.x
* openpyxl library

Install dependency:

```bash id="t6p1zk"
pip install openpyxl
```

---

## Output Example

The script generates:

```text id="8f4zqy"
name_of_all_designs.xlsx
```

with all image details professionally structured.

---

## Author

Built for fast, clean, and professional image inventory management with Excel automation.
