# Suyuan Auto Report

Automated PDF report generation tool customized for Nantong Suyuan Pathology & Medical Laboratory.

## Overview

This tool automatically generates professional PDF reports from Excel test data for food allergy testing. It supports multiple project types (IgG-F96-1, IgG-F64-1, IgG-F32-1) and creates beautifully formatted reports with patient information, test results, and categorized food allergen data.

## Features

- **Multiple Project Types**: Supports IgG-F96-1 (96 items), IgG-F64-1 (64 items), and IgG-F32-1 (32 items)
- **Automatic Classification**: Categorizes food allergens into groups (meat, dairy, seafood, grains, fruits, vegetables, nuts, seasonings)
- **Allergy Level Assessment**: Automatically calculates and displays allergy levels (normal, mild, moderate, severe)
- **Professional PDF Output**: Generates high-quality PDF reports with custom templates
- **Batch Processing**: Processes multiple patient records from a single Excel file
- **Custom Signatures**: Supports personalized signatures for inspectors and reviewers

## Requirements

- Python 3.7+
- Required Python packages:
  - pandas
  - openpyxl
  - jinja2
- WeasyPrint (included in `bin/` directory for Windows)

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/Lambowmm/suyuan-auto-report.git
   cd suyuan-auto-report
   ```

2. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Preparing Your Data

1. Place your test data Excel file named `TestResult.xlsx` in the root directory
2. The Excel file should follow this column structure:
   - Column 1: (Reserved)
   - Column 2: Test Time
   - Column 3: Project Type (IgG-F96-1, IgG-F64-1, or IgG-F32-1)
   - Column 4: Patient ID
   - Column 5: Patient Name
   - Column 6: Gender
   - Column 7: Age
   - Column 15: Inspector Name
   - Column 16: Reviewer Name
   - Column 17+: Food allergen test results

### Running the Report Generator

**Python Script Mode:**
```bash
python generate_reports.py
```

**Executable Mode (Windows):**
Double-click the `generate_reports.exe` file (if you have built the executable)

### Output

Generated PDF reports will be saved in the `output_reports/` directory with filenames in the format:
```
{PatientName}_{ProjectType}_Report.pdf
```

## Project Structure

```
suyuan-auto-report/
├── bin/                    # Binary executables
│   └── weasyprint.exe     # PDF generation tool
├── images/                # Image assets for reports
│   ├── sign/             # Signature files (*.bmp)
│   ├── Back.jpg
│   ├── CommonSymptoms.png
│   ├── Cover.jpg
│   ├── foodGuide.png
│   ├── levelInformation.png
│   ├── logo.png
│   ├── process.jpg
│   └── replaceFood*.png
├── templates/             # HTML templates for reports
│   ├── IgG-F32-Template.html
│   ├── IgG-F64-Template.html
│   └── IgG-F96-Template.html
├── generate_reports.py    # Main report generation script
├── TestResult.xlsx        # Input data file (not included)
└── output_reports/        # Generated PDF reports (created automatically)
```

## Supported Project Types

| Project Type | Number of Food Items | Template File |
|--------------|---------------------|---------------|
| IgG-F96-1    | 96 items           | IgG-F96-Template.html |
| IgG-F64-1    | 64 items           | IgG-F64-Template.html |
| IgG-F32-1    | 32 items           | IgG-F32-Template.html |

## Allergy Level Thresholds

The tool automatically classifies allergen levels based on concentration values:

- **Normal**: < 50 U/mL
- **Mild**: 50-99 U/mL
- **Moderate**: 100-199 U/mL
- **Severe**: ≥ 200 U/mL

## Food Categories

Food allergens are automatically categorized into:

- 肉类 (Meat)
- 蛋奶类 (Eggs & Dairy)
- 水产类 (Seafood)
- 谷物类 (Grains)
- 水果类 (Fruits)
- 蔬菜类 (Vegetables)
- 坚果类 (Nuts)
- 调味类 (Seasonings)

## Customization

### Adding Custom Signatures

1. Create a BMP image file with the signature
2. Name it `{InspectorName}.bmp` or `{ReviewerName}.bmp`
3. Place it in the `images/sign/` directory
4. The tool will automatically use it when that inspector/reviewer is listed in the Excel file

### Modifying Templates

HTML templates are located in the `templates/` directory. You can customize the report layout by editing these files. They use Jinja2 templating syntax.

## Troubleshooting

### Common Issues

1. **Excel file not found**: Ensure `TestResult.xlsx` is in the root directory
2. **WeasyPrint error**: Verify that `weasyprint.exe` is in the `bin/` directory
3. **Missing columns**: Check that your Excel file has all required columns in the correct positions
4. **Template not found**: Ensure all template files are in the `templates/` directory

### Error Messages

- **"未找到 Excel 文件"**: Excel file not found in the expected location
- **"无法执行命令"**: WeasyPrint executable not found or not working
- **"跳过不支持的项目类型"**: The project type in the Excel file is not supported
- **"模板文件不存在"**: Template file is missing

## Development

### Building from Source

To create a standalone executable:

```bash
pip install pyinstaller
pyinstaller --onefile --add-data "templates;templates" --add-data "images;images" --add-data "bin;bin" generate_reports.py
```

### Code Structure

The main script `generate_reports.py` is organized into several sections:

1. **Constants**: Configuration values and thresholds
2. **Path Configuration**: Dynamic path resolution for different run modes
3. **Data Classification**: Food category mappings
4. **Utility Functions**: Helper functions for data processing
5. **Data Processing**: Functions for extracting and processing Excel data
6. **PDF Generation**: HTML to PDF conversion using WeasyPrint
7. **Main Logic**: Report generation workflow

## License

This project is proprietary software customized for Nantong Suyuan Pathology & Medical Laboratory.

## Author

Created for Nantong Suyuan Pathology & Medical Laboratory

## Support

For issues or questions, please contact the laboratory's IT support team.
