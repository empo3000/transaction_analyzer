# Transaction Analyzer

A powerful command-line tool for analyzing household transaction data from bank and credit card statements. Parse Excel files, filter transactions by merchant and date, and generate comprehensive reports with summary statistics and visual charts.

## Features

- **Multi-format Support**: Parse standard Excel files and unusual vertical/stacked formats (e.g., Allkort Mastercard statements)
- **Swedish Localization**: Built-in support for Swedish number formats (`1 234,56`) and date formats (`YYYY-MM-DD`, `DD-MM-YYYY`)
- **Flexible Filtering**: Filter transactions by merchant names, date ranges, or both
- **Interactive Selection**: Choose merchants interactively or specify them via command line
- **Comprehensive Reporting**: Generate detailed Excel reports with:
  - Summary statistics by merchant (count, total, average, min, max, percentage)
  - Monthly spending trends
  - Yearly spending summaries
  - Visual charts (pie charts, bar charts, line charts)
- **Batch Processing**: Process single files or entire directories of Excel files

## Installation

### Prerequisites

- Python 3.7 or higher
- pip (Python package manager)

### Quick Install

```bash
# Clone the repository
git clone https://github.com/your-username/transaction_analyzer.git
cd transaction_analyzer

# Create a virtual environment (recommended)
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Install as a Package

```bash
pip install -e .
```

## Requirements

The project dependencies are listed in `requirements.txt`:

- `pandas>=2.0.0` - Data manipulation and analysis
- `openpyxl>=3.0.0` - Excel file reading and writing
- `inquirer>=3.0.0` - Interactive terminal prompts (optional)

Install them with:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
# Analyze a single Excel file
python transaction_analyzer.py path/to/transactions.xlsx

# Analyze all Excel files in a directory
python transaction_analyzer.py path/to/transactions_folder/

# Output to a custom file
python transaction_analyzer.py transactions.xlsx --output my_report.xlsx
```

### Command Line Options

```
Usage: transaction_analyzer.py [OPTIONS] INPUT

Arguments:
  INPUT    Path to Excel file or directory containing Excel files

Options:
  -o, --output FILE    Output Excel file path (default: transaction_report.xlsx)
  -s, --start DATE     Start date (ISO 8601: YYYY-MM-DD)
  -e, --end DATE       End date (ISO 8601: YYYY-MM-DD)
  -a, --all            Include all merchants (skip selection)
  -m, --merchants TEXT Comma-separated list of merchant names to include
  -h, --help           Show this help message
```

### Examples

1. **Analyze all transactions from a specific date range:**
   ```bash
   python transaction_analyzer.py transactions.xlsx --start 2024-01-01 --end 2024-12-31
   ```

2. **Include only specific merchants:**
   ```bash
   python transaction_analyzer.py transactions.xlsx --merchants "ICA,Systembolaget,Lidl"
   ```

3. **Analyze all files in a directory with all merchants:**
   ```bash
   python transaction_analyzer.py data/ --all --output full_report.xlsx
   ```

4. **Date range with specific merchants:**
   ```bash
   python transaction_analyzer.py data/ --start 2024-01-01 --end 2024-06-30 --merchants "ICA,Coop"
   ```

5. **Interactive merchant selection (default behavior):**
   ```bash
   python transaction_analyzer.py transactions.xlsx
   # Will prompt you to select merchants from a list
   ```

## Input File Formats

### Supported Formats

The tool supports two main Excel file formats:

#### 1. Standard Excel Format

Files with normal column headers like:
- `Köpdatum` (Purchase Date)
- `Inköpsställe` (Merchant)
- `Ort` (Location)
- `Utländskt belopp` (Foreign Amount)
- `Valuta` (Currency)
- `Kurs` (Exchange Rate)
- `Belopp SEK` (Amount in SEK)

#### 2. Vertical/Stacked Format (Allkort Mastercard)

Unusual format where all data is in a single column with merged headers:
- `Köpdatum Inköpsställe` (Date and Merchant combined)
- `Ort` (Location)
- `Utländskt belopp Valuta` (Foreign Amount and Currency)
- `Kurs Belopp SEK` (Exchange Rate and SEK Amount)

The tool automatically detects and handles this format.

### Sample Data

The repository includes a sample data file:
- `data/Allkort_Mastercard_2026-11-10.xlsx` - Example of the vertical format

## Output Report

The generated Excel report (`transaction_report.xlsx` by default) contains multiple sheets:

### 1. Summary by Merchant

Comprehensive statistics for each selected merchant:
- **Count**: Number of transactions
- **Total SEK**: Total amount spent (in Swedish Krona)
- **Average SEK**: Average transaction amount
- **Min SEK**: Minimum transaction amount
- **Max SEK**: Maximum transaction amount
- **Primary Currency**: Most common currency for this merchant
- **Primary Location**: Most common location for this merchant
- **Percentage**: Percentage of total spending

### 2. Monthly Spending

Monthly aggregation showing spending trends over time:
- **Month**: Year and month (YYYY-MM format)
- **Total SEK**: Total spending for that month

### 3. Yearly Spending

Yearly aggregation for multi-year analysis:
- **Year**: Year
- **Total SEK**: Total spending for that year

### 4. Transaction Data

Raw transaction data for the selected merchants and date range, including all columns from the source files.

### 5. Charts

Visual representations of your spending:
- **Pie Chart**: Spending distribution by merchant (top 10)
- **Bar Chart**: Spending by merchant (top 15)
- **Line Chart**: Monthly spending over time

### 6. Info

Metadata about the report:
- Generation timestamp
- Number of selected merchants
- Date range used
- Total number of transactions
- Total amount spent

## Data Parsing

### Swedish Number Format

The tool automatically parses Swedish-formatted numbers:
- `1 234,56` → `1234.56`
- `-4 200,00` → `-4200.00`
- `8,45` → `8.45`

### Swedish Date Format

Supported date formats:
- `YYYY-MM-DD` (ISO 8601)
- `DD-MM-YYYY`
- `YYYY/MM/DD`
- `DD/MM/YYYY`

## Interactive Mode

When no `--merchants` or `--all` option is provided, the tool launches an interactive selection:

```
Available Merchants (Inköpsställe):
============================================================
  1. ICA
  2. Coop
  3. Lidl
  4. Systembolaget
  5. Apoteket
  0. Select all
  a. Select none

Enter merchant numbers separated by commas (e.g., 1,3,5-7):
> 1,3,5
```

You can select:
- Individual merchants by number (e.g., `1,3,5`)
- Ranges of merchants (e.g., `1-5`)
- All merchants (`0`)
- No merchants (`a`)

## Use Cases

### Household Budget Analysis

Track spending across different categories:
```bash
python transaction_analyzer.py bank_statements/ --start 2024-01-01 --end 2024-12-31
```

### Monthly Expense Tracking

Monitor monthly spending trends:
```bash
python transaction_analyzer.py transactions.xlsx --output monthly_report.xlsx
```

### Merchant-Specific Analysis

Analyze spending at specific stores:
```bash
python transaction_analyzer.py data/ --merchants "ICA,Coop,Willys" --start 2024-01-01
```

### Year-Over-Year Comparison

Compare spending across years:
```bash
python transaction_analyzer.py all_transactions.xlsx --output yearly_comparison.xlsx
```

## Directory Structure

```
transaction_analyzer/
├── transaction_analyzer.py  # Main application script
├── README.md                # This documentation
├── requirements.txt         # Python dependencies
├── data/                    # Sample data files
│   └── Allkort_Mastercard_2026-11-10.xlsx
└── transaction_report.xlsx  # Sample output report (if generated)
```

## Troubleshooting

### Common Issues

1. **ImportError: No module named 'pandas'**
   ```bash
   pip install -r requirements.txt
   ```

2. **File not found or not an Excel file**
   - Ensure the path is correct
   - Verify the file has `.xlsx` or `.xls` extension

3. **Date parsing errors**
   - Ensure dates in your file match one of the supported formats
   - The tool will skip transactions with unparseable dates

4. **Inquirer not available**
   - The tool falls back to standard input if inquirer is not installed
   - For full functionality: `pip install inquirer`

5. **Vertical format detection**
   - If your file uses the vertical format but isn't detected, ensure it has the expected headers
   - You can force vertical format parsing by modifying the `load_excel_file` function

### Debug Mode

For verbose output, modify the script to add debug prints or run with:
```bash
python -v transaction_analyzer.py your_file.xlsx
```

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

### Development Setup

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Running Tests

To be added. Please include tests with your contributions.

## License

This project is open source and available under the [MIT License](LICENSE).

## Contact

For questions or support, please open an issue on the GitHub repository.

---

**Transaction Analyzer** - Simplifying household expense analysis since 2024.
