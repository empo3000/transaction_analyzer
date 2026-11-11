#!/usr/bin/env python3
"""
Transaction Analyzer for Household Costs
"""
import pandas as pd
import os
import sys
from datetime import datetime
from pathlib import Path
import argparse

try:
    import inquirer
    from inquirer.themes import GreenPassion
    HAS_INQUIRER = True
except ImportError:
    HAS_INQUIRER = False

def parse_swedish_number(value):
    """Parse Swedish formatted numbers like '-4 200,00' -> -4200.00"""
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(' ', '').replace(',', '.')
    try:
        return float(text)
    except ValueError:
        return None

def parse_swedish_date(date_str):
    """Parse date string in ISO 8601 or Swedish format."""
    if date_str is None or pd.isna(date_str):
        return None
    if isinstance(date_str, datetime):
        return date_str
    text = str(date_str).strip()
    for fmt in ['%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%Y']:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None

def load_allkort_vertical_format(file_path):
    """Load Allkort Mastercard Excel file with unusual vertical/stacked format."""
    import openpyxl
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    # Get all values from column 1
    all_values = []
    for row in range(1, ws.max_row + 1):
        cell = ws.cell(row=row, column=1)
        val = cell.value if cell.value else ''
        all_values.append(str(val).strip())
    
    # Identify blocks by their headers
    blocks = {}
    current_block = None
    
    for i, val in enumerate(all_values):
        if val in ['Köpdatum Inköpsställe', 'Ort', 'Utländskt belopp Valuta', 'Kurs Belopp SEK']:
            current_block = val
            blocks[current_block] = []
        elif val:  # Skip empty strings
            if current_block:
                blocks[current_block].append((i+1, val))
    
    # Extract data from each block
    date_merchant_data = blocks.get('Köpdatum Inköpsställe', [])
    location_data = blocks.get('Ort', [])
    foreign_data = blocks.get('Utländskt belopp Valuta', [])
    sek_data = blocks.get('Kurs Belopp SEK', [])
    
    # Determine number of transactions
    num_transactions = max(len(date_merchant_data), len(location_data), len(foreign_data), len(sek_data))
    
    # Parse each block into transactions
    transactions = []
    
    for i in range(num_transactions):
        tx = {}
        
        # Parse Date and Merchant from Block 1
        if i < len(date_merchant_data):
            _, value = date_merchant_data[i]
            # Split date and merchant (date is first 10 chars: YYYY-MM-DD)
            date_str = value[:10]
            merchant = value[11:] if len(value) > 11 else ''
            tx['Köpdatum'] = parse_swedish_date(date_str)
            tx['Inköpsställe'] = merchant.strip()
        
        # Parse Location from Block 2
        if i < len(location_data):
            _, value = location_data[i]
            tx['Ort'] = value
        
        # Parse Foreign amount, Currency, Rate from Block 3
        if i < len(foreign_data):
            _, value = foreign_data[i]
            # Split by spaces - format: "-4 200,00 INR 0,102090" or "-8,45 EUR 11,22366"
            parts = value.split()
            if len(parts) >= 3:
                # Join amount parts (might have spaces: "-4 200,00")
                amount_str = ' '.join(parts[:-2])
                currency = parts[-2]
                rate_str = parts[-1]
                tx['Utländskt belopp'] = parse_swedish_number(amount_str)
                tx['Valuta'] = currency
                tx['Kurs'] = parse_swedish_number(rate_str)
        
        # Parse SEK amount from Block 4
        if i < len(sek_data):
            _, value = sek_data[i]
            tx['Belopp SEK'] = parse_swedish_number(value)
        
        transactions.append(tx)
    
    # Create DataFrame
    df = pd.DataFrame(transactions)
    return df


def load_excel_file(file_path):
    """Load a single Excel file and parse the data."""
    # First try standard pandas read
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        df.columns = df.columns.str.strip()
        
        # Check if this is the unusual vertical format (single column with merged headers)
        # This format has all data in one column with headers like "Köpdatum Inköpsställe"
        if len(df.columns) == 1:
            col_name = df.columns[0]
            # Check if column name suggests merged headers (contains multiple Swedish column names)
            if ('Köpdatum' in col_name and 'Inköpsställe' in col_name) or \
               ('Ort' in col_name) or \
               ('Utländskt' in col_name and 'belopp' in col_name) or \
               ('Kurs' in col_name and 'Belopp' in col_name and 'SEK' in col_name):
                # This is the vertical format, use special loader
                return load_allkort_vertical_format(file_path)
        
        # Standard format - process columns
        expected_cols = ['Köpdatum', 'Inköpsställe', 'Ort', 'Utländskt belopp', 'Valuta', 'Kurs', 'Belopp SEK']
        if 'Köpdatum' in df.columns:
            df['Köpdatum'] = df['Köpdatum'].apply(parse_swedish_date)
        if 'Belopp SEK' in df.columns:
            df['Belopp SEK'] = df['Belopp SEK'].apply(parse_swedish_number)
        if 'Utländskt belopp' in df.columns:
            df['Utländskt belopp'] = df['Utländskt belopp'].apply(parse_swedish_number)
        if 'Kurs' in df.columns:
            df['Kurs'] = df['Kurs'].apply(parse_swedish_number)
        return df
    except Exception as e:
        # If standard read fails, try the vertical format loader
        print(f"Warning: Standard Excel read failed ({e}). Trying vertical format loader...")
        return load_allkort_vertical_format(file_path)

def load_data(input_path):
    """Load data from a file or directory of Excel files."""
    input_path = Path(input_path)
    all_dfs = []
    if input_path.is_file():
        if input_path.suffix.lower() in ['.xlsx', '.xls']:
            df = load_excel_file(input_path)
            df['_source_file'] = input_path.name
            all_dfs.append(df)
        else:
            print(f"Error: {input_path} is not an Excel file")
            sys.exit(1)
    elif input_path.is_dir():
        excel_files = list(input_path.glob('*.xlsx')) + list(input_path.glob('*.xls'))
        if not excel_files:
            print(f"Error: No Excel files found in {input_path}")
            sys.exit(1)
        for excel_file in excel_files:
            df = load_excel_file(excel_file)
            df['_source_file'] = excel_file.name
            all_dfs.append(df)
            print(f"Loaded: {excel_file.name}")
    else:
        print(f"Error: {input_path} does not exist")
        sys.exit(1)
    combined = pd.concat(all_dfs, ignore_index=True)
    essential_cols = ['Köpdatum', 'Inköpsställe', 'Belopp SEK']
    combined = combined.dropna(subset=[col for col in essential_cols if col in combined.columns])
    return combined

def select_merchants_interactive(transactions):
    """Interactive terminal selection of merchants."""
    merchants = sorted(transactions['Inköpsställe'].unique().tolist())
    if not HAS_INQUIRER:
        print("\n" + "="*60)
        print("Available Merchants (Inköpsställe):")
        print("="*60)
        for i, merchant in enumerate(merchants, 1):
            print(f"  {i}. {merchant}")
        print(f"  0. Select all")
        print(f"  a. Select none")
        print("\nEnter merchant numbers separated by commas (e.g., 1,3,5-7):")
        selection = input("> ").strip()
        if selection.lower() == '0':
            return merchants
        if selection.lower() == 'a':
            return []
        selected = []
        for part in selection.split(','):
            part = part.strip()
            if '-' in part:
                start, end = part.split('-')
                start, end = int(start.strip()), int(end.strip())
                selected.extend(merchants[i-1] for i in range(start, end+1))
            elif part:
                idx = int(part) - 1
                if 0 <= idx < len(merchants):
                    selected.append(merchants[idx])
        return list(set(selected))
    theme = GreenPassion()
    questions = [inquirer.Checkbox('merchants', message="Select merchants:", choices=merchants)]
    answers = inquirer.prompt(questions, theme=theme)
    return answers['merchants'] if answers else []

def filter_by_date_range(transactions, start_date=None, end_date=None):
    """Filter transactions by date range (ISO 8601 format)."""
    if start_date is None and end_date is None:
        return transactions
    if 'Köpdatum' not in transactions.columns:
        print("Warning: No 'Köpdatum' column found. Cannot filter by date.")
        return transactions
    filtered = transactions.copy()
    if start_date:
        start_dt = datetime.fromisoformat(start_date)
        filtered = filtered[filtered['Köpdatum'] >= start_dt]
    if end_date:
        end_dt = datetime.fromisoformat(end_date)
        filtered = filtered[filtered['Köpdatum'] <= end_dt]
    return filtered

def generate_summary(transactions, selected_merchants):
    """Generate summary statistics for selected merchants."""
    if not selected_merchants:
        print("No merchants selected. No summary to generate.")
        return pd.DataFrame()
    filtered = transactions[transactions['Inköpsställe'].isin(selected_merchants)].copy()
    if filtered.empty:
        print("No transactions found for selected merchants.")
        return pd.DataFrame()
    summary_by_merchant = filtered.groupby('Inköpsställe').agg({
        'Belopp SEK': ['count', 'sum', 'mean', 'min', 'max'],
        'Valuta': lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else 'N/A',
        'Ort': lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else 'N/A',
    }).round(2)
    summary_by_merchant.columns = ['Count', 'Total SEK', 'Average SEK', 'Min SEK', 'Max SEK', 'Primary Currency', 'Primary Location']
    summary_by_merchant = summary_by_merchant.sort_values('Total SEK', ascending=False)
    total_all = summary_by_merchant['Total SEK'].sum()
    summary_by_merchant['Percentage'] = (summary_by_merchant['Total SEK'] / total_all * 100).round(2)
    if 'Köpdatum' in filtered.columns:
        filtered['Year'] = filtered['Köpdatum'].dt.year
        filtered['Year-Month'] = filtered['Köpdatum'].dt.to_period('M').astype(str)
        monthly_summary = filtered.groupby('Year-Month')['Belopp SEK'].sum().reset_index()
        monthly_summary.columns = ['Month', 'Total SEK']
        monthly_summary = monthly_summary.sort_values('Month')
        yearly_summary = filtered.groupby('Year')['Belopp SEK'].sum().reset_index()
        yearly_summary.columns = ['Year', 'Total SEK']
        yearly_summary = yearly_summary.sort_values('Year')
    else:
        monthly_summary = pd.DataFrame()
        yearly_summary = pd.DataFrame()
    return {'by_merchant': summary_by_merchant, 'monthly': monthly_summary, 'yearly': yearly_summary, 'all_transactions': filtered}

def create_excel_report(summaries, output_path, selected_merchants, start_date=None, end_date=None):
    """Create an Excel workbook with summary tables and charts."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.chart import PieChart, LineChart, BarChart, Reference
    from openpyxl.utils.dataframe import dataframe_to_rows
    wb = Workbook()
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    ws_summary = wb.create_sheet("Summary by Merchant")
    ws_monthly = wb.create_sheet("Monthly Spending")
    ws_yearly = wb.create_sheet("Yearly Spending")
    ws_data = wb.create_sheet("Transaction Data")
    ws_charts = wb.create_sheet("Charts")
    header_font = Font(name='Calibri', bold=True, size=12)
    header_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    def write_dataframe(ws, df, start_row=1, start_col=1):
        if df.empty:
            return
        for col_idx, col_name in enumerate(df.columns, start=start_col):
            cell = ws.cell(row=start_row, column=col_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='center')
        for row_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=start_row + 1):
            for col_idx, value in enumerate(row, start=start_col):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = border
                if isinstance(value, (int, float)):
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right')
                else:
                    cell.alignment = Alignment(horizontal='left')
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = min(adjusted_width, 50)
    write_dataframe(ws_summary, summaries['by_merchant'])
    if not summaries['monthly'].empty:
        write_dataframe(ws_monthly, summaries['monthly'])
    if not summaries['yearly'].empty:
        write_dataframe(ws_yearly, summaries['yearly'])
    write_dataframe(ws_data, summaries['all_transactions'])
    def create_pie_chart():
        if summaries['by_merchant'].empty:
            return None
        df_sorted = summaries['by_merchant'].sort_values('Total SEK', ascending=False)
        top_n = min(10, len(df_sorted))
        df_top = df_sorted.head(top_n)
        pie = PieChart()
        pie.title = f"Spending by Merchant (Top {top_n})"
        pie.style = 13
        categories = Reference(ws_summary, min_col=1, min_row=2, max_row=2 + len(df_top) - 1, max_col=1)
        values = Reference(ws_summary, min_col=2, min_row=2, max_row=2 + len(df_top) - 1, max_col=2)
        pie.add_data(values, titles_from_data=True)
        pie.set_categories(categories)
        return pie
    def create_line_chart():
        if summaries['monthly'].empty:
            return None
        line = LineChart()
        line.title = "Monthly Spending Over Time"
        line.style = 13
        line.y_axis.title = "Amount (SEK)"
        line.x_axis.title = "Month"
        months = Reference(ws_monthly, min_col=1, min_row=2, max_row=2 + len(summaries['monthly']) - 1, max_col=1)
        values = Reference(ws_monthly, min_col=2, min_row=2, max_row=2 + len(summaries['monthly']) - 1, max_col=2)
        line.add_data(values, titles_from_data=True)
        line.set_categories(months)
        line.y_axis.major_gridlines = True
        return line
    def create_bar_chart():
        if summaries['by_merchant'].empty:
            return None
        df_sorted = summaries['by_merchant'].sort_values('Total SEK', ascending=True)
        top_n = min(15, len(df_sorted))
        df_top = df_sorted.tail(top_n)
        bar = BarChart()
        bar.title = f"Spending by Merchant (Top {top_n})"
        bar.style = 13
        bar.y_axis.title = "Amount (SEK)"
        bar.x_axis.title = "Merchant"
        merchants_ref = Reference(ws_summary, min_col=1, min_row=2, max_row=2 + len(df_top) - 1, max_col=1)
        values = Reference(ws_summary, min_col=2, min_row=2, max_row=2 + len(df_top) - 1, max_col=2)
        bar.add_data(values, titles_from_data=True)
        bar.set_categories(merchants_ref)
        bar.x_axis.label_rotation = 45
        return bar
    pie_chart = create_pie_chart()
    line_chart = create_line_chart()
    bar_chart = create_bar_chart()
    if pie_chart:
        ws_charts.add_chart(pie_chart, "A1")
    if bar_chart:
        ws_charts.add_chart(bar_chart, "A20")
    if line_chart:
        ws_charts.add_chart(line_chart, "A40")
    ws_info = wb.create_sheet("Info")
    ws_info.append(["Transaction Analysis Report"])
    ws_info.append([])
    ws_info.append(["Generated:", datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
    ws_info.append(["Selected Merchants:", len(selected_merchants)])
    ws_info.append(["Date Range:", f"{start_date or 'N/A'} to {end_date or 'N/A'}"])
    ws_info.append(["Total Transactions:", len(summaries['all_transactions'])])
    if not summaries['by_merchant'].empty:
        total_sek = summaries['by_merchant']['Total SEK'].sum()
        ws_info.append(["Total Amount (SEK):", f"{total_sek:,.2f}"])
    ws_info.column_dimensions['A'].width = 25
    ws_info.column_dimensions['B'].width = 50
    wb.save(output_path)
    print(f"\nReport saved to: {output_path}")
    return output_path

def main():
    parser = argparse.ArgumentParser(description='Analyze transaction data and generate household cost reports.')
    parser.add_argument('input', type=str, help='Path to Excel file or directory containing Excel files')
    parser.add_argument('--output', '-o', type=str, default='transaction_report.xlsx', help='Output Excel file path')
    parser.add_argument('--start', '-s', type=str, default=None, help='Start date (ISO 8601: YYYY-MM-DD)')
    parser.add_argument('--end', '-e', type=str, default=None, help='End date (ISO 8601: YYYY-MM-DD)')
    parser.add_argument('--all', '-a', action='store_true', help='Include all merchants')
    parser.add_argument('--merchants', '-m', type=str, default=None, help='Comma-separated list of merchant names')
    args = parser.parse_args()
    print("Loading transaction data...")
    transactions = load_data(args.input)
    print(f"Loaded {len(transactions)} transactions")
    if args.start or args.end:
        transactions = filter_by_date_range(transactions, args.start, args.end)
        print(f"Filtered to {len(transactions)} transactions in date range")
    if args.all:
        selected_merchants = transactions['Inköpsställe'].unique().tolist()
        print(f"Including all {len(selected_merchants)} merchants")
    elif args.merchants:
        selected_merchants = [m.strip() for m in args.merchants.split(',')]
        valid_merchants = transactions['Inköpsställe'].unique().tolist()
        invalid = [m for m in selected_merchants if m not in valid_merchants]
        if invalid:
            print(f"Warning: Merchants not found: {invalid}")
        selected_merchants = [m for m in selected_merchants if m in valid_merchants]
        print(f"Including {len(selected_merchants)} specified merchants")
    else:
        selected_merchants = select_merchants_interactive(transactions)
        if not selected_merchants:
            print("No merchants selected. Exiting.")
            sys.exit(0)
        print(f"\nSelected {len(selected_merchants)} merchants:")
        for m in selected_merchants:
            print(f"  - {m}")
    print("\nGenerating summary...")
    summaries = generate_summary(transactions, selected_merchants)
    if summaries['by_merchant'].empty:
        print("No data to summarize. Exiting.")
        sys.exit(0)
    print("\n" + "="*70)
    print("SUMMARY BY MERCHANT")
    print("="*70)
    print(summaries['by_merchant'].to_string())
    if not summaries['monthly'].empty:
        print("\n" + "="*70)
        print("MONTHLY SPENDING")
        print("="*70)
        print(summaries['monthly'].to_string())
    if not summaries['yearly'].empty:
        print("\n" + "="*70)
        print("YEARLY SPENDING")
        print("="*70)
        print(summaries['yearly'].to_string())
    print("\nCreating Excel report...")
    create_excel_report(summaries, args.output, selected_merchants, args.start, args.end)
    print("\nDone!")

if __name__ == '__main__':
    main()