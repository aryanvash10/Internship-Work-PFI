import pandas as pd
import requests
import os
from datetime import datetime, timedelta
import numpy as np
import time
import calendar

def download_excel_file(url, save_path, max_retries=3):
    """
    Download Excel file from URL and save it locally with retry mechanism
    
    Args:
        url (str): URL of the Excel file
        save_path (str): Path where to save the file
        max_retries (int): Maximum number of retry attempts
    
    Returns:
        bool: True if download successful, False otherwise
    """
    for attempt in range(max_retries):
        try:
            print(f"    Downloading attempt {attempt + 1}/{max_retries}: {os.path.basename(save_path)}")
            response = requests.get(url, timeout=30)
            response.raise_for_status()  # Raises an HTTPError for bad responses
            
            with open(save_path, 'wb') as file:
                file.write(response.content)
            print(f"    ✓ Downloaded: {os.path.basename(save_path)}")
            return True
            
        except requests.exceptions.RequestException as e:
            print(f"    ✗ Attempt {attempt + 1} failed: {str(e)}")
            if attempt < max_retries - 1:
                time.sleep(2)  # Wait 2 seconds before retry
            continue
        except Exception as e:
            print(f"    ✗ Unexpected error: {str(e)}")
            break
    
    print(f"    ✗ Failed to download after {max_retries} attempts")
    return False

def extract_date_from_filename(filename):
    """
    Extract date from filename like 'capacity2-Northern-2025-06.xls'
    
    Args:
        filename (str): Filename to extract date from
    
    Returns:
        str: Date in DD-MM-YYYY format
    """
    try:
        # Split filename and extract year and month
        parts = filename.split('-')
        year = parts[2]
        month = parts[3].split('.')[0]  # Remove .xls extension
        
        # Create date string (assuming last day of month)
        if month in ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']:
            # Get actual last day of the month
            last_day = calendar.monthrange(int(year), int(month))[1]
            return f"{last_day:02d}-{month}-{year}"
    except Exception as e:
        print(f"Error extracting date from filename: {str(e)}")
        return "01-01-2025"  # Default date

def extract_region_from_filename(filename):
    """
    Extract region from filename like 'capacity2-Northern-2025-06.xls'
    
    Args:
        filename (str): Filename to extract region from
    
    Returns:
        str: Region name
    """
    try:
        parts = filename.split('-')
        return parts[1]  # Northern, Eastern, Western, Southern, etc.
    except Exception as e:
        print(f"Error extracting region from filename: {str(e)}")
        return "Unknown"

def process_excel_file(file_path, date, region):
    """
    Process a single Excel file and convert to required format
    
    Args:
        file_path (str): Path to Excel file
        date (str): Date string
        region (str): Region name
    
    Returns:
        pandas.DataFrame: Processed data
    """
    try:
        # Read Excel file
        df = pd.read_excel(file_path, header=None)
        
        print(f"    Processing {region} region file...")
        print(f"      File shape: {df.shape}")
        
        # Initialize column positions
        coal_col = None
        lignite_col = None
        gas_col = None
        diesel_col = None
        nuclear_col = None
        hydro_col = None
        res_col = None
        total_col = None
        state_col = 1  # Default state column
        sector_col = None  # To be detected
        
        header_rows = set()
        
        # Search for header labels across first 15 rows
        for idx in range(min(15, len(df))):
            row = df.iloc[idx]
            for col_idx, cell in enumerate(row):
                if pd.isna(cell):
                    continue
                cell_str = str(cell).upper().strip()
                
                if cell_str == 'STATE':
                    state_col = col_idx
                    header_rows.add(idx)
                
                if 'OWNERSHIP/SECTOR' in cell_str or cell_str == 'SECTOR':
                    sector_col = col_idx
                    header_rows.add(idx)
                
                if 'COAL' in cell_str and 'LIGNITE' not in cell_str:
                    if coal_col is None:
                        coal_col = col_idx
                        header_rows.add(idx)
                
                if 'LIGNITE' in cell_str:
                    if lignite_col is None:
                        lignite_col = col_idx
                        header_rows.add(idx)
                
                if 'GAS' in cell_str:
                    if gas_col is None:
                        gas_col = col_idx
                        header_rows.add(idx)
                
                if 'DIESEL' in cell_str:
                    if diesel_col is None:
                        diesel_col = col_idx
                        header_rows.add(idx)
                
                if 'NUCLEAR' in cell_str:
                    if nuclear_col is None:
                        nuclear_col = col_idx
                        header_rows.add(idx)
                
                if 'HYDRO' in cell_str:
                    if hydro_col is None:
                        hydro_col = col_idx
                        header_rows.add(idx)
                
                if 'RES' in cell_str :
                    if res_col is None:
                        res_col = col_idx
                        header_rows.add(idx)
                
                if 'GRAND' in cell_str or ('TOTAL' in cell_str and col_idx > 8):
                    if total_col is None:
                        total_col = col_idx
                        header_rows.add(idx)
        
        # Fallback for sector_col based on format (lignite presence indicates newer format)
        if sector_col is None:
            sector_col = 2 if lignite_col is None else 3
        
        # Determine start row for data
        start_row = max(header_rows) + 1 if header_rows else 5
        
        # Initialize result list
        result_data = []
        current_state = None
        
        for idx in range(start_row, len(df)):
            row = df.iloc[idx]
            
            # Get state and sector using detected columns
            state_or_number = str(row.iloc[state_col]).strip() if pd.notna(row.iloc[state_col]) else "nan"
            sector = str(row.iloc[sector_col]).strip() if pd.notna(row.iloc[sector_col]) else "nan"
            
            # Skip empty rows
            if state_or_number == 'nan' and sector == 'nan':
                continue
            
            # Check if this is a state name row (no sector info)
            if sector == 'nan' or sector == '' or sector == 'None':
                if 'Total of' in state_or_number or 'TOTAL OF' in state_or_number.upper():
                    # This is a total row for the current state, skip it
                    continue
                elif state_or_number != 'nan' and not state_or_number.isdigit():
                    # Skip regional total rows and invalid entries like URLs
                    if state_or_number.upper() in ['NORTHERN', 'EASTERN', 'WESTERN', 'SOUTHERN', 'NORTH EASTERN'] or state_or_number.startswith('http'):
                        current_state = None
                        continue
                    # This is a new state name
                    current_state = state_or_number
                    continue
            
            # Process sector data rows
            if current_state and sector.upper() in ['STATE SECTOR', 'PVT SECTOR', 'CENTRAL SECTOR']:
                try:
                    # Map sectors
                    sector_mapping = {
                        'STATE SECTOR': 'State',
                        'PVT SECTOR': 'Private', 
                        'CENTRAL SECTOR': 'Central'
                    }
                    
                    def safe_numeric_convert(value):
                        """Safely convert value to numeric, handling various edge cases"""
                        if pd.isna(value):
                            return 0
                        val_str = str(value).strip()
                        if val_str in ['nan', '', 'None']:
                            return 0
                        try:
                            return float(val_str)
                        except:
                            return 0
                    
                    # Extract values from identified column positions, default to 0 if column is missing
                    coal = safe_numeric_convert(row.iloc[coal_col] if coal_col is not None and coal_col < len(row) else 0)
                    lignite = safe_numeric_convert(row.iloc[lignite_col] if lignite_col is not None and lignite_col < len(row) else 0)
                    gas = safe_numeric_convert(row.iloc[gas_col] if gas_col is not None and gas_col < len(row) else 0)
                    diesel = safe_numeric_convert(row.iloc[diesel_col] if diesel_col is not None and diesel_col < len(row) else 0)
                    nuclear = safe_numeric_convert(row.iloc[nuclear_col] if nuclear_col is not None and nuclear_col < len(row) else 0)
                    hydro = safe_numeric_convert(row.iloc[hydro_col] if hydro_col is not None and hydro_col < len(row) else 0)
                    res = safe_numeric_convert(row.iloc[res_col] if res_col is not None and res_col < len(row) else 0)
                    grand_total = safe_numeric_convert(row.iloc[total_col] if total_col is not None and total_col < len(row) else 0)
                    
                    # Always calculate thermal total as sum of components
                    thermal_total = coal + lignite + gas + diesel
                    
                    # Calculate grand total if not available or zero
                    if grand_total == 0:
                        grand_total = thermal_total + nuclear + hydro + res
                    
                    result_data.append({
                        'Date': date,
                        'Region': region,
                        'State': current_state,
                        'Sector': sector_mapping[sector.upper()],
                        'Coal': coal,
                        'Lignite': lignite,
                        'Gas': gas,
                        'Diesel': diesel,
                        'Thermal Total': thermal_total,
                        'Nuclear': nuclear,
                        'Hydro': hydro,
                        'RES': res,
                        'Total': grand_total
                    })
                    
                except Exception as e:
                    continue
        
        processed_df = pd.DataFrame(result_data)
        
        # Remove any rows where State is same as Region (regional totals)
        if not processed_df.empty:
            before_filter = len(processed_df)
            processed_df = processed_df[processed_df['State'] != processed_df['Region']]
            after_filter = len(processed_df)
        
        print(f"      Extracted {len(processed_df)} records from {region}")
        return processed_df
        
    except Exception as e:
        print(f"    Error processing Excel file {file_path}: {str(e)}")
        return pd.DataFrame()

def generate_urls_for_month_year(year, month):
    """
    Generate URLs for all regions for a specific month and year
    
    Args:
        year (str): Year (e.g., '2025')
        month (str): Month in uppercase (e.g., 'JUN')
    
    Returns:
        list: List of tuples (url, filename, region)
    """
    # All 5 regions including North Eastern
    regions = ['Northern', 'Eastern', 'Western', 'Southern', 'North Eastern']
    
    month_num = {
        'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'MAY': '05', 'JUN': '06',
        'JUL': '07', 'AUG': '08', 'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12'
    }
    
    urls = []
    for region in regions:
        filename = f"capacity2-{region}-{year}-{month_num[month]}.xls"
        url = f"https://npp.gov.in/public-reports/cea/monthly/installcap/{year}/{month}/{filename}"
        urls.append((url, filename, region))
    
    return urls

def generate_date_range(start_year, start_month, end_year, end_month):
    """
    Generate list of (year, month) tuples for the given date range
    
    Args:
        start_year (int): Starting year
        start_month (int): Starting month (1-12)
        end_year (int): Ending year
        end_month (int): Ending month (1-12)
    
    Returns:
        list: List of (year, month_name) tuples
    """
    months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 
              'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    
    date_range = []
    current_year = start_year
    current_month = start_month
    
    while (current_year < end_year) or (current_year == end_year and current_month <= end_month):
        date_range.append((str(current_year), months[current_month - 1]))
        
        current_month += 1
        if current_month > 12:
            current_month = 1
            current_year += 1
    
    return date_range

def check_data_availability(year, month):
    """
    Check if data is available for a specific month/year by trying to download one file
    
    Args:
        year (str): Year
        month (str): Month abbreviation
    
    Returns:
        bool: True if data is available, False otherwise
    """
    # Test with Northern region
    month_num = {
        'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'MAY': '05', 'JUN': '06',
        'JUL': '07', 'AUG': '08', 'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12'
    }
    
    test_filename = f"capacity2-Northern-{year}-{month_num[month]}.xls"
    test_url = f"https://npp.gov.in/public-reports/cea/monthly/installcap/{year}/{month}/{test_filename}"
    
    try:
        response = requests.head(test_url, timeout=10)
        return response.status_code == 200
    except:
        return False

def process_month_data(year, month, work_dir):
    """
    Process data for a specific month and year
    
    Args:
        year (str): Year
        month (str): Month abbreviation
        work_dir (str): Working directory
    
    Returns:
        pandas.DataFrame: Processed data for the month
    """
    print(f"\n{'='*60}")
    print(f"Processing data for {month} {year}")
    print(f"{'='*60}")
    
    # Check if data is available
    if not check_data_availability(year, month):
        print(f"⚠️  Data not available for {month} {year} - skipping")
        return pd.DataFrame()
    
    # Generate URLs for all regions
    urls = generate_urls_for_month_year(year, month)
    
    # Initialize monthly dataframe
    monthly_data = pd.DataFrame()
    
    # Process each region
    successful_downloads = 0
    for url, filename, region in urls:
        print(f"\n  Processing {region} region...")
        
        # Download file
        file_path = os.path.join(work_dir, filename)
        
        if download_excel_file(url, file_path):
            successful_downloads += 1
            # Extract date from filename
            date = extract_date_from_filename(filename)
            
            # Process the Excel file
            df = process_excel_file(file_path, date, region)
            
            if not df.empty:
                monthly_data = pd.concat([monthly_data, df], ignore_index=True)
                print(f"    ✓ Processed {len(df)} records from {region}")
            else:
                print(f"    ⚠️  No data extracted from {region}")
                
            # Clean up downloaded file to save space
            try:
                os.remove(file_path)
            except:
                pass
        else:
            print(f"    ✗ Failed to download {region} data")
    
    # Only add All India totals if we have data from multiple regions
    if not monthly_data.empty and successful_downloads >= 3:
        print(f"\n  Calculating All India totals...")
        
        # Group by sector and sum all values
        all_india_totals = monthly_data.groupby('Sector').agg({
            'Coal': 'sum',
            'Lignite': 'sum',
            'Gas': 'sum', 
            'Diesel': 'sum',
            'Thermal Total': 'sum',
            'Nuclear': 'sum',
            'Hydro': 'sum',
            'RES': 'sum',
            'Total': 'sum'
        }).round(2)
        
        # Get the date from the first row
        date = monthly_data.iloc[0]['Date']
        
        # Create All India rows
        all_india_rows = []
        for sector in ['State', 'Private', 'Central']:
            if sector in all_india_totals.index:
                totals = all_india_totals.loc[sector]
                
                all_india_rows.append({
                    'Date': date,
                    'Region': 'All India',
                    'State': 'ALL INDIA',
                    'Sector': sector,
                    'Coal': totals['Coal'],      
                    'Lignite': totals['Lignite'], 
                    'Gas': totals['Gas'],
                    'Diesel': totals['Diesel'],
                    'Thermal Total': totals['Thermal Total'],
                    'Nuclear': totals['Nuclear'],
                    'Hydro': totals['Hydro'],
                    'RES': totals['RES'],
                    'Total': totals['Total']
                })
        
        # Add overall All India total
        overall_totals = all_india_totals.sum()
        
        all_india_rows.append({
            'Date': date,
            'Region': 'All India',
            'State': 'ALL INDIA',
            'Sector': 'Total',
            'Coal': overall_totals['Coal'],      
            'Lignite': overall_totals['Lignite'], 
            'Gas': overall_totals['Gas'],
            'Diesel': overall_totals['Diesel'],
            'Thermal Total': overall_totals['Thermal Total'],
            'Nuclear': overall_totals['Nuclear'],
            'Hydro': overall_totals['Hydro'],
            'RES': overall_totals['RES'],
            'Total': overall_totals['Total']
        })
        
        # Convert to DataFrame and append to monthly_data
        all_india_df = pd.DataFrame(all_india_rows)
        monthly_data = pd.concat([monthly_data, all_india_df], ignore_index=True)
        
        print(f"  ✓ All India totals added")
    
    print(f"\n  📊 Summary for {month} {year}:")
    print(f"     - Successful downloads: {successful_downloads}/5 regions")
    print(f"     - Total records processed: {len(monthly_data)}")
    
    return monthly_data

def main():
    """
    Main function to process NPP data with date range loop
    """
    # Set working directory
    work_dir = r"E:\VRDK 2\NPPWork"
    
    # Create directory if it doesn't exist
    os.makedirs(work_dir, exist_ok=True)
    os.chdir(work_dir)
    
    # ⛔⛔⛔ CONFIGURATION PARAMETERS - MODIFY THESE AS NEEDED ⛔⛔⛔
    
    # Option 1: Process specific date range
    START_YEAR = 2018
    START_MONTH = 1      # January
    END_YEAR = 2025
    END_MONTH = 7       # December
    
    # Option 2: Process from start year to current month (uncomment to use)
    # current_date = datetime.now()
    # START_YEAR = 2018
    # START_MONTH = 1
    # END_YEAR = current_date.year
    # END_MONTH = current_date.month
    
    # Option 3: Process only current year (uncomment to use)
    # current_date = datetime.now()
    # START_YEAR = current_date.year
    # START_MONTH = 1
    # END_YEAR = current_date.year
    # END_MONTH = current_date.month
    
    print(f"🚀 Starting NPP Data Processing Loop")
    print(f"📅 Date Range: {START_MONTH:02d}/{START_YEAR} to {END_MONTH:02d}/{END_YEAR}")
    print(f"📁 Working Directory: {work_dir}")
    
    # Generate date range
    date_range = generate_date_range(START_YEAR, START_MONTH, END_YEAR, END_MONTH)
    total_months = len(date_range)
    
    print(f"📊 Total months to process: {total_months}")
    
    # Initialize master dataframe
    all_data = pd.DataFrame()
    processed_months = 0
    successful_months = 0
    
    # Process each month in the date range
    for i, (year, month) in enumerate(date_range, 1):
        print(f"\n🔄 Progress: {i}/{total_months} months")
        
        try:
            monthly_data = process_month_data(year, month, work_dir)
            processed_months += 1
            
            if not monthly_data.empty:
                all_data = pd.concat([all_data, monthly_data], ignore_index=True)
                successful_months += 1
                print(f"  ✅ Successfully processed {month} {year}")
            else:
                print(f"  ⚠️  No data found for {month} {year}")
                
        except Exception as e:
            print(f"  ❌ Error processing {month} {year}: {str(e)}")
            continue
        
        # Add small delay to be respectful to server
        time.sleep(1)
    
    # Save final results
    if not all_data.empty:
        # Generate output filename with date range
        output_filename = f"complete_npp_data_{START_YEAR}_{END_YEAR}.csv"
        all_data.to_csv(output_filename, index=False)
        
        print(f"\n{'='*80}")
        print(f"🎉 PROCESSING COMPLETE!")
        print(f"{'='*80}")
        print(f"📈 Final Statistics:")
        print(f"   - Total months processed: {processed_months}/{total_months}")
        print(f"   - Successful months: {successful_months}")
        print(f"   - Total records: {len(all_data):,}")
        print(f"   - Output file: {output_filename}")
        
        # Summary by region
        print(f"\n📊 Data Summary by Region:")
        region_summary = all_data.groupby('Region')['Total'].agg(['count', 'sum']).round(2)
        region_summary.columns = ['Records', 'Total_Capacity_MW']
        print(region_summary)
        
        # Summary by year
        print(f"\n📅 Data Summary by Year:")
        all_data['Year'] = pd.to_datetime(all_data['Date'], format='%d-%m-%Y').dt.year
        year_summary = all_data.groupby('Year').size()
        print(year_summary)
        
        print(f"\n💾 Data saved successfully to: {output_filename}")
        
    else:
        print(f"\n❌ No data was processed successfully")
        print(f"   - Check your internet connection")
        print(f"   - Verify that NPP website is accessible")
        print(f"   - Consider adjusting the date range")

# Run the enhanced script
if __name__ == "__main__":
    main()