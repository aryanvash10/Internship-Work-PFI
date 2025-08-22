# NPP Data Processing Script

## Overview
This Python script downloads, processes, and aggregates power capacity data from the National Power Portal (NPP) of India. It retrieves monthly installed capacity reports in Excel format from the NPP website, processes them to extract relevant data, and compiles the results into a single CSV file. The script handles data for multiple regions, sectors, and fuel types, and includes error handling and retry mechanisms for robust operation.

## Features
- **Automated Data Download**: Downloads Excel files from the NPP website for specified date ranges and regions.
- **Data Extraction**: Processes Excel files to extract power capacity data by region, state, sector, and fuel type (Coal, Lignite, Gas, Diesel, Nuclear, Hydro, RES).
- **Error Handling**: Includes retry mechanisms for failed downloads and robust parsing of Excel files with varying formats.
- **Data Aggregation**: Compiles data into a structured CSV file, including regional and All India totals.
- **Progress Tracking**: Provides detailed console output for monitoring download and processing progress.
- **Summary Statistics**: Generates summaries by region and year for quick insights.

## Prerequisites
To run the script, you need the following Python packages:
- `pandas`
- `requests`
- `numpy`
- `openpyxl` (for Excel file processing)

Install the dependencies using pip:
```bash
pip install pandas requests numpy openpyxl
```

## Usage
1. **Clone the Repository**:
   ```bash
   git clone https://github.com/your-username/your-repo-name.git
   cd your-repo-name
   ```

2. **Configure the Script**:
   - Open the `main()` function in the script and modify the configuration parameters (`START_YEAR`, `START_MONTH`, `END_YEAR`, `END_MONTH`) to specify the date range for data processing.
   - Set the `work_dir` variable to your desired working directory where temporary Excel files will be saved and the final CSV output will be stored.

3. **Run the Script**:
   ```bash
   python npp_data_processing.py
   ```

4. **Output**:
   - The script will download Excel files, process them, and save the results to a CSV file named `complete_npp_data_{START_YEAR}_{END_YEAR}.csv` in the specified working directory.
   - Console output will show progress, errors, and summary statistics.

## Configuration
The script provides three options for specifying the date range in the `main()` function:
- **Option 1**: Process a specific date range (default: 2018–2025).
- **Option 2**: Process from a start year to the current month (commented out).
- **Option 3**: Process only the current year (commented out).

Modify these parameters as needed:
```python
START_YEAR = 2018
START_MONTH = 1  # January
END_YEAR = 2025
END_MONTH = 7    # July
```

## Directory Structure
- **Working Directory**: Temporary Excel files are downloaded here and deleted after processing. The final CSV output is saved here.
- **Output File**: Named `complete_npp_data_{START_YEAR}_{END_YEAR}.csv`.

## Data Source
The script fetches data from the National Power Portal (NPP) at URLs like:
```
https://npp.gov.in/public-reports/cea/monthly/installcap/{year}/{month}/capacity2-{region}-{year}-{month_num}.xls
```

## Output Format
The output CSV contains the following columns:
- **Date**: Date in DD-MM-YYYY format (last day of the month).
- **Region**: Region name (Northern, Eastern, Western, Southern, North Eastern, All India).
- **State**: State name or "ALL INDIA" for national totals.
- **Sector**: Sector type (State, Private, Central, Total).
- **Coal**: Coal-based capacity (MW).
- **Lignite**: Lignite-based capacity (MW).
- **Gas**: Gas-based capacity (MW).
- **Diesel**: Diesel-based capacity (MW).
- **Thermal Total**: Sum of Coal, Lignite, Gas, and Diesel (MW).
- **Nuclear**: Nuclear-based capacity (MW).
- **Hydro**: Hydro-based capacity (MW).
- **RES**: Renewable Energy Sources capacity (MW).
- **Total**: Total capacity (MW).

## Notes
- The script assumes internet connectivity and accessibility to the NPP website.
- Temporary Excel files are deleted after processing to save disk space.
- The script includes a 1-second delay between monthly processing to avoid overwhelming the server.
- If data for a specific month is unavailable, the script skips it and continues with the next month.
- All India totals are calculated only if data from at least three regions is successfully processed.

## Troubleshooting
- **Download Failures**: Check your internet connection or verify if the NPP website is accessible.
- **Excel Processing Errors**: Ensure the Excel files follow the expected format. The script is designed to handle variations, but unexpected formats may cause issues.
- **Missing Data**: If no data is processed, try adjusting the date range or checking the NPP website for data availability.

## Contributing
Contributions are welcome! Please submit a pull request or open an issue for bug reports, feature requests, or improvements.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.