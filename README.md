# Attendance Analysis System

A comprehensive Python-based data analysis pipeline for processing, analyzing, and visualizing attendance data across multiple constituencies and branches. This system provides automated data parsing, statistical analysis, and chart generation capabilities.

## Overview

This system consists of three main components:
1. **Data Parser** - Extracts structured data from formatted text input
2. **Statistical Analyzer** - Performs comprehensive attendance analytics
3. **Visualization Generator** - Creates professional charts and reports

## Files Description

### 1. `attendance_parser_final.py`
**Purpose**: Parses raw attendance data from formatted text input and converts it to structured Excel format.

**Key Features**:
- Regex-based parsing of constituency and branch data
- Automatic data validation and cleaning
- Excel export with append functionality
- Basic statistical summaries

**Input Format**:
```
*CONSTITUENCY_NAME*
👉🏾Branch Name - attendance/target
👉🏾Another Branch - attendance/target
```

### 2. `attendance_analysis_python.py`
**Purpose**: Comprehensive statistical analysis of attendance data with multi-level aggregations.

**Key Features**:
- Monthly attendance averages by branch and constituency
- Performance ranking and identification
- Attendance rate calculations
- Multi-sheet Excel export with detailed analytics
- Top/bottom performer identification

### 3. `attendance_charts_generator.py`
**Purpose**: Generates professional bar charts and visual reports from analyzed data.

**Key Features**:
- Constituency-level comparative charts
- Branch-level detailed visualizations
- Automatic chart embedding in Excel files
- Customizable chart layouts and styling

## Installation

### Requirements
```bash
pip install pandas numpy matplotlib openpyxl pathlib
```

### Dependencies
- **pandas**: Data manipulation and analysis
- **numpy**: Numerical computing
- **matplotlib**: Chart generation
- **openpyxl**: Excel file handling
- **pathlib**: File path management

## Usage Workflow

### Step 1: Data Parsing
```python
from attendance_parser_final import main

# Configure parameters
start_row = None  # First run: None, subsequent: last row number
pastor = "Pastor Name"
month = "July"
week = "5"
year = "2025"

# Raw data input (see format example below)
raw_data = """
*CONSTITUENCY_NAME*
👉🏾Branch 1 - 10|15
👉🏾Branch 2 - 8|12
"""

# Execute parsing
main(raw_data, pastor, month, week, year, start_row)
```

### Step 2: Statistical Analysis
```python
from attendance_analysis_python import calculate_monthly_attendance_averages

# Analyze parsed data
results = calculate_monthly_attendance_averages(
    input_file="prayer_changes_everything_attendance_data.xlsx",
    output_file="monthly_attendance_analysis_results.xlsx"
)
```

### Step 3: Chart Generation
```python
from attendance_charts_generator import create_attendance_charts

# Generate visualizations
output_file = create_attendance_charts("monthly_attendance_analysis_results.xlsx")
```

## Configuration Guide

### First Time Setup
1. In `attendance_parser_final.py`, set:
   - `start_row = None`
   - Update `collection_date` (line 181)
   - Set appropriate `pastor`, `month`, `week`, `year` values

2. Ensure raw data follows the required format with `-` separators

### Subsequent Runs
1. Open the generated Excel file to find the last row number
2. Set `start_row` to the next available row number
3. Update date parameters as needed

## Output Files

### Primary Outputs
- `prayer_changes_everything_attendance_data.xlsx` - Raw parsed data
- `monthly_attendance_analysis_results.xlsx` - Comprehensive analytics
- `monthly_attendance_analysis_results_with_charts.xlsx` - Final report with visualizations

### Excel Sheet Structure
1. **Constituency_Monthly** - Aggregated constituency-level data
2. **Branch_Monthly** - Detailed branch-level analysis
3. **Top_Performers** - Highest performing branches by month
4. **Low_Performers** - Lowest performing branches by month
5. **Original_Data** - Clean source data for reference
6. **Constituency_Monthly_Charts** - Visual comparisons by constituency
7. **Branch_Monthly_Charts** - Detailed branch performance charts

## Data Structure

### Input Data Fields
- Constituency name (enclosed in asterisks)
- Branch name
- Actual attendance
- Target attendance

### Output Data Fields
- **Constituency**: Geographic/organizational grouping
- **Branch**: Individual location/unit
- **Pastor**: Responsible leader
- **Attendance**: Actual attendance count
- **Target**: Expected attendance
- **Attendance_Rate**: Percentage achievement
- **Month/Week/Year**: Temporal identifiers

## Analysis Features

### Statistical Metrics
- Monthly attendance averages
- Attendance rate calculations
- Performance rankings
- Trend identification
- Comparative analysis

### Visualization Types
- Constituency comparison charts
- Branch performance graphs
- Target vs. actual comparisons
- Monthly trend analysis

## Troubleshooting

### Common Issues

**File Not Found Error**
```python
# Ensure correct file paths
input_file = "correct_path/prayer_changes_everything_attendance_data.xlsx"
```

**Data Format Issues**
- Verify constituency names are enclosed in asterisks: `*NAME*`
- Ensure branch lines start with 👉🏾 emoji
- Check attendance format: `attendance|target` or `attendance/target`

**Chart Generation Errors**
- Verify matplotlib backend compatibility
- Ensure sufficient memory for large datasets
- Check Excel file permissions

### Data Validation
The system includes automatic validation for:
- Missing attendance values
- Invalid numeric formats
- Incomplete constituency/branch information

## Performance Considerations

### Large Datasets
- The system efficiently handles datasets with 1000+ records
- Memory usage scales linearly with data size
- Chart generation time increases with constituency/branch count

### Optimization Tips
- Use categorical data types for repeated string values
- Consider data chunking for very large datasets
- Implement caching for repeated analyses

## Contributing

When modifying the scripts:
1. Maintain consistent data validation patterns
2. Follow pandas best practices for data manipulation
3. Ensure backward compatibility with existing data files
4. Test with various data sizes and formats

## License

This project is designed for internal data analysis use. Ensure compliance with your organization's data handling policies.

## Support

For issues or questions:
1. Check the troubleshooting section
2. Verify data format compliance
3. Review error messages for specific guidance
4. Consider data volume and system resources