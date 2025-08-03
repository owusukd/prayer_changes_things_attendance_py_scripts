import re
import pandas as pd
#from datetime import datetime
from typing import List, Dict

def parse_attendance_data(data: str, pastor: str, month: str, week: str, year: str) -> List[Dict[str, str]]:
    """
    Parse attendance data from the given text format.
    
    Args:
        data (str): Raw text data containing attendance information
        
    Returns:
        List[Dict]: List of dictionaries containing parsed attendance data
    """
    
    # Split data into lines and clean up
    lines = [line.strip() for line in data.split('\n') if line.strip()]
    
    parsed_data = []
    current_constituency = None
    
    for line in lines:
        # Check if line is a constituency header (enclosed in asterisks)
        if line.startswith('*') and line.endswith('*'):
            # Extract constituency name, removing asterisks
            current_constituency = line.strip('*')
        
        # Check if line contains branch attendance data (starts with 👉🏾)
        elif line.startswith('👉🏾') and current_constituency:
            # Parse branch attendance line
            branch_data = parse_branch_line(line, current_constituency, pastor, month, week, year)
            if branch_data:
                parsed_data.append(branch_data)
    
    return parsed_data

def parse_branch_line(line: str, constituency: str, pastor: str, month: str, week: str, year: str) -> Dict[str, str]:
    """
    Parse a single branch attendance line.
    
    Args:
        line (str): Line containing branch attendance data
        constituency (str): Current constituency name
        
    Returns:
        Dict: Parsed branch data or None if parsing fails
    """
    
    # Remove the emoji and clean up the line
    line = line.replace('👉🏾', '').strip()
    
    # Use regex to extract branch name and attendance numbers
    # Pattern looks for: branch_name - attendance/expected_attendance
    pattern = r'^(.+?)\s*-\s*(\d+)\s*[/|]\s*(\d+)$'
    match = re.search(pattern, line)
    
    if match:
        branch_name = match.group(1).strip()
        attendance = int(match.group(2))
        expected_attendance = int(match.group(3))
        
        return {
            'Constituency': constituency,
            'Branch': branch_name,
            'Pastor': pastor,
            'Attendance': attendance,
            'Target': expected_attendance,
            'Attendance_rate': round((attendance / expected_attendance) * 100, 2),
            'Month': month,
            'Week': week,
            'Year': year
        }
    
    return None

def analyze_attendance_data(parsed_data: List[Dict]) -> Dict:
    """
    Perform basic analysis on the parsed attendance data.
    
    Args:
        parsed_data (List[Dict]): Parsed attendance data
        
    Returns:
        Dict: Analysis results
    """
    
    if not parsed_data:
        return {}
    
    df = pd.DataFrame(parsed_data)
    
    analysis = {
        'total_branches': len(df),
        'total_attendance': df['Attendance'].sum(),
        'total_expected': df['Target'].sum(),
        'overall_attendance_rate': round((df['Attendance'].sum() / df['Target'].sum()) * 100, 2),
        'constituency_summary': df.groupby('Constituency').agg({
            'Attendance': 'sum',
            'Target': 'sum',
            'Branch': 'count'
        }).rename(columns={'Branch': 'Branch_count'}).round(2),
        'top_performing_branches': df.nlargest(5, 'Attendance_rate')[['Branch', 'Constituency', 'Attendance_rate']],
        'lowest_performing_branches': df.nsmallest(5, 'Attendance_rate')[['Branch', 'Constituency', 'Attendance_rate']]
    }
    
    # Calculate attendance rate for each constituency
    constituency_rates = []
    for constituency in df['Constituency'].unique():
        constituency_data = df[df['Constituency'] == constituency]
        rate = (constituency_data['Attendance'].sum() / constituency_data['Target'].sum()) * 100
        constituency_rates.append({'Constituency': constituency, 'Attendance_rate': round(rate, 2)})
    
    analysis['constituency_rates'] = sorted(constituency_rates, key=lambda x: x['Attendance_rate'], reverse=True)
    
    return analysis

def main(raw_data, pastor, month, week, year, start_row=None):
    """
    Main function to demonstrate the parser with the provided data.
    """
    
    parsed_data = parse_attendance_data(raw_data, pastor, month, week, year)
    
    # Convert to DataFrame for easier viewing
    df = pd.DataFrame(parsed_data)
    
    # Display results
    print("=== PARSED ATTENDANCE DATA ===")
    print(df.to_string(index=False))
    print(f"\nTotal records parsed: {len(df)}")
    
    # Perform analysis
    analysis = analyze_attendance_data(parsed_data)
    
    print("\n=== ANALYSIS RESULTS ===")
    print(f"Total branches: {analysis['total_branches']}")
    print(f"Total attendance: {analysis['total_attendance']}")
    print(f"Total expected: {analysis['total_expected']}")
    print(f"Overall attendance rate: {analysis['overall_attendance_rate']}%")
    
    print("\n=== CONSTITUENCY SUMMARY ===")
    print(analysis['constituency_summary'])
    
    print("\n=== CONSTITUENCY ATTENDANCE RATES ===")
    for item in analysis['constituency_rates']:
        print(f"{item['Constituency']}: {item['Attendance_rate']}%")
    
    print("\n=== TOP 5 PERFORMING BRANCHES ===")
    print(analysis['top_performing_branches'][['Branch', 'Constituency', 'Attendance_rate']].to_string(index=False))
    
    print("\n=== LOWEST 5 PERFORMING BRANCHES ===")
    print(analysis['lowest_performing_branches'][['Branch', 'Constituency', 'Attendance_rate']].to_string(index=False))
    
    #### Save to excel file
    ## First use of this code
    # with pd.ExcelWriter('prayer_changes_everything_attendance_data.xlsx', mode='w', engine='openpyxl') as writer:
    #     df.to_excel(writer, sheet_name='Sheet1', index=False, header=True)
    ## Subsequent use 
    with pd.ExcelWriter('prayer_changes_everything_attendance_data.xlsx', mode='a', if_sheet_exists="overlay", engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=start_row)
        
    print("\nData saved to 'prayer_changes_everything_attendance_data.xlsx'")
    

if __name__ == "__main__":
    start_row = 161 # first run set to 'None'. open saved excel file after first run for the row number to append from: it is same as the last row number
    pastor = "NA"  # Replace with your actual pastor's name
    month = "July"  # Replace with your actual month
    week = "5"  # Replace with your actual week
    year = "2025"  # Replace with your actual year
    
    raw_data ="""
    *CONSTITUENCY_NAME*
    👉🏾Branch 1 - 10|15
    👉🏾Branch 2 - 8|12"""

    main(raw_data, pastor, month, week, year, start_row)
