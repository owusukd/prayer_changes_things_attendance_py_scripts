import pandas as pd
import numpy as np
from pathlib import Path

def calculate_monthly_attendance_averages(input_file, output_file='monthly_attendance_averages.xlsx'):
    """
    Calculate monthly attendance averages for each year, branch, and constituency
    and export results to Excel with multiple sheets.
    
    Parameters:
    input_file (str): Path to the input Excel file
    output_file (str): Path for the output Excel file
    """
    
    # Read the Excel file
    print(f"Reading data from {input_file}...")
    df = pd.read_excel(input_file)
    
    # Display basic info about the dataset
    print(f"Dataset shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    print(f"Date range: {df['Year'].min()}-{df['Year'].max()}")
    print(f"Months: {sorted(df['Month'].unique())}")
    print(f"Constituencies: {sorted(df['Constituency'].unique())}")
    print(f"Total branches: {df['Branch'].nunique()}")
    
    # Clean the data - remove rows with missing key values
    df_clean = df.dropna(subset=['Constituency', 'Branch', 'Attendance'])
    print(f"Clean dataset shape: {df_clean.shape}")
    
    # 1. Monthly averages by Branch (most granular level)
    print("\nCalculating monthly averages by branch...")
    branch_monthly = df_clean.groupby(['Year', 'Month', 'Constituency', 'Branch']).agg({
        'Attendance': 'mean',
        'Target': 'max'
    }).round(2)
    
    # Flatten column names
    branch_monthly.columns = ['Monthly_Attendance_Avg', 'Target']
    branch_monthly = branch_monthly.reset_index()
    
    # Calculate attendance rate
    branch_monthly['Attendance_Rate'] = (
        branch_monthly['Monthly_Attendance_Avg'] / branch_monthly['Target'] * 100
    ).round(2)
    
    # 2. Monthly averages by Constituency
    print("Calculating monthly averages by constituency...")
    constituency_monthly = df_clean.groupby(['Year', 'Month', 'Constituency']).agg({
        'Week': 'nunique',
        'Attendance': 'sum',
        'Branch': 'nunique',
        'Target': 'sum'
    }).round(2)
    constituency_monthly.columns = ['Reports_Count', 'Total_Attendance', 'Unique_Branches', 'Total_Target']
    constituency_monthly = constituency_monthly.reset_index()
    

    constituency_monthly['Monthly_Attendance_Avg'] = (
        constituency_monthly['Total_Attendance'] / constituency_monthly['Reports_Count']
    ).round(2)
    
    
    constituency_monthly['Target'] = (
        constituency_monthly['Total_Target'] / constituency_monthly['Reports_Count']
    ).round(0)
    
    # Calculate attendance rate
    constituency_monthly['Attendance_Rate'] = (
        constituency_monthly['Monthly_Attendance_Avg'] / constituency_monthly['Target'] * 100
    ).round(2)
    
    # Flatten column names
    constituency_monthly.drop(['Reports_Count', 'Total_Attendance', 'Total_Target'], axis=1, inplace=True)
   
    
    # # 3. Overall monthly averages
    # print("Calculating overall monthly averages...")
    # overall_monthly = df_clean.groupby(['Year', 'Month']).agg({
    #     'Attendance': ['mean', 'sum', 'count'],
    #     'Constituency': 'nunique',
    #     'Branch': 'nunique',
    #     'Target': 'mean'
    # }).round(2)
    
    # # Flatten column names
    # overall_monthly.columns = ['Monthly_Attendance_Avg', 'Total_Attendance', 'Reports_Count', 'Unique_Constituencies', 'Unique_Branches', 'Avg_Target']
    # overall_monthly = overall_monthly.reset_index()
    
    # # Calculate attendance rate
    # overall_monthly['Attendance_Rate'] = (
    #     overall_monthly['Monthly_Attendance_Avg'] / overall_monthly['Avg_Target'] * 100
    # ).round(2)
    
    # # 4. Summary statistics by constituency (across all months)
    # print("Calculating constituency summary statistics...")
    # constituency_summary = df_clean.groupby('Constituency').agg({
    #     'Attendance': ['mean', 'median', 'std', 'min', 'max', 'sum'],
    #     'Branch': 'nunique',
    #     'Target': 'mean'
    # }).round(2)
    
    # # Flatten column names
    # constituency_summary.columns = ['Mean_Attendance', 'Median_Attendance', 'Std_Attendance', 
    #                                'Min_Attendance', 'Max_Attendance', 'Total_Attendance', 
    #                                'Unique_Branches', 'Avg_Target']
    # constituency_summary = constituency_summary.reset_index()
    
    # 5. Top performing branches
    # print("Identifying top performing branches...")
    # top_branches = branch_monthly.nlargest(20, 'Monthly_Attendance_Avg')[
    #     ['Year', 'Month', 'Constituency', 'Branch', 'Monthly_Attendance_Avg', 'Attendance_Rate']
    # ]
    
    print("\nIdentifying top performing branches...")
    top_performing_branches = branch_monthly.groupby('Month').apply(lambda x: x.nlargest(5,'Monthly_Attendance_Avg')).reset_index(drop=True)[
        ['Year', 'Month', 'Constituency', 'Branch', 'Monthly_Attendance_Avg', 'Attendance_Rate']
    ]
    print("\nIdentifying low performing branches...")    
    low_performing_branches = branch_monthly.groupby('Month').apply(lambda x: x.nsmallest(5,'Monthly_Attendance_Avg')).reset_index(drop=True)[
        ['Year', 'Month', 'Constituency', 'Branch', 'Monthly_Attendance_Avg', 'Attendance_Rate']
    ]
    
    # # 6. Month-over-month comparison (if multiple months exist)
    # if len(df_clean['Month'].unique()) > 1:
    #     print("Calculating month-over-month changes...")
    #     # Pivot to get months as columns for comparison
    #     mom_comparison = branch_monthly.pivot_table(
    #         index=['Year', 'Constituency', 'Branch'], 
    #         columns='Month', 
    #         values='Monthly_Attendance_Avg'
    #     ).reset_index()
        
    #     # Calculate percentage change between months (assuming chronological order)
    #     months = sorted(df_clean['Month'].unique())
    #     if len(months) == 2:
    #         mom_comparison['Change_Amount'] = (
    #             mom_comparison[months[1]] - mom_comparison[months[0]]
    #         ).round(2)
    #         mom_comparison['Change_Percent'] = (
    #             (mom_comparison[months[1]] - mom_comparison[months[0]]) / 
    #             mom_comparison[months[0]] * 100
    #         ).round(2)
    # else:
    #     mom_comparison = pd.DataFrame()  # Empty if only one month
    
    # Export to Excel with multiple sheets
    print(f"\nExporting results to {output_file}...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write each analysis to a separate sheet
        # overall_monthly.to_excel(writer, sheet_name='Overall_Monthly', index=False)
        constituency_monthly.to_excel(writer, sheet_name='Constituency_Monthly', index=False)
        branch_monthly.to_excel(writer, sheet_name='Branch_Monthly', index=False)
        # constituency_summary.to_excel(writer, sheet_name='Constituency_Summary', index=False)
        top_performing_branches.to_excel(writer, sheet_name='Top_Performers', index=False)
        low_performing_branches.to_excel(writer, sheet_name='Low_Performers', index=False)
        
        # if not mom_comparison.empty:
        #     mom_comparison.to_excel(writer, sheet_name='Month_Over_Month', index=False)
        
        # Also include the original cleaned data for reference
        df_clean.to_excel(writer, sheet_name='Original_Data', index=False)
    
    print(f"Analysis complete! Results saved to {output_file}")
    
    # Display summary results
    print(f"\n{'='*60}")
    print("SUMMARY RESULTS")
    print(f"{'='*60}")
    
    # print("\nOverall Monthly Averages:")
    # print(overall_monthly.to_string(index=False))
    
    print("\nTop Performing Branches in each month:")
    print(top_performing_branches.head(15).to_string(index=False))
    print(low_performing_branches.head(15).to_string(index=False))
    
    # print("\nConstituency Summary (sorted by mean attendance):")
    # constituency_summary_sorted = constituency_summary.sort_values('Mean_Attendance', ascending=False)
    # print(constituency_summary_sorted.to_string(index=False))
    
    return {
        # 'overall_monthly': overall_monthly,
        'constituency_monthly': constituency_monthly,
        'branch_monthly': branch_monthly,
        # 'constituency_summary': constituency_summary,
        # 'top_branches': top_branches,
        # 'mom_comparison': mom_comparison
    }

# Example usage
if __name__ == "__main__":
    # Replace with your actual file path
    input_file = "prayer_changes_everything_attendance_data.xlsx"
    output_file = "monthly_attendance_analysis_results.xlsx"
    
    # Run the analysis
    try:
        results = calculate_monthly_attendance_averages(input_file, output_file)
        print("\n✅ Analysis completed successfully!")
        print(f"📊 Results exported to: {output_file}")
        print(f"📝 The Excel file contains {len(results)} different analysis sheets")
        
    except FileNotFoundError:
        print(f"❌ Error: Could not find input file '{input_file}'")
        print("Please ensure the file path is correct and the file exists.")
        
    except Exception as e:
        print(f"❌ Error during analysis: {str(e)}")
        print("Please check your data format and try again.")

# Additional utility function for quick analysis
def quick_summary(input_file):
    """Quick summary of attendance data without full export"""
    df = pd.read_excel(input_file)
    
    print("QUICK DATA SUMMARY")
    print("="*40)
    print(f"Total records: {len(df)}")
    print(f"Date range: {df['Month'].iloc[0]} - {df['Month'].iloc[-1]} {df['Year'].iloc[0]}")
    print(f"Constituencies: {df['Constituency'].nunique()}")
    print(f"Branches: {df['Branch'].nunique()}")
    print(f"Average attendance: {df['Attendance'].mean():.2f}")
    print(f"Total attendance: {df['Attendance'].sum()}")
    
    # Top 5 branches by average attendance
    top_branches = df.groupby('Branch')['Attendance'].mean().nlargest(5)
    print("\nTop 5 branches by average attendance:")
    for branch, avg in top_branches.items():
        print(f"  {branch}: {avg:.1f}")

