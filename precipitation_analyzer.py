import pandas as pd
import os
import calendar # For leap year check

def is_leap_year(year):
    """Checks if a given year is a leap year."""
    return calendar.isleap(year)

def process_excel_files(input_folder, output_processed_folder, output_summary_folder):
    """
    Processes Excel and CSV files to calculate column averages and save results.

    This enhanced function dynamically identifies the starting point of year data
    by looking for "DAY" or "days" headers. It then processes subsequent columns
    as years to calculate averages.

    Args:
        input_folder (str): Path to the folder containing original Excel/CSV files.
        output_processed_folder (str): Path to the folder to save processed files.
        output_summary_folder (str): Path to the folder to save the summary of averages.
    """

    # Create output directories if they don't exist
    os.makedirs(output_processed_folder, exist_ok=True)
    os.makedirs(output_summary_folder, exist_ok=True)

    summary_data = [] # To store year and average for the summary file

    for filename in os.listdir(input_folder):
        # Check for both Excel and CSV file extensions
        if filename.endswith(('.xlsx', '.xls', '.csv')):
            filepath = os.path.join(input_folder, filename)
            try:
                # Read the file based on its extension
                if filename.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(filepath)
                    print(f"Reading Excel file: {filename}")
                elif filename.endswith('.csv'):
                    df = pd.read_csv(filepath)
                    print(f"Reading CSV file: {filename}")

                # Initialize a list to hold the averages for the current file
                current_file_averages = []
                average_row_data = {} # To store averages for the new row

                # Determine the starting index for year columns
                # We'll look for 'DAY' or 'days' and start processing from the column after it.
                start_col_index_for_years = -1 # Default to -1, meaning not found yet

                for col_index, col_name in enumerate(df.columns):
                    # Convert column name to lowercase for case-insensitive comparison
                    lower_col_name = str(col_name).lower()
                    if lower_col_name == 'day' or lower_col_name == 'days':
                        start_col_index_for_years = col_index + 1 # Years start from the next column
                        break # Found our marker, no need to check earlier columns

                # If 'DAY' or 'days' marker was not found, assume years start from the 3rd column (index 2)
                # as a fallback to the previous logic.
                if start_col_index_for_years == -1:
                    print(f"Warning: Neither 'DAY' nor 'days' header found in '{filename}'. Assuming year columns start from the 3rd column (index 2).")
                    start_col_index_for_years = 2

                # Iterate through columns from the determined starting index for years
                for col_index, col_name in enumerate(df.columns):
                    if col_index >= start_col_index_for_years:
                        try:
                            # Attempt to convert column name to an integer year.
                            # If the column header is not a valid year, this will raise a ValueError,
                            # and the column will be skipped.
                            year = int(col_name)
                            # Calculate the average of the numerical data in the column
                            # .dropna() to ignore any non-numeric or empty cells
                            column_data = pd.to_numeric(df[col_name], errors='coerce').dropna()
                            if not column_data.empty:
                                col_average = column_data.mean()
                                current_file_averages.append({
                                    'Year': year,
                                    'Average': col_average,
                                    'Leap Year': is_leap_year(year)
                                })
                                average_row_data[col_name] = col_average
                            else:
                                print(f"Warning: Column '{col_name}' in '{filename}' contains no valid numerical data. Skipping average calculation.")
                                average_row_data[col_name] = None # Or some placeholder
                        except ValueError:
                            # This block handles cases where the column header is not a valid integer year.
                            print(f"Warning: Column header '{col_name}' in '{filename}' is not a valid year. Skipping average calculation for this column.")
                            average_row_data[col_name] = None # Or some placeholder
                        except Exception as e:
                            print(f"Error processing column '{col_name}' in '{filename}': {e}")
                            average_row_data[col_name] = None # Or some placeholder
                    else:
                        # For columns before the year data starts, just add them to the average_row_data
                        # with their original values or None if they are not relevant for averaging.
                        # This ensures the new row aligns correctly.
                        average_row_data[col_name] = None # Or df.iloc[-1][col_name] if you want to copy last row data

                # Add a new row with averages to the DataFrame for the processed file
                # Create a new DataFrame for the average row, aligning columns
                average_df = pd.DataFrame([average_row_data], columns=df.columns)
                df_with_average = pd.concat([df, average_df], ignore_index=True)

                # Save the processed file (maintaining original extension for now, can be changed if needed)
                output_filepath = os.path.join(output_processed_folder, filename)
                if filename.endswith(('.xlsx', '.xls')):
                    df_with_average.to_excel(output_filepath, index=False)
                elif filename.endswith('.csv'):
                    df_with_average.to_csv(output_filepath, index=False)
                print(f"Processed '{filename}' and saved to '{output_filepath}'")

                # Add data to the overall summary
                for avg_info in current_file_averages:
                    summary_data.append({
                        'File': filename,
                        'Year': avg_info['Year'],
                        'Average Data': avg_info['Average'],
                        'Is Leap Year': avg_info['Leap Year']
                    })

            except Exception as e:
                print(f"Error processing file '{filename}': {e}")

    # Save the consolidated summary of years and averages
    if summary_data:
        summary_df = pd.DataFrame(summary_data)
        summary_output_filepath = os.path.join(output_summary_folder, 'all_files_years_and_averages.csv')
        summary_df.to_csv(summary_output_filepath, index=False)
        print(f"\nSummary of years and averages saved to '{summary_output_filepath}'")
    else:
        print("\nNo data processed to create a summary file.")

# --- How to use the code ---
if __name__ == "__main__":
    # IMPORTANT: Replace this with the actual path to your input folder
    input_folder_path = r"C:\Users\aaa\Desktop\all 30s"
    output_processed_folder_path = os.path.join(input_folder_path, "processed_excel_files")
    output_summary_folder_path = os.path.join(input_folder_path, "summary_averages")

    process_excel_files(input_folder_path, output_processed_folder_path, output_summary_folder_path)
    print("\nScript execution completed. Check the specified output folders.")
