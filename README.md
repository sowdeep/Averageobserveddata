# Averageobserveddata
this repository consist of python file that will average the observed data of the DHM station of years for multiple stations at once 
place all the data in the base directory and run the python ; it will give 2 folders , processed_excel_files and summary_averages , there it will created same csv file but with average on it last row ; it counts leap year too .. and averages and in summary folder it will give a single csv file of stations years avg data 
✅


The precipitation_analyzer.py script processes Excel and CSV files to calculate column averages, specifically for yearly precipitation data.
Here's a breakdown of what it does:
It reads Excel and CSV files from a specified input folder.
It dynamically identifies columns containing year data by looking for "DAY" or "days" headers.
For each identified year column, it calculates the average of the numerical data within that column.
It also determines if a year is a leap year.
A new row containing these calculated averages is added to the original DataFrame for each processed file.
The modified files (with the average row) are saved to a processed_excel_files subfolder.
Finally, it creates a summary CSV file named all_files_years_and_averages.csv in a summary_averages subfolder, which consolidates the file name, year, calculated average, and leap year status for all processed files.
