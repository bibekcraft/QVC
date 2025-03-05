import pandas as pd

# Load the data from the master file
df = pd.read_excel('mster- Barcode.xlsx')

# Check if the columns are correct
if 'SINO' in df.columns and 'UNIQUENUMBER' in df.columns:
    # Open the report and check file
    with open('report.txt', 'w') as report, open('checking_duplicates.txt', 'w') as check_file:
        report.write("Report for Data Splitting into Lots\n\n")
        check_file.write("Checking for Duplicates (SINO or UNIQUENUMBER)\n\n")

        # Check for duplicates in SINO or UNIQUENUMBER
        duplicate_sino = df[df.duplicated(subset='SINO', keep=False)]
        duplicate_uniquenum = df[df.duplicated(subset='UNIQUENUMBER', keep=False)]

        if not duplicate_sino.empty:
            check_file.write("Duplicates found in SINO:\n")
            check_file.write(duplicate_sino.to_string(index=False))
            check_file.write("\n\n")

        if not duplicate_uniquenum.empty:
            check_file.write("Duplicates found in UNIQUENUMBER:\n")
            check_file.write(duplicate_uniquenum.to_string(index=False))
            check_file.write("\n\n")
        
        # Split the data into 30 lots, each containing 1000 rows
        num_rows = len(df)
        lot_size = 1000

        for i in range(30):
            # Calculate the start and end indices for each lot
            start_idx = i * lot_size
            end_idx = start_idx + lot_size

            # Slice the dataframe for the current lot
            lot_df = df.iloc[start_idx:end_idx]

            # Save the lot as a new Excel file
            lot_df.to_excel(f'lot{i+1}.xlsx', index=False)

            # Get the first and last SINO numbers for the lot
            first_sino = lot_df['SINO'].iloc[0]
            last_sino = lot_df['SINO'].iloc[-1]
            
            # Write information to the report
            report.write(f'Lot {i+1}:\n')
            report.write(f' - Data from Serial Number {first_sino} to {last_sino}\n')

            # Exclude middle SINO number from report if it's not the first or last entry
            if len(lot_df) > 2:
                middle_sino = lot_df['SINO'].iloc[len(lot_df) // 2]
                report.write(f' - Excluded middle Serial Number: {middle_sino}\n')

            report.write('\n')

            print(f'lot{i+1}.xlsx created with {len(lot_df)} records.')

        print("Report and checking file generated successfully.")
else:
    print('The necessary columns "SINO" and "UNIQUENUMBER" are missing.')
