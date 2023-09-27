import pandas as pd
import os
# Step 1: Read the Excel file into a Pandas DataFrame.
df = pd.read_excel(r'C:\Users\FutureX\Aged.xlsx')

# Step 1: Search “LOST” in Site Name, exclude them in all following reports.
mask = df['Site Name'].str.contains('LOST', case=False)
mask = mask.fillna(False).astype(bool)
file = df[~mask]
outputpath = r'C:\Users\FutureX\Occupant_Accounts\File.csv'  # Full file path

# Step 2: save file
file.to_csv(outputpath, index=False)
# Step 3: Search “occup” in Account Name and copy to a new sheet named 'Occupant_Accounts'.
df = pd.read_csv(r'C:\Users\FutureX\Occupant_Accounts\File.csv')
occupant_df = df[df['Account Name'].str.contains('occup', case=False)]

# Define the output directory ('C:\Users\FutureX') and file name ('Occupant_account.xlsx').
output_directory = os.path.expanduser('~')
outputpath = r'C:\Users\FutureX\Occupant_Accounts\Occupant_Account.csv'  

occupant_df.to_csv(outputpath, index=False)

df = pd.read_csv(r'C:\Users\FutureX\Occupant_Accounts\File.csv')

# Step 2: Filter based on criteria.
filtered_df = df[
    (df['Credit Control Status'].str.contains('Standard credit control procedures'))&
   (df['Account Status'].isin(['OPEN', '[mixed]'])) &
    (df['Overdue Balance'] >= 300)
]
outputpath = r'C:\Users\FutureX\Occupant_Accounts\Disconnection list.csv'  
filtered_df.to_csv(outputpath, index=False)

df = pd.read_csv(r'C:\Users\FutureX\Occupant_Accounts\File.csv')

Pass_to = df[
    (df['Account Status'] == 'CLOSED') &
    (df['Overdue Balance'] >= 100)
]
outputpath = r'C:\Users\FutureX\Occupant_Accounts\To Be pass to Third Party Collection Agent.csv'  
Pass_to.to_csv(outputpath, index=False)

df = pd.read_csv(r'C:\Users\FutureX\Occupant_Accounts\File.csv')
DCA_df = df[df['Credit Control Status'] == 'Outsourced to DCA']

outputpath = r'C:\Users\FutureX\Occupant_Accounts\Outsourced.csv'  
DCA_df.to_csv(outputpath, index=False)

df = pd.read_csv(r'C:\Users\FutureX\Occupant_Accounts\File.csv')
Hardship_df = df[df['Credit Control Status'].str.contains('Payment Plan|Smooth Pay|Hardship', case=False, na=False)]
outputpath = r'C:\Users\FutureX\Occupant_Accounts\Hardship.csv'  
Hardship_df.to_csv(outputpath, index=False)

df = pd.read_csv(r'C:\Users\FutureX\Occupant_Accounts\File.csv')
Call_df = df[
    (df['Credit Control Status'].str.contains('Standard credit control procedures'))]
outputpath = r'C:\Users\FutureX\Occupant_Accounts\Call List.csv'  
Call_df.to_csv(outputpath, index=False)
