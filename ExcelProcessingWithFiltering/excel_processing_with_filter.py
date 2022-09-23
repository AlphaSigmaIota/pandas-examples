# import libraries
import pandas as pd

EXCEL_FILE = "processing_file.xlsx"
SHEET_NAME = "data"

# read data from xlsx from a given sheet into a data frame
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

# filter data where column "Done" has not an X as value
series_unhandled_data = df.query("Done!='X'")
print(f"Length of whole data: {df.shape[0]}")
print(f"Length of undone data: {len(series_unhandled_data)}")

# loop trough series with the unhandled data
for i in range(0, len(series_unhandled_data)):
    # get the row index of the data
    row_index = series_unhandled_data.index[i]

    # get the values of the columns you want to work with
    first_name = df.FirstName.T[row_index]
    last_name = df.LastName.T[row_index]

    print(f"Firstname: {first_name} - Lastname: {last_name}")
    df.Done.T[row_index] = "X"

    # save dataframe to excel
    df.to_excel(EXCEL_FILE.replace(".xlsx", "_done.xlsx"), sheet_name=SHEET_NAME, index=False)
