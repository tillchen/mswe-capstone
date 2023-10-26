import pandas as pd
from docx import Document

# 1. read csv file
def read_csv_file(file_path):
    try:
        # Read the CSV file into a DataFrame
        df = pd.read_csv(file_path)
        # Keep only the specified fields (columns)
        fields = ["Title", "Funder", "Deadline", "Amount", "Eligibility", "Abstract", "More Information"]
        filtered_df = df[fields]
        # Return the DataFrame
        return filtered_df
    except FileNotFoundError:
        print("File not found!")
        return None
    except KeyError:
        print("Some specified fields do not exist in the CSV file!")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# 2. convert the csv file to word file
def df_to_word(data_frame, file_path):
    # Create a new Word document
    doc = Document()

    # Add a table to the document with the same number of columns as the DataFrame
    table = doc.add_table(rows=1, cols=len(data_frame.columns))

    # Get the cells of the header row in the table
    hdr_cells = table.rows[0].cells

    # Fill in the header row with the column names of the DataFrame
    for i, column_name in enumerate(data_frame.columns):
        hdr_cells[i].text = column_name

    # Loop over each row in the DataFrame
    for index, row in data_frame.iterrows():
        # Add a new row to the table
        row_cells = table.add_row().cells
        # Fill in the cells of the row with the values from the DataFrame
        for i, value in enumerate(row):
            row_cells[i].text = str(value)

    # Save the Word document to the specified file path
    doc.save(file_path)

if __name__ == "__main__":
    # 1. read csv file
    file_path = "sample_data/opps_export.csv"
    data_frame = read_csv_file(file_path)
    if data_frame is not None:
        print(data_frame.head())  # Print first 5 rows of the DataFrame

        # 2. convert csv file to word file
        word_file_path = "output_word/output.docx"
        df_to_word(data_frame, word_file_path)
