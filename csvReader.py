import pandas as pd

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

# Example usage
if __name__ == "__main__":
    file_path = "sample_data/opps_export.csv"
    data_frame = read_csv_file(file_path)
    if data_frame is not None:
        print(data_frame.head())  # Print first 5 rows of the DataFrame
