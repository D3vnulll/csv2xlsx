import pandas as pd
from tkinter import Tk, filedialog

def select_csv_file():
    root = Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(title="Select CSV file", filetypes=[("CSV files", "*.csv")])

    if not file_path:
        print("No file selected. Exiting.")
        return None

    return file_path

def get_custom_excel_name():
    custom_name = input("Enter a custom name for the output Excel file (without extension): ")
    return custom_name

def select_output_path():
    root = Tk()
    root.withdraw()  # Hide the main window

    output_path = filedialog.askdirectory(title="Select output directory")

    if not output_path:
        print("No output directory selected. Using the current working directory.")
        return ''

    return output_path

def csv_to_excel(csv_file, excel_file):
    # Read the CSV file into a Pandas DataFrame
    df = pd.read_csv(csv_file, encoding='utf-8')

    # Write the DataFrame to an Excel file
    df.to_excel(excel_file, index=False)

    print(f"Conversion successful. CSV file '{csv_file}' has been converted to Excel file '{excel_file}'.")

def main():
    # Allow the user to select the input CSV file
    csv_file_path = select_csv_file()

    if csv_file_path:
        # Get a custom name for the output Excel file
        custom_excel_name = get_custom_excel_name()
        if not custom_excel_name:
            custom_excel_name = 'output'

        # Allow the user to select the output directory
        output_directory = select_output_path()

        # Specify the output Excel file
        excel_file_path = f"{output_directory}/{custom_excel_name}.xlsx" if output_directory else f"{custom_excel_name}.xlsx"

        # Convert CSV to Excel
        csv_to_excel(csv_file_path, excel_file_path)

if __name__ == "__main__":
    main()
