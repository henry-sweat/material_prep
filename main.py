
import pandas as pd
from datetime import datetime
from pathlib import Path
import os
import config


# CONSTANTS
DATE_FORMAT = '%m-%d-%Y'
MAX_HOURS = 40.0
MACHINE_OVERFLOW = {'B7': 'D25', 'D24': 'C17'}
OUTPUT_FOLDER_PATH = 'output/'


# DEFINITIONS
def get_order_date():
    while True:
        date_input = input("Please enter the material order date (MM-DD-YYYY): ")
        try:
            return datetime.strptime(date_input, DATE_FORMAT).date()
        except ValueError:
            print("That's not a valid date. Please try again.")


def get_order_files(arg_directory_path, arg_order_date):
    # Create a list to store file details
    file_details = []
    # Iterate over each file
    for file_path in Path(arg_directory_path).glob('*'):
        # Get the file modification timestamp
        modified_timestamp = file_path.stat().st_mtime
        # Convert the timestamp to a datetime object
        modified_date = datetime.fromtimestamp(modified_timestamp)
        # Check if the modified date matches the order date
        if modified_date.date() == arg_order_date:
            # Append the file details to the list
            file_details.append(str(arg_directory_path) + file_path.name)
    return file_details


def reassign_overflow(arg_df_machine, arg_overflow_machine):
    # Iterate over each row in df_machine to see if sum < MAX_HOURS
    for idx in range(0, len(arg_df_machine) - 1):
        if arg_df_machine['Est. Total Hrs'].iloc[0:idx].sum() < MAX_HOURS:
            continue
        # If sum is > MAX_HOURS, then change machine of remaining rows to overflow machine
        arg_df_machine.iloc[idx:len(arg_df_machine), 3] = arg_overflow_machine
        break
    return arg_df_machine


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # Get user input for order date
    order_date = get_order_date()

    # Get order files based on the order date
    order_files = get_order_files(config.DIRECTORY_PATH, order_date)

    # Combine order files into single DataFrame
    df_order = pd.concat([pd.read_csv(file_path,
                                      usecols=[1, 2, 3],
                                      names=['Part Number', 'Description', 'Quantity'],
                                      skiprows=1) for file_path in order_files], ignore_index=True)

    # Load Matrix Spreadsheet into dataframe
    df_matrix = pd.read_excel('matrix.xlsx', usecols=[1, 2, 3, 6])

    # Merge df_order with df_matrix on Part Number (vlookup)
    df_order = pd.merge(df_order, df_matrix, on='Part Number', how='left')

    # Create Est. Total Hrs column based on qty and runtime
    df_order['Est. Total Hrs'] = (df_order['Quantity'] * df_order['Run Time'] / 60).round(2)
    # Drop Run Time column
    df_order.drop('Run Time', axis=1, inplace=True)
    # Set the order for the product column
    df_order['Product'] = pd.Categorical(df_order['Product'],
                                         categories=[config.PRODUCTS[0], config.PRODUCTS[1], config.PRODUCTS[2]],
                                         ordered=True)
    # Sort by Machine, Product, then Est. Total Hrs
    df_order = df_order.sort_values(by=['Machine', 'Product', 'Est. Total Hrs'],
                                    ascending=[True, True, False]).reset_index(drop=True)

    # Iterate over each machine to assess workload and to offload work, if necessary
    for machine in df_order['Machine'].unique():
        # Filter df_order for machine, sort by est. total hrs descending
        df_machine = df_order[df_order['Machine'] == machine].sort_values(by='Est. Total Hrs', ascending=False)

        # Skip if sum of est hours on machine < MAX_HOURS
        if df_machine['Est. Total Hrs'].sum() < MAX_HOURS:
            continue

        if machine in MACHINE_OVERFLOW:
            # Manual handling of copper overflow by product
            if df_machine[df_machine['Product'] == config.PRODUCTS[0]]['Est. Total Hrs'].sum() < MAX_HOURS:
                df_machine.loc[(df_machine['Product'] == config.PRODUCTS[1]) |
                               (df_machine['Product'] == config.PRODUCTS[2]), 'Machine'] = MACHINE_OVERFLOW[machine]
                df_order.update(df_machine)
                continue
            # If no manual handling required, run overflow function
            df_machine = reassign_overflow(df_machine, MACHINE_OVERFLOW[machine])
            df_order.update(df_machine)

    # Final formatting of df_order
    df_order = df_order.sort_values(by=['Machine', 'Est. Total Hrs'], ascending=[True, False]).reset_index(drop=True)
    df_order['MFG QTY'] = ''
    df_order['Sign-off'] = ''

    # Create the new folder
    os.makedirs(f'{OUTPUT_FOLDER_PATH}{order_date}', exist_ok=True)

    # Export manifests to excel
    for machine in df_order['Machine'].unique():
        df_order[df_order['Machine'] == machine].to_excel(f'{OUTPUT_FOLDER_PATH}{order_date}/{machine}.xlsx',
                                                          columns=['Part Number', 'Description', 'Quantity',
                                                                   'Est. Total Hrs', 'MFG QTY', 'Sign-off'],
                                                          sheet_name=machine,
                                                          index=False, engine='openpyxl')
