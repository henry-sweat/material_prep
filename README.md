# Material Prep
This Python script is designed to process raw material orders. It reads order files and a Matrix spreadsheet, merges them, calculates estimated total hours, handles machine overflow, and finally exports the data to excel files.

## Getting Started
### Prerequisites
* Python 3.6 or Above
* pandas library
* openpyxl library

You can install the necessary packages using pip:

```python
pip install pandas openpyxl
```

### Configuration
Create a **config.py** file in the same directory as the main script with the following parameters:

```python
DIRECTORY_PATH = 'path/to/your/order/files/'
PRODUCTS = ['Product1', 'Product2', 'Product3']
```

### Usage
To run the script, simply execute it from the command line:

```python
python main.py
```

The script will prompt you to enter the material order date in the format **MM-DD-YYYY**. It then locates the relevant order files in the directory specified in **config.py**, combines these into a single pandas DataFrame, and processes the orders according to certain logic rules.

The processed order data is finally exported to individual excel files per machine, within a new folder named after the order date under 'output/' directory.

### Code Breakdown
This Python script involves the following:
1. get_order_date(): Captures user input for order date, validating the format.

2. get_order_files(arg_directory_path, arg_order_date): Returns a list of order files that match the given date. 

3. reassign_overflow(arg_df_machine, arg_overflow_machine): Reassigns tasks to overflow machines if a machine's total estimated hours exceed a given threshold.

4. Main execution:

   * Reads and combines data from order files and a Matrix spreadsheet.
   * Calculates 'Est. Total Hrs' and sorts the orders.
   * Processes each machine's workload and reassigns tasks to overflow machines as necessary.
   * Writes the processed data to separate excel files for each machine.

### Data Format
* Order files are expected to be CSV files with at least the columns 'Part Number', 'Description', 'Quantity' present.
* The Matrix spreadsheet ('matrix.xlsx') should contain at least the columns 'Part Number', 'Product', 'Run Time'.
* The final excel files contain the columns 'Part Number', 'Description', 'Quantity', 'Est. Total Hrs', 'MFG QTY', 'Sign-off'.

Please note that the machine overflow handling is currently hardcoded in the constant **MACHINE_OVERFLOW**. For extended use or for handling a different set of machines, this would need to be updated accordingly.

### Limitations
The script assumes that all input files are properly formatted and error-free, and does not handle possible exceptions that might occur when reading these files.