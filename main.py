import format_data # Extracts and formats the new DPP Data
import merge_data # Merges new DPP workorders to the master sheet

"""
The main function will call the format_data and merge_data functions
    - This will format the new DPP workorders and append them to the master sheet
    - A temporary copy of the master sheet is created in the merge function
    - It will be deleted if the merge is successful
"""
def execute_script(master_sheet_filepath, raw_dpp_filepath):

    print('Executing script...')
    print('Processing new DPP data...')

    # Extract and format the new DPP data
    formatted_dpp_filepath = format_data.format(raw_dpp_filepath)

    print('Formatting complete. Merging data...')

    # Merge the formatted DPP data into the master sheet
    merge_data.merge(master_sheet_filepath, formatted_dpp_filepath)

    print('Script execution complete.')

if __name__ == '__execute_script__':
    master_sheet_filepath = "workorders_2024-06-19.xlsx"
    raw_dpp_filepath = "test.xlsx"
    execute_script(master_sheet_filepath, raw_dpp_filepath)