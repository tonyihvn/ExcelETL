import pandas as pd

def transform_excel(old_file_path, output_path, mapping, new_headers):
    """
    Transforms an Excel file based on a many-to-one column mapping.

    Args:
        old_file_path (str): Path to the source Excel file.
        output_path (str): Path to save the transformed Excel file.
        mapping (dict): A dictionary mapping old column names to new column names. 
                        Example: {'Source A': 'id', 'Source B': 'id', 'Source C': 'location'}
        new_headers (list): A list of all headers for the output file.
    """
    # Read and concatenate all sheets from the source file
    xls = pd.ExcelFile(old_file_path)
    if len(xls.sheet_names) > 1:
        all_sheets_data = [pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names]
        old_data = pd.concat(all_sheets_data, ignore_index=True)
    else:
        old_data = pd.read_excel(old_file_path)

    # Create an empty DataFrame with the desired new headers
    new_data = pd.DataFrame(columns=new_headers)

    # Invert the mapping to group old columns by their target new column
    # inverted_mapping will be like: {'id': ['Source A', 'Source B'], 'location': ['Source C']}
    inverted_mapping = {}
    for old_col, new_col in mapping.items():
        if new_col not in inverted_mapping:
            inverted_mapping[new_col] = []
        inverted_mapping[new_col].append(old_col)

    # Process each new column
    for new_col, old_cols in inverted_mapping.items():
        # Filter out old columns that don't actually exist in the source data
        valid_old_cols = [col for col in old_cols if col in old_data.columns]
        if not valid_old_cols:
            continue

        # Create a new series for the new column
        # Use .bfill(axis=1) to find the first non-null value from the right across the selected columns
        # This effectively merges the data from multiple source columns into one
        new_series = old_data[valid_old_cols].bfill(axis=1).iloc[:, 0]
        new_data[new_col] = new_series

    # Ensure all new headers are present, even if they were not mapped
    for col in new_headers:
        if col not in new_data.columns:
            new_data[col] = pd.Series(dtype='object')

    # Reorder columns to match the specified new_headers list
    new_data = new_data[new_headers]

    # Save the transformed data
    new_data.to_excel(output_path, index=False)