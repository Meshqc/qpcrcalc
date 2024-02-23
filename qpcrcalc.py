import pandas as pd

excel_file = '2024-02-15_112154)2ndrun.xls' #file should be in same folder as qpcrcalc.py
sheet_name = 'Results'
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Select relevant columns
relevant_columns = ['Sample Name', 'Target Name', 'CT']
df = df[relevant_columns]

# Create a dictionary to store the mapping of Sample Name to Target Name and CT
sample_mapping = {}

# Iterate through each row in the dataframe
for index, row in df.iterrows():
    sample_name = row['Sample Name']
    target_name = row['Target Name']
    ct_value = row['CT']
    
    # Check if sample_name already exists in the dictionary
    if sample_name in sample_mapping:
        sample_mapping[sample_name].append((target_name, ct_value))
    else:
        sample_mapping[sample_name] = [(target_name, ct_value)]

# Sort the sample names alphabetically
sorted_sample_names = sorted(sample_mapping.keys())

# Create lists to store printed data
output_data = []

# Perform calculations for each sample
for sample_name in sorted_sample_names:
    sample_data = {"Sample Name": sample_name}
    hprt_ct = None
    for target, ct in sample_mapping[sample_name]:
        if target == 'HPRT':  # Assuming 'HPRT' is the control target
            hprt_ct = ct
            sample_data["HPRT"] = ct  # Store HPRT CT value
            continue  # Skip further processing for HPRT
        if hprt_ct is not None:
            delta_ct = ct - hprt_ct
            result = round(2 ** (-delta_ct), 5)  # Round to 5 digits after decimal
            sample_data[target] = result  # Store result for the target
    output_data.append(sample_data)

# Create a DataFrame from the printed data
df_output = pd.DataFrame(output_data)

# Save DataFrame to Excel
output_excel_file = 'output_data.xlsx'  # Specify the output file name
df_output.to_excel(output_excel_file, index=False) #new output file will be extracted in same folder as qpcrcalc.py file 
