import pandas as pd
import numpy as np
import sys
from force_results import process_force_results as pfr
pd.options.mode.chained_assignment = None

#Reads excel file and creats input dataframe
file_path_in = r"C:\Users\silvam3530\ARCADIS\ASC HS2 Project Portal - 03 - Project Execution\07 - C3 Section\05 - Boddington Cutting\01 - BR - Bridges\Banbury Rd OB (Boddington)\02 Working\01 Structures\03 Calcs\20. SSI - Geotech\2020-08-17" + "\\"
file_name_in = "Loads from Midas"
extension_in = ".xlsx"

# Possible forces to extract = ['Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz','[SRSS] Fx, My', '[SRSS] Fx, Mz', '[SRSS] My, Mz']
# Copy and paste below as needed
forces_to_extract = ['Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']

# Load file from path above 
try:
	df_in = pd.read_excel(file_path_in + file_name_in + extension_in)
except FileNotFoundError:
	print("The file selected could not be found. Please make sure you have the correct path, file name, and extension.")

df_out = pd.DataFrame()

if 'Component' in list(df_in.head().columns):
	print("Concurrent results are being extracted.")
else:
	print("You have not provided concurrent results. This can be done via the menu 'View by Max Value Item...' in Midas.")

has_separator = all(list(df_in['Load'].str.contains('_')))

if has_separator == True:
	df_out = pfr.extract_results_per_combination(df_in, forces_to_extract)
else:
	print("You have not provided the '_' separator for combinations. All loadcases will be treated has not belonging to any combination.")
	df_out = pfr.extract_combinationless_results(df_in, forces_to_extract)

"""
# Creates and saves outputfile
file_path_out = file_path_in
file_name_out = file_name_in + "_MinMax"
extension_out = extension_in
df_out.to_excel(file_path_out + file_name_out + extension_out)
"""