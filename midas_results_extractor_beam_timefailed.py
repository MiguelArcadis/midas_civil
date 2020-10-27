import pywinauto as pwa
from pywinauto import keyboard
from pywinauto.timings import always_wait_until
from time import sleep
import pandas as pd
import numpy as np
import ctypes
from force_results import process_force_results as pfr

file_path_in = r"C:\Users\silvam3530\ARCADIS\ASC HS2 Project Portal - 03 - Project Execution\07 - C3 Section\05 - Boddington Cutting\01 - BR - Bridges\Banbury Rd OB (Boddington)\02 Working\01 Structures\02 Models\03. Combs&Results" + "\\"
file_name_in = "Results Extractor"
extension_in = ".xlsx"

# File path to save results files (processed loads)
file_path_out = file_path_in
file_name_out = "Extracted Results"
extension_out = ".xlsx"

source_df = pd.read_excel(file_path_in + file_name_in + extension_in, sheet_name='Extract Results')

source_df.dropna(axis=0, how='all', inplace=True)

# Midas elements to extract
# Midas elements to extract
source_df['Element no.'] = source_df['Element no.'].astype(int)
ELEMENTS_TO_EXTRACT = source_df['Element no.'].tolist()

# Midas parts to extract
ELEMENT_PARTS_TO_EXTRACT = source_df['Node'].tolist()

# Midas forces to extract
all_force_cols_names = ['Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']
force_extract = []
for col_name in all_force_cols_names:
	if not source_df[col_name].isnull().values.all():
		force_extract.append(col_name)
FORCES_TO_EXTRACT = list(dict.fromkeys(force_extract))

# Midas combinations to extract
comb_extract = []
list_comb_keys = ['Env 1', 'Env 2', 'Env 3', 'Env 4', 'Env 5', 'Env 6', 'Env 7', 'Env 8', 'Env 9', 'Env 10']

for index, row in source_df['Element no.'].items():
	for key in list_comb_keys:
		if pd.isna(source_df[key].iloc[index]) == False:
			current_comb = source_df[key].iloc[index]
			idx = combs_df[combs_df['Combination Envelopes'] == current_comb].index
			comb_extract.append(combs_df['Combinations'].loc[idx].values[0].replace("'","").replace(" ","").split(","))

COMBS_TO_EXTRACT = [item for sublist in comb_extract for item in sublist]
COMBS_TO_EXTRACT = list(dict.fromkeys(COMBS_TO_EXTRACT))

# Create {Element no: Element Name} dictionary
elem_names = source_df['Element name'].tolist()
elem_num_names = dict(zip(elem_extract, elem_names))

# Midas forces to extract
all_force_cols_names = ['Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']
force_extract = []
for col_name in all_force_cols_names:
	if not source_df[col_name].isnull().values.all():
		force_extract.append(col_name)
FORCES_TO_EXTRACT = list(dict.fromkeys(force_extract))


column_names = ["Elem", "Load", "Part", "Component", "Axial (kN)", "Shear-y (kN)", "Shear-z (kN)", "Torsion (kN*m)", "Moment-y (kN*m)", "Moment-z (kN*m)"]

df_results = pd.read_clipboard(header=None, names=column_names)

# Extracts results per combination from Midas table containing all results
df_out = pfr.extract_beam_results_per_combination(df_midas_results, FORCES_TO_EXTRACT, ELEMENTS_TO_EXTRACT, ELEMENT_PARTS_TO_EXTRACT)

# Adds a new column to df_out to map the element no. to element name. Ease of use by user
df_out.insert(loc=0, column='Name', value=np.NaN)
df_out['Name'] = df_out['Element'].map(elem_num_names)

# Saves to desired path
df_out.to_excel(file_path_out + file_name_out + extension_out)

# Write success message when program has finished
ctypes.windll.user32.MessageBoxW(0, "Results have been successfully extracted and saved at your chosen location.", "Midas Auto Extract", 0x1000)

