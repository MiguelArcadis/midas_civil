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
combs_df = pd.read_excel(file_path_in + file_name_in + extension_in, sheet_name='Combinations')

source_df.dropna(axis=0, how='all', inplace=True)

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
elem_num_names = dict(zip(ELEMENTS_TO_EXTRACT, elem_names))


def midas_parser(elements_to_extract, combs_to_extract, element_parts_to_extract):
	"""
	# Check all running apps to find Midas
	windows = pwa.Desktop(backend="win32").windows()
	running_windows = [window.window_text() for window in windows]

	midas_title = ""
	target_title = "Civil"

	for window in running_windows:
		if target_title in window:
			midas_title += window
		else:
			pass

	# Create Midas app coonection and set focus
	app = pwa.Application().connect(title=midas_title)
	app.wait_cpu_usage_lower(threshold=5, timeout=20, usage_interval=2)
	app[midas_title].set_focus()

	# Open results table in Midas from Midas command menu
	keyboard.send_keys('rrrrrrrrrrrr') # Finds Results Tables
	keyboard.send_keys('{ENTER}') # Press enter on results tables
	keyboard.send_keys('{ENTER}') # Press enter on beam
	keyboard.send_keys('f') # Finds forces for results
	keyboard.send_keys('{ENTER}') # Press enter on force

	# Write the elements to extract
	for e in ELEMENTS_TO_EXTRACT:
		keyboard.send_keys(str(e))
		keyboard.send_keys('{SPACE}')

	# Select combinations to extract
	list_item = app.RecordsActivationDialog.ListBox2
	item_wrapper = list_item.wrapper_object()
	for comb in COMBS_TO_EXTRACT:
		item_wrapper.select(comb).set_focus().type_keys('{VK_SPACE}')

	# Select the element parts 
	preselected_parts = ['Part   i', 'Part   j']

	for item in preselected_parts:
		if item in ELEMENT_PARTS_TO_EXTRACT:
			ELEMENT_PARTS_TO_EXTRACT.remove(item)
		else:
			ELEMENT_PARTS_TO_EXTRACT.append(item)

	for part in ELEMENT_PARTS_TO_EXTRACT:
		app.RecordsActivationDialog.ListBox3.select(part).set_focus().type_keys('{VK_SPACE}')

	# Press 'ok' for Midas to calculate results
	app.RecordsActivationDialog.OK.click()
	
	# Wait up to 10 hours for Midas to process results. Check if finised every 30 sec.
	# Once finished, we right-click to go to the concurrent results table and proceed.
	app.wait_cpu_usage_lower(threshold=5, timeout=36000, usage_interval=5)
	app.MidasGenMainFrmClass.set_focus().click_input(button='right')
	app.wait_cpu_usage_lower(threshold=5, timeout=20, usage_interval=2)

	for _ in range(9):
		keyboard.send_keys('{VK_DOWN}')

	keyboard.send_keys('{ENTER}')

	# Select concurrent components to extract. By default we extract all and press 'ok'
	sleep(1)
	concurrent_components = ['Axial', 'Shear-y', 'Shear-z', 'Torsion', 'Moment-y', 'Moment-z']

	for component in concurrent_components:
		app.ResultViewItems.ListBox.select(component).set_focus().type_keys('{VK_SPACE}')

	app.ResultViewItems.OK.set_focus().click_input(button='left')
	
	
	# Wait up to 10 hours for Midas to extract concurrent results. This can take a while.
	app.wait_cpu_usage_lower(threshold=5, timeout=36000, usage_interval=5)
	app.MidasGenMainFrmClass['Result by Max Value-[Beam Force]'].GxMdiFrame.GXWND.set_focus()
	app.wait_cpu_usage_lower(threshold=5, timeout=20, usage_interval=2)

	# Basically we're selecting the entire table with keyboard and copying into clipboard.
	keyboard.send_keys("{VK_RIGHT}")
	sleep(0.5)
	keyboard.send_keys("{VK_SHIFT down}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_SHIFT up}")
	keyboard.send_keys("{VK_CONTROL down}{VK_SHIFT down}{VK_DOWN}{VK_CONTROL up}{VK_SHIFT up}")

	sleep(25)
	keyboard.send_keys(r'^c')

	sleep(10)
	"""

	# We can't select the table headers so we'll create them manually to use in pandas
	column_names = ["Elem", "Load", "Part", "Component", "Axial (kN)", "Shear-y (kN)", "Shear-z (kN)", "Torsion (kN*m)", "Moment-y (kN*m)", "Moment-z (kN*m)"]
	
	# Create pandas dataframe with clipboard content (table data) and column names we just created and return
	df_results = pd.read_clipboard(header=None, names=column_names)

	return df_results

# Assigns midas parser to variable and runs
df_midas_results = midas_parser(ELEMENTS_TO_EXTRACT, COMBS_TO_EXTRACT, FORCES_TO_EXTRACT)

# Extracts results per combination from Midas table containing all results
df_out = pfr.extract_beam_results_per_combination(df_midas_results, FORCES_TO_EXTRACT, ELEMENTS_TO_EXTRACT, ELEMENT_PARTS_TO_EXTRACT)

# Adds a new column to df_out to map the element no. to element name. Ease of use by user
df_out.insert(loc=0, column='Name', value=np.NaN)
df_out['Name'] = df_out['Element'].map(elem_num_names)

# Saves to desired path
df_out.to_excel(file_path_out + file_name_out + extension_out)

# Write success message when program has finished
ctypes.windll.user32.MessageBoxW(0, "Results have been successfully extracted and saved at your chosen location.", "Midas Auto Extract", 0x1000)