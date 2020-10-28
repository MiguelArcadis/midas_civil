import pandas as pd
import numpy as np
import re
import sys
pd.options.mode.chained_assignment = None


def extract_beam_results_per_combination(df_in, forces_to_extract, elements_to_extract, element_parts_to_extract):
	df_out = pd.DataFrame()
	#Process dataframe to extract combination, loadcase, part and node columns. Renames all cols at the end
	df_in['Element Part'] = df_in['Part'].apply(lambda item: re.sub(r'\[.*?\]', '', item))
	df_in['Node'] = df_in['Part'].apply(lambda item: re.sub(r'(.*?)\[(.*?)\]', lambda x: str(x.group(2)), item))
	df_in['Combination'] = df_in['Load'].apply(lambda item: re.sub(r'^(.*?)\_(.*?)$', lambda x: str(x.group(1)), item))
	df_in['Loadcase'] = df_in['Load'].apply(lambda item: re.sub(r'^(.*?)\_(.*?)$', lambda x: str(x.group(2)), item))

	df_in.drop(['Load', 'Part'], axis=1, inplace=True)
	df_in = df_in.rename(columns={'Elem': 'Element', 'Element Part': 'Part',
								  'Axial (kN)': 'Fx', 'Shear-y (kN)': 'Fy', 'Shear-z (kN)': 'Fz',
								  'Torsion (kN*m)': 'Mx', 'Moment-y (kN*m)': 'My', 'Moment-z (kN*m)': 'Mz'})

	#Adds a new column to display leading force to user on the output
	df_in["Leading Force"] = np.nan
	#Re-orders the dataframe
	df_in = df_in[['Element', 'Part', 'Node', 'Combination', 'Loadcase', 'Leading Force', 'Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']]
	#Adds new columns to extract possible critical results that are not min and max
	df_in['[SRSS] Fx My'] = (df_in['Fx'] ** 2 + df_in['My'] ** 2) ** 0.5
	df_in['[SRSS] Fx Mz'] = (df_in['Fx'] ** 2 + df_in['Mz'] ** 2) ** 0.5
	df_in['[SRSS] My Mz'] = (df_in['My'] ** 2 + df_in['Mz'] ** 2) ** 0.5

	unique_combinations = list(df_in['Combination'].unique())

	# Creates dict to map parts to extract from "Part   i" to "I" etc
	mapped_element_parts_to_extract = {"Part   i": "I", "Part 1/4": "1/4", "Part 2/4": "2/4", "Part 3/4": "3/4", "Part   j": "J"}
	element_parts_to_extract = [mapped_element_parts_to_extract[k] for k in element_parts_to_extract]

	for element_part_idx, element in enumerate(elements_to_extract):
		for combination in unique_combinations:
			df_temp = df_in[(df_in['Element'] == element) & (df_in['Part'] == element_parts_to_extract[element_part_idx]) & (df_in['Combination'] == combination)]
			for force in forces_to_extract:
				row_max = df_temp[force].idxmax()
				row_min = df_temp[force].idxmin()
				if 'SRSS' in force:
					df_temp['Leading Force'][row_max] = "[MAX] " + force
					df_out = df_out.append(df_temp.loc[[row_max]])
				else:
					df_temp['Leading Force'][row_max] = "[MAX] " + force
					df_temp['Leading Force'][row_min] = "[MIN] " + force
					df_out = df_out.append(df_temp.loc[[row_max]])
					df_out = df_out.append(df_temp.loc[[row_min]])
	# Makes sure columns with SRSS are not given back to user
	df_out = df_out[['Element', 'Part', 'Node', 'Combination', 'Loadcase', 'Leading Force', 'Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']]
	return df_out

def extract_elink_results_per_combination(df_in, forces_to_extract):
	df_out = pd.DataFrame()
	#Process dataframe to extract combination, loadcase, part and node columns. Renames all cols at the end
	df_in['Combination'] = df_in['Load'].apply(lambda item: re.sub(r'^(.*?)\_(.*?)$', lambda x: str(x.group(1)), item))
	df_in['Loadcase'] = df_in['Load'].apply(lambda item: re.sub(r'^(.*?)\_(.*?)$', lambda x: str(x.group(2)), item))

	df_in.drop(['Load'], axis=1, inplace=True)
	df_in = df_in.rename(columns={'Elem': 'Element',
								  'Axial (kN)': 'Fx', 'Shear-y (kN)': 'Fy', 'Shear-z (kN)': 'Fz',
								  'Torsion (kN*m)': 'Mx', 'Moment-y (kN*m)': 'My', 'Moment-z (kN*m)': 'Mz'})
	#Adds a new column to display leading force to user on the output
	df_in["Leading Force"] = np.nan
	#Re-orders the dataframe
	df_in = df_in[['Element', 'Node', 'Combination', 'Loadcase', 'Leading Force', 'Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']]
	#Adds new columns to extract possible critical results that are not min and max
	df_in['[SRSS] Fx, My'] = (df_in['Fx'] ** 2 + df_in['My'] ** 2) ** 0.5
	df_in['[SRSS] Fx, Mz'] = (df_in['Fx'] ** 2 + df_in['Mz'] ** 2) ** 0.5
	df_in['[SRSS] My, Mz'] = (df_in['My'] ** 2 + df_in['Mz'] ** 2) ** 0.5

	unique_elements = list(df_in['Element'].unique())
	unique_combinations = list(df_in['Combination'].unique())

	for element in unique_elements:
		for combination in unique_combinations:
			df_temp = df_in[(df_in['Element'] == element) & (df_in['Combination'] == combination)]
			for force in forces_to_extract:
				row_max = df_temp[force].idxmax()
				row_min = df_temp[force].idxmin()
				if 'SRSS' in force:
					df_temp['Leading Force'][row_max] = "[MAX] " + force
					df_out = df_out.append(df_temp.loc[[row_max]])
				else:
					df_temp['Leading Force'][row_max] = "[MAX] " + force
					df_temp['Leading Force'][row_min] = "[MIN] " + force
					df_out = df_out.append(df_temp.loc[[row_max]])
					df_out = df_out.append(df_temp.loc[[row_min]])
	# Makes sure columns with SRSS are not given back to user
	df_out = df_out[['Element', 'Node', 'Combination', 'Loadcase', 'Leading Force', 'Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']]
	return df_out


def extract_composite_beam_stresses_per_combination(df_in, forces_to_extract):
	df_out = pd.DataFrame()

	#Process dataframe to extract combination, loadcase, part and node columns. Renames all cols at the end
	df_in['Element Part'] = df_in['Part']
	df_in['Combination'] = df_in['Load'].apply(lambda item: re.sub(r'^(.*?)\_(.*?)$', lambda x: str(x.group(1)), item))
	df_in['Loadcase'] = df_in['Load'].apply(lambda item: re.sub(r'^(.*?)\_(.*?)$', lambda x: str(x.group(2)), item))
	df_in.drop(['Load', 'Part', 'DOF', 'Axial (kN/m^2)', 'Bend(+y)(kN/m^2)', 'Bend(-y)(kN/m^2)', 'Bend(+z)(kN/m^2)', 'Bend(-z)(kN/m^2)', 'Cb(min/max)(kN/m^2)'], axis=1, inplace=True)

	df_in = df_in[['Element', 'Section Part', 'Element Part', 'Combination', 'Loadcase', 'Cb1(-y+z)(kN/m^2)', 'Cb2(+y+z)(kN/m^2)', 'Cb3(+y-z)(kN/m^2)', 'Cb4(-y-z)(kN/m^2)']]

	df_in = df_in.rename(columns={'Cb1(-y+z)(kN/m^2)':'TopLeft (MPa)', 'Cb2(+y+z)(kN/m^2)': 'TopRight (MPa)', 'Cb3(+y-z)(kN/m^2)': 'BottomRight (MPa)', 'Cb4(-y-z)(kN/m^2)': 'BottomLeft (MPa)'})

	df_in["Location"] = np.nan

	unique_elements = list(df_in['Element'].unique())
	unique_combinations = list(df_in['Combination'].unique())
	unique_section_parts = [1, 2]
	stress_locs_to_extract = ['TopLeft (MPa)', 'TopRight (MPa)', 'BottomRight (MPa)', 'BottomLeft (MPa)']

	section_parts_dict = {1: 'Steel', 2: 'Concrete'}

	for element in unique_elements:
		for combination in unique_combinations:
			for part in unique_section_parts:
				for stress_loc in stress_locs_to_extract:
					df_temp = df_in[(df_in['Element'] == element) & (df_in['Combination'] == combination) & (df_in['Section Part'] == part)]

					row_max = df_temp[stress_loc].idxmax()
					row_min = df_temp[stress_loc].idxmin()
					df_temp['Location'][row_max] = "[MAX] " + stress_loc
					df_temp['Location'][row_min] = "[MIN] " + stress_loc
					df_out = df_out.append(df_temp.loc[[row_max]])
					df_out = df_out.append(df_temp.loc[[row_min]])

					

					with pd.option_context('display.max_rows', None, 'display.max_columns', None):
						print(df_temp.head())

	df_out['Section Part'] = df_out['Section Part'].map(section_parts_dict)
	df_out = df_out[['Element', 'Section Part', 'Element Part', 'Combination', 'Loadcase', 'Location', 'TopLeft (MPa)', 'TopRight (MPa)', 'BottomRight (MPa)', 'BottomLeft (MPa)']]

	for stress_loc in stress_locs_to_extract:
		df_out[stress_loc] = df_out[stress_loc].div(1000).round(1)

	return df_out


def extract_beam_combinationless_results(df_in, forces_to_extract):
	df_out = pd.DataFrame()
	#Process dataframe to extract combination, loadcase, part and node columns. Renames all cols at the end
	df_in['Element Part'] = df_in['Part'].apply(lambda item: re.sub(r'\[.*?\]', '', item))
	df_in['Node'] = df_in['Part'].apply(lambda item: re.sub(r'(.*?)\[(.*?)\]', lambda x: str(x.group(2)), item))
	df_in.drop(['Part'], axis=1, inplace=True)
	df_in = df_in.rename(columns={'Elem': 'Element', 'Element Part': 'Part',
								  'Axial (kN)': 'Fx', 'Shear-y (kN)': 'Fy', 'Shear-z (kN)': 'Fz',
							  	  'Torsion (kN*m)': 'Mx', 'Moment-y (kN*m)': 'My', 'Moment-z (kN*m)': 'Mz'})
	#Adds a new column to display leading force to user on the output
	df_in["Leading Force"] = np.nan
	#Re-orders the dataframe
	df_in = df_in[['Element', 'Part', 'Node', 'Load', 'Leading Force',
			   	   'Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']]
	#Adds new columns to extract possible critical results that are not min and max
	df_in['[SRSS] Fx, My'] = (df_in['Fx'] ** 2 + df_in['My'] ** 2) ** 0.5
	df_in['[SRSS] Fx, Mz'] = (df_in['Fx'] ** 2 + df_in['Mz'] ** 2) ** 0.5
	df_in['[SRSS] My, Mz'] = (df_in['My'] ** 2 + df_in['Mz'] ** 2) ** 0.5

	unique_elements = list(df_in['Element'].unique())

	for element in unique_elements:
		df_temp = df_in[(df_in['Element'] == element)]
		for force in forces_to_extract:
			row_max = df_temp[force].idxmax()
			row_min = df_temp[force].idxmin()
			if 'SRSS' in force:
				df_temp['Leading Force'][row_max] = "[MAX] " + force
				df_out = df_out.append(df_temp.loc[[row_max]])
			else:
				df_temp['Leading Force'][row_max] = "[MAX] " + force
				df_temp['Leading Force'][row_min] = "[MIN] " + force
				df_out = df_out.append(df_temp.loc[[row_max]])
				df_out = df_out.append(df_temp.loc[[row_min]])
	# Makes sure columns with SRSS are not given back to user
	df_out = df_out[['Element', 'Part', 'Node', 'Load', 'Leading Force',
					 'Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']]
	print("Warning: Loadcases have not been defined using the _ separator. All loads are treated as belonging to a single combination.")

	return df_out
