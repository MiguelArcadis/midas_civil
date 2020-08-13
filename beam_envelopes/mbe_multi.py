import pandas as pd
import numpy as np
import sys
pd.options.mode.chained_assignment = None

#Reads excel file and creats input dataframe
file_path_in = r"C:\Users\silvam3530\ARCADIS\ASC HS2 Project Portal - 03 - Project Execution\07 - C3 Section\05 - Boddington Cutting\01 - BR - Bridges\Banbury Rd OB (Boddington)\02 Working\01 Structures\03 Calcs\07. Piles\Pier Piles" + "\\"
file_name_in = "Forces"
extension_in = ".xlsx"

forces_to_extract = ['Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz',
					 '[SRSS] Fx, My', '[SRSS] Fx, Mz', '[SRSS] My, Mz']

try:
	df_in = pd.read_excel(file_path_in + file_name_in + extension_in)
except FileNotFoundError:
	print("The file selected could not be found. Please make sure you have the correct path, file name, and extension.")
	sys.exit()

df_out = pd.DataFrame()

try:
	#Process dataframe to extract combination, loadcase, part and node columns. Renames all cols at the end
	df_in[['Element Part','Node']] = df_in['Part'].str.replace(r']$', '').str.split('[', expand = True)
	df_in[['Combination','Loadcase']] = df_in['Load'].str.split('_', expand = True)
	df_in.drop(['Load', 'Part'], axis=1, inplace=True)
	df_in = df_in.rename(columns={'Elem': 'Element', 'Element Part': 'Part',
								  'Axial (kN)': 'Fx', 'Shear-y (kN)': 'Fy', 'Shear-z (kN)': 'Fz',
								  'Torsion (kN*m)': 'Mx', 'Moment-y (kN*m)': 'My', 'Moment-z (kN*m)': 'Mz'})
	#Adds a new column to display leading force to user on the output
	df_in["Leading Force"] = np.nan
	#Re-orders the dataframe
	df_in = df_in[['Element', 'Part', 'Node', 'Combination', 'Loadcase', 'Leading Force',
				   'Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']]
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
	df_out = df_out[['Element', 'Part', 'Node', 'Combination', 'Loadcase', 'Leading Force',
				   	 'Fx', 'Fy', 'Fz', 'Mx', 'My', 'Mz']]

except ValueError:
	try:
		#Process dataframe to extract combination, loadcase, part and node columns. Renames all cols at the end
		df_in[['Element Part','Node']] = df_in['Part'].str.replace(r']$', '').str.split('[', expand = True)
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

	except:
		print("Something went wrong. Check files contents are correct and that you've deleted empty first row.")
		sys.exit()


except:
	print("Something went wrong. Check files contents are correct and that you've deleted empty first row.")

finally:
	# Creates and saves outputfile
	file_path_out = file_path_in
	file_name_out = file_name_in + "_BE"
	extension_out = extension_in
	df_out.to_excel(file_path_out + file_name_out + extension_out)