import pandas as pd
import numpy as np
import itertools
from load_combs import combination_functions as mlc_func
pd.set_option('display.max_columns', 20)

#Reads excel file and creats input dataframe
file_path_in = r"C:\Users\silvam3530\Desktop" + "\\"
file_name_in = "BRB - Load Comb 4"
extension_in = ".xlsx"

file_path_out = file_path_in
file_name_out = file_name_in
extension_out = extension_in

#Creates and processes input dataframe to rename columns and fill blanks
xls_file = pd.ExcelFile(file_path_in + file_name_in + extension_in)
df_in = pd.read_excel(xls_file, sheet_name = 'Combinations')
df_in = df_in.fillna(method='ffill', axis=1)
df_in.columns = pd.Series([np.nan if 'Unnamed:' in x else x for x in df_in.columns.values]).ffill().values.flatten()

#Creates output dataframe with unique columns and clears all rows of content except first row that contains type of Loadcase for Midas
df_out = pd.DataFrame()
df_out = df_in.loc[:,~df_in.columns.duplicated()]
df_out = df_out.iloc[:2]

list_of_comb_types = df_in.iloc[0].values.tolist()[1:][::2]

#Places all rows of the input dataframe in a list of lists.
list_of_combs = df_in.iloc[2:].values.tolist()

list_of_all_permutations = []

for comb in list_of_combs:
	comb_permutations = []
	initial_comb_name = []
	initial_comb_name.append(comb[0]) #Saves comb name in a list for later naming
	
	del comb[0] #Deletes comb name from the other list for data processing
	index_main_variable = None

	# Finds main variable and its index, removes the marking x and makes that values a float
	for index, item in enumerate(comb):
		if isinstance(item, str) == True:
			index_main_variable = int(index / 2)
			comb[index] = float(item[1:])

	# Runs all the functions to create combinations and filter out unnecessary results

	comb_upper, comb_lower = mlc_func.zigzag(comb)

	current_list = [None] * min(len(comb_upper), len(comb_lower))
	mlc_func.get_combinations(current_list, comb_upper, comb_lower, 0, comb_permutations)
	mlc_func.filter_permanent_groups(comb_permutations, list_of_comb_types)
	mlc_func.filter_main_variable_equal_0(comb_permutations, index_main_variable)
	comb_permutations = mlc_func.remove_duplicates(comb_permutations)

	# Adds a new column with the combination name and returns the new dataframe
	final_comb_name = []
	for i in range(len(comb_permutations)):
	    final_comb_name.append(initial_comb_name[0] + '|' + str(i + 1))

	for i, item in enumerate(comb_permutations):
		item.insert(0, final_comb_name[i])
	
	list_of_all_permutations.extend(comb_permutations)


df_out = df_out.append(pd.DataFrame(list_of_all_permutations, columns = df_out.columns))

mct_string = mlc_func.df_to_mct(df_out)

mlc_func.combs_list_to_excel(df_out, file_path_out, file_name_out, extension_out)

text_file = open(file_path_out + file_name_out + '_MCT_shell' + ".txt", "w")
text_file.write(mct_string)
text_file.close()


df_out.to_excel(file_path_out + file_name_out + '_MLC1' + extension_out)
