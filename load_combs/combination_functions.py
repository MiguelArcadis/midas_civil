import numpy as np
from  more_itertools import unique_everseen
import xlsxwriter

def zigzag(seq):
    return seq[::2], seq[1::2]

def get_combinations(curr_list, list1, list2, index, overall_list):
    if index == len(list1) or index == len(list2):
        overall_list.append(curr_list[:])
        return

    curr_list[index] = list1[index]
    get_combinations(curr_list, list1, list2, index+1, overall_list)
    curr_list[index] = list2[index]
    get_combinations(curr_list, list1, list2, index+1, overall_list)

def filter_permanent_groups(comb_permutations, list_of_comb_types):
    permanent_group_indices = [index for index, comb_type in enumerate(list_of_comb_types) if comb_type == "PERM.g"] #Gets indices of PERMg loadcases
    items_to_delete = []
    
    for permutation_index, permutation in enumerate(comb_permutations):
        value_comparison = []
        for index in permanent_group_indices:
            value_comparison.append(permutation[index])

        permanent_group_unfavourable = all(value >= 1 for value in value_comparison if value != 0)
        permanent_group_favourable = all(value < 1 for value in value_comparison if value != 0)

        if (permanent_group_unfavourable == False and permanent_group_favourable == False):
            items_to_delete.append(permutation_index)

    for item in sorted(items_to_delete, reverse = True):
        del comb_permutations[item]

    return comb_permutations

def filter_main_variable_equal_0(comb_permutations, index_main_variable):
    if index_main_variable == None:
        return comb_permutations
        
    else:
        items_to_delete = []
        for i, permutation in enumerate(comb_permutations):
            if permutation[index_main_variable] == 0:
                items_to_delete.append(i)

        for item in sorted(items_to_delete, reverse = True):
            del comb_permutations[item]
            
        return comb_permutations

def remove_duplicates(comb_permutations):
    unique_comb_permutations = []
    for comb in comb_permutations:
        if comb not in unique_comb_permutations:
            unique_comb_permutations.append(comb)
    return unique_comb_permutations

def df_to_mct(df):
    full_array = np.array(df)

    loadcases = np.array(df.columns.values)[1:]
    combinations = full_array[2:, 0]
    load_types = full_array[1, 1:]
    factors = full_array[2:, 1:]

    mct_string = '*LOADCOMB    ; Combinations\n; NAME=NAME, KIND, ACTIVE, bES, iTYPE, DESC, iSERV-TYPE, nLCOMTYPE, nSEISTYPE   ; line 1\n;      ANAL1, LCNAME1, FACT1, ...                                               ; from line 2'

    #Creates load combinations
    for row_index, row in enumerate(factors):
        combination = combinations[row_index]
        mct_string = mct_string + '\n' + '   NAME=' + str(combination) +', GEN, ACTIVE, 0, 0, , 0, 0, 0\n'

        for factor_index, factor in enumerate(row):
            load_type = load_types[factor_index]
            loadcase = loadcases [factor_index]

            if factor_index != 0:
                mct_string = mct_string + ', '
            mct_string = mct_string + str(load_type) + ', ' + str(loadcase) + ', ' + str(factor)

    #Creates combination envelopes
    combination_envelopes = []
    for comb in combinations:
        combination_envelopes.append(comb.split('_')[0])
        combination_envelopes.append(comb.split('|')[0])

    combination_envelopes = list(unique_everseen(combination_envelopes))
 
    combination_groups = []
    for envelope in combination_envelopes:
        group = []
        for comb in combinations:
            if envelope in comb:
                group.append(comb)
        combination_groups.append(group)

    for group_index, group in enumerate(combination_groups):
        mct_string = mct_string + '\n' + '   NAME=' + str(combination_envelopes[group_index]) +', GEN, ACTIVE, 0, 1, , 0, 0, 0\n'
        for comb_index, comb in enumerate(group):
            if comb_index != 0:
                mct_string = mct_string + ', '
            mct_string = mct_string + 'CB' + ', ' + str(comb) + ', ' + '1.0'

    return mct_string


def list_string_converter(lst):
    final_string = ', '.join("'{0}'".format(l) for l in lst)
    final_string = final_string.rstrip(',')
    return final_string

def combs_list_to_excel(df, file_path_out, file_name_out, extension_out):
    full_array = np.array(df)
    combinations = full_array[2:, 0]

    combination_envelopes = []
    for comb in combinations:
        combination_envelopes.append(comb.split('_')[0])
        combination_envelopes.append(comb.split('|')[0])

    combination_envelopes = list(unique_everseen(combination_envelopes))
 
    combination_groups = []
    for envelope in combination_envelopes:
        group = []
        for comb in combinations:
            if envelope in comb:
                group.append(str(comb) + '(CB:max)')
                group.append(str(comb) + '(CB:min)')
        combination_groups.append(group)

    with xlsxwriter.Workbook(file_path_out + file_name_out + '_MLC2' + extension_out) as workbook:
        worksheet = workbook.add_worksheet()

        for row in range(len(combination_envelopes)):
            worksheet.write(row, 0, combination_envelopes[row])
            worksheet.write(row, 1, list_string_converter(combination_groups[row]))

