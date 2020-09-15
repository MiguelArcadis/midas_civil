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


def char_range(c1, c2):
    """Generates the characters from `c1` to `c2`, inclusive."""
    for c in range(ord(c1), ord(c2)+1):
        yield chr(c)

def comb_extractor(df, file_path_out, file_name_out, extension_out):
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

    # Create new excel file for midas_result_extraction
    with xlsxwriter.Workbook(file_path_out + file_name_out + extension_out) as workbook:
        # Create the relevant sheets on the excel file
        ws_readme = workbook.add_worksheet("ReadMe!")
        ws_extract = workbook.add_worksheet("Extract Results")
        ws_elements = workbook.add_worksheet("Elements")
        ws_combinations = workbook.add_worksheet("Combinations")

        ws_readme.hide_gridlines(2)
        ws_extract.hide_gridlines(2)
        ws_elements.hide_gridlines(2)
        ws_combinations.hide_gridlines(2)

        # Set formats for title, input cells and locked cells
        title_format = workbook.add_format({'bold': True,
                                            'font_size': 12,
                                            'bg_color': '#E57022',
                                            'font_color': '#FFFFFF',
                                            'border': 1
                                            })

        input_format = workbook.add_format({'bg_color': '#FFF7F1',
                                            'border': 1})

        locked_format = workbook.add_format({'bg_color': '#DFE0DF',
                                             'border': 1})

        ### READ ME ###
        merge_format_h1 = workbook.add_format({'bold': 1,
                                               'font_size': 16,
                                               'font_color': '#FFFFFF',
                                               'bg_color': '#E57022',
                                               'valign': 'vcenter',
                                               'align': 'center'})

        merge_format_h2 = workbook.add_format({'bold': 0,
                                               'underline': True,
                                               'font_size': 12,
                                               'font_color': '#000000',
                                               'bg_color': '#FCD5B4',
                                               'valign': 'vcenter',
                                               'align': 'center'})

        merge_format_t = workbook.add_format({'valign': 'vcenter',
                                              'align': 'left',
                                              'text_wrap': True})
        
        ws_readme.merge_range('A1:P1', 'Midas Results Extractor', merge_format_h1)
        ws_readme.merge_range('A2:P2', 'Introduction', merge_format_h2)
        ws_readme.merge_range('A3:P4','This spreadsheet is used to hold all the combinations and combination envelopes as well as a matrix of which element and node correspond to which part of the structure. With this information we can, any time that is needed, extract all the relevant load combinations for a given element, process them and give concurrent results to the user.', merge_format_t)
        ws_readme.merge_range('A5:P5', 'Spreadsheet Structure', merge_format_h2)

        ws_readme.write('A6', "Elements Tab")
        ws_readme.write('B7', "Element Name: Write relevant name to identify element (e.g. Main Girder Sagging Span 1)")
        ws_readme.write('B8', "Element Number: Midas number corresponding to the element defined")
        ws_readme.write('B9', "Node: Midas node for the element defined")
        ws_readme.write('A10', "Combinations Tab")
        ws_readme.write('B11', "Combination Envelopes: The combination envelopes generated from the spreadsheet")
        ws_readme.write('B12', "Combinations: The individual combinations that make up these envelopes")
        ws_readme.write('A13', "Extract Results Tab")
        ws_readme.write('B14', "Element Name: Select from drop-down menu (Fills in as defined in Elements Tab)")
        ws_readme.write('B15', "Element No.: Auto-filled")
        ws_readme.write('B16', "Node: Auto-filled")
        ws_readme.write('B17', "Fx, Fy, Fz, Mx, My, Mz: Select 'x' from drop-down if you want force extracted")
        ws_readme.write('B18', "Env 1 to Env 10: Select envelope from drop-down. The concurrent combinations that make up the envelope will be extracted, not the envelope itself.")

        ws_readme.merge_range('A19:P19', 'Cell Style Key', merge_format_h2)
        ws_readme.write('A20', "", title_format)
        ws_readme.write('B20', "Title cell. Do not edit.")
        ws_readme.write('A21', "", input_format)
        ws_readme.write('B21', "Input cell. Input as required.")
        ws_readme.write('A22', "", locked_format)
        ws_readme.write('B22', "Automatic Cell. Automatically filled or pre-filled. Do not change")

        ws_readme.merge_range('A23:P23', 'How to extract', merge_format_h2)
        ws_readme.write('B24', '1) Have relevant Midas file open (only 1 Midas file open)')
        ws_readme.write('B25', '2) Close all Midas tabs. Have only the main window in Midas opened')
        ws_readme.write('B26', '3) Make sure excel file is closed (it cannot be read by python if open)')
        ws_readme.write('B27', '4) Run python script. Make sure you select correct input file and desired output location')

        ### ELEMENTS SHEET ###
        ws_elements.write(0, 0, "Element name", title_format)
        ws_elements.write(0, 1, "Element number", title_format)
        ws_elements.write(0, 2, "Node", title_format)
        ws_elements.data_validation('C2:C200', {'validate': 'list', 'source': ['Part   i', 'Part 1/4', 'Part 2/4', 'Part 3/4', 'Part   j']})
       
        ws_elements.set_column('A2:A200', width=20, cell_format=input_format)
        ws_elements.set_column('B2:B200', width=20, cell_format=input_format)
        ws_elements.set_column('C2:C200', width=10, cell_format=input_format)

        ### COMBINATIONS SHEET ###
        for row in range(len(combination_envelopes)):
            ws_combinations.write(0, 0, "Combination Envelopes", title_format)
            ws_combinations.write(0, 1, "Combinations", title_format)
            ws_combinations.write(row + 1, 0, combination_envelopes[row], locked_format)
            ws_combinations.write(row + 1, 1, list_string_converter(combination_groups[row]), locked_format)
        ws_combinations.set_column('A:A', width=25, cell_format=locked_format)
        ws_combinations.set_column('B:B', width=255, cell_format=locked_format)

        ### EXTRACT RESULTS SHEET ###
        ws_extract.write(0, 0, "Element name", title_format)
        ws_extract.data_validation('A2:A200', {'validate': 'list', 'source': '=Elements!$A$2:$A$200'})
        ws_extract.write(0, 1, "Element no.", title_format)
        ws_extract.write(0, 2, "Node", title_format)
        for row in range(2, 200):
            ws_extract.write_formula(f'B{row}', f'=IFERROR(VLOOKUP(A{row}, Elements!$A$2:$C$200, 2, FALSE),"")')
            ws_extract.write_formula(f'C{row}', f'=IFERROR(VLOOKUP(A{row}, Elements!$A$2:$C$200, 3, FALSE),"")')

        ws_extract.write(0, 3, "Fx", title_format)
        ws_extract.write(0, 4, "Fy", title_format)
        ws_extract.write(0, 5, "Fz", title_format)
        ws_extract.write(0, 6, "Mx", title_format)
        ws_extract.write(0, 7, "My", title_format)
        ws_extract.write(0, 8, "Mz", title_format)
        ws_extract.data_validation('D2:I200', {'validate': 'list', 'source': ['x']})

        for column in range(9,19):
            ws_extract.write(0, column, f"Env {column - 8}", title_format)

        for letter in char_range('J', 'S'):
            ws_extract.data_validation(f'{letter}2:{letter}200', {'validate': 'list', 'source': '=Combinations!$A$2:$A$200'})

        ws_extract.set_column('A:A', width=15, cell_format=input_format)
        ws_extract.set_column('B:C', width=12, cell_format=locked_format)
        ws_extract.set_column('D:I', width=2.5, cell_format=input_format)
        ws_extract.set_column('J:S', width=10, cell_format=input_format)

        

        

        
        
        





