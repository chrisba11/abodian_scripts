import os
import re



def add_door_material():
    """

    """
    # This is to ask for the directory path at the command prompt
    parent_path = input('What is the path for the directory where the Materials.dat & New_Material.txt files are located? ')
    print('\n---\n')
    # generates path for location of Materials.dat file and opens the file
    mat1_file_path = parent_path + '\\Materials.dat'
    with open(mat1_file_path, "rt") as f:

        # grabs entire Materials.dat file contents
        curr_mat_list = f.readlines()

    # generates path for location of New_Material.txt file and opens the file
    mat2_file_path = parent_path + '\\New_Material.txt'
    with open(mat2_file_path, "rt") as f:

        # grabs entire New_Material.txt file contents
        full_new_mat_file = f.readlines()

    # list of indices that are okay having blank strings
    blank_okay_idx = [31, 46, 49, 52, 55, 61, 64, 67, 70]

    # pulls out only the relevant inputs for variables
    # removes newline char at end of each line
    new_mat = [None]
    for i in range(10, 72, 3):
        if full_new_mat_file[i][:-1] == '' and i not in blank_okay_idx:
            print('\nCHECK YOUR INPUT FOR BLANK LINES')
            print('\nOnly questions 8, 16, and 21 can be left blank.\n')
            return
        new_mat.append(full_new_mat_file[i][:-1])

    # removes any special characters from user input
    for i in range(len(new_mat)):
        if new_mat[i] is not None and new_mat[i] != '':
            new_mat[i] = re.sub("[@%&*'\"!?#~`<>\^\\\$\[\]\{\}\|\(\)]", '', new_mat[i])
            new_mat[i] = new_mat[i].strip()

    interior_materials = [
        ['Maple 34', 'CM Hardrock Maple 3/4 WF275 PRZ', '19.19986', 'Matching'],
        ['Maple 58', 'CM Hardrock Maple 5/8 WF275 PRZ', '16.10106', 'Matching'],
        ['PF Ply 34', 'CM Birch PF 3/4', '17.80032', 'Matching'],
        ['PF Ply 58', 'CM Birch PF 5/8', '15.8', 'Matching'],
        ['Storm 34', 'Storm Gray 3/4 S565 PAN', '19.2', 'Matching'],
        ['Storm 58', 'Storm Gray 5/8 S565 PAN', '16.2', 'Matching'],
        ['White 34', 'CM White 3/4 W100 PRZ', '19.19986', 'White STN101'],
        ['White 58', 'CM White 5/8 W100 PRZ', '16.10106', 'White STN101']
        ]

    # filling out variables with input read from New_Material.txt file
    sheet_name = new_mat[1]
    sheet_width = str(round(float(new_mat[2]) * 25.4, 4))
    sheet_length = str(round(float(new_mat[3]) * 25.4, 4))
    sheet_thick = str(round(float(new_mat[4]) * 25.4, 4))
    has_grain = new_mat[5]
    has_grain = "True" if has_grain[0] == 'y' or has_grain[0] == 'Y' else "False"
    two_sided = new_mat[6]
    two_sided = "True" if two_sided[0] == 'y' or two_sided[0] == 'Y' else "False"
    waste_percent = new_mat[7]
    optmize = "True"
    sheet_comment = new_mat[8]
    case_mat_band_name = f'{new_mat[9]} Banding'
    band_temp_name = new_mat[10]
    cab_temp_name = new_mat[11]

    case_band_exists = False
    door_band_exists = False
    band_flag = False
    case_band_name = new_mat[12]
    door_band_name = new_mat[17]
    
    birch_58_ADJ_exists = False
    birch_58_BACK_exists = False
    birch_58_GABLE_exists = False
    birch_58_PARTS_exists = False
    birch_58_PARTS_BSAW_exists = False
    birch_58_SHELF_exists = False
    
    birch_34_ADJ_exists = False
    birch_34_BACK_exists = False
    birch_34_GABLE_exists = False
    birch_34_PARTS_exists = False
    birch_34_PARTS_BSAW_exists = False
    birch_34_SHELF_exists = False

    hardrock_58_ADJ_exists = False
    hardrock_58_BACK_exists = False
    hardrock_58_GABLE_exists = False
    hardrock_58_PARTS_exists = False
    hardrock_58_PARTS_BSAW_exists = False
    hardrock_58_SHELF_exists = False

    hardrock_34_ADJ_exists = False
    hardrock_34_BACK_exists = False
    hardrock_34_GABLE_exists = False
    hardrock_34_PARTS_exists = False
    hardrock_34_PARTS_BSAW_exists = False
    hardrock_34_SHELF_exists = False
    
    storm_58_ADJ_exists = False
    storm_58_BACK_exists = False
    storm_58_GABLE_exists = False
    storm_58_PARTS_exists = False
    storm_58_PARTS_BSAW_exists = False
    storm_58_SHELF_exists = False
    
    storm_34_ADJ_exists = False
    storm_34_BACK_exists = False
    storm_34_GABLE_exists = False
    storm_34_PARTS_exists = False
    storm_34_PARTS_BSAW_exists = False
    storm_34_SHELF_exists = False
    
    white_58_ADJ_exists = False
    white_58_BACK_exists = False
    white_58_GABLE_exists = False
    white_58_PARTS_exists = False
    white_58_PARTS_BSAW_exists = False
    white_58_SHELF_exists = False
    
    white_34_ADJ_exists = False
    white_34_BACK_exists = False
    white_34_GABLE_exists = False
    white_34_PARTS_exists = False
    white_34_PARTS_BSAW_exists = False
    white_34_SHELF_exists = False
    
    panel_flag = False
    hardrock_58_name = interior_materials[1][1]
    hardrock_34_name = interior_materials[0][1]
    birch_58_name = interior_materials[3][1]
    birch_34_name = interior_materials[2][1]
    storm_58_name = interior_materials[5][1]
    storm_34_name = interior_materials[4][1]
    white_58_name = interior_materials[7][1]
    white_34_name = interior_materials[6][1]
    
    for line in curr_mat_list:
        if line == '  <PanelStockMaterials>\n':
            panel_flag = True
        if panel_flag == True:
            start_idx = line.find('Name=') + 6
            end_idx = line.find('" Quan=')
            line_name = line[start_idx:end_idx]
            
            # check if material matches any of the materials in int_mat_part_types
            if line_name == f'{birch_58_name} ADJ [{case_mat_band_name}]':
                birch_58_ADJ_exists = True
                print(f'Birch PF Ply 5/8 ADJ case material exists - Using Existing Material')
                birch_58_ADJ = line
            if line_name == f'{birch_34_name} ADJ [{case_mat_band_name}]':
                birch_34_ADJ_exists = True
                print(f'Birch PF Ply 3/4 ADJ case material exists - Using Existing Material')
                birch_34_ADJ = line
            if line_name == f'{hardrock_58_name} ADJ [{case_mat_band_name}]':
                hardrock_58_ADJ_exists = True
                print(f'Hardrock 5/8 ADJ case material exists - Using Existing Material')
                hardrock_58_ADJ = line
            if line_name == f'{hardrock_34_name} ADJ [{case_mat_band_name}]':
                hardrock_34_ADJ_exists = True
                print(f'Hardrock 3/4 ADJ case material exists - Using Existing Material')
                hardrock_34_ADJ = line
            if line_name == f'{storm_58_name} ADJ [{case_mat_band_name}]':
                storm_58_ADJ_exists = True
                print(f'Storm 5/8 ADJ case material exists - Using Existing Material')
                storm_58_ADJ = line
            if line_name == f'{storm_34_name} ADJ [{case_mat_band_name}]':
                storm_34_ADJ_exists = True
                print(f'Storm 3/4 ADJ case material exists - Using Existing Material')
                storm_34_ADJ = line
            if line_name == f'{white_58_name} ADJ [{case_mat_band_name}]':
                white_58_ADJ_exists = True
                print(f'White 5/8 ADJ case material exists - Using Existing Material')
                white_58_ADJ = line
            if line_name == f'{white_34_name} ADJ [{case_mat_band_name}]':
                white_34_ADJ_exists = True
                print(f'White 3/4 ADJ case material exists - Using Existing Material')
                white_34_ADJ = line

            if line_name == f'{birch_58_name} BACK [{case_mat_band_name}]':
                birch_58_BACK_exists = True
                print(f'Birch PF Ply 5/8 BACK case material exists - Using Existing Material')
                birch_58_BACK = line
            if line_name == f'{birch_34_name} BACK [{case_mat_band_name}]':
                birch_34_BACK_exists = True
                print(f'Birch PF Ply 3/4 BACK case material exists - Using Existing Material')
                birch_34_BACK = line
            if line_name == f'{hardrock_58_name} BACK [{case_mat_band_name}]':
                hardrock_58_BACK_exists = True
                print(f'Hardrock 5/8 BACK case material exists - Using Existing Material')
                hardrock_58_BACK = line
            if line_name == f'{hardrock_34_name} BACK [{case_mat_band_name}]':
                hardrock_34_BACK_exists = True
                print(f'Hardrock 3/4 BACK case material exists - Using Existing Material')
                hardrock_34_BACK = line
            if line_name == f'{storm_58_name} BACK [{case_mat_band_name}]':
                storm_58_BACK_exists = True
                print(f'Storm 5/8 BACK case material exists - Using Existing Material')
                storm_58_BACK = line
            if line_name == f'{storm_34_name} BACK [{case_mat_band_name}]':
                storm_34_BACK_exists = True
                print(f'Storm 3/4 BACK case material exists - Using Existing Material')
                storm_34_BACK = line
            if line_name == f'{white_58_name} BACK [{case_mat_band_name}]':
                white_58_BACK_exists = True
                print(f'White 5/8 BACK case material exists - Using Existing Material')
                white_58_BACK = line
            if line_name == f'{white_34_name} BACK [{case_mat_band_name}]':
                white_34_BACK_exists = True
                print(f'White 3/4 BACK case material exists - Using Existing Material')
                white_34_BACK = line

            if line_name == f'{birch_58_name} GABLE [{case_mat_band_name}]':
                birch_58_GABLE_exists = True
                print(f'Birch PF Ply 5/8 GABLE case material exists - Using Existing Material')
                birch_58_GABLE = line
            if line_name == f'{birch_34_name} GABLE [{case_mat_band_name}]':
                birch_34_GABLE_exists = True
                print(f'Birch PF Ply 3/4 GABLE case material exists - Using Existing Material')
                birch_34_GABLE = line
            if line_name == f'{hardrock_58_name} GABLE [{case_mat_band_name}]':
                hardrock_58_GABLE_exists = True
                print(f'Hardrock 5/8 GABLE case material exists - Using Existing Material')
                hardrock_58_GABLE = line
            if line_name == f'{hardrock_34_name} GABLE [{case_mat_band_name}]':
                hardrock_34_GABLE_exists = True
                print(f'Hardrock 3/4 GABLE case material exists - Using Existing Material')
                hardrock_34_GABLE = line
            if line_name == f'{storm_58_name} GABLE [{case_mat_band_name}]':
                storm_58_GABLE_exists = True
                print(f'Storm 5/8 GABLE case material exists - Using Existing Material')
                storm_58_GABLE = line
            if line_name == f'{storm_34_name} GABLE [{case_mat_band_name}]':
                storm_34_GABLE_exists = True
                print(f'Storm 3/4 GABLE case material exists - Using Existing Material')
                storm_34_GABLE = line
            if line_name == f'{white_58_name} GABLE [{case_mat_band_name}]':
                white_58_GABLE_exists = True
                print(f'White 5/8 GABLE case material exists - Using Existing Material')
                white_58_GABLE = line
            if line_name == f'{white_34_name} GABLE [{case_mat_band_name}]':
                white_34_GABLE_exists = True
                print(f'White 3/4 GABLE case material exists - Using Existing Material')
                white_34_GABLE = line

            if line_name == f'{birch_58_name} PARTS [{case_mat_band_name}]':
                birch_58_PARTS_exists = True
                print(f'Birch PF Ply 5/8 PARTS case material exists - Using Existing Material')
                birch_58_PARTS = line
            if line_name == f'{birch_34_name} PARTS [{case_mat_band_name}]':
                birch_34_PARTS_exists = True
                print(f'Birch PF Ply 3/4 PARTS case material exists - Using Existing Material')
                birch_34_PARTS = line
            if line_name == f'{hardrock_58_name} PARTS [{case_mat_band_name}]':
                hardrock_58_PARTS_exists = True
                print(f'Hardrock 5/8 PARTS case material exists - Using Existing Material')
                hardrock_58_PARTS = line
            if line_name == f'{hardrock_34_name} PARTS [{case_mat_band_name}]':
                hardrock_34_PARTS_exists = True
                print(f'Hardrock 3/4 PARTS case material exists - Using Existing Material')
                hardrock_34_PARTS = line
            if line_name == f'{storm_58_name} PARTS [{case_mat_band_name}]':
                storm_58_PARTS_exists = True
                print(f'Storm 5/8 PARTS case material exists - Using Existing Material')
                storm_58_PARTS = line
            if line_name == f'{storm_34_name} PARTS [{case_mat_band_name}]':
                storm_34_PARTS_exists = True
                print(f'Storm 3/4 PARTS case material exists - Using Existing Material')
                storm_34_PARTS = line
            if line_name == f'{white_58_name} PARTS [{case_mat_band_name}]':
                white_58_PARTS_exists = True
                print(f'White 5/8 PARTS case material exists - Using Existing Material')
                white_58_PARTS = line
            if line_name == f'{white_34_name} PARTS [{case_mat_band_name}]':
                white_34_PARTS_exists = True
                print(f'White 3/4 PARTS case material exists - Using Existing Material')
                white_34_PARTS = line

            if line_name == f'{birch_58_name} PARTS [{case_mat_band_name}] BSAW':
                birch_58_PARTS_BSAW_exists = True
                print(f'Birch PF Ply 5/8 PARTS_BSAW case material exists - Using Existing Material')
                birch_58_PARTS_BSAW = line
            if line_name == f'{birch_34_name} PARTS [{case_mat_band_name}] BSAW':
                birch_34_PARTS_BSAW_exists = True
                print(f'Birch PF Ply 3/4 PARTS_BSAW case material exists - Using Existing Material')
                birch_34_PARTS_BSAW = line
            if line_name == f'{hardrock_58_name} PARTS [{case_mat_band_name}] BSAW':
                hardrock_58_PARTS_BSAW_exists = True
                print(f'Hardrock 5/8 PARTS_BSAW case material exists - Using Existing Material')
                hardrock_58_PARTS_BSAW = line
            if line_name == f'{hardrock_34_name} PARTS [{case_mat_band_name}] BSAW':
                hardrock_34_PARTS_BSAW_exists = True
                print(f'Hardrock 3/4 PARTS_BSAW case material exists - Using Existing Material')
                hardrock_34_PARTS_BSAW = line
            if line_name == f'{storm_58_name} PARTS [{case_mat_band_name}] BSAW':
                storm_58_PARTS_BSAW_exists = True
                print(f'Storm 5/8 PARTS_BSAW case material exists - Using Existing Material')
                storm_58_PARTS_BSAW = line
            if line_name == f'{storm_34_name} PARTS [{case_mat_band_name}] BSAW':
                storm_34_PARTS_BSAW_exists = True
                print(f'Storm 3/4 PARTS_BSAW case material exists - Using Existing Material')
                storm_34_PARTS_BSAW = line
            if line_name == f'{white_58_name} PARTS [{case_mat_band_name}] BSAW':
                white_58_PARTS_BSAW_exists = True
                print(f'White 5/8 PARTS_BSAW case material exists - Using Existing Material')
                white_58_PARTS_BSAW = line
            if line_name == f'{white_34_name} PARTS [{case_mat_band_name}] BSAW':
                white_34_PARTS_BSAW_exists = True
                print(f'White 3/4 PARTS_BSAW case material exists - Using Existing Material')
                white_34_PARTS_BSAW = line

            if line_name == f'{birch_58_name} SHELF [{case_mat_band_name}]':
                birch_58_SHELF_exists = True
                print(f'Birch PF Ply 5/8 SHELF case material exists - Using Existing Material')
                birch_58_SHELF = line
            if line_name == f'{birch_34_name} SHELF [{case_mat_band_name}]':
                birch_34_SHELF_exists = True
                print(f'Birch PF Ply 3/4 SHELF case material exists - Using Existing Material')
                birch_34_SHELF = line
            if line_name == f'{hardrock_58_name} SHELF [{case_mat_band_name}]':
                hardrock_58_SHELF_exists = True
                print(f'Hardrock 5/8 SHELF case material exists - Using Existing Material')
                hardrock_58_SHELF = line
            if line_name == f'{hardrock_34_name} SHELF [{case_mat_band_name}]':
                hardrock_34_SHELF_exists = True
                print(f'Hardrock 3/4 SHELF case material exists - Using Existing Material')
                hardrock_34_SHELF = line
            if line_name == f'{storm_58_name} SHELF [{case_mat_band_name}]':
                storm_58_SHELF_exists = True
                print(f'Storm 5/8 SHELF case material exists - Using Existing Material')
                storm_58_SHELF = line
            if line_name == f'{storm_34_name} SHELF [{case_mat_band_name}]':
                storm_34_SHELF_exists = True
                print(f'Storm 3/4 SHELF case material exists - Using Existing Material')
                storm_34_SHELF = line
            if line_name == f'{white_58_name} SHELF [{case_mat_band_name}]':
                white_58_SHELF_exists = True
                print(f'White 5/8 SHELF case material exists - Using Existing Material')
                white_58_SHELF = line
            if line_name == f'{white_34_name} SHELF [{case_mat_band_name}]':
                white_34_SHELF_exists = True
                print(f'White 3/4 SHELF case material exists - Using Existing Material')
                white_34_SHELF = line

        if line == '  </PanelStockMaterials>\n':
            panel_flag = False
        if line == '  <EdgeBandingMaterials>\n':
            band_flag = True
        if band_flag == True:
            start_idx = line.find('Name=') + 6
            end_idx = line.find('" Quan=')
            line_name = line[start_idx:end_idx]
            if line_name == case_band_name:
                case_band_exists = True
                print('Case banding exists - Using Existing Material')
                case_band = line
            if line_name == door_band_name:
                door_band_exists = True
                print('Door banding exists - Using Existing Material')
                door_band = line
        if case_band_exists == True and door_band_exists == True:
            break
        if line == '  </EdgeBandingMaterials>\n':
            break
    
    if case_band_exists == False:
        if new_mat[13] == '' or new_mat[14] == '' or new_mat[15] == '':
            print('\nCHECK YOUR INPUT FOR BLANK LINES')
            print('\nNew case banding materials must have width, length, and type.\n')
            return
        case_band_width = str(round(float(new_mat[13]) * 25.4, 4))
        case_band_length = str(round(float(new_mat[14]) * 12 * 25.4, 4))
        case_band_type = new_mat[15]
        case_band_comment = new_mat[16]

    if door_band_exists == False:
        if new_mat[18] == '' or new_mat[19] == '' or new_mat[20] == '':
            print('\nCHECK YOUR INPUT FOR BLANK LINES')
            print('\nNew door banding materials must have width, length, and type.\n')
            return
        door_band_width = str(round(float(new_mat[18]) * 25.4, 4))
        door_band_length = str(round(float(new_mat[19]) * 12 * 25.4, 4))
        door_band_type = new_mat[20]
        door_band_comment = new_mat[21]

    # creates folder where new material templates will be placed
    folder_name = sheet_name.replace('/', '-')
    folder_path = os.path.join(parent_path, folder_name)
    os.mkdir(folder_path)

    # creates lines for new door materials, one CNC (standard) and one BSAW
    door_sheet_CNC = f'    <Material Name="{sheet_name} [Matching Banding]" '
    door_sheet_CNC_doors = f'    <Material Name="{sheet_name} DOORS [Matching Banding]" '
    door_sheet_BSAW_doors = f'    <Material Name="{sheet_name} DOORS [Matching Banding] BSAW" '
    door_sheet_CNC_parts = f'    <Material Name="{sheet_name} PARTS [Matching Banding]" '
    door_sheet_BSAW_parts = f'    <Material Name="{sheet_name} PARTS [Matching Banding] BSAW" '

    door_sheet_mid = \
        f'Quan="1" ' \
        f'W="{sheet_width}" ' \
        f'L="{sheet_length}" ' \
        f'Thick="{sheet_thick}" '
    
    door_sheet_CNC_trim = f'WTrim="6.35" LTrim="6.35" '
    
    door_sheet_BSAW_trim = f'WTrim="12" LTrim="12" '

    door_sheet_end = \
        f'HasGrain="{has_grain}" ' \
        f'TwoSided="{two_sided}" ' \
        f'CostEach="0" ' \
        f'MarkupPercentage="0" ' \
        f'AddOnCost="0" ' \
        f'Speed="100" ' \
        f'ImageFile="" ' \
        f'Comment="{sheet_comment}" ' \
        f'WastePercentage="{waste_percent}" ' \
        f'Optimize="{optmize}" ' \
        f'BandType="0" />\n'

    door_sheet_CNC = door_sheet_CNC + door_sheet_mid + door_sheet_CNC_trim + door_sheet_end
    door_sheet_CNC_doors = door_sheet_CNC_doors + door_sheet_mid + door_sheet_CNC_trim + door_sheet_end
    door_sheet_BSAW_doors = door_sheet_BSAW_doors + door_sheet_mid + door_sheet_BSAW_trim + door_sheet_end
    door_sheet_CNC_parts = door_sheet_CNC_parts + door_sheet_mid + door_sheet_CNC_trim + door_sheet_end
    door_sheet_BSAW_parts = door_sheet_BSAW_parts + door_sheet_mid + door_sheet_BSAW_trim + door_sheet_end
    
    # creates lines for each of the case materials with new material banding
    if birch_58_ADJ_exists == False:
        birch_58_ADJ = \
            f'    <Material Name="{birch_58_name} ADJ [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="15.8" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_58_BACK_exists == False:
        birch_58_BACK = \
            f'    <Material Name="{birch_58_name} BACK [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="15.8" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_58_GABLE_exists == False:
        birch_58_GABLE = \
            f'    <Material Name="{birch_58_name} GABLE [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="15.8" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_58_PARTS_exists == False:
        birch_58_PARTS = \
            f'    <Material Name="{birch_58_name} PARTS [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="15.8" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_58_PARTS_BSAW_exists == False:
        birch_58_PARTS_BSAW = \
            f'    <Material Name="{birch_58_name} PARTS [{case_mat_band_name}] BSAW" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="15.8" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_58_SHELF_exists == False:
        birch_58_SHELF = \
            f'    <Material Name="{birch_58_name} SHELF [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="15.8" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_34_ADJ_exists == False:
        birch_34_ADJ = \
            f'    <Material Name="{birch_34_name} ADJ [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="17.80032" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_34_BACK_exists == False:
        birch_34_BACK = \
            f'    <Material Name="{birch_34_name} BACK [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="17.80032" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_34_GABLE_exists == False:
        birch_34_GABLE = \
            f'    <Material Name="{birch_34_name} GABLE [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="17.80032" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_34_PARTS_exists == False:
        birch_34_PARTS = \
            f'    <Material Name="{birch_34_name} PARTS [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="17.80032" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_34_PARTS_BSAW_exists == False:
        birch_34_PARTS_BSAW = \
            f'    <Material Name="{birch_34_name} PARTS [{case_mat_band_name}] BSAW" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="17.80032" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if birch_34_SHELF_exists == False:
        birch_34_SHELF = \
            f'    <Material Name="{birch_34_name} SHELF [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1231.9" ' \
            f'L="2451.1" ' \
            f'Thick="17.80032" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_58_ADJ_exists == False:
        hardrock_58_ADJ = \
            f'    <Material Name="{hardrock_58_name} ADJ [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_58_BACK_exists == False:
        hardrock_58_BACK = \
            f'    <Material Name="{hardrock_58_name} BACK [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_58_GABLE_exists == False:
        hardrock_58_GABLE = \
            f'    <Material Name="{hardrock_58_name} GABLE [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_58_PARTS_exists == False:
        hardrock_58_PARTS = \
            f'    <Material Name="{hardrock_58_name} PARTS [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_58_PARTS_BSAW_exists == False:
        hardrock_58_PARTS_BSAW = \
            f'    <Material Name="{hardrock_58_name} PARTS [{case_mat_band_name}] BSAW" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_58_SHELF_exists == False:
        hardrock_58_SHELF = \
            f'    <Material Name="{hardrock_58_name} SHELF [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_34_ADJ_exists == False:
        hardrock_34_ADJ = \
            f'    <Material Name="{hardrock_34_name} ADJ [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_34_BACK_exists == False:
        hardrock_34_BACK = \
            f'    <Material Name="{hardrock_34_name} BACK [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_34_GABLE_exists == False:
        hardrock_34_GABLE = \
            f'    <Material Name="{hardrock_34_name} GABLE [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_34_PARTS_exists == False:
        hardrock_34_PARTS = \
            f'    <Material Name="{hardrock_34_name} PARTS [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_34_PARTS_BSAW_exists == False:
        hardrock_34_PARTS_BSAW = \
            f'    <Material Name="{hardrock_34_name} PARTS [{case_mat_band_name}] BSAW" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if hardrock_34_SHELF_exists == False:
        hardrock_34_SHELF = \
            f'    <Material Name="{hardrock_34_name} SHELF [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2460.625" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="True" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_58_ADJ_exists == False:
        storm_58_ADJ = \
            f'    <Material Name="{storm_58_name} ADJ [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="16.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_58_BACK_exists == False:
        storm_58_BACK = \
            f'    <Material Name="{storm_58_name} BACK [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="16.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_58_GABLE_exists == False:
        storm_58_GABLE = \
            f'    <Material Name="{storm_58_name} GABLE [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="16.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_58_PARTS_exists == False:
        storm_58_PARTS = \
            f'    <Material Name="{storm_58_name} PARTS [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="16.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_58_PARTS_BSAW_exists == False:
        storm_58_PARTS_BSAW = \
            f'    <Material Name="{storm_58_name} PARTS [{case_mat_band_name}] BSAW" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="16.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_58_SHELF_exists == False:
        storm_58_SHELF = \
            f'    <Material Name="{storm_58_name} SHELF [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="16.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_34_ADJ_exists == False:
        storm_34_ADJ = \
            f'    <Material Name="{storm_34_name} ADJ [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="19.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_34_BACK_exists == False:
        storm_34_BACK = \
            f'    <Material Name="{storm_34_name} BACK [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="19.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_34_GABLE_exists == False:
        storm_34_GABLE = \
            f'    <Material Name="{storm_34_name} GABLE [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="19.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_34_PARTS_exists == False:
        storm_34_PARTS = \
            f'    <Material Name="{storm_34_name} PARTS [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="19.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_34_PARTS_BSAW_exists == False:
        storm_34_PARTS_BSAW = \
            f'    <Material Name="{storm_34_name} PARTS [{case_mat_band_name}] BSAW" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="19.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if storm_34_SHELF_exists == False:
        storm_34_SHELF = \
            f'    <Material Name="{storm_34_name} SHELF [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1244.6" ' \
            f'L="2463.8" ' \
            f'Thick="19.2" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_58_ADJ_exists == False:
        white_58_ADJ = \
            f'    <Material Name="{white_58_name} ADJ [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_58_BACK_exists == False:
        white_58_BACK = \
            f'    <Material Name="{white_58_name} BACK [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_58_GABLE_exists == False:
        white_58_GABLE = \
            f'    <Material Name="{white_58_name} GABLE [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_58_PARTS_exists == False:
        white_58_PARTS = \
            f'    <Material Name="{white_58_name} PARTS [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_58_PARTS_BSAW_exists == False:
        white_58_PARTS_BSAW = \
            f'    <Material Name="{white_58_name} PARTS [{case_mat_band_name}] BSAW" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_58_SHELF_exists == False:
        white_58_SHELF = \
            f'    <Material Name="{white_58_name} SHELF [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="16.10106" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_34_ADJ_exists == False:
        white_34_ADJ = \
            f'    <Material Name="{white_34_name} ADJ [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_34_BACK_exists == False:
        white_34_BACK = \
            f'    <Material Name="{white_34_name} BACK [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_34_GABLE_exists == False:
        white_34_GABLE = \
            f'    <Material Name="{white_34_name} GABLE [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_34_PARTS_exists == False:
        white_34_PARTS = \
            f'    <Material Name="{white_34_name} PARTS [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_34_PARTS_BSAW_exists == False:
        white_34_PARTS_BSAW = \
            f'    <Material Name="{white_34_name} PARTS [{case_mat_band_name}] BSAW" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'

    if white_34_SHELF_exists == False:
        white_34_SHELF = \
            f'    <Material Name="{white_34_name} SHELF [{case_mat_band_name}]" ' \
            f'Quan="1" ' \
            f'W="1243.012" ' \
            f'L="2462.212" ' \
            f'Thick="19.19986" ' \
            f'WTrim="6.35" ' \
            f'LTrim="6.35" ' \
            f'HasGrain="False" ' \
            f'TwoSided="True" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="100" ' \
            f'ImageFile="" ' \
            f'Comment="" ' \
            f'WastePercentage="20" ' \
            f'Optimize="True" ' \
            f'BandType="0" />\n'


    # creates banding materials if they don't already exist
    if case_band_exists == False:
        case_band = \
            f'    <Material Name="{case_band_name}" ' \
            f'Quan="1" ' \
            f'W="{case_band_width}" ' \
            f'L="{case_band_length}" ' \
            f'Thick="0" ' \
            f'WTrim="0" ' \
            f'LTrim="0" ' \
            f'HasGrain="False" ' \
            f'TwoSided="False" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="0" ' \
            f'ImageFile="" ' \
            f'Comment="{case_band_comment}" ' \
            f'WastePercentage="0" ' \
            f'Optimize="False" ' \
            f'BandType="{case_band_type}" />\n'

    if door_band_exists == False:    
        door_band = \
            f'    <Material Name="{door_band_name}" ' \
            f'Quan="1" ' \
            f'W="{door_band_width}" ' \
            f'L="{door_band_length}" ' \
            f'Thick="0" ' \
            f'WTrim="0" ' \
            f'LTrim="0" ' \
            f'HasGrain="False" ' \
            f'TwoSided="False" ' \
            f'CostEach="0" ' \
            f'MarkupPercentage="0" ' \
            f'AddOnCost="0" ' \
            f'Speed="0" ' \
            f'ImageFile="" ' \
            f'Comment="{door_band_comment}" ' \
            f'WastePercentage="0" ' \
            f'Optimize="False" ' \
            f'BandType="{door_band_type}" />\n'
    
            
    # get the end of the sheet materials list
    curr_sheet_mat_end_idx = curr_mat_list.index('  </PanelStockMaterials>\n')

    # get the end of the banding material list
    curr_band_list_end_idx = curr_mat_list.index('  </EdgeBandingMaterials>\n')

    updated_material_list = curr_mat_list[:curr_sheet_mat_end_idx]

    updated_material_list.append(door_sheet_CNC)
    updated_material_list.append(door_sheet_CNC_doors)
    updated_material_list.append(door_sheet_BSAW_doors)
    updated_material_list.append(door_sheet_CNC_parts)
    updated_material_list.append(door_sheet_BSAW_parts)

    if birch_58_ADJ_exists == False:
        updated_material_list.append(birch_58_ADJ)
    if birch_58_BACK_exists == False:
        updated_material_list.append(birch_58_BACK)
    if birch_58_GABLE_exists == False:
        updated_material_list.append(birch_58_GABLE)
    if birch_58_PARTS_exists == False:
        updated_material_list.append(birch_58_PARTS)
    if birch_58_PARTS_BSAW_exists == False:
        updated_material_list.append(birch_58_PARTS_BSAW)
    if birch_58_SHELF_exists == False:
        updated_material_list.append(birch_58_SHELF)

    if birch_34_ADJ_exists == False:
        updated_material_list.append(birch_34_ADJ)
    if birch_34_BACK_exists == False:
        updated_material_list.append(birch_34_BACK)
    if birch_34_GABLE_exists == False:
        updated_material_list.append(birch_34_GABLE)
    if birch_34_PARTS_exists == False:
        updated_material_list.append(birch_34_PARTS)
    if birch_34_PARTS_BSAW_exists == False:
        updated_material_list.append(birch_34_PARTS_BSAW)
    if birch_34_SHELF_exists == False:
        updated_material_list.append(birch_34_SHELF)

    if hardrock_58_ADJ_exists == False:
        updated_material_list.append(hardrock_58_ADJ)
    if hardrock_58_BACK_exists == False:
        updated_material_list.append(hardrock_58_BACK)
    if hardrock_58_GABLE_exists == False:
        updated_material_list.append(hardrock_58_GABLE)
    if hardrock_58_PARTS_exists == False:
        updated_material_list.append(hardrock_58_PARTS)
    if hardrock_58_PARTS_BSAW_exists == False:
        updated_material_list.append(hardrock_58_PARTS_BSAW)
    if hardrock_58_SHELF_exists == False:
        updated_material_list.append(hardrock_58_SHELF)

    if hardrock_34_ADJ_exists == False:
        updated_material_list.append(hardrock_34_ADJ)
    if hardrock_34_BACK_exists == False:
        updated_material_list.append(hardrock_34_BACK)
    if hardrock_34_GABLE_exists == False:
        updated_material_list.append(hardrock_34_GABLE)
    if hardrock_34_PARTS_exists == False:
        updated_material_list.append(hardrock_34_PARTS)
    if hardrock_34_PARTS_BSAW_exists == False:
        updated_material_list.append(hardrock_34_PARTS_BSAW)
    if hardrock_34_SHELF_exists == False:
        updated_material_list.append(hardrock_34_SHELF)

    if white_58_ADJ_exists == False:
        updated_material_list.append(white_58_ADJ)
    if white_58_BACK_exists == False:
        updated_material_list.append(white_58_BACK)
    if white_58_GABLE_exists == False:
        updated_material_list.append(white_58_GABLE)
    if white_58_PARTS_exists == False:
        updated_material_list.append(white_58_PARTS)
    if white_58_PARTS_BSAW_exists == False:
        updated_material_list.append(white_58_PARTS_BSAW)
    if white_58_SHELF_exists == False:
        updated_material_list.append(white_58_SHELF)

    if white_34_ADJ_exists == False:
        updated_material_list.append(white_34_ADJ)
    if white_34_BACK_exists == False:
        updated_material_list.append(white_34_BACK)
    if white_34_GABLE_exists == False:
        updated_material_list.append(white_34_GABLE)
    if white_34_PARTS_exists == False:
        updated_material_list.append(white_34_PARTS)
    if white_34_PARTS_BSAW_exists == False:
        updated_material_list.append(white_34_PARTS_BSAW)
    if white_34_SHELF_exists == False:
        updated_material_list.append(white_34_SHELF)

    if storm_58_ADJ_exists == False:
        updated_material_list.append(storm_58_ADJ)
    if storm_58_BACK_exists == False:
        updated_material_list.append(storm_58_BACK)
    if storm_58_GABLE_exists == False:
        updated_material_list.append(storm_58_GABLE)
    if storm_58_PARTS_exists == False:
        updated_material_list.append(storm_58_PARTS)
    if storm_58_PARTS_BSAW_exists == False:
        updated_material_list.append(storm_58_PARTS_BSAW)
    if storm_58_SHELF_exists == False:
        updated_material_list.append(storm_58_SHELF)

    if storm_34_ADJ_exists == False:
        updated_material_list.append(storm_34_ADJ)
    if storm_34_BACK_exists == False:
        updated_material_list.append(storm_34_BACK)
    if storm_34_GABLE_exists == False:
        updated_material_list.append(storm_34_GABLE)
    if storm_34_PARTS_exists == False:
        updated_material_list.append(storm_34_PARTS)
    if storm_34_PARTS_BSAW_exists == False:
        updated_material_list.append(storm_34_PARTS_BSAW)
    if storm_34_SHELF_exists == False:
        updated_material_list.append(storm_34_SHELF)

    updated_material_list += curr_mat_list[curr_sheet_mat_end_idx:curr_band_list_end_idx]
    
    # adds new banding materials to banding list if they don't already exist
    if case_band_exists == False:
        updated_material_list.append(case_band)

    if door_band_exists == False:  
        updated_material_list.append(door_band)
    
    updated_material_list += curr_mat_list[curr_band_list_end_idx:]

    new_mat_path = folder_path + '\\Materials.dat'
    with open(new_mat_path, "wt") as f:
        f.writelines(updated_material_list)


    # creates all of the banding templates for each interior material
    for interior_band in [
        ['MC', '0.5mm Apple Spice 4274 DOL'],
        ['PC', '0.5mm Birch PVC'],
        ['SC', '0.5mm Storm Gray S565 PAN'],
        ['WC', '0.5mm White STN101 TEK']
        ]:

        band_temp = \
            f'2\n' \
            f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
            f'<MaterialTemplate Name="02 - {band_temp_name} - {interior_band[0]}" Type="4" SymbolForLabels="{sheet_name[0:2].upper()}">\n' \
            f'  <MaterialReference PartType="EdgeBand" Mat="{case_band_name}" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="EdgeBand2" Mat="{door_band_name}" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="EdgeBand3" Mat="{interior_band[1]}" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="EdgeBand4" Mat="None EB4" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'</MaterialTemplate>'

        band_temp_path = folder_path + f'\\02 - {band_temp_name} - {interior_band[0]}.BandTmp'

        with open(band_temp_path, "wt") as f:
            f.writelines(band_temp)



    # creates the finished interior cabinet templates
    fin_int_case_temp_02 = \
        f'2\n' \
        f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
        f'<MaterialTemplate Name="02 - {band_temp_name} .75" Type="3" SymbolForLabels="">\n' \
        f'  <MaterialReference PartType="Top" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Bottom" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FEnd" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UEnd" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="AppliedUE" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Shelf" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="AdjustableShelf" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="RolloutShelf" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FixedShelf" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Partition" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Divider" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Frame" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FramelessRail" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Nailer" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Stretcher" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Cleat" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Toe" Mat="Raw Toe Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinExterior" Mat="{sheet_name} PARTS [Matching Banding] BSAW" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinInteriorBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinExterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="ToeSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="EdgeBand" Mat="Banding" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Sleeper" Mat="Raw Toe Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Glass" Mat="1/8 Glass" MatThick="3.175" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Plastic" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Metal" Mat="Mirror 1/4" MatThick="6.35" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Special" Mat="{sheet_name} PARTS [Matching Banding] BSAW" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="ClosetRod" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'</MaterialTemplate>'

    fin_int_case_temp_12 = \
        f'2\n' \
        f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
        f'<MaterialTemplate Name="12 - {band_temp_name} .75" Type="3" SymbolForLabels="">\n' \
        f'  <MaterialReference PartType="Top" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Bottom" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FEnd" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UEnd" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="AppliedUE" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Shelf" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="AdjustableShelf" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="RolloutShelf" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FixedShelf" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Partition" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Divider" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Frame" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FramelessRail" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Nailer" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Stretcher" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Cleat" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Toe" Mat="Raw Toe Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinExterior" Mat="{sheet_name} PARTS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinInteriorBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinExterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="ToeSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="EdgeBand" Mat="Banding" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Sleeper" Mat="Raw Toe Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Glass" Mat="1/8 Glass" MatThick="3.175" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Plastic" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Metal" Mat="Mirror 1/4" MatThick="6.35" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Special" Mat="{sheet_name} PARTS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="ClosetRod" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'</MaterialTemplate>'

    fin_int_case_temp_path_02 = folder_path + f'\\02 - {band_temp_name} .75.CabTmp'
    fin_int_case_temp_path_12 = folder_path + f'\\12 - {band_temp_name} .75.CabTmp'

    with open(fin_int_case_temp_path_02, "wt") as f:
        f.writelines(fin_int_case_temp_02)

    with open(fin_int_case_temp_path_12, "wt") as f:
        f.writelines(fin_int_case_temp_12)

    # creates the standard cabinet templates
    for interior_mat in interior_materials:

        cab_temp_05 = \
            f'2\n' \
            f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
            f'<MaterialTemplate Name="05 - {interior_mat[0]} [{cab_temp_name}]" Type="3" SymbolForLabels="">\n' \
            f'  <MaterialReference PartType="Top" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Bottom" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FEnd" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UEnd" Mat="{interior_mat[1]} GABLE [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="AppliedUE" Mat="{interior_mat[1]} GABLE [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Shelf" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="AdjustableShelf" Mat="{interior_mat[1]} ADJ [{interior_mat[3]} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="RolloutShelf" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FixedShelf" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Partition" Mat="{interior_mat[1]} GABLE [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Divider" Mat="{interior_mat[1]} GABLE [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UBack" Mat="{interior_mat[1]} BACK [{interior_mat[3]} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FBack" Mat="{interior_mat[1]} BACK [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Frame" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FramelessRail" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Nailer" Mat="{interior_mat[1]} BACK [{interior_mat[3]} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Stretcher" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Cleat" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Toe" Mat="Raw Toe Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinExterior" Mat="{sheet_name} PARTS [Matching Banding] BSAW" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinInteriorBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinExterior" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinInterior" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="ToeSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="EdgeBand" Mat="Banding" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Sleeper" Mat="Raw Toe Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Glass" Mat="1/8 Glass" MatThick="3.175" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Plastic" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Metal" Mat="Mirror 1/4" MatThick="6.35" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Special" Mat="{sheet_name} PARTS [Matching Banding] BSAW" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="ClosetRod" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'</MaterialTemplate>'

        cab_temp_15 = \
            f'2\n' \
            f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
            f'<MaterialTemplate Name="15 - {interior_mat[0]} [{cab_temp_name}]" Type="3" SymbolForLabels="">\n' \
            f'  <MaterialReference PartType="Top" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Bottom" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FEnd" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UEnd" Mat="{interior_mat[1]} GABLE [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="AppliedUE" Mat="{interior_mat[1]} GABLE [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Shelf" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="AdjustableShelf" Mat="{interior_mat[1]} ADJ [{interior_mat[3]} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="RolloutShelf" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FixedShelf" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Partition" Mat="{interior_mat[1]} GABLE [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Divider" Mat="{interior_mat[1]} GABLE [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UBack" Mat="{interior_mat[1]} BACK [{interior_mat[3]} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FBack" Mat="{interior_mat[1]} BACK [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Frame" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FramelessRail" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Nailer" Mat="{interior_mat[1]} BACK [{interior_mat[3]} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Stretcher" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Cleat" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Toe" Mat="Raw Toe Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinExterior" Mat="{sheet_name} PARTS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinInteriorBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinExterior" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinInterior" Mat="{interior_mat[1]} SHELF [{case_mat_band_name}]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="ToeSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="EdgeBand" Mat="Banding" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Sleeper" Mat="Raw Toe Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Glass" Mat="1/8 Glass" MatThick="3.175" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Plastic" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Metal" Mat="Mirror 1/4" MatThick="6.35" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Special" Mat="{sheet_name} PARTS [Matching Banding] BSAW" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="ClosetRod" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'</MaterialTemplate>'

        cab_temp_path_05 = folder_path + f'\\05 - {interior_mat[0]} [{cab_temp_name}].CabTmp'
        cab_temp_path_15 = folder_path + f'\\15 - {interior_mat[0]} [{cab_temp_name}].CabTmp'

        with open(cab_temp_path_05, "wt") as f:
            f.writelines(cab_temp_05)

        with open(cab_temp_path_15, "wt") as f:
            f.writelines(cab_temp_15)


    # creates the door template
        door_temp = \
            f'2\n' \
            f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
            f'<MaterialTemplate Name="02 - {band_temp_name}" Type="1" SymbolForLabels="">\n' \
            f'  <MaterialReference PartType="Door" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="DoorPanel" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="DoorFrame" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Drawer" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="DrawerPanel" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="DrawerFrame" Mat="{sheet_name} DOORS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="PanelizedEnd" Mat="{sheet_name} PARTS [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'</MaterialTemplate>'

        door_temp_path = folder_path + f'\\02 - {band_temp_name}.DoorTmp'

        with open(door_temp_path, "wt") as f:
            f.writelines(door_temp)

    print('\n---------------------------------------------\nNew Materials & Templates Added Successfully!\n---------------------------------------------\n')



add_door_material()

input("Press Enter to Close")
