import os
from typing import Optional



def add__door_material():
    """

    """
    # This is to ask for the directory path at the command prompt
    parent_path = input('What is the path for the directory where the Materials.dat file is located? ')

    mat_file_path = parent_path + '\\Materials.dat'
    f = open(mat_file_path, "rt")

    # grabs entire material.dat file contents
    curr_mat_list = f.readlines()

    # get the end of the sheet materials list
    curr_sheet_mat_end_idx = curr_mat_list.index('  </PanelStockMaterials>\n')

    # variables to capture start and end of banding material list
    curr_band_list_start_idx = curr_mat_list.index('  <EdgeBandingMaterials>\n') + 1
    curr_band_list_end_idx = curr_mat_list.index('  </EdgeBandingMaterials>\n')

    # grabs only the banding materials
    curr_band_list = curr_mat_list[curr_band_list_start_idx:curr_band_list_end_idx]


    # test sheet variables
    sheet_name = 'TEST MATERIAL 3/4 T24736'
    sheet_width = '1244.6'
    sheet_length = '2743.2'
    sheet_thick = '19.05'
    has_grain = 'True'
    two_sided = 'True'
    waste_percent = '20'
    optmize = 'True'
    sheet_comment = 'Nothing to see here'
    sheet_case_mat_name = 'TEST MATERIAL'

    # test case band variables
    case_band_exists = 'False'
    case_band_name = '0.5mm NEW CASE BANDING'
    case_band_width = '22.225'
    case_band_length = '91440.0'
    case_band_type = '0'
    case_band_comment = 'BANDING FOR CASE'

    # test door band variables
    door_band_exists = 'False'
    door_band_name = '1mm NEW DOOR BANDING'
    door_band_width = '22.225'
    door_band_length = '91440.0'
    door_band_type = '0'
    door_band_comment = 'BANDING FOR DOOR'


    # # user input sheet related variables
    # sheet_name = input('What do you want to name the sheet material? ')
    # sheet_width = str(float(input('What is the sheet width in inches? (48.5 format) ')) * 25.4)
    # sheet_length = str(float(input('What is the sheet length in inches? (96.5 format) ')) *25.4)
    # sheet_thick = str(float(input('What is the sheet thickness in inches? (0.75 format) ')) * 25.4)
    # has_grain = input('Does the material have grain? (Y or N) ')
    # has_grain = "True" if has_grain[0] == 'y' or has_grain[0] == 'Y' else "False"
    # two_sided = input('Is this material 2-sided? (Y or N) ')
    # two_sided = "True" if two_sided[0] == 'y' or two_sided[0] == 'Y' else "False"
    # waste_percent = input('What is the waste percentage? (20 = 20%) ')
    # optmize = input('Will this material be cut on the CNC? (Y or N) ')
    # optmize = "True" if optmize[0] == 'y' or optmize[0] == 'Y' else "False"
    # sheet_comment = input('What comment would you like to add to this material? ')
    # sheet_case_mat_name = input('What should the banding name be on the case material? (Exclude "Banding") ')

    # #banding related variables
    # case_band_exists = input('\nDoes the case banding already exist? (Y or N) ')
    # case_band_exists = 'True' if case_band_exists[0] == 'y' or case_band_exists[0] == 'Y' else 'False'
    # case_band_name = input('What is the name of the case banding? ')
    # case_band_width = input('What is the width of the case banding in inches? (0.875 format) ')
    # case_band_length = input('What is the length of the case banding in feet? ')
    # case_band_type = input('What is the case banding type? (0 = Roll, 1 = Strip) ')
    # case_band_comment = input('What comment would you like to add to this case banding? ')

    # door_band_exists = input('\nDoes the door banding already exist? (Y or N) ')
    # door_band_name = input('What is the name of the door banding? ')
    # door_band_width = input('What is the width of the door banding in inches? (0.875 format) ')
    # door_band_length = input('What is the length of the door banding in feet? ')
    # door_band_thick = input('What is the door banding thickness in millimeters? ')
    # door_band_type = input('What is the door banding type? (0 = Roll, 1 = Strip) ')
    # door_band_comment = input('What comment would you like to add to this door banding? ')


    # # creates folder where new material templates will be placed
    # folder_name = sheet_name.replace('/', '-')
    # folder_path = os.path.join(parent_path, folder_name)
    # os.mkdir(folder_path)

    # creates lines for new door materials, one CNC (standard) and one BSAW
    door_sheet_CNC = f'    <Material Name="{sheet_name} [Matching Banding]" '

    door_sheet_BSAW = f'    <Material Name="{sheet_name} [Matching Banding] BSAW" '

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

    door_sheet_CNC = door_sheet_CNC + door_sheet_mid + door_sheet_mid + door_sheet_CNC_trim + door_sheet_end
    door_sheet_BSAW = door_sheet_BSAW + door_sheet_mid + door_sheet_mid + door_sheet_BSAW_trim + door_sheet_end
    
    # creates lines for each of the case materials with new material banding
    birch_58 = \
        f'    <Material Name="CM Birch PF 5/8 [{sheet_case_mat_name} Banding]" ' \
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

    birch_34 = \
        f'    <Material Name="CM Birch PF 3/4 [{sheet_case_mat_name} Banding]" ' \
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

    fog_grey_58 = \
        f'    <Material Name="CM Fog Grey 5/8 SF213 PRZ [{sheet_case_mat_name} Banding]" ' \
        f'Quan="1" ' \
        f'W="1243.012" ' \
        f'L="2462.212" ' \
        f'Thick="15.875" ' \
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

    fog_grey_34 = \
        f'    <Material Name="CM Fog Grey 3/4 SF213 PRZ [{sheet_case_mat_name} Banding]" ' \
        f'Quan="1" ' \
        f'W="1243.012" ' \
        f'L="2462.212" ' \
        f'Thick="19.05" ' \
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

    hardrock_58 = \
        f'    <Material Name="CM Hardrock Maple 5/8 WF275 PRZ [{sheet_case_mat_name} Banding]" ' \
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

    hardrock_34 = \
        f'    <Material Name="CM Hardrock Maple 3/4 WF275 PRZ [{sheet_case_mat_name} Banding]" ' \
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

    white_58 = \
        f'    <Material Name="CM White 5/8 W100 PRZ [{sheet_case_mat_name} Banding]" ' \
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

    white_34 = \
        f'    <Material Name="CM White 3/4 W100 PRZ [{sheet_case_mat_name} Banding]" ' \
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

    storm_58 = \
        f'    <Material Name="Storm Gray 5/8 S565 PAN [{sheet_case_mat_name} Banding]" ' \
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

    storm_34 = \
        f'    <Material Name="Storm Gray 3/4 S565 PAN [{sheet_case_mat_name} Banding]" ' \
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


    # creates banding materials if they don't already exist
    if case_band_exists == 'False':
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

    if door_band_exists == 'False':    
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
    

    updated_material_list = curr_mat_list[:curr_sheet_mat_end_idx]
    updated_material_list.append(door_sheet_CNC)
    updated_material_list.append(door_sheet_BSAW)
    updated_material_list.append(birch_58)
    updated_material_list.append(birch_34)
    updated_material_list.append(fog_grey_58)
    updated_material_list.append(fog_grey_34)
    updated_material_list.append(hardrock_58)
    updated_material_list.append(hardrock_34)
    updated_material_list.append(white_58)
    updated_material_list.append(white_34)
    updated_material_list.append(storm_58)
    updated_material_list.append(storm_34)
    updated_material_list += curr_mat_list[curr_sheet_mat_end_idx:curr_band_list_end_idx]
    
    # adds new banding materials to banding list if they don't already exist
    if case_band_exists == 'False':
        updated_material_list.append(case_band)

    if door_band_exists == 'False':  
        updated_material_list.append(door_band)
    updated_material_list += curr_mat_list[curr_band_list_end_idx:]

    f = open('Material_New.dat', "wt")
    f.writelines(updated_material_list)
    f.close()






add__door_material()
