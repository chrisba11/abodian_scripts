import os
import re



def add_door_material():
    """

    """
    # This is to ask for the directory path at the command prompt
    parent_path = input('What is the path for the directory where the Materials.dat & New_Material.txt files are located? ')

    # generates path for location of Materials.dat file and opens the file
    mat1_file_path = parent_path + '\\Materials.dat'
    f = open(mat1_file_path, "rt")

    # grabs entire Materials.dat file contents
    curr_mat_list = f.readlines()
    f.close()

    # generates path for location of New_Material.txt file and opens the file
    mat2_file_path = parent_path + '\\New_Material.txt'
    f = open(mat2_file_path, "rt")

    # grabs entire New_Material.txt file contents
    full_new_mat_file = f.readlines()
    f.close()

    # list of indices that are okay having blank strings
    blank_okay_idx = [8, 16, 21]

    # pulls out only the relevant inputs for variables
    new_mat = [None]
    for i in range(10, 72, 3):
        if full_new_mat_file[i] == '' and i not in blank_okay_idx:
            print('CHECK YOUR INPUT FOR BLANK LINES')
            print('\nOnly questions 8, 16, and 21 can be left blank.')
            return
        new_mat.append(full_new_mat_file[i][:-1])

    print(new_mat)

    # removes any special characters from user input
    for user_response in new_mat:
        if user_response is not None and user_response != '':
            user_response = re.sub("[@%&*'\"!?#~`<>\^\\\$\[\]\{\}\|\(\)]", '', user_response)

    # filling out variables with input read from New_Material.txt file
    sheet_name = new_mat[1]
    sheet_width = str(float(new_mat[2]) * 25.4)
    sheet_length = str(float(new_mat[3]) * 25.4)
    sheet_thick = str(float(new_mat[4]) * 25.4)
    has_grain = new_mat[5]
    has_grain = "True" if has_grain[0] == 'y' or has_grain[0] == 'Y' else "False"
    two_sided = new_mat[6]
    two_sided = "True" if two_sided[0] == 'y' or two_sided[0] == 'Y' else "False"
    waste_percent = new_mat[7]
    optmize = "True"
    sheet_comment = new_mat[8]
    case_mat_band_name = new_mat[9]
    band_temp_name = new_mat[10]
    cab_temp_name = new_mat[11]

    #banding related variables
    # case_band_exists = 
    # case_band_exists = 'True' if case_band_exists[0] == 'y' or case_band_exists[0] == 'Y' else 'False'
    case_band_exists = 'False'
    case_band_name = [12]
    case_band_width = str(float(new_mat[13]) * 25.4)
    case_band_length = str(float(new_mat[14]) * 12 * 25.4)
    case_band_type = new_mat[15]
    case_band_comment = new_mat[16]

    # door_band_exists = input('\nDoes the door banding already exist? (Y or N) ')
    # door_band_exists = 'True' if case_band_exists[0] == 'y' or case_band_exists[0] == 'Y' else 'False'
    door_band_exists = 'False'
    door_band_name = new_mat[17]
    door_band_width = str(float(new_mat[18]) * 25.4)
    door_band_length = str(float(new_mat[19]) * 12 * 25.4)
    door_band_type = new_mat[20]
    door_band_comment = new_mat[21]


    # creates folder where new material templates will be placed
    folder_name = sheet_name.replace('/', '-')
    folder_path = os.path.join(parent_path, folder_name)
    os.mkdir(folder_path)

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

    door_sheet_CNC = door_sheet_CNC + door_sheet_mid + door_sheet_CNC_trim + door_sheet_end
    door_sheet_BSAW = door_sheet_BSAW + door_sheet_mid + door_sheet_BSAW_trim + door_sheet_end
    
    # creates lines for each of the case materials with new material banding
    birch_58 = \
        f'    <Material Name="CM Birch PF 5/8 [{case_mat_band_name} Banding]" ' \
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
        f'    <Material Name="CM Birch PF 3/4 [{case_mat_band_name} Banding]" ' \
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
        f'    <Material Name="CM Fog Grey 5/8 SF213 PRZ [{case_mat_band_name} Banding]" ' \
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
        f'    <Material Name="CM Fog Grey 3/4 SF213 PRZ [{case_mat_band_name} Banding]" ' \
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
        f'    <Material Name="CM Hardrock Maple 5/8 WF275 PRZ [{case_mat_band_name} Banding]" ' \
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
        f'    <Material Name="CM Hardrock Maple 3/4 WF275 PRZ [{case_mat_band_name} Banding]" ' \
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
        f'    <Material Name="CM White 5/8 W100 PRZ [{case_mat_band_name} Banding]" ' \
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
        f'    <Material Name="CM White 3/4 W100 PRZ [{case_mat_band_name} Banding]" ' \
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
        f'    <Material Name="Storm Gray 5/8 S565 PAN [{case_mat_band_name} Banding]" ' \
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
        f'    <Material Name="Storm Gray 3/4 S565 PAN [{case_mat_band_name} Banding]" ' \
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
    
            
    # get the end of the sheet materials list
    curr_sheet_mat_end_idx = curr_mat_list.index('  </PanelStockMaterials>\n')

    # get the end of the banding material list
    curr_band_list_end_idx = curr_mat_list.index('  </EdgeBandingMaterials>\n')

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

    new_mat_path = folder_path + '\\Materials.dat'
    f = open(new_mat_path, "wt")
    f.writelines(updated_material_list)
    f.close()


    # creates all of the banding templates for each interior material
    for interior_band in [
        ['GC', '0.5mm Fog Grey SF213 PRZ'],
        ['MC', '0.5mm Hardrock Maple WF275 PRZ'],
        ['PC', '0.5mm Birch PVC'],
        ['SC', '0.5mm Storm Gray S565 PAN'],
        ['WC', '0.5mm White W100 PRZ']
        ]:

        band_temp = \
            f'2\n' \
            f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
            f'<MaterialTemplate Name="02 - {band_temp_name} - {interior_band[0]}" Type="4" SymbolForLabels="{band_temp_name[0:2]}">\n' \
            f'  <MaterialReference PartType="EdgeBand" Mat="{case_band_name}" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="EdgeBand2" Mat="{door_band_name}" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="EdgeBand3" Mat="{interior_band[1]}" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="EdgeBand4" Mat="None EB4" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'</MaterialTemplate>'

        band_temp_path = folder_path + f'\\02 - {band_temp_name} - {interior_band[0]}.BandTmp'

        f = open(band_temp_path, "wt")
        f.writelines(band_temp)
        f.close()


    # creates the finished interior cabinet templates
    fin_int_case_temp_02 = \
        f'2\n' \
        f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
        f'<MaterialTemplate Name="02 - {band_temp_name} .75" Type="3" SymbolForLabels="">\n' \
        f'  <MaterialReference PartType="Top" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Bottom" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FEnd" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
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
        f'  <MaterialReference PartType="FramelessRail" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Nailer" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Stretcher" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Cleat" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Toe" Mat="Ladderbox Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinExterior" Mat="{sheet_name} [Matching Banding] BSAW" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinInteriorBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinExterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="ToeSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="EdgeBand" Mat="Banding" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Sleeper" Mat="Ladderbox Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Glass" Mat="1/8 Glass" MatThick="3.175" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Plastic" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Metal" Mat="Mirror 1/4" MatThick="6.35" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Special" Mat="{sheet_name} [Matching Banding] BSAW" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="ClosetRod" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'</MaterialTemplate>'

    fin_int_case_temp_12 = \
        f'2\n' \
        f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
        f'<MaterialTemplate Name="12 - {band_temp_name} .75" Type="3" SymbolForLabels="">\n' \
        f'  <MaterialReference PartType="Top" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Bottom" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FEnd" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
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
        f'  <MaterialReference PartType="FramelessRail" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Nailer" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Stretcher" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Cleat" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Toe" Mat="Ladderbox Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinExterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinInteriorBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinExterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="FinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="UfinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="ToeSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="EdgeBand" Mat="Banding" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Sleeper" Mat="Ladderbox Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Glass" Mat="1/8 Glass" MatThick="3.175" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Plastic" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Metal" Mat="Mirror 1/4" MatThick="6.35" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="Special" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
        f'  <MaterialReference PartType="ClosetRod" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
        f'</MaterialTemplate>'

    fin_int_case_temp_path_02 = folder_path + f'\\02 - {band_temp_name} .75.CabTmp'
    fin_int_case_temp_path_12 = folder_path + f'\\12 - {band_temp_name} .75.CabTmp'

    f = open(fin_int_case_temp_path_02, "wt")
    f.writelines(fin_int_case_temp_02)
    f.close()

    f = open(fin_int_case_temp_path_12, "wt")
    f.writelines(fin_int_case_temp_12)
    f.close()


    # creates the standard cabinet templates
    for interior_mat in [
        ['Fog Grey 34', 'CM Fog Grey 3/4 SF213 PRZ', '19.05'],
        ['Fog Grey 58', 'CM Fog Grey 5/8 SF213 PRZ', '15.875'],
        ['Maple 34', 'CM Hardrock Maple 3/4 WF275 PRZ', '19.19986'],
        ['Maple 58', 'CM Hardrock Maple 5/8 WF275 PRZ', '16.10106'],
        ['PF Ply 34', 'CM Birch PF 3/4', '17.80032'],
        ['PF Ply 58', 'CM Birch PF 5/8', '15.8'],
        ['Storm 34', 'Storm Gray 3/4 S565 PAN', '19.2'],
        ['Storm 58', 'Storm Gray 5/8 S565 PAN', '16.2'],
        ['White 34', 'CM White 3/4 W100 PRZ', '19.19986'],
        ['White 58', 'CM White 5/8 W100 PRZ', '16.10106']
        ]:

        cab_temp_05 = \
            f'2\n' \
            f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
            f'<MaterialTemplate Name="05 - {interior_mat[0]} [{cab_temp_name}]" Type="3" SymbolForLabels="">\n' \
            f'  <MaterialReference PartType="Top" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Bottom" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FEnd" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UEnd" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="AppliedUE" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Shelf" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="AdjustableShelf" Mat="{interior_mat[1]} [Matching Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="RolloutShelf" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FixedShelf" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Partition" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Divider" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UBack" Mat="{interior_mat[1]} [Matching Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FBack" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Frame" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FramelessRail" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Nailer" Mat="{interior_mat[1]} [Matching Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Stretcher" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Cleat" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Toe" Mat="Ladderbox Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinExterior" Mat="{sheet_name} [Matching Banding] BSAW" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinInteriorBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinExterior" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinInterior" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="ToeSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="EdgeBand" Mat="Banding" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Sleeper" Mat="Ladderbox Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Glass" Mat="1/8 Glass" MatThick="3.175" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Plastic" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Metal" Mat="Mirror 1/4" MatThick="6.35" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Special" Mat="{sheet_name} [Matching Banding] BSAW" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="ClosetRod" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'</MaterialTemplate>'

        cab_temp_15 = \
            f'2\n' \
            f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
            f'<MaterialTemplate Name="15 - {interior_mat[0]} [{cab_temp_name}]" Type="3" SymbolForLabels="">\n' \
            f'  <MaterialReference PartType="Top" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Bottom" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FEnd" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UEnd" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="AppliedUE" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Shelf" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="AdjustableShelf" Mat="{interior_mat[1]} [Matching Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="RolloutShelf" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FixedShelf" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Partition" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Divider" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UBack" Mat="{interior_mat[1]} [Matching Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FBack" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Frame" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FramelessRail" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Nailer" Mat="{interior_mat[1]} [Matching Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Stretcher" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Cleat" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Toe" Mat="Ladderbox Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinExterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinInterior" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinInteriorBack" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinExterior" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinInterior" Mat="{interior_mat[1]} [{case_mat_band_name} Banding]" MatThick="{interior_mat[2]}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="FinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="UfinSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="ToeSkin" Mat="NONE" MatThick="19.05" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="EdgeBand" Mat="Banding" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Sleeper" Mat="Ladderbox Material 5/8" MatThick="16.2" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Glass" Mat="1/8 Glass" MatThick="3.175" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Plastic" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Metal" Mat="Mirror 1/4" MatThick="6.35" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Special" Mat="{sheet_name} [Matching Banding] BSAW" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="ClosetRod" Mat="" MatThick="0" MatWall="" MatWallThick="0" />\n' \
            f'</MaterialTemplate>'

        cab_temp_path_05 = folder_path + f'\\05 - {interior_mat[0]} [{cab_temp_name}].CabTmp'
        cab_temp_path_15 = folder_path + f'\\15 - {interior_mat[0]} [{cab_temp_name}].CabTmp'

        f = open(cab_temp_path_05, "wt")
        f.writelines(cab_temp_05)
        f.close()

        f = open(cab_temp_path_15, "wt")
        f.writelines(cab_temp_15)
        f.close()


    # creates the door template
        door_temp = \
            f'2\n' \
            f'<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n' \
            f'<MaterialTemplate Name="02 - {band_temp_name}" Type="1" SymbolForLabels="">\n' \
            f'  <MaterialReference PartType="Door" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="DoorPanel" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="DoorFrame" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="Drawer" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="DrawerPanel" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="DrawerFrame" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'  <MaterialReference PartType="PanelizedEnd" Mat="{sheet_name} [Matching Banding]" MatThick="{sheet_thick}" MatWall="" MatWallThick="0" />\n' \
            f'</MaterialTemplate>'

        door_temp_path = folder_path + f'\\02 - {band_temp_name}.DoorTmp'

        f = open(door_temp_path, "wt")
        f.writelines(door_temp)
        f.close()



add_door_material()
