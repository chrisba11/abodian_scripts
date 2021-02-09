import os
from typing import Optional



def add__door_material():
    """

    """
    # This is to ask for the directory path at the command prompt
    parent_path = input('What is the path for the directory where the Materials.dat file is located? ')


    # test sheet variables
    sheet_name = 'TEST MATERIAL 3/4 T24736'
    sheet_width = '49'
    sheet_length = '108'
    sheet_thick = '.75'
    width_trim = '.25'
    length_trim = '.25'
    has_grain = 'True'
    two_sided = 'True'
    waste_percent = '20'
    optmize = 'True'
    sheet_comment = 'Nothing to see here'

    # user input sheet related variables
    # sheet_name = input('What do you want to name the sheet material? ')
    # sheet_width = str(int(input('What is the sheet width in inches? (48.5 format) ')) * 25.4)
    # sheet_length = str(int(input('What is the sheet length in inches? (96.5 format) ')) *25.4)
    # sheet_thick = str(int(input('What is the sheet thickness in inches? (0.75 format) ')) * 25.4)
    # width_trim = str(int(input('How much trim should be removed from each long edge? (0.25 format)  ')) * 25.4)
    # length_trim = str(int(input('How much trim should be removed from each short edge? (0.25 format)  ')) * 25.4)
    # has_grain = input('Does the material have grain? (Y or N) ')
    # has_grain = "True" if has_grain == 'y' or has_grain == 'Y' else "False"
    # two_sided = input('Is this material 2-sided? (Y or N) ')
    # two_sided = "True" if two_sided == 'y' or two_sided == 'Y' else "False"
    # waste_percent = input('What is the waste percentage? (20 = 20%) ')
    # optmize = input('Will this material be cut on the CNC? (Y or N) ')
    # optmize = "True" if optmize == 'y' or optmize == 'Y' else "False"
    # sheet_comment = input('What comment would you like to add to this material? ')

    #banding related variables
    # case_band_exists = input('\nDoes the case banding already exist? (Y or N) ')
    # case_band_name = input('What is the name of the case banding? ')
    # case_band_width = input('What is the width of the case banding in inches? (0.875 format) ')
    # case_band_length = input('What is the length of the case banding in feet? ')
    # case_band_thick = input('What is the case banding thickness in millimeters? ')
    # case_band_type = input('What is the case banding type? (0 = Roll, 1 = Strip) ')
    # case_band_comment = input('What comment would you like to add to this case banding? ')

    # door_band_exists = input('\nDoes the door banding already exist? (Y or N) ')
    # door_band_name = input('What is the name of the door banding? ')
    # door_band_width = input('What is the width of the door banding in inches? (0.875 format) ')
    # door_band_length = input('What is the length of the door banding in feet? ')
    # door_band_thick = input('What is the door banding thickness in millimeters? ')
    # door_band_type = input('What is the door banding type? (0 = Roll, 1 = Strip) ')
    # door_band_comment = input('What comment would you like to add to this door banding? ')


    folder_name = sheet_name.replace('/', '-')
    folder_path = os.path.join(parent_path, folder_name)
    os.mkdir(folder_path)

    

    sheet_mat1 = \
        f'    <Material Name="{sheet_name} [Matching Banding]" '

    sheet_mat2 = \
        f'    <Material Name="{sheet_name} [Matching Banding] BSAW" '

    sheet_mat_end = \
        f'Quan="1" W="{sheet_width}" ' \
        f'L="{sheet_length}" ' \
        f'Thick="{sheet_thick}" ' \
        f'WTrim="{width_trim}" ' \
        f'LTrim="{length_trim}" ' \
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
        f'BandType="0" />'

    sheet_mat1 += sheet_mat_end
    sheet_mat2 += sheet_mat_end

    
    print(sheet_mat1)
    print(sheet_mat2)



add__door_material()
