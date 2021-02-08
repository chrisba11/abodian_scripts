import os

# This is to ask for the directory path at the command prompt

def add__door_material():
    """

    """
    dir_path = input('In what directory is the Materials.dat file located? ')
    
    # sheet related variables
    sheet_name = input('What do you want to name the sheet material? ')
    sheet_width = input('What is the sheet width in inches? (48.5 format) ')
    sheet_length = input('What is the sheet length in inches? (96.5 format) ')
    sheet_thick = input('What is the sheet thickness in inches? (0.75 format) ')
    width_trim = input('How much trim should be removed from each long edge? ')
    length_trim = input('How much trim should be removed from each short edge? ')
    has_grain = input('Does the material have grain? (Y or N) ')
    two_sided = input('Is this material 2-sided? (Y or N) ')
    waste_percent = input('What is the waste percentage? (20 = 20%')
    optmize = input('Will this material be cut on the CNC? (Y or N) ')
    sheet_comment = input('What comment would you like to add to this material? ')

    #banding related variables
    case_band_name = input('What is the name of the case banding? ')
    case_band_width = input('What is the width of the case banding in inches? (0.875 format) ')
    case_band_length = input('What is the length of the case banding in feet? ')
    case_band_thick = input('What is the case banding thickness in millimeters? ')
    case_band_type = input('What is the case banding type? (0 = Roll, 1 = Strip) ')
    case_band_comment = input('What comment would you like to add to this case banding? ')

    door_band_name = input('What is the name of the door banding? ')
    door_band_width = input('What is the width of the door banding in inches? (0.875 format) ')
    door_band_length = input('What is the length of the door banding in feet? ')
    door_band_thick = input('What is the door banding thickness in millimeters? ')
    door_band_type = input('What is the door banding type? (0 = Roll, 1 = Strip) ')
    door_band_comment = input('What comment would you like to add to this door banding? ')






add__door_material()
