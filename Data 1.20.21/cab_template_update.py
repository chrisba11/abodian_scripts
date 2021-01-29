import os

# This is to get the directory that the program  
# is currently running in
dir_path = os.path.dirname(os.path.realpath(__file__)) 


def update_file(new_text, line_of_code, end_of_filename):
    for root, dirs, files in os.walk(dir_path): 
        for file in files:  
    
            # creates a list of all lines in a file with '.CabTmp' extension
            # replaces specified index in the list with the text supplied
            # writes over file with new list content
            # closes the file
            if file.endswith(end_of_filename):
                f = open(file, "rt")
                content = f.readlines()
                content[line_of_code - 1] = new_text

                f = open(file, "wt")
                f.writelines(content)
                f.close


replacement_text = '  <MaterialReference PartType="Metal" Mat="Mirror 1/4" MatThick="6.35" MatWall="" MatWallThick="0" />\n'

# update_file(replacement_text, 35, '.CabTmp')