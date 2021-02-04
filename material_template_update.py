import os


# This is to get the directory that the program is currently running in
# dir_path = os.path.dirname(os.path.realpath(__file__))

# This is to ask for the directory path at the command prompt
dir_path = input('What is the directory path? ')


def update_with_start_string(new_text, line_of_code, start_of_filename):
    """
    Updates the specified line of the file with the new_text string
    Applies only to files ending with the start_of_filename string specified
    """
    for root, dirs, files in os.walk(dir_path): 
        for file in files:  
    
            # opens a filename starting with start_of_filename parameter
            # creates a list of all lines in the file
            # replaces specified index in the list with the text supplied
            # writes over file with new list content
            # closes the file
            # moves to next filename with matching string at beginning and repeats

            if file.startswith(start_of_filename):
                f = open(file, "rt")
                content = f.readlines()
                content[line_of_code - 1] = new_text

                f = open(file, "wt")
                f.writelines(content)
                f.close()


def update_with_end_string(new_text, line_of_code, end_of_filename):
    """
    Updates the specified line of the file with the new_text string
    Applies only to files ending with the end_of_filename string specified
    """
    for root, dirs, files in os.walk(dir_path): 
        for file in files:  
    
            # opens a filename ending with end_of_filename parameter
            # creates a list of all lines in the file
            # replaces specified index in the list with the text supplied
            # writes over file with new list content
            # closes the file
            # moves to next filename with matching string at end and repeats

            if file.endswith(end_of_filename):
                f = open(file, "rt")
                content = f.readlines()
                content[line_of_code - 1] = new_text

                f = open(file, "wt")
                f.writelines(content)
                f.close()


def update_with_start_and_end_strings(new_text, line_of_code, insertion_point, start_of_filename, end_of_filename):
    """
    Updates the specified line of the file with the new_text string
    Applies only to files ending with the end_of_filename string specified
    """
    for root, dirs, files in os.walk(dir_path): 
        for file in files:  

            if file.startswith(start_of_filename):
                if file.endswith(end_of_filename):
                    f = open(file, "rt")
                    content = f.readlines()
                    existing_line = content[line_of_code-1]
                    idx = existing_line.find(insertion_point)
                    new_line = existing_line[:idx] + new_text + existing_line[idx:]
                    
                    content[line_of_code-1] = new_line

                    f = open(file, "wt")
                    f.writelines(content)
                    f.close()


def copy_materials(starting_line, ending_line, append_string, new_filename):
    """
    This function looks at the Materials.dat file inside the given directory
    It reads all of the content of that file into the content variable
    Then it looks at the lines specified by starting_line and ending_line
    For those lines only, it duplicates the line and appends the append_string after the first "]" character
    It writes all of the content from the original and the new content to a new file called new_filename
    """
    
    for root, dirs, files in os.walk(dir_path): 
        for file in files:  

            if file.startswith('Materials.dat'):
                f = open(file, "rt")
                content = f.readlines()
                new_content = content[:starting_line-1]
                
                for line in content[starting_line-1:ending_line]:

                    idx = line.index(']')

                    edited_line = line[:idx+1] + append_string + line[idx+1:]
                    new_content.append(line)
                    new_content.append(edited_line)

                new_content += content[ending_line:]

                f = open(new_filename, "wt")
                f.writelines(new_content)
                f.close()


def edit_materials_2_params(filename, search_string, param1, value1, param2, value2, param3, new_filename):
    """
    This function looks at the file matching filename inside the given directory
    It reads all of the content of that file into the content variable
    For each line of the content array, it looks for the search_string in the line
    If if finds the search string, it replaces the values for param1 and param2 with value1 and value2
    Then it adds the line to the new_content array
    If it does not find the search string, it appends the original line to the new_content array
    When complete, it writes all of the new content to the new_filename file
    """

    for root, dirs, files in os.walk(dir_path): 
        for file in files:  

            if file.startswith(filename):
                f = open(file, "rt")
                content = f.readlines()
                new_content = []
                
                for line in content:
                    match = line.find(search_string)
                    if match != -1:
                        idx1 = line.find(param1)
                        idx2 = line.find(param2)
                        idx3 = line.find(param3)

                        new_line=line[:idx1]
                        new_line = new_line + param1 + '"' + str(value1) + '" '
                        new_line = new_line + param2 + '"' + str(value2) + '" '
                        new_line = new_line + line[idx3:]
                        
                        new_content.append(new_line)
                    else:
                        new_content.append(line)
                            

                f = open(new_filename, "wt")
                f.writelines(new_content)
                f.close()



def update_multiple_lines_with_start_and_end_strings(new_text, lines_of_code_array, insertion_point_start, insertion_point_end, start_of_filename, end_of_filename):
    """
    Updates the specified line of the file with the new_text string
    Applies only to files ending with the start_of_filename and
    end_of_filename string specified
    """
    for root, dirs, files in os.walk(dir_path): 
        for file in files:  

            if file.startswith(start_of_filename):
                if file.endswith(end_of_filename):
                    full_path = dir_path + '\\' + file
                    f = open(full_path, "rt")
                    content = f.readlines()

                    for line_number in lines_of_code_array:
                        existing_line = content[line_number-1]
                        idx1 = existing_line.find(insertion_point_start)
                        idx2 = existing_line.find(insertion_point_end)
                        new_line = existing_line[:idx1+1] + new_text + existing_line[idx2:]
                        
                        content[line_number-1] = new_line

                    f = open(full_path, "wt")
                    f.writelines(content)
                    f.close()


def copy_mat_to_different_line_in_template(line_to_copy, line_to_replace, trailing_text):
    """
    Copies everything after the part type from line_to_copy
    and replaces everything on line_to_replace after the part type
    with the copy of that material with trailing_text on the end of the material name
    """
    for root, dirs, files in os.walk(dir_path): 
        for file in files:  

            full_path = dir_path + '\\' + file
            f = open(full_path, "rt")

            content = f.readlines()
            copy_string = content[line_to_copy-1]

            part_idx1 = copy_string.find("=")
            part_idx2 = copy_string.find(" Mat")

            mat_idx2 = copy_string.find('" MatThick=')

            paste_string = copy_string[:part_idx1]
            paste_string += content[line_to_replace-1][part_idx1:part_idx2]
            paste_string += copy_string[part_idx2:mat_idx2]
            paste_string += trailing_text
            paste_string += copy_string[mat_idx2:]

            content[line_to_replace-1] = paste_string

            f = open(full_path, "wt")
            f.writelines(content)
            f.close()



    
# replacement_text = '  <MaterialReference PartType="AdjustableShelf" Mat="CM Fog Grey 3/4 SF213 PRZ [Matching Banding]" MatThick="19.05" MatWall="" MatWallThick="0" />\n'
# update_with_start_string(replacement_text, 10, '05 - Fog Grey 34')

# copy_materials(31,89, ' BSAW', 'Materials2.dat')

# update_with_start_and_end_strings(' BSAW', 36, '" MatThick', '05 - ', '.CabTmp')

# edit_materials_2_params('Materials.dat', 'BSAW', 'WTrim=', 12, 'LTrim=', 12, 'HasGrain', 'Edited_Materials.dat')

# update_multiple_lines_with_start_and_end_strings('Matching Banding', [15, 19], '[', ']', '05 - ', '.CabTmp')

# copy_mat_to_different_line_in_template(24, 23, " BSAW")
