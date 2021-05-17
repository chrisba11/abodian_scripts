import os


# This is to ask for the directory path at the command prompt
dir_path = input('What is the full directory path? ')


def zero_spash_height():
    for root, dirs, files in os.walk(dir_path):
        
        for file in files:

            if file.endswith('.des'):

                full_path = dir_path + '\\' + file
                f = open(full_path, "rt")
                content = f.readlines()
                splash_start_idx = content[3].find('CT_Splash=') + 11
                splash_end_idx = content[3].find('" CT_Face=')
                new_line = content[3][:splash_start_idx] + '0' + content[3][splash_end_idx:]
                content[3] = new_line

                f = open(full_path, "wt")
                f.writelines(content)
                f.close()

zero_spash_height()
