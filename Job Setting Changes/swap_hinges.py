import os


# This is to ask for the directory path at the command prompt
dir_path = input('What is the full directory path of the job you want to change? ')

hinge_type = None

# ask the user for the hinge type we're swapping to
while hinge_type not in ['blum', 'dtc', 'hettich']:
    hinge_type = input('Which hinge are we switching to? Please type "blum", "dtc", or "hettich". ').lower()


# dict to hold room names and replacement counts for each hinge type
room_dict = {}

# dict of all hinges being considered for swapping from each manufacturer
all_hinges = {
    'blum': {
        'bifold': 'Blum Bi-fold 60 / 0mm - 79T853180',
        'blind': 'Blum 95 Blind / 3mm - 79B958180',
        'corner90-58': 'Blum 170 / 3mm - 71T658180',
        'corner90-34': 'Blum 170 / 0mm - 71T658180',
    },
    'dtc': {
        'bifold': 'DTC Bi-fold 135 / 2mm - 105-H425N',
        'blind': 'DTC C-80 Blind 90 / 2mm - 105-C96J6A5NH',
        'corner90-58': 'DTC 165 / 2mm - 105-A505N',
        'corner90-34': 'DTC 165 / 0mm - 105-A505N',
    },
    'hettich': {
        'bifold': 'Hettich - Sensys Bi-fold 50-65 / 3mm - 9090116',
        'blind': 'Hettich - Sensys Blind 95 / 3mm - 9088045',
        'corner90-58': 'Hettich - Sensys 165 / 3mm - 9099621',
        'corner90-34': 'Hettich - Sensys 165 / 0mm - 9099621',
    },
}

h_edge_b_dict = {
    'blum': '23.1',
    'dtc': '17',
    'hettich': '30'
}

# removes dict of selected brand from all hinges dict
# stores brand of hinges swapping to in new dict variable
swap_to = all_hinges.pop(hinge_type)

# all_hinges dict now only contains the hinges that will be replaced

def swap_hinges():
    # variable for total hinges to swap
    total_to_replace = 0
    # variable for total hinges not swapped that should have been
    # this should always be 0 as long as the .replace() method works correctly
    total_not_replaced = 0

    for root, dirs, files in os.walk(dir_path):     
        for file in files:
            if file.endswith('.des'):

                full_path = dir_path + '\\' + file
                with open(full_path, "rt") as f:
                    content = f.read()

                    print('---------------')
                    print('Room:', file)
                    print('---------------')

                    # iterate through each hinge type (currently 4 types)
                    for hinge in swap_to:
                        # cycle through each of the remaining hinge brands
                        # swap anything that matches the other hinges with the swap to version
                        for brand in all_hinges:
                            hinge_being_replaced = all_hinges[brand][hinge]
                            qty_to_replace = content.count(hinge_being_replaced)

                            # if there are hinges with this name in the file
                            if qty_to_replace > 0:
                                # set the new hinge to the current iteration of swap_to hinge
                                new_hinge = swap_to[hinge]
                                # replace the current hinge string with the new hinge string
                                content = content.replace(hinge_being_replaced, new_hinge)
                                # check to see if any did not get replaced
                                qty_not_replaced = content.count(hinge_being_replaced)
                                
                                if qty_not_replaced == 0:
                                    print(
                                        '  > Replaced (' + str(qty_to_replace) + ') "' + str(hinge_being_replaced) +
                                        '" with "' + str(new_hinge) + '".')
                                else:
                                    print(
                                        '  > Error: found (' + str(qty_to_replace) + ') "' + str(hinge_being_replaced) +
                                        '", but (' + str(qty_not_replaced) + ') were not replaced.'
                                    )

                                total_to_replace += qty_to_replace
                                total_not_replaced += qty_not_replaced



                with open(full_path, "wt") as f:
                    f.write(content)

    
    print('Total hinges replaced:', total_to_replace)
    print('Total hinges missed:', total_not_replaced)

def swap_hedgeb():
    for root, dirs, files in os.walk(dir_path):     
        for file in files:
            if file.endswith('-JobParms.dat'):

                full_path = dir_path + '\\' + file
                with open(full_path, "rt") as f:
                    content = f.readlines()

                    line_num = 0
                    for line in content:
                        line_num += 1
                        if line.startswith('HEdgeB'):
                            content[line_num + 1] = h_edge_b_dict[hinge_type] + '\n'
                            break
                
                with open(full_path, "wt") as f:
                    f.writelines(content)

    print('\nHEdgeB = ' + h_edge_b_dict[hinge_type] + '\n')

swap_hinges()
swap_hedgeb()

input("Press Enter to Close")