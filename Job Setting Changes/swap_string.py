import os


# This is to ask for the directory path at the command prompt
dir_path = input('What is the full directory path of the job you want to change? ')

hinge_type = None

# ask the user for the hinge type we're swapping to
string_to_replace = input('What string are we looking for? ')
new_string = input('What string are we replacing it with? ')

def swap_string():
    total_to_replace = 0
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

                    qty_to_replace = content.count(string_to_replace)

                    # if there are hinges with this name in the file
                    if qty_to_replace > 0:
                        # set the new hinge to the current iteration of swap_to hinge
                        content = content.replace(string_to_replace, new_string)
                        # check to see if any did not get replaced
                        qty_not_replaced = content.count(string_to_replace)
                        
                        if qty_not_replaced == 0:
                            print(
                                '  > Replaced (' + str(qty_to_replace) + ') "' + str(string_to_replace) +
                                '" with "' + str(new_string) + '".')
                        else:
                            print(
                                '  > Error: found (' + str(qty_to_replace) + ') "' + str(string_to_replace) +
                                '", but (' + str(qty_not_replaced) + ') were not replaced.'
                            )

                        total_to_replace += qty_to_replace
                        total_not_replaced += qty_not_replaced



                    f = open(full_path, "wt")
                    f.write(content)
                    f.close()

    print('Total strings replaced:', total_to_replace)
    print('Total strings missed:', total_not_replaced)

          

swap_string()
