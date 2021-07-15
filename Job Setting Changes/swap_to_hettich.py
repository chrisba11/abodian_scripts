import os


# This is to ask for the directory path at the command prompt
dir_path = input('What is the full directory path of the job you want to change? ')

string_to_replace = '"Blum 95 Blind / 3mm - 79B958180"'
new_string = '"Hettich - Sensys Blind 95/3mm Plate - 9088045 / 9071667"'

total_count = 0
room_dict = {}
hinge_dict = {
    '"Blum 95 Blind / 3mm - 79B958180"': {},
    '"Blum Bi-fold 60 / 0mm - 79T853180"': {},
    '"Blum 170 / 3mm - 71T658180"': {},
    # make this have all of the hinges to swap and all of the hinges they should swap to
}

def swap_hinge():
    for root, dirs, files in os.walk(dir_path):
        
        for file in files:

            if file.endswith('.des'):

                full_path = dir_path + '\\' + file
                f = open(full_path, "rt")
                content = f.read()
                
                room_dict[full_path] = {'count_before': content.count(string_to_replace)}
                content = content.replace(string_to_replace, new_string)
                room_dict[full_path]['count_after'] = content.count(string_to_replace)

                f = open(full_path, "wt")
                f.write(content)
                f.close()

    for room in room_dict:
        if room_dict[room]['count_before'] != 0:
            print("")
            print(room + ":", room_dict[room]['count_before'], 'found')
            print(room + ":", room_dict[room]['count_before'] - room_dict[room]['count_after'], 'replaced')
            

swap_hinge()
