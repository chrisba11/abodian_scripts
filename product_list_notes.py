import os
from xlwt import Workbook


# This is to ask for the directory path at the command prompt
dir_path = input('What is the full directory path? ')


def product_list_with_notes():
    """
    Generates a file listing all of the products in the job that have special notes
    Opens each of the .des files inside a job directory and looks at every product in the room
    If the product has notes, it adds that product to the list
    Then it writes the list of products with notes to a new file
    """
    dir_path_lst = dir_path.split('\\')
    job_name = dir_path_lst[-1]

    xml_char_ents = [
        ['&quot;', '"'],
        ['&#xD;&#xA;', '\n'],
        ['&amp;', '&'],
        ['&apos;', "'"],
        ['&gt;', '>'],
        ['&lt;', '<']
    ]

    prod_list = []
    
    for root, dirs, files in os.walk(dir_path):
        curr = 0
        
        for file in files:

            if file.endswith('.des'):

                full_path = dir_path + '\\' + file
                f = open(full_path, "rt")
                content = f.readlines()
                rm_start_idx = content[2].find('Name=') + 6
                rm_end_idx = content[2].find('" RoomNosDirty=')
                room_name = content[2][rm_start_idx:rm_end_idx]
                prod_list.append([room_name,[]])
                                
                for line in content:
                    if line.startswith('    <Product '):
                        note_start_idx = line.find('Notes=') + 6
                        note_end_idx = line.find('" Price=')
                        note_intro = line[note_start_idx:note_start_idx + 2]

                        if note_intro != '""':
                            prod_start_idx = line.find('ProdName=') + 10
                            prod_end_idx = line.find('" IDTag=')
                            prod_name = line[prod_start_idx:prod_end_idx]
                            note = line[note_start_idx + 1:note_end_idx]

                            for char in xml_char_ents:
                                note = note.replace(char[0], char[1])
                            
                            prod_list[curr][1].append([prod_name, note])

                curr += 1

    # print statement to see clean list of lists (requires import json)
    # print(json.dumps(prod_list, indent=4))

    wb = Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    row = 0
    col = 0
    for room in prod_list:
        if room[1] != []:
            sheet1.write(row, 0, room[0])
            row += 1
            for prod in room[1]:
                sheet1.write(row, col, prod[0])
                col += 1
                sheet1.write(row, col, prod[1])
                col -= 1
                row += 1

            row += 1

    
    save_name = job_name + '.xls'
    full_save_name = os.path.join(dir_path, save_name)
    wb.save(full_save_name)

    
product_list_with_notes()
