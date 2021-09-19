import os
from openpyxl.workbook import Workbook
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.worksheet.pagebreak import Break


# This is to ask for the directory path at the command prompt
dir_path = input('What is the full directory path for the job? ')


def hinge_boring_report():
    """

    """
    dir_path_lst = dir_path.split('\\')
    job_name = dir_path_lst[-1]

    product_dict = {}

# #    product_dict sample
#     product_dict = {
#         "Room1": {
#             "MatDoorTemplate": "02 - 5pc Painted - In House MDF (1)",
#             "Products": {
#                 "C4": {
#                     "UniqueID": "3101352",
#                     "Name": "Base Sink",
#                     "MatOR": None,
#                     "Doors": [
#                         {
#                             "DoorStyle": "Shaker Door",
#                             "Name": "Door(L)",
#                             "ReportName": "Door(L)",
#                             "Comment": "Cabinet Door",
#                             "Quan": 1,
#                             "W": 454.0125,
#                             "H": 771.5,
#                             "HingeCenterLines": [101.6, 542.9],
#                             "HingeEdge": "Left",
#                             "HingeType": "DTC C-80 110 / 2mm - 105-C80A675NF",
#                             "Horizontal": False,
#                         },
#                         {
#                             "DoorStyle": "Shaker Door",
#                             "Name": "Door(R)",
#                             "ReportName": "Door(R)",
#                             "Comment": "Cabinet Door",
#                             "Quan": 1,
#                             "W": 454.0125,
#                             "H": 771.5,
#                             "HingeCenterLines": [101.6, 542.9],
#                             "HingeEdge": "Right",
#                             "HingeType": "DTC C-80 110 / 2mm - 105-C80A675NF",
#                             "Horizontal": False,
#                         },
#                     ]
#                 },
#                 "C6": {
#                     "UniqueID": "3101389",
#                     "Name": "Base Corner Left - Susan (L)",
#                     "MatOR": None,
#                     "Doors": [
#                         {
#                             "DoorStyle": "Shaker Door",
#                             "Name": "Door(L)",
#                             "ReportName": "Door(L)",
#                             "Comment": "Cabinet Door",
#                             "Quan": 1,
#                             "W": 454.0125,
#                             "H": 771.5,
#                             "HingeCenterLines": [101.6, 542.9],
#                             "HingeEdge": "Left",
#                             "HingeType": "DTC C-80 110 / 2mm - 105-C80A675NF",
#                             "Horizontal": False,
#                         },
#                         {
#                             "DoorStyle": "Shaker Door",
#                             "Name": "Door(L)",
#                             "ReportName": "Door(L)",
#                             "Comment": "Cabinet Door",
#                             "Quan": 1,
#                             "W": 454.0125,
#                             "H": 771.5,
#                             "HingeCenterLines": [101.6, 542.9],
#                             "HingeEdge": "Right",
#                             "HingeType": "DTC C-80 110 / 2mm - 105-C80A675NF",
#                             "Horizontal": False,
#                         },
#                     ]
#                 },
#             },
#         },
#     }

    
    for root, dirs, files in os.walk(dir_path):
        
        for file in files:

            if file.endswith('.des'):

                full_path = dir_path + '\\' + file
                f = open(full_path, "rt")
                content = f.readlines()
                f.close()

                # room settings
                room_num = file[4:file.find('.')]
                mat_temp_start_idx = content[4].find('MatDoorTemplate=') + 17
                mat_temp_end_idx = content[4].find('" MatDrawerTemplate=')
                room_mat_door_template = content[4][mat_temp_start_idx:mat_temp_end_idx]

                # add room to product list dictionary
                product_dict[room_num] = {
                    "MatDoorTemplate": room_mat_door_template,
                    "Products": []
                }

                # set products var to current room's products list
                products = product_dict[room_num]["Products"]

                # declaring variables that are used in multiple scopes
                doors = None

                for line in content:
                    if line.startswith('    <Product '):
                        # product number
                        start_idx = line.find('CabNo=') + 7
                        end_idx = line.find('" Numbered=')
                        prod_num = line[start_idx:end_idx]
                        
                        # product number prefix
                        start_idx = end_idx + 12
                        num_prefix = 'C' if line[start_idx:start_idx + 1] == 'T' else 'N'
                        
                        # unique id
                        start_idx = line.find('UniqueID=') + 10
                        end_idx = line.find('" OrderID=')
                        unique_id = line[start_idx:end_idx]

                        # product name
                        start_idx = line.find('ProdName=') + 10
                        end_idx = line.find('" IDTag=')
                        prod_name = line[start_idx:end_idx]

                        # door material override
                        start_idx = line.find('MatDoorOR=') + 11
                        end_idx = line.find('" MatDrawerOR=')
                        door_mat_or = line[start_idx:end_idx] if line[start_idx:end_idx] != '' else None

                        # add product to room's product list
                        products[num_prefix + prod_num] = {
                            "UniqueID": unique_id,
                            "Name": prod_name,
                            "MatOR": door_mat_or,
                            "Doors": []
                        }

                        # set doors var to current product's door list
                        doors = products[num_prefix + prod_num]["Doors"]


                    if line.startswith('        <ProductDoor '):
                        door = {}

                        # door style
                        start_idx = line.find('DoorStyle=') + 11
                        end_idx = line.find('" W=')
                        door_style = line[start_idx:end_idx]

                        # width
                        start_idx = end_idx + 5
                        end_idx = line.find('" H=')
                        width = float(line[start_idx:end_idx])
                        
                        # height
                        start_idx = end_idx + 5
                        end_idx = line.find('" Oversize=')
                        height = float(line[start_idx:end_idx])

                        # hinge center lines
                        start_idx = line.find('HingeCenterLines=') + 18
                        end_idx = line.find('" HingeEdge=')
                        hinge_centers = line[start_idx:end_idx].split('#')
                        hinge_centers = [float(i) for i in hinge_centers]

                        # hinge edge
                        start_idx = end_idx + 13
                        end_idx = line.find('" HingeType=')
                        hinge_edge = line[start_idx:end_idx]

                        # hinge type
                        start_idx = end_idx + 13
                        end_idx = line.find('" IsDrawerFront=')
                        hinge_type = line[start_idx:end_idx]

                        # horizontal grain?
                        start_idx = line.find('IsHorizGrain=') + 14
                        end_idx = line.find('" DoorType=')
                        horizontal = line[start_idx:end_idx]
                        horizontal = True if horizontal == 'True' else False


                        # add to door dict
                        door["DoorStyle"] = door_style
                        door["W"] = width
                        door["H"] = height
                        door["HingeCenterLines"] = hinge_centers
                        door["HingeEdge"] = hinge_edge
                        door["HingeType"] = hinge_type
                        door["IsHorizGrain"] = horizontal



                    if line.startswith('          <DoorProdPart ') or line.startswith('            <DoorProdPart '):
                        # door name
                        start_idx = line.find('Name=') + 6
                        end_idx = line.find('" ReportName=')
                        door_name = line[start_idx:end_idx]
                        
                        # report name
                        start_idx = line.find('ReportName=') + 12
                        end_idx = line.find('" UsageType=')
                        report_name = line[start_idx:end_idx]

                        # comment

                        # quantity

                        # 

                    

    
    for template in temp_dict:
        for door_style in temp_dict[template]:
            temp_dict[template][door_style].sort(key=lambda x: int(x[5]))
            temp_dict[template][door_style].sort(key=lambda x: int(x[3]))
            temp_dict[template][door_style].reverse()
            temp_dict[template][door_style].sort(key=lambda x: x[2])
            temp_dict[template][door_style].sort(key=lambda x: x[1])
            temp_dict[template][door_style].reverse()

    # print('--------')        
    # for template in temp_dict:
    #     if temp_dict[template]:
    #         print('Template Name:', template)
    #         for door_style in temp_dict[template]:
    #             print('   Door Style:', door_style)
    #             for door in temp_dict[template][door_style]:
    #                 print('     ', door[0] + ',', str(door[1]) + ',', str(door[2]) + ',', door[3])
    # print('--------')


    wb = Workbook()
    sheet1 = wb.active
    row = 1
    col = 1
    
    
    for template in temp_dict:
        if temp_dict[template]:

            sheet1.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 3)
            sheet1.cell(row, col, job_name + ' - Door List')
            sheet1.cell(row, col).alignment = Alignment(vertical='top')
            sheet1.cell(row, col).font = Font(size=12, bold=True, italic=True)
            sheet1.row_dimensions[row].height = 40
            row += 1

            sheet1.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 3)
            sheet1.cell(row, col, template)
            sheet1.cell(row, col).alignment = Alignment(wrapText=True, horizontal='left')
            sheet1.cell(row, col).font = Font(size=14, bold=True, underline='single')
            row += 1
            page_total = 0

            for door_style in temp_dict[template]:
                sheet1.cell(row, col, door_style)
                sheet1.cell(row, col).alignment = Alignment(wrapText=True, vertical='top', indent=1.0)
                sheet1.cell(row, col).font = Font(size=12, bold=True, italic=True)
                sheet1.cell(row, col).border = Border(bottom=Side(style='thin', color='000000'))
                col += 1
                sheet1.cell(row, col, "W")
                sheet1.cell(row, col).alignment = Alignment(wrapText=True, vertical='center', indent=1.0)
                sheet1.cell(row, col).font = Font(size=10, bold=True, italic=True)                
                sheet1.cell(row, col).border = Border(bottom=Side(style='thin', color='000000'))
                col += 1
                sheet1.cell(row, col, "H")
                sheet1.cell(row, col).alignment = Alignment(wrapText=True, vertical='center', indent=1.0)
                sheet1.cell(row, col).font = Font(size=10, bold=True, italic=True)                
                sheet1.cell(row, col).border = Border(bottom=Side(style='thin', color='000000'))
                col += 1
                sheet1.cell(row, col, "Cab #")
                sheet1.cell(row, col).alignment = Alignment(wrapText=True, vertical='center', indent=1.0)
                sheet1.cell(row, col).font = Font(size=10, bold=True, italic=True)                
                sheet1.cell(row, col).border = Border(bottom=Side(style='thin', color='000000'))
                row += 1
                col -= 3
                qty = 0

                for door in temp_dict[template][door_style]:
                    sheet1.cell(row, col, door[0])
                    sheet1.cell(row, col).alignment = Alignment(wrapText=True, vertical='top', horizontal='left', indent=4.0)
                    sheet1.cell(row, col).font = Font(size=11)
                    sheet1.cell(row, col).border = Border(bottom=Side(style='thin', color='D4D4D4'))
                    col += 1
                    sheet1.cell(row, col, round(door[1],1))
                    sheet1.cell(row, col).alignment = Alignment(wrapText=True, vertical='top', horizontal='left', indent=1.0)
                    sheet1.cell(row, col).font = Font(size=11)
                    sheet1.cell(row, col).border = Border(bottom=Side(style='thin', color='D4D4D4'))
                    col += 1
                    sheet1.cell(row, col, round(door[2],1))
                    sheet1.cell(row, col).alignment = Alignment(wrapText=True, vertical='top', horizontal='left', indent=1.0)
                    sheet1.cell(row, col).font = Font(size=11)
                    sheet1.cell(row, col).border = Border(bottom=Side(style='thin', color='D4D4D4'))
                    col += 1
                    sheet1.cell(row, col, 'R' + door[3] + door[4] + door[5])
                    sheet1.cell(row, col).alignment = Alignment(wrapText=True, vertical='top', horizontal='left', indent=1.0)
                    sheet1.cell(row, col).font = Font(size=11)
                    sheet1.cell(row, col).border = Border(bottom=Side(style='thin', color='D4D4D4'))
                    col -= 3
                    row += 1
                    qty += 1

                row -= 1
                sheet1.cell(row, col).border = Border(bottom=Side(style='thin', color='000000'))
                sheet1.cell(row, col + 1).border = Border(bottom=Side(style='thin', color='000000'))
                sheet1.cell(row, col + 2).border = Border(bottom=Side(style='thin', color='000000'))
                sheet1.cell(row, col + 3).border = Border(bottom=Side(style='thin', color='000000'))
                row += 1
                sheet1.cell(row, col, "(" + str(qty) + ")")
                sheet1.cell(row, col).alignment = Alignment(wrapText=True, vertical='top', horizontal='left', indent=5.0)
                sheet1.cell(row, col).font = Font(size=10, italic=True)
                page_total += qty
                row += 2

            sheet1.cell(row, col, "Material Qty: " + str(page_total))
            sheet1.cell(row, col).alignment = Alignment(wrapText=True, vertical='top', horizontal='left', indent=2.0)
            sheet1.cell(row, col).font = Font(size=11, italic=True, bold=True)
            page_break = Break(id=row)
            sheet1.row_breaks.append(page_break)
            row += 1


    sheet1.column_dimensions['A'].width = 30
    sheet1.column_dimensions['B'].width = 10
    sheet1.column_dimensions['C'].width = 20
    sheet1.column_dimensions['D'].width = 28

    sheet1.oddFooter.right.text = "Page &[Page] of &N"
    sheet1.oddFooter.right.size = 9
    sheet1.oddFooter.right.color = "000000"
    
    print_area = 'A1:D' + str(row - 1)
    sheet1.print_area = print_area
    
    save_name = job_name + ' - Door List' + '.xlsx'
    full_save_name = os.path.join(dir_path, save_name)
    try:
        wb.save(full_save_name)
    except PermissionError:
        print("\nSAVE FAILED\nYou will need to close the open file before it can be saved.")

    os.startfile(full_save_name)
    
hinge_boring_report()
