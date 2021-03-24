import os

state_path = "C:\\Mozaik\\State.dat"
sales_dir = "E:\\Axon Cabinets\\Projects - Project Documents\\Mozaik Files\\Jobs - Burnaby Sales\n"
contents = []

with open(state_path, "rt") as f:
    contents = f.readlines()
    contents[16] = sales_dir

with open(state_path, "wt") as f:
    f.writelines(contents)
