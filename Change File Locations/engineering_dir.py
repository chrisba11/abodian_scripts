import os

state_path = "C:\\Mozaik\\State.dat"
eng_dir = "E:\\Axon Cabinets\\Projects - Project Documents\\Mozaik Files\\Jobs - Engineering\n"
contents = []

with open(state_path, "rt") as f:
    contents = f.readlines()
    contents[16] = eng_dir

with open(state_path, "wt") as f:
    f.writelines(contents)
