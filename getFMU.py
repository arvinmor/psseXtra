import os
from shutil import copy
from shutil import rmtree
fname = 'fmulist.txt'   # define the list in the text file
fmuFolder = 'FMU'   # the name of the folder which have all the fmus
fmuSelectedFolder = 'FMUselected'   # name of the destination folder which you would like to copy your selected FMUs
if os.path.exists(fmuSelectedFolder):
    rmtree(fmuSelectedFolder)
os.mkdir(fmuSelectedFolder);
with open(fname) as f:
    fmulist = f.readlines()
# you may also want to remove whitespace characters like `\n` at the end of each line
fmulist = [x.strip() for x in fmulist]
src_files = os.listdir(fmuFolder)
for fmu in fmulist:
    fmu = fmu + '.fmu';
    if fmu in src_files:
        copy(fmuFolder+'/'+fmu, fmuSelectedFolder+'/')