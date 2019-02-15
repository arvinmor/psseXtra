## Convert SAV file to RAW file
import psse_tools;
savfile=r'Model_2B-Haveen_12092011.sav';
rawfile=r'Model_2B-Haveen_12092011.raw';
n=psse_tools.sav2raw(savfile,rawfile);