import os,sys
import csv
psspath=r'C:\Program Files (x86)\PTI\PSSEUniversity33\PSSBIN'
sys.path.append(psspath)
os.environ['PATH'] += ';' + psspath
import psspy
import redirect
import random
redirect.psse2py()
psspy.psseinit(10000)
import excelpy
