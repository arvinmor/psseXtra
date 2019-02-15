import os,sys
psspath=r'c:\program files\pti\psse32\pssbin'
sys.path.append(psspath)
os.environ['PATH'] += ';' + psspath

import psspy
import redirect
redirect.psse2py()
psspy.psseinit(10000)
psspy.case(r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\HQ_PSSmodel\nseieee_jan_23_13_0714.sav')
psspy.fnsl(
    options1=0,
    options5=0
    )
ierr, volt = psspy.abusreal(-1, string="PU")
ierr=psspy.load_data_3(i=1020,
                       intgar1=0
                       )
ierr, buses = psspy.abusint(-1, string="NUMBER")
psspy.fnsl(
    options1=0,
    options5=0
    )
ierr, voltd = psspy.abusreal(-1, string="PU")
n=len(volt[0][:])
dv=[[0]*n]
for i in range(0,n-1):
    dv[0][i]=voltd[0][i]-volt[0][i]

import excelpy
x1=excelpy.workbook(r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\outdata1.xls',sheet="Feuil1",overwritesheet=False, mode='w')
x1.show()
x1.show_alerts(False)
x1=excelpy.workbook(r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\outdata1.xls',sheet="Feuil1",overwritesheet=False, mode='w')
x1.set_cell('e1','Bus No.')
x1.set_cell('f1','Vnom (p.u)')
x1.set_cell('g1','Vdist (p.u)')
x1.set_cell('h1','Vd (p.u)')
x1.set_range(2, 'e', zip(*buses))
x1.set_range(2, 'f', zip(*volt))
x1.set_range(2, 'g', zip(*voltd))
x1.set_range(2, 'h', zip(*dv))
