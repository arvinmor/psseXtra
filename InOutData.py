import os,sys
psspath=r'c:\program files\pti\psse32\pssbin'
sys.path.append(psspath)
os.environ['PATH'] += ';' + psspath

import psspy
import redirect
redirect.psse2py()
psspy.psseinit(10000)
psspy.case(r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\HQ_PSSmodel\HQ1_shuntadded.sav')
psspy.fnsl(
    options1=0,
    options5=0
    )
ierr, volt = psspy.abusreal(-1, string="PU")        #voltage at buses in normal condition

ierr=psspy.load_data_3(i=303,
                       REALAR2=600
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
x1=excelpy.workbook(r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\inoutdata.xls',sheet="Feuil1",overwritesheet=True, mode='w')
x1.show()
x1.show_alerts(False)
x1.set_cell('a1','Bus No.')
x1.set_cell('b1','Vnom (p.u)')
x1.set_cell('c1','Vdist (p.u)')
x1.set_cell('d1','Vd (p.u)')
x1.set_range(2, 'a', zip(*buses))
x1.set_range(2, 'b', zip(*volt))
x1.set_range(2, 'c', zip(*voltd))
x1.set_range(2, 'd', zip(*dv))
