### Initialization
import os,sys
import csv
psspath=r'c:\program files\pti\psse32\pssbin'
sys.path.append(psspath)
os.environ['PATH'] += ';' + psspath
import psspy
import redirect
import random
redirect.psse2py()
psspy.psseinit(10000)
import excelpy

### Functions to be used by the main routine

##  Function to convert tuple elements to list elements
def listit(t):
    return list(map(listit, t)) if isinstance(t, (list, tuple)) else t

##  Function to import data for pilots and capacitors from Excel to Python
def exldataimport(capdataf,pilotdataf):
    capf=open(capdataf,'rb')
    capdata=csv.reader(capf,delimiter=';')
    capdata=[[eval(row[0]),eval(row[5]),eval(row[6])] for row in capdata]
    capf.close()
    lcdata=len(capdata)
    capbus=[0]*lcdata
    capstep=[0]*lcdata
    capQ=[0]*lcdata
    for i in range(lcdata):
        capbus[i]=capdata[i][0]
        capstep[i]=capdata[i][1]
        capQ[i]=capdata[i][2]
    capstep=listit(capstep)
    capQ=listit(capQ)
    pilotf=open(pilotdataf,'rb')
    pilotdata=csv.reader(pilotf,delimiter=';')
    pilotdata=[eval(row[0]) for row in pilotdata]
    pilotf.close()
    return capbus,capstep,capQ,pilotdata;

##  Function for applying random distrubance on all loads of the network with a specific percentage
def randdist(argin):
    filename=argin[0]
    percentage=argin[1]
    mode=argin[2]
    shuntctrlmode=0
    if mode!='arbitrary':
        shuntctrlmode=argin[3]
    psspy.case(filename)
    ierr, Qload=psspy.aloadcplx(-1, string="MVAACT")    
    Qload=[x.imag for x in Qload[0]]   
    ierr, tlbuses = psspy.aloadint(string='NUMBER')
    nbus=len(tlbuses[0])
    if mode=='random':
        zoseq=[0,1]*nbus
        sb=random.sample(zoseq,nbus)
    elif mode=='arbitrary':
        abus=argin[3]      # argin[3] in this case is the arbitrary buses to apply the disturbance
        sb=[0]*nbus
        for i in range(len(abus)):
            sb[tlbuses[0].index(abus[i])]=1            
    else:
        sb=[1]*nbus
    for j in range(nbus):
        if sb[j]==1:            
            Qd=Qload[j]+percentage/100.0*abs(Qload[j])
            ierr=psspy.load_data_3(i=tlbuses[0][j],REALAR2=Qd)
    if shuntctrlmode==1:
        for i in tlbuses[0]:
            ierr = psspy.switched_shunt_data_3(i, intgar9=1)
    psspy.fnsl(options1=0,options5=0)
    ierr, voltd = psspy.abusreal(-1, string="PU")        #voltage at buses after disturbance
    argout=[]
    argout.append(voltd)
    argout.append(Qload)
    argout.append(tlbuses)
    return argout;

##  Function for random switching of capacitor banks
def capselect(capbus,step,Q):
    nc=len(capbus)
    zoseq=[0,1]*nc
    cs=random.sample(zoseq,nc)
    scb=[]     # Selected Capacitor Bus
    ss=[]       # Selected Step
    qss=[]      # Q(Mvar) of selected step
    for i in range(nc):
        if cs[i]==1:
            if isinstance(step[i],int):
                nscs=0
                s=random.randint(0,step[i])
                css=0
                q=Q[i]
            else:
                nscs=len(step[i])
                css=random.randint(0,nscs-1)
                s=random.randint(0,step[i][css])
                q=Q[i][css]
            scb.append(capbus[i])
            ss.append(s)
            qss.append(q)
    return [scb,ss,qss];

### Begining of the main routine

capfile=r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\shuntcapdata.csv'
pilotfile=r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\pilotdata.csv'
[capbus,capstep,capQ,pilot]=exldataimport(capfile,pilotfile)
##print 'capacitor buses=',capbus
##print 'steps for capacitors=',capstep
##print 'MVar per each step=',capQ
##print 'Pilot nodes=',pilot

filename=r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\HQ_PSSmodel\HQ1_shuntadded.sav'
start=0;
stop=40;
stp=5;
#distpercent=range(start,stop,stp)
distpercent=[0,10] # 27.511 % of distrubance seems to be the maximum possible disturbance on the whole loads
ndist=len(distpercent)
vd=[]
for i in range(ndist):
    argout=randdist([filename,distpercent[i],'all',1]) # the last argin is for switching shunts status
    voltd=argout[0]
    Qload=argout[1]
    tlbuses=argout[2]
    vd.append(voltd)

##for i in range(ndist):
##    argout=randdist([filename,distpercent[i],'all',1])
##    voltd=argout[0]
##    Qload=argout[1]
##    tlbuses=argout[2]
##    vd.append(voltd)

x1=excelpy.workbook(r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\inoutdata.xls',sheet="Feuil1",overwritesheet=True, mode='w')
x1.show()
x1.show_alerts(False)
ierr, baseV = psspy.abusreal(-1, string='BASE')
ierr, buses = psspy.abusint(-1, string="NUMBER")
x1.set_cell('a1','Bus No.')
x1.set_range(2, 'a', zip(*buses))
x1.set_cell('b1','Base Voltage(kv)')
x1.set_range(2, 'b', zip(*baseV))

##distpercent=distpercent*2
##ndist=ndist*2
for i in range(ndist):
    x1.set_cell((1,i+3),'dist '+ str(distpercent[i])+' %')
    x1.set_range(2, i+3, zip(*vd[i]))

# the following code shows the effect of capacitor swithching on the disturbed system
C=308
capidx=tlbuses[0].index(C)
  
Qcapidx= Qload[capidx]   
Qd=Qcapidx-2*380
ierr=psspy.load_data_3(i=C,REALAR2=Qd)

psspy.fnsl(options1=0,options5=0)
ierr, voltd = psspy.abusreal(-1, string="PU")        #voltage at buses in normal condition
x1.set_cell((1,ndist+3),'cap308')
x1.set_range(2, ndist+3, zip(*voltd)) 
##[scb,ss,qss]=capselect(capbus,capstep,capQ)
##print 'Selected Capacitor Bus', scb
##print 'Selected Step', ss
##print 'Q(Mvar) of selected step',qss
