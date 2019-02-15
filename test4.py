### Initialization
import os,sys
import csv
psspath=r'C:\Program Files (x86)\PTI\PSSEUniversity32\PSSBIN'
sys.path.append(psspath)
os.environ['PATH'] += ';' + psspath
import psspy
import redirect
import random
redirect.psse2py()
psspy.psseinit(10000)
import excelpy

################################################################################################################
################################################################################################################
################################################################################################################
################################################################################################################
## This Section controls the program
## flags turn on/off parts of the code

flag1=0; ## STEP1- Python: Run TEST4.py program to generate data from PSSE
flag4=0; ## STEP4- Python: Apply Random Disturbance on the network
flag7=1; ## STEP7- Apply the Control to the disturbed model in Python

### Begining of the main routine

if flag1==1:
    capfile=r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\shuntcapdata1.csv'
    pilotfile=r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\pilotdata.csv'
    [capbus,capstep,capQ,pilot]=exldataimport(capfile,pilotfile)
    ##print 'capacitor buses=',capbus
    ##print 'steps for capacitors=',capstep
    ##print 'MVar per each step=',capQ
    ##print 'Pilot nodes=',pilot

    npn=len(pilot)
    ncap=len(capbus)
    filename=r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\HQ_PSSmodel\HQ1_shuntadded.sav'

    ### ### ### ### ### ### 
    distpercent=20 #% of distrubance seems to be the maximum possible disturbance on the whole loads
    ### ### ### ### ### ### 

    ndist=10
    nswitch=20

    [volt, voltd, vpn, vpnd, dvpn, switch]=randdist([filename,distpercent,'random',0,capbus,capstep,capQ,pilot,ndist,nswitch])

    x1=excelpy.workbook(r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\test4data.xls',sheet="Feuil1",overwritesheet=True, mode='w')
    x1.show()
    x1.show_alerts(False)

    r0=2
    c0=1

    ## Initialzation data for MATLAB
    x1.set_cell((1, 1),npn)
    x1.set_cell((1, 2),ncap)
    x1.set_cell((1, 3),ndist)
    x1.set_cell((1, 4),nswitch)
    x1.set_cell((1, 5),r0)
    x1.set_cell((1, 6),c0)

    x1.set_range(r0+1, c0, zip(*[['vpnd']*npn]))
    x1.set_range(r0+npn+1, c0, zip(*[['vpndc']*npn]))
    x1.set_range(r0+2*npn+1, c0, zip(*[['dvpn']*npn]))
    x1.set_range(r0+3*npn+1, c0, zip(*[['Cap Bus']*ncap]))

    x1.set_range(r0+1, c0+1, zip(*[pilot]))
    x1.set_range(r0+npn+1, c0+1, zip(*[pilot]))
    x1.set_range(r0+2*npn+1, c0+1, zip(*[pilot]))
    x1.set_range(r0+3*npn+1, c0+1, zip(*[capbus]))

    for i in range(ndist*nswitch):          # an error appears when ndist*nswitch exceeds 256=maximum No. of columns in excel 2003
        x1.set_cell((r0, i+c0+2),'case '+ str(i))
        x1.set_range(r0+1, i+c0+2, zip(*[vpn[i]]))
        x1.set_range(r0+npn+1, i+c0+2, zip(*[vpnd[i]]))
        x1.set_range(r0+2*npn+1, i+c0+2, zip(*[dvpn[i]]))
        x1.set_range(r0+3*npn+1, i+c0+2, zip(*[switch[i]]))
    x1.save()
    ##x2=excelpy.workbook(r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\test4voltage.xls',sheet="Feuil1",overwritesheet=True, mode='w')
    ##x2.show()
    ##x2.show_alerts(False)
    ##
    ##ierr, baseV = psspy.abusreal(-1, string='BASE')
    ##ierr, buses = psspy.abusint(-1, string="NUMBER")
    ##
    ##nbus= len(buses[0])
    ##
    ##x2.set_cell('a1','Bus No.')
    ##x2.set_range(2, 'a', zip(*buses))
    ##x2.set_cell('b1','Base Voltage(kv)')
    ##x2.set_range(2, 'b', zip(*baseV))
    ##
    ##x2.set_range(3+nbus,'a', zip(*buses))
    ##x2.set_range(3+nbus,'b', zip(*baseV))
    ##
    ##for i in range(ndist*nswitch):
    ##    x2.set_cell((1,i+3),'case '+ str(i))
    ##    x2.set_range(2, i+3, zip(*volt[i]))
    ##    x2.set_range(3+nbus, i+3, zip(*voltd[i]))

##############################################################
##############################################################
##############################################################
## STEP4- Apply Random Disturbance on the network and save it in test4dist.xls


if flag4==1:
    nswitch=1;
    ndist=1;
    distpercent=0;
    [volt, voltd, vpn, vpnd, dvpn, switch]=randdist([filename,distpercent,'random',0,capbus,capstep,capQ,pilot,ndist,nswitch])

    distpercent=40;
    [volt1, voltd1, vpn1, vpnd1, dvpn1, switch1]=randdist([filename,distpercent,'random',0,capbus,capstep,capQ,pilot,ndist,nswitch])

    x2=excelpy.workbook(r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\test4dist.xls',sheet="Feuil1",overwritesheet=True, mode='w')
    x2.show()
    x2.show_alerts(False)

    r0=1
    c0=1
    x2.set_range(r0, c0, zip(*vpn))
    x2.set_range(r0, c0+1, zip(*vpn1))
    x2.save()

##############################################################
##############################################################
##############################################################
## STEP7- Apply the Control to the disturbed model in Python



if flag7==1:
    ierr, Qload=psspy.aloadcplx(-1, string="MVAACT")    # Qload is the amount of apparent power (P+jQ) for all loads
    Qload=[x.imag for x in Qload[0]]                    # We just need Reactive power (Q) so we find imaginary part of apparent power
    ierr, tlbuses = psspy.aloadint(string='NUMBER')     # tlbuses is all loads' bus nomber, it includes also the compensators that we have added as loads to the network

    ctrl=[0,1,-1,0,0,0,0,2,-1,1,0,2,0,1,0,1]
    Qctrl=[0 for x in range(ncap)]
    for k in range(len(ctrl)):        
        Qctrl[k]=Qload[tlbuses[0].index(capbus[k])]-ctrl[k]*capQ[k]
        ierr=psspy.load_data_3(i=capbus[k],REALAR2=Qctrl[k])
    psspy.fdns(OPTIONS1=0,OPTIONS5=0,OPTIONS6=1)               # do power flow on the disturbed system
    ierr, voltdc = psspy.abusreal(-1, string="PU")       # and measure voltage at all buses after disturbance (there is no option in PSS to measure the voltage of just one bus)
    
    x2=excelpy.workbook(r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\final_result.xls',sheet="Feuil1",overwritesheet=True, mode='w')
    x2.show()
    x2.show_alerts(False)

    x2.set_range(2, 1, zip(*volt[0]))
    x2.set_range(2, 2, zip(*volt1[0]))
    x2.set_range(2, 3, zip(*[voltdc[0]]))


################################################################################################################
################################################################################################################
################################################################################################################
################################################################################################################
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
    capbus=[0 for i in range(lcdata)]
    capstep=[0 for i in range(lcdata)]
    capQ=[0 for i in range(lcdata)]
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

    filename=argin[0]                                   # initializing input arguments
    maxdist=argin[1]
    mode=argin[2]
    shuntctrlmode=argin[3]
    capbus=argin[4]
    capstep=argin[5]
    capQ=argin[6]
    pilot=argin[7]
    ndist=argin[8]
    nswitch=argin[9]
    reloadfile=1;                                       # By default the network data file will be reloaded at each iteration
    if len(argin)==11:
        reloadfile=argin[10]
    psspy.case(filename)
    ierr, Qload=psspy.aloadcplx(-1, string="MVAACT")    # Qload is the amount of apparent power (P+jQ) for all loads
    Qload=[x.imag for x in Qload[0]]                    # We just need Reactive power (Q) so we find imaginary part of apparent power
    ierr, tlbuses = psspy.aloadint(string='NUMBER')     # tlbuses is all loads' bus nomber, it includes also the compensators that we have added as loads to the network
    nbus=len(tlbuses[0])                                # nbus is No. of all load buses
    ncap=len(capbus)
    ierr, busn = psspy.abusint(-1, string='NUMBER')

    npn=len(pilot)

    volt=[]
    voltd=[]
    vpn=[[0 for x in range(npn)] for y in range(ndist*nswitch)]
    vpnd=[[0 for x in range(npn)] for y in range(ndist*nswitch)]
    dvpn=[[0 for x in range(npn)] for y in range(ndist*nswitch)]
    switch=[[0 for x in range(ncap)] for y in range(ndist*nswitch)]
    for i in range(ndist):                              # in this loop we generate ndist random distrurbane cases and apply nswitch control actions on each case
        if reloadfile==1:
            psspy.case(filename)
        percentage=random.random()* maxdist             # choose randomly the amount of disturbance with maximum of maxdist
        zoseq=[0,1]*nbus
        if mode=='random': sb=random.sample(zoseq,nbus)  # choose randomly loads to apply the disturbance
        if mode=='all': sb=[1]*nbus
        capQd=[0 for x in range(ncap)]
        for j in range(nbus):                           # applying the disturbance
            if sb[j]==1:
##                if not(tlbuses[0][j] in capbus):          # we make sure that no dist. applied on capacitor buses (which are considered also as loads)            
                Qd=Qload[j]+percentage/100.0*abs(Qload[j])
                ierr=psspy.load_data_3(i=tlbuses[0][j],REALAR2=Qd)
            if tlbuses[0][j] in capbus:
                capidx=capbus.index(tlbuses[0][j])
                if sb[j]==1:
                    capQd[capidx]=Qd
                else:
                    capQd[capidx]=Qload[j]
        if shuntctrlmode==1:                            # by this option we can unlock all compensators over the network
            for j in tlbuses[0]:
                ierr = psspy.switched_shunt_data_3(j, intgar9=1)

        print '###### DIST('+str(i)+ ') ########'
        psspy.fdns(OPTIONS1=0,OPTIONS5=0,OPTIONS6=1)               # do power flow on the disturbed system
        ierr, v = psspy.abusreal(-1, string="PU")       # and measure voltage at all buses after disturbance (there is no option in PSS to measure the voltage of just one bus)

        for j in range(nswitch):                        # now we apply random cap. switchings nswitch times
            [scb,ss,qss]=capselect(capbus,capstep,capQ)

            print '### ADD CAP ###'
            for k in range(len(scb)):
                scbidx=capbus.index(scb[k])
                switch[i*nswitch+j][scbidx]=ss[k]
                
                capQd[scbidx]=capQd[scbidx]-ss[k]*qss[k]
                ierr=psspy.load_data_3(i=scb[k],REALAR2=capQd[scbidx])

            print '###### DIST('+str(i)+') , '+'SWITCHCASE('+ str(i*nswitch+j) + ') ########'
            psspy.fdns(OPTIONS1=0,OPTIONS5=0,OPTIONS6=1)               # do power flow on the disturbed system
            ierr, vd = psspy.abusreal(-1, string="PU")       # and measure voltage at all buses after disturbance (there is no option in PSS to measure the voltage of just one bus)
            voltd.append(vd)
            volt.append(v)
            for k in range(npn):                        # measuring vpn, and vpnd as outputs
                pnidx=busn[0].index(pilot[k])
                vpn[i*nswitch+j][k]=v[0][pnidx]
                vpnd[i*nswitch+j][k]=vd[0][pnidx]
                dvpn[i*nswitch+j][k]=vpnd[i*nswitch+j][k]-vpn[i*nswitch+j][k]
            print '### REMOVE CAP ###'
            for k in range(len(scb)):                            # after cap switchings we remove their effect for next switching
                scbidx=capbus.index(scb[k])
                capQd[scbidx]=capQd[scbidx]+ss[k]*qss[k]
                ierr=psspy.load_data_3(i=scb[k],REALAR2=capQd[scbidx])

    return volt,voltd,vpn,vpnd,dvpn,switch;

##  Function for random switching of capacitor banks
def capselect(capbus,step,Q):
    nc=len(capbus)
    zoseq=[0,1]*nc
    cs=random.sample(zoseq,nc)
    scb=[]     # Selected Capacitor Buses
    ss=[]       # Selected Steps
    qss=[]      # Q(Mvar) of selected steps
    for i in range(nc):
        if cs[i]==1:
            if isinstance(step[i],int):
                nscs=0
                if step[i]>=0: s=random.randint(0,step[i])
                else: s=random.randint(step[i],0)
                css=0
                q=Q[i]
            else:
                nscs=len(step[i])
                css=random.randint(0,nscs-1)
                if step[i][css]>=0: s=random.randint(0,step[i][css])
                else: s=random.randint(step[i][css],0)
                q=Q[i][css]
            scb.append(capbus[i])
            ss.append(s)
            qss.append(q)
    return [scb,ss,qss];
