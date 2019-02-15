from test4 import randdist,exldataimport
capfile=r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\shuntcapdata.csv'
pilotfile=r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\IREQ\PythonProgs\pilotdata.csv'
[capbus,capstep,capQ,pilot]=exldataimport(capfile,pilotfile)

filename=r'C:\Documents and Settings\xw0419\Mes documents\Mon Projet\Simulation\HQ_PSSmodel\HQ1_shuntadded.sav'
nswitch=0;
ndist=1;
distpercent=20;
[volt, voltd, vpn, vpnd, dvpn, switch]=randdist([filename,distpercent,'random',0,capbus,capstep,capQ,pilot,ndist,nswitch])
