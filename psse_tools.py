# -*- coding: utf-8 -*-
"""
Created on Fri Jun 12 13:04:11 2015

@author: arvin
"""
import os,sys;

pssepath=r'C:\Program Files (x86)\PTI\PSSEUniversity33\PSSBIN'   # change to the address of your PSSBIN directory on your computer
sys.path.append(pssepath);
os.environ['PATH'] += ';' + pssepath;

import caspy; 

def Write_Case_Identification_Data(f,pssmsc,pssrev,psstit):
    data='0,%(SBASE)#9.2f,%(PSSVER)#3d,%(XFRRAT)#2d,%(NXFRAT)#2d,%(BASFRQ)#6.2f' % pssmsc;
    f.write(data);
    f.write('     / PSS®E-%(PSSVER)s    '  % pssmsc);
    f.write(' '+pssrev.get('DATE'));
    for v in psstit.get('TITLE'):
        f.write('\n%s' % v);
    for v in psstit.get('HEDING'):
        f.write('\n%s' % v);
def Write_Bus_Data(f,pssbus,nbus):
    f.write('\n');
    keys=['NUM', 'NAME', 'BASKV', 'IDE', 'AREA', 'ZONE', 'OWNER', 'VM', 'VA', 'NMAXV', 'NMINV', 'EMAXV', 'EMINV']; 
    for i in range(nbus):
        for key in keys:
            v=pssbus.get(key)[i];
            if key=='NUM':
                item='%#6d' %v;
            elif key=='NAME':
                item=",'%#12s'" %v;
            elif key=='BASKV':
                item=',%#9.4f' %v;
            elif key=='IDE':
                item=',%#1d' %v;
            elif key=='AREA':
                item=',%#4d' %v;
            elif key=='ZONE':
                item=',%#4d' %v;
            elif key=='OWNER':
                item=',%#4d' %v;
            elif key=='VM':
                item=',%#7.5f' %v;
            elif key=='VA':
                item=',%#9.4f' %v;
            elif key=='NMAXV':
                item=',%#7.5f' %v;
            elif key=='NMINV':
                item=',%#7.5f' %v;
            elif key=='EMAXV':
                item=',%#7.5f' %v;
            elif key=='EMINV':
                item=',%#7.5f' %v;              
            f.write(item);
        f.write('\n');
    f.write('0 / END OF BUS DATA');
def Write_Load_Data(f,psslod,nload,sbase):
    f.write(', BEGIN LOAD DATA')
    f.write('\n');
    keys=['NUM','ID', 'STATUS', 'AREA', 'ZONE','LOAD', 'OWNER', 'LDSCALE', 'LDINT']; 
    for i in range(nload):
        for key in keys:
            v=psslod.get(key)[i];    
            if key=='NUM':
                 item='%#6d' %v;               
            elif key=='ID':
                item=",'%#2s'" %v;
            elif key=='STATUS':
                 item=',%#1d' %v; 
            elif key=='AREA':
                item=',%#4d' %v;
            elif key=='ZONE':
                item=',%#4d' %v;
            elif key=='LOAD':
                item=',%#10.3f,%#10.3f,%#10.3f,%#10.3f,%#10.3f,%#10.3f' %(v[0]*sbase,v[1]*sbase,v[2],v[3],v[4],v[5]);
            elif key=='OWNER':
                item=',%#4d' %v;
            elif key=='LDSCALE':
                item=',%#1d' %v;
            elif key=='LDINT':
                item=',%#1d' %v;
            f.write(item);
        f.write('\n');
    f.write('0 / END OF LOAD DATA');
def Write_Fixed_Bus_Shunt_Data(f,pssfsh,nbushn,sbase):
    f.write(', BEGIN FIXED SHUNT DATA')
    f.write('\n');
    keys=['NUM', 'ID', 'STATUS', 'SHUNT'];
    for i in range(nbushn):
        for key in keys:
            v=pssfsh.get(key)[i];    
            if key=='NUM':
                 item='%#6d' %v;               
            elif key=='ID':
                item=",'%#2s'" %v;
            elif key=='STATUS':
                 item=',%#1d' %v; 
            elif key=='SHUNT':
                item=',%#10.3f,%#10.3f' % (v.real*sbase,v.imag*sbase);
            f.write(item);
        f.write('\n');
    f.write('0 / END OF FIXED SHUNT DATA');
def Write_Generator_Data(f,pssgen,ngen,sbase):
    f.write(',  BEGIN GENERATOR DATA')
    f.write('\n');
    keys=['NUM','IDE','PG','QG','QT','QB','VS','IREG','MBASE','ZSORCE','XTRAN','GTAP','STAT','RMPCT','PT','PB','OWNER','OWNPCT'];
#    ,'WMOD','WPF']; 
    for i in range(ngen):
        for key in keys:
            v=pssgen.get(key)[i];    
            if key=='NUM':
                 item='%#6d' %v;               
            elif key=='IDE':
                item=",'%#2s'" %v;    
            elif key=='PG':
                item=",%#10.3f" % (v*sbase); 
            elif key=='QG':
                item=",%#10.3f" % (v*sbase); 
            elif key=='QT':
                item=",%#10.3f" % (v*sbase); 
            elif key=='QB':
                item=",%#10.3f" % (v*sbase); 
            elif key=='VS':
                item=",%#7.5f" %v; 
            elif key=='IREG':
                item=",%#6s" %v; 
            elif key=='MBASE':
                item=",%#10.3f" %v; 
            elif key=='ZSORCE':
                item=",%#11.5E,%#11.5E" %(v.real,v.imag); 
            elif key=='XTRAN':
                item=",%#11.5E,%#11.5E" %(v.real,v.imag); 
            elif key=='GTAP':
                item=",%#7.5f" %v; 
            elif key=='STAT':
                item=",%#1s" %v; 
            elif key=='RMPCT':
                item=",%#7.1f" %v; 
            elif key=='PT':
                item=",%#10.3f" % (v*sbase); 
            elif key=='PB':
                item=",%#10.3f" % (v*sbase); 
            elif key=='OWNER':
                owner=v;
                item='';
            elif key=='OWNPCT':
                j=0;
                item='';
                while j<=3:
                    if owner[j] <> 0:
                        item=item+",%#4s,%#6.4f" %(owner[j],v[j]);
                    j+=1;
#            elif key=='WMOD':
#                item=",'%#2s'" %v; 
#            elif key=='WPF':
#                item=",'%#2s'" %v; 
            f.write(item);
        f.write('\n');
    f.write('0 / END OF GENERATOR DATA');
def Write_Non_Transformer_Branch_Data(f,pssbrn,nlin,ntrfmr):
    offset=nlin-ntrfmr;
    f.write(',  BEGIN BRANCH DATA')
    f.write('\n');
    keys=['FRMBUS','TOBUS','CKT','RX','B','RATEA','RATEB','RATEC','GBI','GBJ','STAT','METBUS','LINLEN','OWNER','OWNPCT'];
    for i in range(offset):
        for key in keys:
            v=pssbrn.get(key)[i];  
            if key=='FRMBUS':
                item='%#6d' %v;               
            if key=='TOBUS':
                item=',%#6d' %v;
            elif key=='CKT':
                item=",'%#2s'" %v; 
            elif key=='RX':
                item=",%#11.5E,%#11.5E" %(v.real,v.imag); 
            elif key=='B':
                item=",%#10.5f" %v; 
            elif key=='RATEA':
                item=",%#8.2f" %v; 
            elif key=='RATEB':
                item=",%#8.2f" %v; 
            elif key=='RATEC':
                item=",%#8.2f" %v; 
            elif key=='GBI':
                item=",%#10.5f,%#10.5f" %(v.real,v.imag); 
            elif key=='GBJ':
                item=",%#10.5f,%#10.5f" %(v.real,v.imag); 
            elif key=='STAT':
                item=",%#1s" %v; 
            elif key=='METBUS':
                if v==pssbrn.get('FRMBUS')[i]:
                    item=",1";
                else:
                    item=',2';
            elif key=='LINLEN':
                item=",%#7.2f" %v; 
            elif key=='OWNER':
                owner=v;
                item='';
            elif key=='OWNPCT':
                j=0;
                item='';
                while j<=3:
                    if owner[j] <> 0:
                        item=item+",%#4s,%#6.4f" %(owner[j],v[j]);
                    j+=1;
            f.write(item);
        f.write('\n');
    f.write('0 / END OF BRANCH DATA');    
def Write_Transformer_Data(f,psstrn,pssbrn,ntrfmr,nlin):
    offset=nlin-ntrfmr;
    f.write(',  BEGIN TRANSFORMER DATA')
    f.write('\n');
    for i in range(ntrfmr):
        keys=['FRMBUS','TOBUS','K','CKT','CW','CZ','CM','MAG1','MAG2','METBUS','TRNAME','STAT','OWNER','OWNPCT','VECGRP'];
        for key in keys:
            if key=='FRMBUS':
                v=pssbrn.get(key)[offset+i];                    
                item='%#6d' %v;               
            if key=='TOBUS':
                v=pssbrn.get(key)[offset+i];
                item=',%#6d' %v;
            if key=='K':
                v=0;
                item=',%#6d' %v;
            elif key=='CKT':
                v=pssbrn.get(key)[offset+i];
                item=",'%#2s'" %v; 
            elif key=='CW':
                v=1;
                item=',%#1d' %v; 
            elif key=='CZ':
                v=1;
                item=',%#1d' %v;
            elif key=='CM':
                v=1;
                item=',%#1d' %v;
            elif key=='MAG1':
                v=0;
                item=',%#11.5E' %v;
            elif key=='MAG2':
                v=0;
                item=',%#11.5E' %v; 
            elif key=='METBUS':
                v=pssbrn.get(key)[offset+i];
                if v==pssbrn.get('FRMBUS')[offset+i]:
                    item=",2";
                else:
                    item=',1';
            elif key=='TRNAME':
                v=psstrn.get(key)[i];
                item=",'%12s'" %v; 
            elif key=='STAT':
                v=pssbrn.get(key)[offset+i];
                item=",%#1s" %v; 
            elif key=='OWNER':
                owner=pssbrn.get(key)[offset+i];
                item='';
            elif key=='OWNPCT':
                v=pssbrn.get(key)[offset+i];
                j=0;
                item='';
                while j<=3:
                    item=item+",%#4s,%#6.4f" %(owner[j],v[j]);
                    j+=1;
            elif key=='VECGRP':
                v=psstrn.get('VECGRP')[i];
                item=",'%12s'" %v; 
            f.write(item);
        f.write('\n');
        keys=['RXTRAN','SBASE1'];
        for key in keys:
            v=psstrn.get(key)[i]; 
            if key=='RXTRAN':                   
                item='%#11.5E,%#11.5E' %(v.real,v.imag);               
            if key=='SBASE1':                   
                item=',%#9.2f' %v;
            f.write(item);
        f.write('\n');
        keys=['WIND1','NOMV1','ANG1','RATEA','RATEB','RATEC','CNTL','CONBUS','RMAX','RMIN','VMAX','VMIN','NTAPS','TABLE','CR1','CX1','CNXA1'];
        for key in keys:
            if key=='WIND1':                   
                v=psstrn.get(key)[i]; 
                item='%#7.5f' %v;               
            if key=='NOMV1':                   
                v=psstrn.get(key)[i]; 
                item=',%#8.3f' %v;
            if key=='ANG1':                   
                v=psstrn.get(key)[i]; 
                item=',%#8.3f' %v;
            if key=='RATEA':                   
                v=pssbrn.get(key)[offset+i];
                item=',%#8.2f' %v;
            if key=='RATEB':                   
                v=pssbrn.get(key)[offset+i];
                item=',%#8.2f' %v;
            if key=='RATEC':                   
                v=pssbrn.get(key)[offset+i];
                item=',%#8.2f' %v;
            if key=='CNTL':                   
                v=psstrn.get(key)[i]; 
                item=',%#2d' %v;
            if key=='CONBUS':                   
                v=psstrn.get(key)[i]; 
                item=',%#7d' %v;
            if key=='RMAX':                   
                v=psstrn.get(key)[i]; 
                item=',%#8.5f' %v;
            if key=='RMIN':                   
                v=psstrn.get(key)[i]; 
                item=',%#8.5f' %v;
            if key=='VMAX':                   
                v=psstrn.get(key)[i]; 
                item=',%#8.5f' %v;
            if key=='VMIN':                   
                v=psstrn.get(key)[i]; 
                item=',%#8.5f' %v;
            if key=='NTAPS':                   
                v=psstrn.get(key)[i]; 
                item=',%#4d' %v;
            if key=='TABLE':
                v=psstrn.get(key)[i]; 
                item=',%#2d' %v;                
            if key=='CR1':
                v=0;                  
                item=',%#8.5f' %v;
            if key=='CX1':                   
                v=0;                  
                item=',%#8.5f' %v;
            if key=='CNXA1':                   
                v=0;                  
                item=',%#7.3f' %v;
            f.write(item);
        f.write('\n');
        keys=['WIND2','NOMV2'];
        for key in keys:
            v=psstrn.get(key)[i]; 
            if key=='WIND2':                   
                item='%#7.5f' %v;               
            if key=='NOMV2':                   
                item=',%#8.3f' %v;
            f.write(item);
        f.write('\n');
    f.write('0 / END OF TRANSFORMER DATA');        
def sav2raw(*args):
    if len(args)==1:
        savfile=args[0];
        rawfile=savfile[0:savfile.find(".")]+".raw";
    else:
        savfile=args[0];
        rawfile=args[1];       
    net = caspy.Savecase(savfile);
    f=open(rawfile,'w');
    Write_Case_Identification_Data(f,net.pssmsc,net.pssrev,net.psstit);
    Write_Bus_Data(f,net.pssbus,net.psssiz.get('NBUS'));
    Write_Load_Data(f,net.psslod,net.psssiz.get('NLOAD'),net.pssmsc.get('SBASE'));
    Write_Fixed_Bus_Shunt_Data(f,net.pssfsh,net.psssiz.get('NBUSHN'),net.pssmsc.get('SBASE'));
    Write_Generator_Data(f,net.pssgen,net.psssiz.get('NGEN'),net.pssmsc.get('SBASE'));
    Write_Non_Transformer_Branch_Data(f,net.pssbrn,net.psssiz.get('NLIN'),net.psssiz.get('NTRFMR'));
    Write_Transformer_Data(f,net.psstrn,net.pssbrn,net.psssiz.get('NTRFMR'),net.psssiz.get('NLIN'));
    f.close;  
    return(net.psssiz);