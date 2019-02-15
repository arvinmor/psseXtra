# -*- coding: utf-8 -*-
"""
Created on Fri Jun 12 13:04:11 2015

@author: arvin
"""
class psse_tools:
    def __init__(self,pssepath):
        import os,sys;
        self.pssepath=pssepath;
        sys.path.append(pssepath);
        os.environ['PATH'] += ';' + pssepath;
    def sav2raw(*args,self.pssepath):
        if len(args)==1:
            savfile=args(1);
            rawfile=args(1);
        else:
            savfile=args(1);
            rawfile=args(2);
        import caspy   
        sample = caspy.Savecase(savfile);
        pss2dc=sample.pss2dc #returns Two-Terminal dc Transmission Line Data.
        pss3ix=sample.pss3ix # builds Indices to 3-Winding Transformer Data.
        pss3wt=sample.pss3wt # returns 3-Winding Transformer Data.
        #vpssabx=sample.vpssabx # returns Nominal Values of Adjusted Branch Data.
        pssain=sample.pssain # returns Area Interchange Data Data.
        pssald=sample.pssald # returns Nominal Values of Adjusted Load Data.
        pssatr=sample.pssatr # returns Area Transaction Data.
        pssbrn=sample.pssbrn # returns Branch Data.
        pssbus=sample.pssbus # returns Bus Data.
        psscls=sample.psscls # returns Closes the PSS(R)E saved case opened by PSSOPN.
        pssfct=sample.pssfct # returns FACTS Device Data.
        pssfsh=sample.pssfsh # returns Fixed Shunt Data.
        pssgen=sample.pssgen # returns Generator Data.
        psslod=sample.psslod # returns Load Data.
        pssmdc=sample.pssmdc # returns Multi-Terminal dc Transmission Line Data.
        pssmsc=sample.pssmsc # returns Miscellaneous Data, SBASE, Freq. , PSSE VER.,...
        pssmsl=sample.pssmsl # returns Multisection Line Data.
        pssopn=sample.pssopn # returns Opens a PSS(R)E saved case for use by the other data extraction routines.
        pssown=sample.pssown # returns Owner Names.
        pssrev=sample.pssrev # returns release identification information from Saved Case Data Extraction library.
        psssiz=sample.psssiz # returns Get Saved Case Sizes
        psstic=sample.psstic # returns Transformer Impedance Correction Tables Data.
        psstit=sample.psstit # returns headings and title data.
        psstrn=sample.psstrn # returns 2-Winding Transformer Data.
        pssvsc=sample.pssvsc # returns Voltage Source Converter dc Line Data.
        psswsh=sample.psswsh # returns Switched Shunt Data.
        pssznm=sample.pssznm # returns Zone Names.
        pssind=sample.pssind # returns Induction Machine Data.
    return sample;  