# SCENARIO01- THIS SCRIPT IS USED TO RUN IEEE 39 BUS DYNAMIC SIMULATION
import os,sys

# PSSE initialization
psspath=r'C:\Program Files (x86)\PTI\PSSEUniversity33\PSSBIN'   # change to the address of your PSSBIN directory on your computer
sys.path.append(psspath)
os.environ['PATH'] += ';' + psspath
import psspy
import dyntools
import redirect
import random
import excelpy

redirect.psse2py()
psspy.psseinit(10000)

# required files
raw_file=r'OPAL50.raw';
dyr_file=r'OPAL50.dyr';

# generated files
file_name=r'39b_V2';
out_file=r'PSSE_'+file_name+r'.out';
psse_xls_file=r'PSSE_'+file_name+r'.xlsx';

# Read load flow and dynamic files
psspy.read(0,raw_file)
psspy.dyre_new([1,1,1,1],dyr_file,"","","")

# Remove old xls files
if os.path.isfile(psse_xls_file)==True:
    os.remove(psse_xls_file);

# Convert Generator and Loads
psspy.cong(0)
psspy.conl(0,1,1,[1,0],[0.0, 100.0,0.0, 100.0])
psspy.conl(0,1,2,[1,0],[0.0, 100.0,0.0, 100.0])
psspy.conl(0,1,3,[1,0],[0.0, 100.0,0.0, 100.0])

# Order Network for matrix operation
psspy.ordr(0)

# Factorize Admittance maatrix
psspy.fact()

# Solve switching study network solutions
psspy.tysl()

# Save sav file after conversion

trans_file = raw_file[0:raw_file.find(".")]+"_trans.sav"
psspy.save(trans_file)


# Select Channels
psspy.delete_all_plot_channels()
psspy.chsb(0,1,[-1,-1,-1,1,1,0])
psspy.chsb(0,1,[-1,-1,-1,1,2,0])
psspy.voltage_and_angle_channel([-1,-1,-1,6],["",""])
psspy.voltage_and_angle_channel([-1,-1,-1,7],["",""])
psspy.voltage_and_angle_channel([-1,-1,-1,9],["",""])
psspy.voltage_and_angle_channel([-1,-1,-1,13],["",""])
psspy.voltage_and_angle_channel([-1,-1,-1,14],["",""])
psspy.voltage_and_angle_channel([-1,-1,-1,39],["",""])
psspy.branch_p_channel([-1,-1,-1,6,7],r"""1""","")
psspy.branch_p_channel([-1,-1,-1,13,14],r"""1""","")
psspy.branch_p_channel([-1,-1,-1,39,9],r"""1""","")
psspy.branch_mva_channel([-1,-1,-1,5,8],r"""1""","")
psspy.branch_mva_channel([-1,-1,-1,6,5],r"""1""","")
psspy.branch_mva_channel([-1,-1,-1,4,5],r"""1""","")
psspy.load_array_channel([-1,1,71],r"""1""","")
psspy.load_array_channel([-1,2,71],r"""1""","")
psspy.load_array_channel([-1,1,8],r"""1""","")
psspy.load_array_channel([-1,2,8],r"""1""","")
psspy.chsb(0,1,[-1,-1,-1,1,5,0])
psspy.chsb(0,1,[-1,-1,-1,1,27,0])
    
psspy.strt(0,out_file);

# Run simulation
DELTA=0.005;
psspy.dynamics_solution_param_2(realar3=DELTA); # set sample time, DELTA sec.


# Simulation Procedure
##1.	System initialization and run simulation to t = 1 s.
psspy.run(0, 1,0,0,0);
##2.	Apply 3-phase fault at line B6 – B7 at t = 1s.
psspy.dist_branch_fault(6, 7, '1',3,0.0,[75.625e-5,0]);
##3.	Continue simulation to t = 1.1s.
psspy.run(0, 1.1,0,0,0);
##4.	Trip line B6 – B7 and clear fault at t = 1.1s.
psspy.dist_clear_fault();
psspy.dist_branch_trip(6, 7, '1');
##5.	Continue simulation to t = 6.1s.
psspy.run(0, 6.1,0,0,0);
##6.	Reschedule generator with run-up M1 from 250 MW to 500 MW and run-down M2 from about 573 MW to 300 MW, and trip load at B7 at t = 6.1s at t = 6.1s
psspy.increment_gref(30, '1', 0.25);
psspy.increment_gref(31, '1', -0.273);
psspy.dist_branch_trip(7, 71, '1');
##7.	Continue simulation to t = 30s.
psspy.run(0, 30,0,0,0);

### Reconnect Load
##psspy.busdata2(jbus,INTGAR1=1);
##psspy.DIST_BRANCH_CLOSE(ibus, jbus, id);

# Data Retrieval
achnf = dyntools.CHNF(out_file);

# Export Data to Excel
achnf.xlsout();

##id_xls_COMPARE.close();
