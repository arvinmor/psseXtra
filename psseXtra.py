import os, sys
def configure_paths_for_psse():
    # PSSE initialization
    psspath=r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSBIN'   # change to the address of your PSSBIN directory on your computer
    sys.path.append(psspath)
    os.environ['PATH'] +=';'+psspath
    psspath=r'C:\Program Files (x86)\PTI\PSSEXplore34\PSSPY27'   # change to the address of your PSSBIN directory on your computer
    sys.path.append(psspath)
    os.environ['PATH'] +=';'+psspath

configure_paths_for_psse()
import psspy              # it sets new path for psspy
import redirect
import dyntools
map_Signal2Channel = {'DELTA': 1, 'PELEC': 2, 'QELEC': 3, 'ETERM': 4, 'EFD': 5, 'PMECH': 6, 'SPEED': 7,
                      'XADIFD': 8,
                      'ECOMP': 9, 'VOTHSG': 10, 'VREF': 11, 'BSFREQ': 12, 'VMAG': 13, 'VMAG_VANG': 14,
                      'PLINE': 15
    , 'PLINE_QLINE': 16, 'SLINE': 17, 'RX': 18, 'ITERM': 21,
                      'MACHINE_RX': 22, 'VUEL': 23, 'VOEL': 24, 'PLOAD': 25, 'QLOAD': 26,
                      'GREF': 27, 'LCREF': 28}


class PSSESimulation:
    def __init__(self, raw_file, dyr_file):
        self.raw_file = raw_file
        self.dyr_file = dyr_file
        self.achnf = None
        self.out_file = self._change_extention('.out')
        self._initialize_psse()
    def _initialize_psse(self):
        redirect.psse2py()
        psspy.psseinit(1000)

        # Read load flow and dynamic files
        psspy.read(0, self.raw_file)
        psspy.dyre_new([1, 1, 1, 1], self.dyr_file, "", "", "")
        psspy.fdns([0, 0, 0, 1, 1, 0, 0, 0])
        # Convert Generator and Loads
        psspy.cong(0)
        # Order Network for matrix operation
        psspy.ordr(0)
        # Factorize Admittance matrix
        psspy.fact()
        # Solve switching study network solutions
        psspy.tysl()

    def _compute_achnf(self):
        if self.achnf is None:
            self.achnf = dyntools.CHNF(self.out_file);

    def get_result(self):
        self._compute_achnf()
        # Data Retrieval
        ch_data = self.achnf.chandata;
        ch_id = self.achnf.chanid;
        return ch_data, ch_id
    def get_result_csv(self):
        self._compute_achnf()
        # Data Retrieval
        ch_data = self.achnf.chandata;
        ch_id = self.achnf.chanid;
        csv_out = self._save_csv(ch_data, ch_id)
        return csv_out

    def _change_extention(self, extention):
        # generated files
        path = os.path.abspath(self.dyr_file)
        directory = os.path.dirname(path)
        fname = os.path.basename(path)

        file_name, _ = os.path.splitext(fname)
        out_file = 'PSSE_{0}{1}'.format(file_name,extention);
        return os.path.join(directory, out_file)

    def select_channel(self, signal_list):
        psspy.delete_all_plot_channels()

        for i in signal_list.keys():
            bus_list = signal_list.get(i)[0]
            if bus_list:
                psspy.bsys(sid=i, numbus=len(bus_list), buses=bus_list)
            variable_list = signal_list.get(i)[1]
            for variable in variable_list:
                f = None
                if variable == 'STATE':
                    f = psspy.state_channel
                elif variable == 'VAR':
                    f = psspy.var_channel

                if f is not None:
                    for j in bus_list:
                        f(status=[-1, j], ident=variable + str(j))
                else:
                    list_ = [-1, -1, -1, 1, map_Signal2Channel.get(variable), 1]
                    if bus_list:
                        psspy.chsb(i, 0, list_)
                    else:
                        psspy.chsb(0, 1, list_)
        psspy.strt(0, self.out_file)

    def _save_csv(self,ch_data, ch_id):
        import csv
        csv_file = self._change_extention('.csv')

        first_row=ch_id.get(ch_id.keys()[0])
        data=ch_data.get(ch_data.keys()[0])

        with open(csv_file, 'wb') as f:  # Just use 'w' mode in 3.x
            def swap(list_, index1, index2):
                foo = list_[index2]
                list_[index2] = list_[index1]
                list_[index1] = foo

            data_matrix = [i for i in data.values()]
            time_index=len(data_matrix)-1
            list_keys = list(first_row.keys())
            swap(list_keys, time_index, 0)
            w = csv.DictWriter(f, list_keys)
            w.writerow(first_row)
            # w = csv.DictWriter(f, data.keys())
            # w.writerow(data)
            swap(data_matrix, time_index, 0)
            # data_matrix.reverse()
            data_matrix=map(list, zip(*data_matrix))
            w = csv.writer(f)
            for i in data_matrix:
                w.writerow(i)

        return csv_file

    def plot(self, csvfile, mat):
        import os
        cwd = os.getcwd()
        import numpy as np
        import matplotlib.pyplot as plt
        import scipy.io

        # matfile= os.path.join(os.path.dirname(__file__), 'ePH.mat')
        localmat = mat
        # matfile = 'C:\Zhiqi_Local\\allwork\project\python project\ePH.mat'
        # csvfile = 'C:\Zhiqi_Local\\allwork\project\python project\\123.csv'
        psse = np.genfromtxt(csvfile, delimiter=",", skip_header=1)
        varname = np.genfromtxt(csvfile, delimiter=",", skip_footer=2)
        eph1 = scipy.io.loadmat(localmat)
        ll=str(eph1.keys())
        need='s_{0}_{1}_{2}'.format(16, 'gen_36_1', 'delta')
        t='s_{0}'.format('time')
       # print psse[6, :]
      #  print eph1[t]
      #  print eph1[need]
        plt.plot(psse[:, 0], psse[:, 6], 'r')
        plt.plot(eph1[t][0,:], eph1[need][0,:], 'b')
        plt.show()