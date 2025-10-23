"""
pyTNMR is a python library for controlling TECMAG's TNMR software

The software is designed so that only one file is open at a time.
This file is tracked as self.current

pyTNMR was designed in this manner to prevent having large numbers
of open files which can create issues over large runs.

Michael Malone mwmalone@gmail.com
"""

"""CHANGE LOG

2025-10-22 actual debugging
2025-09-18 The great _ reformat
2025-09-13 Better logging and folder creation
2022-07-22 setTable and getTable added - ARA
2020-12-10 SpyctraV5 Overhaul
2018-04-01 Overhaul
"""

import os
from time import perf_counter as time, sleep
import win32com.client
import ntpath
import shutil

import sys
sys.path.append('C://code//spyctra_V6//')

from spyctra import spyctra
from TNT import read as read_TNT

class TNMR(object):
    def __init__(self, path, unique=True, running=1):
        """TNMR Object initialization

        As is TNMR tradition, pyTNMR expects an
        open .tnt file when initiated

        Args:
            path: The default directory all data and
                associtated log files will be stored in.
            unique: if False will overwrite data in path
                    if True will append unique number to path
            *running: binary flag determining whether
                sequences are actually run. Useful for
                debugging sequences
        """
        self.unique = unique
        self.running = running

        self.root = os.getcwd() + '\\' 
        self.exp_path = self.get_exp_path(path.replace('/', '\\'))
        suffix = self.copy_source()

        self.tnmr = win32com.client.Dispatch("NTNMR.Application")
        #print(dir(self.tnmr))
        self.current = self.tnmr.GetActiveDocPath
        
        self.pyTNMR_log_file = self.init_log('pyTNMR', suffix)
        self.exp_log_file = self.init_log('exp', suffix)

        self.report()

        input('Ready? ')


    def close(self, filename):
        """
        Closes the active file

        Is not needed in scripts since opening a new file
        will close the current file
        """

        filename = self.name_check(filename)

        self.pyTNMR_log(f'  CLOSED {filename}')
        self.tnmr.CloseFile(filename)


    def copy_source(self):
        source = sys.argv[0]

        print(f'{source = }')

        source_name = ntpath.basename(source)

        print(f'{source_name = }')

        souce_destination = self.exp_path + source_name

        c = 0

        while os.path.exists(souce_destination):
            c += 1
            souce_destination = f'{self.exp_path}{source_name[:-3]}-v{c}.py'

        print(f'{souce_destination = }')

        shutil.copyfile(source, souce_destination)
        
        if c > 0:
            suffix = f'-v{c}'
        else:
            suffix = ''
        
        return suffix


    def get_exp_path(self, path):
        if path[-1] != '\\':
            path += '\\'

        path0 = path

        if self.unique:
            c = 0
            while os.path.isdir(path0[:-1] + f'_{c}\\'):
                c += 1

            exp_path = path0[:-1] + f'_{c}\\'          
            print(f'Making experiment directory {exp_path}')
            os.makedirs(exp_path)
            
            return exp_path
        else:
            if os.path.isdir(path0):
                print(f'{path0 = } exists')
            else:
                print(f'Making experiment directory {path0}')
                
                os.makedirs(path0)
            
            return path0


    def get_param(self, param):
        value = self.tnmr.GetNMRParameter(param)
        self.pyTNMR_log(f'    - {param} has value {value}')

        return value


    def get_table(self, table):
        value = self.tnmr.GetTable(table)
        self.pyTNMR_log(f'    - {table} has values {value}')

        return value

    
    def init_log(self, log_name, suffix):
        log_file = f'{self.exp_path}{log_name}_log{suffix}.txt'
        
        with open(f'{log_file}', 'w') as a:
            a.write('')
    
        return log_file



    def log(self, line):
        """
        selectively prints args to experiment_log.txt
        """

        with open(f'{self.exp_log_file}', 'a') as a:
            print(line)
            a.write(line + '\n')


    def name_check(self, filename):
        if filename[-4:] != '.tnt':
            filename += '.tnt'
        if filename.find('/')<0 and filename.find('\\')<0:
            filename = self.root + self.exp_path + filename

        filename = filename.replace('/','\\')
        
        return filename


    def open(self, filename):
        """
        Opens a new file and closes the current file

        Exits program if file not found
        will search for filename in self.exp_path if more detailed path not specified
        in filename
        """

        filename =self.name_check(filename)
        self.pyTNMR_log('\n---------------')
        self.pyTNMR_log(filename[filename.rfind('\\')+1:])
        self.pyTNMR_log(f' OPENING {filename}')

        #Check if file is already open--happens when using loops
        if filename != self.tnmr.GetActiveDocPath:
            self.tnmr.OpenFile(filename)
            self.close(self.current)
            self.current = self.tnmr.GetActiveDocPath

            if filename != self.tnmr.GetActiveDocPath:
                self.pyTNMR_log('    Could not open')
                self.pyTNMR_log(filename)
                self.pyTNMR_log(self.tnmr.GetActiveDocPath)
                input('Exit')
        else:
            self.pyTNMR_log('  Already Opened')

        self.pyTNMR_log()


    def pyTNMR_log(self, line=''):
        """
        prints both to the stdout and the experiment log file in the
        experimental data folder
        """

        with open(f'{self.pyTNMR_log_file}', 'a') as a:
            print(line)
            a.write(line + '\n')


    def read(self, filename):
        return read_TNT(self.exp_path + filename)


    def report(self):
        print()
        print(f'{self.root = }')
        print(f'{self.exp_path = }')
        print(f'{self.running = }')
        print(f'{self.unique = }')


    def reset(self):
        self.pyTNMR_log('Resetting Hardware')
        self.tnmr.Reset


    def restart(self, timeout=3):
        """
        Restarts the TNMR program
        Holdover from old code.
        Kept for future concerns
        """
        self.pyTNMR_log("Crashed! Restarting tnmr")
        os.system("taskkill /F /IM TNMR.exe")
        self.tnmr = win32com.client.Dispatch("TNMR.Application")
        sleep(timeout)
        # for some reason, things are really wonky if self.tnmr is created
        # without a file loaded. So, load a file and re-create that object.
        self.open_file(self.sweep_file)
        self.tnmr = win32com.client.Dispatch("TNMR.Application")


    def save_as(self, filename):
        """
        Saves the current file to the specified path
        """
        filename = self.name_check(filename)
        self.pyTNMR_log(f'  SAVING {self.current}')
        self.pyTNMR_log(f'      AS {filename}')
        self.tnmr.SaveAs(filename)
        self.close(self.current)
        self.current = self.tnmr.GetActiveDocPath
        self.pyTNMR_log()


    def set_param(self, param, value):
        """Set TNMR parameter to the specified value

        args:
            param (str): the TNMR parameter to be changed
            value (str,int,floatt): new parameter value

        """

        self.pyTNMR_log(f'    -Setting {param} to {value}')
        self.tnmr.SetNMRParameter(param, value)


    def set_table(self, table, value):
        """Set TNMR table to ###

        args:
            param (str): the TNMR table to be changed
            value (str): new tables values, must be comma seperated

        """

        self.pyTNMR_log(f'    -Setting {table} to {value}')
        self.tnmr.SetTable(table, value)


    def sleep(self, t):
        if self.running == 1:
            sleep(t)


    def zg(self, *manual_check):
        """ZeroAndGo
        Runs the active TNMR file checking every second for completion

        manual check allows the user to verify after each zg if the
        experiment needs to be repeated
        """
        scans_1D = int(self.tnmr.GetNMRParameter('Scans 1D'))
        points_2D = int(self.tnmr.GetNMRParameter('Points 2D'))
        total_scans = scans_1D*points_2D

        if self.tnmr.GetNMRParameter('Points 3D')*self.tnmr.GetNMRParameter('Points 4D') > 1:
            raise ValueError('ERROR: What is wrong with you? The whole point of pyTNMR is to avoid 3D and 4D tables')

        while True:
            if self.running == 1:
                self.pyTNMR_log('\n RUNNING')
                self.tnmr.ZG
                # Now wait for it to finish.
                t0 = time()
                sleep(.1)
                expected_time_str = self.tnmr.GetNMRParameter("Exp. Elapsed Time")
                expected_time = [int(val) for val in expected_time_str.split(':')]
                expected_time = 3600*expected_time[0] + 60*expected_time[1] + expected_time[2]

                if expected_time == 0:
                    expected_time = 1

                self.pyTNMR_log(f'{expected_time_str = }')
                self.pyTNMR_log(f'{expected_time = }')
                self.pyTNMR_log(f'{total_scans = }')

                done = False
                c = 1
                sleep_inc = max(expected_time/9, 1)

                while not done:
                    sleep(sleepInc) #evervy 9th of data collect
                    done = self.tnmr.CheckAcquisition
                    act_points_1D = int(self.tnmr.GetNMRParameter('Actual Scans 1D'))
                    act_points_2D = int(self.tnmr.GetNMRParameter('Actual Points 2D'))
                    percent_done = 100*(act_points_1D + (act_points_2D-1)*scans1D)/total_scans

                    if (time()-t0)/expected_time > 2*percent_done or (percent_done==0 and time()-t0>20):
                        print(f'Collect failed?')

                    if percent_done>c/10:
                        print(f'{percent_done = :.1f} %')

                        c = percent_done//10+1
            break

        if manual_check:
            print()

            check = input('Happy? ')

            if check in ['n','N','no','No']:
                self.zg(manual_check[0])


class TNMRError(Exception):
    """
    A TNMR-specific exception class
    Placehold and holdover from old code
    """
    pass


def test_suite():
    a = TNMR('data', False, 0)
    a.open(f'{a.root}RO')
    
    for rec_gain in [i+60 for i in range(10)]:
        a.log(f'{rec_gain = }')
        a.set_param('Receiver Gain', rec_gain)
        a.zg()
        a.save_as(f'RO_RecGain{rec_gain}')
        
        
def main():
    test_suite()


if __name__ == '__main__':
    main()
