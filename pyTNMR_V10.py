"""
pyTNMR is a python library for controlling TECMAG's TNMR software

The software is designed so that only one file is open at a time.
This file is tracked as self.current

pyTNMR was designed in this manner to prevent having large numbers
of open files which can create issues over large runs.

No error controls are implemented since no errors
have ever been personally observed.

Michael Malone mwmalone@gmail.com
"""

"""CHANGE LOG

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

        self.root = os.getcwd() + '\\'
        self.path = self.getPath(path, unique)
        self.source = self.copy_source()

        self.tnmr = win32com.client.Dispatch("NTNMR.Application")
        #print(dir(self.tnmr))
        self.current = self.tnmr.GetActiveDocPath
        self.log_file = open(f'{path}log.txt', 'w')
        self.log_file.close()
        self.logger(f'Saving log in {path}log.txt')
        self.logger(f'Initializing TNMR with {self.current}')

        self.data_log = open(f'{path}data_log.txt', 'w')
        self.data_log.close()

        self.running = running

        self.print_self()

        input('Ready? ')


    def copy_source(self):
        source = sys.argv[0]

        print(f'{source = }')

        filename = ntpath.basename(source)

        print(f'{filename = }')

        file_path = path + filename

        c = 0

        while os.path.exists(file_path):
            c += 1
            file_path = f'{path}{filename[:-3]}-v{c}.py'

        print(f'{file_path = }')

        shutil.copyfile(source, file_path)


    def close(self, filename):
        """
        Closes the active file

        Is not needed in scripts since opening a new file
        will close the current file
        """

        filename = self.name_check(filename)

        self.logger(f'  CLOSED {filename}')
        self.tnmr.CloseFile(filename)


    def get_path(self, path, unique):
        if path[-1] != '\\':
            path += '\\'

        path0 = self.root + path + '\\'

        if os.path.isdir(path0):
            print('{path0 = } exists')

            if unique:
                c = 0

                while os.path.isdir(path0):
                    c += 1
                    path0 = self.root + path[:-1] + f'_{c}\\'

                os.makedirs(path0)

                return path0
            else:
                return path0
        else:
            os.makedirs(path0)

        return path0


    def open(self, filename):
        """
        Opens a new file and closes the current file

        Exits program if file not found
        will search for filename in self.path if more detailed path not specified
        in filename
        """

        filename =self.name_check(filename)
        self.logger('\n---------------')
        self.logger(filename[filename.rfind('\\')+1:])
        self.logger(f' OPENING {filename}')

        #Check if file is already open--happens when using loops
        if filename != self.tnmr.GetActiveDocPath:
            self.tnmr.OpenFile(filename)
            self.close(self.current)
            self.current = self.tnmr.GetActiveDocPath

            if filename != self.tnmr.GetActiveDocPath:
                self.logger('    Could not open')
                self.logger(filename)
                self.logger(self.tnmr.GetActiveDocPath)
                input('Exit')
        else:
            self.logger('  Already Opened')

        self.logger()


    def print_self(self):
        print(f'{self.root = }')
        print(f'{self.path = }')
        print(f'{self.running = }')


    def get_spyctra(self, *filename):
        """
        returns tnt spyctra object
        """

        if filename:
            return TNT_reader(self.name_check(filename[0]))
        else:
            return TNT_reader(self.tnmr.GetActiveDocPath)


    def set_table(self, table, value):
        """Set TNMR table to ###

        args:
            param (str): the TNMR table to be changed
            value (str): new tables values, must be comma seperated

        """

        self.logger(f'    -Setting {table} to {value}')
        self.tnmr.SetTable(table, value)


    def set_param(self, param, value):
        """Set TNMR parameter to the specified value

        args:
            param (str): the TNMR parameter to be changed
            value (str,int,floatt): new parameter value

        """

        self.logger(f'    -Setting {param} to {value}')
        self.tnmr.SetNMRParameter(param, value)


    def get_param(self, param):
        value = self.tnmr.GetNMRParameter(param)
        self.logger(f'    - {param} has value {value}')

        return value


    def get_table(self, table):
        value = self.tnmr.GetTable(table)
        self.logger(f'    - {table} has values {value}')

        return value


    def save_as(self, filename):
        """
        Saves the current file to the specified path
        """
        filename =self.name_check(filename)
        self.logger(f'  SAVING {self.current}')
        self.logger(f'      AS {filename}')
        self.tnmr.SaveAs(filename)
        self.close(self.current)
        self.current = self.tnmr.GetActiveDocPath
        self.logger()


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
                self.logger('\n RUNNING')
                self.tnmr.ZG
                # Now wait for it to finish.
                t0 = time()
                sleep(1)
                expected_time_str = self.tnmr.GetNMRParameter("Exp. Elapsed Time")
                expected_time = [int(val) for val in expected_time_str.split(':')]
                expected_time = 3600*expected_time[0] + 60*expected_time[1] + expected_time[2]

                if expected_time == 0:
                    expected_time = 1

                self.logger(f'{expected_time_str = }')
                self.logger(f'{expected_time = }')
                self.logger(f'{total_scans = }')

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


    def read(self, filename):
        return read_TNT(self.path + filename)


    def reset(self):
        self.logger('Resetting Hardware')
        self.tnmr.Reset


    def restart(self, timeout=3):
        """
        Restarts the TNMR program
        Holdover from old code.
        Kept for future concerns
        """
        self.logger("Crashed! Restarting tnmr")
        os.system("taskkill /F /IM TNMR.exe")
        self.tnmr = win32com.client.Dispatch("TNMR.Application")
        sleep(timeout)
        # for some reason, things are really wonky if self.tnmr is created
        # without a file loaded. So, load a file and re-create that object.
        self.open_file(self.sweep_file)
        self.tnmr = win32com.client.Dispatch("TNMR.Application")


    def sleep(self, t):
        if self.running == 1:
            sleep(t)


    def log(self, *args):
        """
        selectively prints args to data_log.txt
        """

        self.data_log = open(f'{self.path}data_log.txt', 'a')

        for arg in args:
            print(arg, end=',', file=self.data_log)

        print(file = self.data_log)

        self.data_log.close()


    def logger(self, *args):
        """
        prints both to the stdout and the experiment log file in the
        experimental data folder
        """

        self.log_file = open(f'{self.path}log.txt', 'a')

        for arg in args:
            print(arg, end=' ')
            print(arg, end=' ', file = self.log_file)

        print()
        print(file = self.log_file)

        self.log_file.close()


    def name_check(self, filename):
        if filename[-4:] != '.tnt':
            filename += '.tnt'
        if filename.find('/')<0 and filename.find('\\')<0:
            filename = self.path + filename

        return filename.replace('/','\\')


class TNMRError(Exception):
    """
    A TNMR-specific exception class
    Placehold and holdover from old code
    """
    pass


def main():
    a = TNMR('./', 0)
    sleep(5)
    print('a')
    a.reset()
    sleep(5)
    a.reset()



if __name__ == '__main__':
    main()



