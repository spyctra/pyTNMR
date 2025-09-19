import sys
sys.path.append('c:/code/spyctra_v6')

from pyTNMRV9 import TNMR
from fitlib import fit

from time import sleep
from math import e, pi
from numpy.random import rand

import numpy as np

import pyvisa
from subprocess import Popen #, PIPE, STDOUT


def compExp(x, amp, df, phi):
    return amp*e**(1j*(-2*pi*df*x + phi))
    
    
def get_CPMG_freq(s0):
    s = s0.copy()
    s.new_count(2048)
    s.decimate()
    
    g = s.copy()
    g.resize(2048)
    g.fft()
    
    df = g.find_df()
    phi = g.find_phi()

    s.shift(1)
    s.resize(37)
    
    p,r = fit(compExp, s.x, s.data[0],
              [ np.abs(s.data[0][0])
               ,df
               ,phi
              ]
              ,guess=0, check=0, result='amp,df,phi')
              
    f1 = round(s0.freq[0] + p[0][1])
    
    if np.abs(f1-2.25e6) > 6000:
        f1 = 2.25e6
    
    return f1/1e6


def main():   
    BpOn = 64.3
    trials = 64
    f1 = 2.25164
        
    a = TNMR('data', unique=False, running=1)

    a.open(f'{a.root}CPMG_Bp_NMR_v3')
    a.set_param('Scans 1D', 1)   
    a.set_param('Last Delay', '1s')   
    a.set_param('Bp', BpOn)   
    a.set_param('Observe Freq.', f1)   
    a.save_as('CPMG_BASE')

    a.open(f'{a.root}FID_Bp_NMR_v3')
    a.set_param('Scans 1D', 1)   
    a.set_param('Last Delay', '1s')   
    a.set_param('Bp', BpOn)   
    a.set_param('Observe Freq.', f1)   
    a.save_as('FID_BASE')
    
    j0 = 0
    
    for j in range(j0, trials):        
        a.open('CPMG_BASE')
        a.zg()
        a.save_as(f'CPMG_BASE_{j}')
        a.sleep(1)

        b = a.get_spyctra()
        f1 = get_CPMG_freq(b)
        a.set_param('Observe Freq.', f1)   
        
        a.open('FID_BASE')
        a.set_param('Observe Freq.', f1)   
        a.zg()
        a.save_as(f'FID_{j}')
        a.sleep(8)


if __name__ == '__main__':
    main()
    
    


