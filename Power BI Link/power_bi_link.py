

import pandas as pd
import numpy as np
import io, re, os, sys, time

from dotnet import *
from dotnet.overrides import type, isinstance, issubclass
from dotnet.commontypes import *

# load PLEXOS assemblies:
plexos_path = 'C:/Program Files (x86)/Energy Exemplar/PLEXOS 8.1/'
add_assemblies(plexos_path)
load_assembly('PLEXOS7_NET.Core')
load_assembly('EEUTILITY')

from PLEXOS7_NET.Core import *
from EEUTILITY.Enums import *
from System import Enum

def is_switch(arg_opt):
    '''
    Check to see if switch arg_opt is defined in the command line
    '''
    return arg_opt in sys.argv

def switch_index(arg_opt):
    '''
    Determine the index at which arg_opt occurs
    '''
    return -1 if not is_switch(arg_opt) else sys.argv.index(arg_opt)

def switch_data(arg_opt):
    '''
    Find the data if any that is associated with switch arg_opt
    '''
    if is_switch(arg_opt):
        idx = switch_index(arg_opt)
        if 0 <= idx < len(sys.argv) - 1:
            if not sys.argv[idx + 1][0] == '-':
                return sys.argv[idx + 1]
    
    return None

def query_data_to_csv(sol, csv_file, sim_phase, coll, period):
    # Run the query
    try:
        if sol.QueryToCSV(csv_file, True, sim_phase, coll, '', '', period, SeriesTypeEnum.Values, ''):
            if is_switch('-v'):
                print('{1} successfully {0} phase output.'.format(str(sim_phase), str(coll)))
        else:
            if is_switch('-v'):
                print('{1} found no {0} phase output.'.format(str(sim_phase), str(coll)))
    except Exception as ex:
        if is_switch('-v'):
            print('{1} is not present in the {0} phase output.'.format(str(sim_phase), str(coll)))


def pull_data(sol_cxn, time_res, arg_opt, default_csv):

    # is this time_res active? If not quit
    if not is_switch(arg_opt): return

    start_time = time.time() # start timer

    # get the csv_file name
    csv_file = switch_data(arg_opt)
    if csv_file is None:
        csv_file = re.sub('\.zip$', '', sys.argv[1]) + default_csv

    # remove the csv_file if it already exists
    if os.path.exists(csv_file): os.remove(csv_file)

    # loop through all relevant collections and phases
    for phase in Enum.GetValues(type(SimulationPhaseEnum)):
        for coll in Enum.GetValues(type(CollectionEnum)):
            query_data_to_csv(sol_cxn, csv_file, phase, coll, time_res)

    print('Completed',str(time_res),'in',time.time() - start_time,'sec')

def main():
    if len(sys.argv) <= 1:
        print('''
Usage:
    python power_bi_link.py <solution_file> [-y [yr_file]]
                                            [-q [qt_file]]
                                            [-m [mn_file]]
                                            [-w [wk_file]]
                                            [-d [dy_file]]
                                            [-h [hr_file]]
                                            [-i [in_file]]
                                            [-f [from_date]]
                                            [-t [to_date]]
        ''')
        return
    
    start_time = time.time()

    # setup and connect to the solution file
    sol_file = sys.argv[1]
    sol_cxn = Solution()
    sol_cxn.Connection(sol_file)
    sol_cxn.DisplayAlerts = False

    # pull data
    pull_data(sol_cxn, PeriodEnum.FiscalYear, '-y', '_annual.csv') # pull annual results
    pull_data(sol_cxn, PeriodEnum.Quarter, '-q', '_quarterly.csv') # pull quarter results
    pull_data(sol_cxn, PeriodEnum.Month, '-m', '_monthly.csv') # pull month results
    pull_data(sol_cxn, PeriodEnum.Week, '-w', '_weekly.csv') # pull week results
    pull_data(sol_cxn, PeriodEnum.Day, '-d', '_daily.csv') # pull day results
    pull_data(sol_cxn, PeriodEnum.Hour, '-h', '_hourly.csv') # pull hourly results
    pull_data(sol_cxn, PeriodEnum.Interval, '-i', '_interval.csv') # pull interval results

    print ('Completed in', time.time() - start_time, 'sec')


if __name__ == '__main__':
    main()
