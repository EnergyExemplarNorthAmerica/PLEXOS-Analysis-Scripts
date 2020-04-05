
import pandas as pd
import numpy as np
import io, re, os, sys, time, clr

from System import Enum, Boolean, DateTime

def is_switch(args, arg_opt):
    '''
    Check to see if switch arg_opt is defined in the command line
    '''
    return arg_opt in args

def switch_index(args, arg_opt):
    '''
    Determine the index at which arg_opt occurs
    '''
    return -1 if not is_switch(args, arg_opt) else args.index(arg_opt)

def switch_data(args, arg_opt):
    '''
    Find the data if any that is associated with switch arg_opt
    '''
    if is_switch(args, arg_opt):
        idx = switch_index(args, arg_opt)
        if 0 <= idx < len(args) - 1:
            if not args[idx + 1][0] == '-':
                return args[idx + 1]
    
    return None

def switch_data_to_date(args, arg_opt):
    '''
    Pull a date from the command line
    '''
    try:
        return DateTime.Parse(switch_data(args, arg_opt))
    except:
        return None

def query_data_to_csv(sol, csv_file, sim_phase, coll, period, date_from, date_to, is_verbose = False):
    # Run the query
    from EEUTILITY.Enums import SeriesTypeEnum
    params = (csv_file, True, sim_phase, coll, '', '', period, SeriesTypeEnum.Values, '', date_from, date_to)
    try:
        if sol.QueryToCSV(*params):
            if is_verbose:
                print('{1} successfully {0} phase output.'.format(str(sim_phase), str(coll)))
        else:
            if is_verbose:
                print('{1} found no {0} phase output.'.format(str(sim_phase), str(coll)))
    except Exception as ex:
        if is_verbose:
            print('{1} is not present in the {0} phase output.'.format(str(sim_phase), str(coll)))


def pull_data(sol_cxn, time_res, args, arg_opt, default_csv):
    from EEUTILITY.Enums import SimulationPhaseEnum, CollectionEnum, PeriodEnum

    # is this time_res active? If not quit
    if not is_switch(args, arg_opt): return

    # date_from and date_to
    date_from = switch_data_to_date(args, '-f')
    date_to = switch_data_to_date(args, '-t')

    start_time = time.time() # start timer

    # get the csv_file name
    csv_file = switch_data(args, arg_opt)
    if csv_file is None:
        csv_file = re.sub('\.zip$', '', args[1]) + default_csv

    # remove the csv_file if it already exists
    if os.path.exists(csv_file): os.remove(csv_file)

    # loop through all relevant collections and phases
    for phase in Enum.GetValues(clr.GetClrType(SimulationPhaseEnum)):
        for coll in Enum.GetValues(clr.GetClrType(CollectionEnum)):
            query_data_to_csv(sol_cxn, csv_file, phase, coll, time_res, date_from, date_to, is_verbose=is_switch(args, '-v'))

    print('Completed',clr.GetClrType(PeriodEnum).GetEnumName(time_res),'in',time.time() - start_time,'sec')

def main():
    from EEUTILITY.Enums import PeriodEnum
    
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
    pull_data(sol_cxn, PeriodEnum.FiscalYear, sys.argv, '-y', '_annual.csv') # pull annual results
    pull_data(sol_cxn, PeriodEnum.Quarter, sys.argv, '-q', '_quarterly.csv') # pull quarter results
    pull_data(sol_cxn, PeriodEnum.Month, sys.argv, '-m', '_monthly.csv') # pull month results
    pull_data(sol_cxn, PeriodEnum.Week, sys.argv, '-w', '_weekly.csv') # pull week results
    pull_data(sol_cxn, PeriodEnum.Day, sys.argv, '-d', '_daily.csv') # pull day results
    pull_data(sol_cxn, PeriodEnum.Hour, sys.argv, '-h', '_hourly.csv') # pull hourly results
    pull_data(sol_cxn, PeriodEnum.Interval, sys.argv, '-i', '_interval.csv') # pull interval results

    print ('Completed in', time.time() - start_time, 'sec')

def set_plexos_path(plexos_path):
    # load PLEXOS assemblies:
    sys.path.append(plexos_path)
    clr.AddReference('PLEXOS7_NET.Core')
    clr.AddReference('EEUTILITY')
    from PLEXOS7_NET.Core import DatabaseCore, Solution, PLEXOSConnect

if __name__ == '__main__':

    set_plexos_path('C:/Program Files (x86)/Energy Exemplar/PLEXOS 8.1/')
    main()
