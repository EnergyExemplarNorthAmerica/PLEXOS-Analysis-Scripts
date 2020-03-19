

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

def query_data_to_csv(sol, csv_file, sim_phase, coll, period):
    # Run the query
    try:
        if sol.QueryToCSV(csv_file, True, sim_phase, coll, '', '', period, SeriesTypeEnum.Values, ''):
            print('{1} successfully {0} phase output.'.format(str(sim_phase), str(coll)))
        else:
            print('{1} found no {0} phase output.'.format(str(sim_phase), str(coll)))
    except Exception as ex:
        print('{1} is not present in the {0} phase output.'.format(str(sim_phase), str(coll)))


def pull_data(sol_cxn, time_res, arg_opt, default_csv):
    # Pull Annual Output
    try:
        idx = sys.argv.index(arg_opt)
        if idx + 1 >= len(sys.argv) or sys.argv[idx + 1][0] == '-':
            csv_file = re.sub('\.zip$', '', sys.argv[1]) + default_csv
        else:
            csv_file = sys.argv[idx + 1]
    except Exception as ex:
        raise ex

    if os.path.exists(csv_file): os.remove(csv_file)
    for phase in Enum.GetValues(type(SimulationPhaseEnum)):
        for coll in Enum.GetValues(type(CollectionEnum)):
            query_data_to_csv(sol_cxn, csv_file, phase, coll, time_res)

def pull_annual(sol_cxn):
    pull_data(sol_cxn, PeriodEnum.FiscalYear, '-y', '_annual.csv')

def pull_monthly(sol_cxn):
    pull_data(sol_cxn, PeriodEnum.Month, '-m', '_monthly.csv')

def main():
    if len(sys.argv) <= 1:
        print('Usage: python power_bi_link.py <solution_file> [-y [yr_file] -m [mn_file]]')
    
    start_time = time.time()

    # setup and connect to the solution file
    sol_file = sys.argv[1]
    sol_cxn = Solution()
    sol_cxn.Connection(sol_file)
    sol_cxn.DisplayAlerts = False
    pull_annual(sol_cxn) # pull annual results
    pull_monthly(sol_cxn) # pull monthly results

    print ('Completed in', time.time() - start_time, 'sec')


if __name__ == '__main__':
    main()
