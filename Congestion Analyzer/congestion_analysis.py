'''
CONGESTION ANALYSIS TOOL

Approach & Idea : TAREK IBRAHIM
Author : TAREK IBRAHIM
Acknowledgments : Energy Exemplar Solution Engineering Team
'''

import csv
import pandas as pd
import os
import time
import sys, re
import csv
import numpy as np
from pandas.io.common import EmptyDataError
from sympy import symbols
import xlsxwriter
from shutil import copyfile
from dotnet import add_assemblies, load_assembly
from dotnet.overrides import type, isinstance, issubclass
from dotnet.commontypes import *

# load PLEXOS assemblies:
plexos_path = 'C:/Program Files (x86)/Energy Exemplar/PLEXOS 8.1/'
add_assemblies(plexos_path)
load_assembly('PLEXOS7_NET.Core')
load_assembly('EEUTILITY')

from PLEXOS7_NET.Core import *
from EEUTILITY.Enums import *

def main(FB, TB, TS, ptdf_data, plexos_db = '', plexos_sol = '', db=None, sol=None):

    # Initialize timer and file io
    start_time = time.time()
    Line = '{}_{}'.format(FB,TB)
    temp_file = '{}_sol.csv'.format(Line)
    results_folder = './Results {}'.format(Line)
    results_excel = '{}.xlsx'.format(Line)

    # create the relevant results folder
    os.makedirs(results_folder, exist_ok=True)

    # Load PLEXOS Input and Output data
    if db is None:
        db = DatabaseCore()
        db.Connection(plexos_db)
        print('Loaded PLEXOS input:', time.time() - start_time)

    if sol is None:
        sol = Solution()
        sol.Connection(plexos_sol)
        print('Loaded PLEXOS solution:', time.time() - start_time)

    # Before we load the PTDF data, let's first find which data we actually want to utilize
    #   The PTDF data file is produced from a different PLEXOS Analysis script.
    with open(ptdf_data) as infile:
        # read the first 3 lines of the file
        #   line1 is the columns header
        #   line2 is a list of "from" busses
        #   line3 is a list of "to" busses
        line1, line2, line3 = infile.readline(), infile.readline(), infile.readline()

        # find only those columns that have the right from and to bus pairs
        idxs = [idx for idx, from_bus, to_bus in zip(line1.split(','), line2.split(','), line3.split(',')) if FB in from_bus and TB in to_bus]

        print('Computed required PTDF columns:', time.time() - start_time)
    
    # Pull only those data that are related to this from-to pair
    ptdf_df = pd.read_csv(ptdf_data, usecols=['0'] + idxs, low_memory=False, skiprows=[1,2])
    ptdf_df.rename(columns = {'0':'child_name'}, inplace=True)
    print('Loaded PTDF data:', time.time() - start_time)

    # Pull a query from the PLEXOS data into a CSV file
    #   below is the signature of the API method
    '''
    Boolean QueryToCSV(
            String strCSVFile,
            Boolean bAppendToFile,
            SimulationPhaseEnum SimulationPhaseId,
            CollectionEnum CollectionId,
            String ParentName,
            String ChildName,
            PeriodEnum PeriodTypeId,
            SeriesTypeEnum SeriesTypeId,
            String PropertyList[ = None],
            Object DateFrom[ = None],
            Object DateTo[ = None],
            String TimesliceList[ = None],
            String SampleList[ = None],
            String ModelName[ = None],
            AggregationEnum AggregationType[ = None],
            String Category[ = None],
            String Separator[ = ,]
            )
    '''
    sol.QueryToCSV(temp_file, False, \
                    SimulationPhaseEnum.STSchedule, \
                    CollectionEnum.SystemNodes, \
                    '', \
                    '', \
                    PeriodEnum.Interval, \
                    SeriesTypeEnum.Periods, \
                    '22')
    injection_df = pd.read_csv(temp_file)
    os.remove(temp_file)
    injection_df = injection_df.drop(injection_df.columns[[0,1,2,3,4,5,6,7,8,9,11,13,14,15,16,17,18,19]], axis = 1)
    print('Queried Net Injections:', time.time() - start_time)
        
    # join injections to ptdfs
    temp_injection_df = pd.DataFrame(columns=['child_name',Line])
    for col in [col for col in ptdf_df if not col == 'index']:
        temp_injection_df['child_name'] = ptdf_df['child_name']
        temp_injection_df[Line] = ptdf_df[col]


    ptdf_df['child_name'] = ptdf_df.applymap(str)
    injection_df['child_name'] = injection_df.applymap(str)

    ptdf_df = ptdf_df.merge(injection_df, on='child_name')

    injection_df = injection_df.merge(temp_injection_df, on='child_name')

    # begin calculations of shift factors and impacts
    Mult = pd.DataFrame()
    Mult[Line] = injection_df.iloc[:,0]

    for z in range(2, len(injection_df.columns)-1):
        row = injection_df.columns[z]
        injection_df[row] = injection_df[row].astype(float)
        injection_df[Line] = injection_df[Line].astype(float)
        injection_df = injection_df.dropna()
        Mult[row] = injection_df[row] * injection_df[Line]
    
    Mult['Shift Factor'] = injection_df[Line]

    # drop columns that are not needed
    # we only need the columns for the transmission path, the timestamp we are querying, and the Shift Factor
    Mult = Mult.drop([col for col in Mult.columns if col not in [Line, TS, 'Shift Factor']], axis = 1)

    # we need another data frame for shift factor analysis: we will sort them differently
    SF = Mult        
    Mult = Mult.sort_values([TS]) # sorting on injections
    SF = SF.sort_values(['Shift Factor']) # sorting by shift factors

    # Part of our impact analysis is to focus on those buses / generators that are most impactful
    #   in relieving or causing congestion
    Merge = Mult.head(50).merge(SF.head(50), on = Line, how = 'outer')
    Merge1 = Mult.tail(50).merge(SF.tail(50), on = Line, how = 'outer')
    Merge = Merge.dropna()
    Merge1 = Merge1.dropna()
    Merge = pd.concat([Merge,Merge1], axis=0)

    # this analysis will relate the generators to the buses in question
    df9 = injection_df
    df9.rename(columns = {'child_name':'Node'}, inplace=True)

    print('Load gen-bus mapping', time.time() - start_time)
    mem = pd.DataFrame()
    try:
        
        for Node in Merge[Line]:
            try:
                for gen in db.GetParentMembers(CollectionEnum.GeneratorNodes, Node):
                    mem = mem.append({'Node': Node}, ignore_index=True)
                    mem = mem.append({'Generator': gen}, ignore_index=True)
            except:
                pass        
    except:
        pass
    mem['Generator'] = mem['Generator'].shift(-1)
    mem = mem.dropna()
    mem = mem.merge(df9, on = 'Node', how = 'inner')
    for col in mem.columns:
        if col != 'Node' and col != TS and col != Line and col != 'Generator':
            mem = mem.drop([col], axis = 1)
                
    # query generation to a temporary file
    print('Loading generation', time.time() - start_time)
    sol.QueryToCSV(temp_file, True, \
                    SimulationPhaseEnum.STSchedule, \
                    CollectionEnum.SystemGenerators, \
                    'System', \
                    '', \
                    PeriodEnum.Interval, \
                    SeriesTypeEnum.Periods, \
                    '1')

    '''
    for col in mem['Generator']:
        try:
            sol.QueryToCSV(temp_file, True, \
                            SimulationPhaseEnum.STSchedule, \
                            CollectionEnum.SystemGenerators, \
                            'System', \
                            col, \
                            PeriodEnum.Interval, \
                            SeriesTypeEnum.Periods, \
                            '1')
        except:
            pass
    '''
    
    # read the results into a data frame
    Gen = pd.read_csv(temp_file)
    os.remove(temp_file)
    Gen = Gen.drop(Gen.columns[[0,1,2,3,4,5,6,7,8,9,11,13,14,15,16,17,18,19]], axis = 1)
    print('Generation loaded', time.time() - start_time)

    # Drop the columns I don't need
    Gen = Gen.drop([col for col in Gen.columns if col not in [TS, 'child_name']], axis = 1)

    # rename the remaining columns
    Gen.rename(columns = {TS:'Generation'}, inplace=True)
    Gen.rename(columns = {'child_name':'Generator'}, inplace=True)

    # Compute some additional impact factors
    mem = mem.merge(Gen, on = 'Generator', how = 'outer')
    mem.rename(columns = {Line:'Shift Factor'}, inplace=True)
    mem['Generation * SF'] = mem['Generation'] * mem['Shift Factor']
    mem = mem.sort_values(['Generation * SF'])
    mem.rename(columns = {TS:'Net Injection'}, inplace=True)
    mem = mem.dropna()
    
    # write results to a spreadsheet
    print('Preparing to write results', time.time() - start_time)
    writer = pd.ExcelWriter(os.path.join(results_folder, '{}.xlsx.'.format(Line)), engine='xlsxwriter')
    Mult.to_excel(writer, sheet_name='NI x SF')
    SF.to_excel(writer, sheet_name='SF')
    mem.to_excel(writer, sheet_name='Avg Cost')
    writer.save()

    # open the spreadsheet for the user
    os.startfile(os.path.join(results_folder,'{}.xlsx'.format(Line)))

    # finish and return handles to the plexos inputs and outputs so they don't need to be reloaded
    print("--- %s seconds ---" % (time.time() - start_time))
    return db, sol


# run this script
if __name__ == '__main__':
    plexos_cxn = None
    plexos_sol = None

    plexos_db = sys.argv[1] # first argument is the plexos input db filename
    ptdf_file = sys.argv[2] # second argument is the ptdf file produced by a different script
    sol_file = sys.argv[3]  # third argument is the plexos solution file

    # arguments past the fourth should appear in triples
    #   4th: from_bus
    #   5th: to_bus
    #   6th: timestamp
    for from_bus, to_bus, timestamp in zip(sys.argv[4::3], sys.argv[5::3], sys.argv[6::3]):

        # the first time through, we haven't loaded the data, but we can hold onto handles to them
        if plexos_cxn is None or ptdf_data is None:
            plexos_cxn, plexos_sol = main(from_bus, to_bus, timestamp, ptdf_file, plexos_db=plexos_db, plexos_sol=sol_file)

        # the second and subsequent times through, we already have loaded the data.
        else:
            main(from_bus, to_bus, timestamp, ptdf_file, db = plexos_cxn, sol = plexos_sol)



