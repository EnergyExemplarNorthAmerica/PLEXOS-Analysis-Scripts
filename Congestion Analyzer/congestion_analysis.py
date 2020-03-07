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

    start_time = time.time()
    Line = '{}_{}'.format(FB,TB)
    temp_file = '{}_sol.csv'.format(Line)
    results_folder = './Results {}'.format(Line)
    os.makedirs(results_folder, exist_ok=True)

    if db is None:
        db = DatabaseCore()

        # Enter the path for your Database:
        db.Connection(plexos_db)
        print('Loaded PLEXOS input:', time.time() - start_time)

    if sol is None:
        sol = Solution()
        sol.Connection(plexos_sol)
        print('Loaded PLEXOS solution:', time.time() - start_time)

    with open(ptdf_data) as infile:
        idxs = [idx for idx, from_bus, to_bus in zip(infile.readline().split(','),infile.readline().split(','),infile.readline().split(',')) if FB in from_bus and TB in to_bus]
        print('Computed required PTDF columns:', time.time() - start_time)

    df = pd.read_csv(ptdf_data, usecols=['0'] + idxs, low_memory=False, skiprows=[1,2])
    df.rename(columns = {'0':'child_name'}, inplace=True)

    print('Loaded PTDF data:', time.time() - start_time)

    '''
    Data.iloc[[0,1]] = Data.iloc[[0,1]].applymap(str)

    for idx1, column in enumerate(Data.iloc[0]):
        
        for idx2, column1 in enumerate(Data.iloc[1]):
            if ((column[:len(FB)] == FB) and (column1[:len(TB)] == TB) and (idx1 == idx2)):
                df = Data.iloc[:, idx1]
                df = pd.concat([Data.iloc[:,1], df], axis=1)  
                break
            else:
                continue

        continue
        break
    '''

    try:
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
        df8 = pd.read_csv(temp_file)
        os.remove(temp_file)
        df8 = df8.drop(df8.columns[[0,1,2,3,4,5,6,7,8,9,11,13,14,15,16,17,18,19]], axis = 1)
        print('Queried Net Injections:', time.time() - start_time)
        
    except:
        pass

    writer = pd.ExcelWriter(os.path.join(results_folder, '{}.xlsx.'.format(Line)), engine='xlsxwriter')
    df1 = pd.DataFrame(columns=['child_name',Line])
    for col in [col for col in df if not col == 'index']:
        df1['child_name'] = df['child_name']
        df1[Line] = df[col]


    df['child_name'] = df.applymap(str)
    df8['child_name'] = df8.applymap(str)

    df = df.merge(df8, on='child_name')

    df8 = df8.merge(df1, on='child_name')

    Mult = pd.DataFrame()
    SF = pd.DataFrame()

    Mult[Line] = df8.iloc[:,0]

    for z in range (2, len(df8.columns)-1):
        row = df8.columns[z]
        df8[row] = df8[row].astype(float)
        df8[Line] = df8[Line].astype(float)
        df8 = df8.dropna()
        Mult[row] = df8[row] * df8[Line]
        

    Mult['Shift Factor'] = df8[Line]

    # Taking the Timestamp of interest from the user:
    # print('The First Time Stamp is :{}, The Last Time Stamp is :{}.'.format(Mult.columns[1],Mult.columns[len(Mult.columns)-2]))
    #Timestamp = str(input("Enter the Time Stamp with the Same Format Above :"))

    for col in Mult.columns:
        if col != Line and col != TS and col != 'Shift Factor':
            Mult = Mult.drop([col], axis = 1)
            
    SF = Mult        

    Mult = Mult.sort_values([TS])
    SF = SF.sort_values(['Shift Factor'])

    Merge = Mult.head(50).merge(SF.head(50), on = Line, how = 'outer')
    Merge1 = Mult.tail(50).merge(SF.tail(50), on = Line, how = 'outer')
    Merge = Merge.dropna()
    Merge1 = Merge1.dropna()
    Merge = pd.concat([Merge,Merge1], axis=0)

    df9 = df8
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
    
    print('Generation loaded', time.time() - start_time)

    Gen = pd.read_csv(temp_file)
    os.remove(temp_file)
    Gen = Gen.drop(Gen.columns[[0,1,2,3,4,5,6,7,8,9,11,13,14,15,16,17,18,19]], axis = 1)

    for col in Gen.columns:
        if col != TS and col != 'child_name':
            Gen = Gen.drop([col], axis = 1)

    Gen.rename(columns = {TS:'Generation'}, inplace=True)
    Gen.rename(columns = {'child_name':'Generator'}, inplace=True)

    mem = mem.merge(Gen, on = 'Generator', how = 'outer')
    mem.rename(columns = {Line:'Shift Factor'}, inplace=True)
    mem['Generation * SF'] = mem['Generation'] * mem['Shift Factor']
    mem = mem.sort_values(['Generation * SF'])
    mem.rename(columns = {TS:'Net Injection'}, inplace=True)
    mem = mem.dropna()
    
    print('Preparing to write results', time.time() - start_time)
    Mult.to_excel(writer, sheet_name='NI x SF')
    SF.to_excel(writer, sheet_name='SF')
    mem.to_excel(writer, sheet_name='Avg Cost')
    writer.save()
    os.startfile(os.path.join(results_folder,'{}.xlsx'.format(Line)))

    print("--- %s seconds ---" % (time.time() - start_time))

    return db, sol

if __name__ == '__main__':
    plexos_cxn = None
    plexos_sol = None

    plexos_db = sys.argv[1]
    ptdf_file = sys.argv[2]
    sol_file = sys.argv[3]

    for from_bus, to_bus, timestamp in zip(sys.argv[4::3], sys.argv[5::3], sys.argv[6::3]):
        if plexos_cxn is None or ptdf_data is None:
            plexos_cxn, plexos_sol = main(from_bus, to_bus, timestamp, ptdf_file, plexos_db=plexos_db, plexos_sol=sol_file)
        else:
            main(from_bus, to_bus, timestamp, ptdf_file, db = plexos_cxn, sol = plexos_sol)



