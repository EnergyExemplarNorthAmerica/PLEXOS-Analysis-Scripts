
import pandas as pd
import numpy as np
import io, re, os, sys, time, clr, json

from System import Enum, Boolean, DateTime
# load PLEXOS assemblies:
#set_plexos_path('C:/Program Files/Energy Exemplar/PLEXOS 9.0 API')
sys.path.append('C:/Program Files/Energy Exemplar/PLEXOS 9.0 API')
clr.AddReference('PLEXOS_NET.Core')
clr.AddReference('EEUTILITY')
clr.AddReference('EEDataSets')
clr.AddReference('EnergyExemplar.PLEXOS.Utility')
from PLEXOS_NET.Core import DatabaseCore, Solution, PLEXOSConnect

from EnergyExemplar.PLEXOS.Utility.Enums import PeriodEnum, SeriesTypeEnum
from PLEXOS_NET.Core import DatabaseCore, Solution, PLEXOSConnect
from EEUTILITY.Enums  import SimulationPhaseEnum, CollectionEnum, ClassEnum
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

def query_data_to_csv(sol, csv_file, sim_phase, coll, period, date_from, date_to, is_verbose = False, property_list = ''):
    # Run the query
    params = (csv_file, True, sim_phase, coll, '', '', period, SeriesTypeEnum.Values, property_list, date_from, date_to)
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
    config_json = switch_data(args, '-c')
    if not config_json is None and os.path.exists(config_json):
        cfg_json_obj = json.load(open(config_json))
        for query in cfg_json_obj['queries']:
            try:
                phase = Enum.Parse(SimulationPhaseEnum, query['phase'])
                coll = sol_cxn.CollectionName2Id(query['parentclass'], query['childclass'], query['collection'])
            except:
                print("Phase:", query['phase'], "Collection:", query['collection'], "Parent Class:", query['parentclass'], "Child Class:",  query['childclass'])
                print(" --> This combination doesn't identify queryable information")
                continue # If the phase or collection are missing or incorrect just skip

            if 'properties' in query:
                # the data in properties field may be a list of property names; we'll pull each individually
                if type(query['properties']) is list:
                    for prop in query['properties']:
                        try:
                            prop_id = str(sol_cxn.PropertyName2EnumId(query['parentclass'], query['childclass'], query['collection'], prop))
                            query_data_to_csv(sol_cxn, csv_file, phase, coll, time_res, date_from, date_to, is_verbose=is_switch(args, '-v'), property_list=prop_id)
                        except:
                            pass
                
                # the data in properties field may be a single property name; we'll just pull it
                elif type(query['properties']) is str:
                    try:
                        prop_id = str(sol_cxn.PropertyName2EnumId(query['parentclass'], query['childclass'], query['collection'], query['properties']))
                        query_data_to_csv(sol_cxn, csv_file, phase, coll, time_res, date_from, date_to, is_verbose=is_switch(args, '-v'), property_list=prop_id)
                    except:
                        pass
                
                # the data in properties field may be poorly formatted
                else:
                    print(query['properties'],'is not valid property information')

            # properties field may be missing; just pull all properties
            else:
                query_data_to_csv(sol_cxn, csv_file, phase, coll, time_res, date_from, date_to, is_verbose=is_switch(args, '-v'))
    else:
        for phase in Enum.GetValues(clr.GetClrType(SimulationPhaseEnum)):
            for coll in Enum.GetValues(clr.GetClrType(CollectionEnum)):
                query_data_to_csv(sol_cxn, csv_file, phase, coll, time_res, date_from, date_to, is_verbose=is_switch(args, '-v'))

    print('Completed',clr.GetClrType(PeriodEnum).GetEnumName(time_res),'in',time.time() - start_time,'sec')

def none_to_empty_list(ret):
    if ret is None:
        return []
    else:
        return ret

def pull_xref(sol_cxn, xref_file):
    #from EEDataSets import t_membershipDataTable, t_objectDataTable, t_classDataTable, t_categoryDataTable, t_collectionDataTable

    # retrieve memberships
    mem_df = pd.DataFrame(columns = ['pclass', 'cclass', 'coll', 'pobj', 'cobj'])
    tbl = sol_cxn.GetDataTable('t_membership', '')
    for row in tbl if 'EEDataSets' in str(type(tbl)) else tbl[0]:
        mem_df.loc[row.membership_id] = [row.parent_class_id, row.child_class_id, row.collection_id, row.parent_object_id, row.child_object_id]

    # retrieve objects
    obj_df = pd.DataFrame(columns = ['oname', 'ocat'])
    tbl = sol_cxn.GetDataTable('t_object', '')
    for row in tbl if 'EEDataSets' in str(type(tbl)) else tbl[0]:
        if row.show:
            obj_df.loc[row.object_id] = [row.name, row.category_id]
    
    # retrieve classes
    cls_df = pd.DataFrame(columns = ['class_name'])
    tbl = sol_cxn.GetDataTable('t_class', '')
    for row in tbl if 'EEDataSets' in str(type(tbl)) else tbl[0]:
        cls_df.loc[row.class_id] = [row.name]

    # retrieve categories
    cat_df = pd.DataFrame(columns = ['category'])
    tbl = sol_cxn.GetDataTable('t_category', '')
    for row in tbl if 'EEDataSets' in str(type(tbl)) else tbl[0]:
        cat_df.loc[row.category_id] = [row.name]

    # retrieve collections
    coll_df = pd.DataFrame(columns = ['collection'])
    tbl = sol_cxn.GetDataTable('t_collection', '')
    for row in tbl if 'EEDataSets' in str(type(tbl)) else tbl[0]:
        coll_df.loc[row.collection_id] = [row.name]

    # Join the metadata into a single xref table

    # join parent and child object names
    mem_df = mem_df.join(obj_df, on='cobj', rsuffix='_child')
    mem_df = mem_df.join(obj_df, on='pobj', rsuffix='_parent')

    # join parent and child class names
    mem_df = mem_df.join(cls_df, on='cclass', rsuffix = '_child')
    mem_df = mem_df.join(cls_df, on='pclass', rsuffix = '_parent')

    # join parent and child categories
    mem_df = mem_df.join(cat_df, on='ocat', rsuffix = '_child')
    mem_df = mem_df.join(cat_df, on='ocat_parent', rsuffix = '_parent')

    # join collections
    mem_df = mem_df.join(coll_df, on='coll', rsuffix = '_parent')

    # reformat the dataframe to look nicer
    mem_df = mem_df.drop(columns = ['cobj', 'pobj', 'cclass', 'pclass', 'ocat', 'ocat_parent', 'coll'])
    mem_df = mem_df[['class_name_parent', 'oname_parent', 'category_parent', 'collection', 'class_name','oname','category']]
    mem_df.columns = ['ParentClass', 'ParentName', 'ParentCategory', 'Collection', 'ChildClass', 'ChildName', 'ChildCategory']
    
    # write the xref file
    mem_df.to_csv(xref_file)


def main():

    
    if len(sys.argv) <= 1:
        print('''
Usage:
    python power_bi_link.py <solution_file> [-x [xref_file]]
                                            [-c [config_json]]
                                            [-y [yr_file]]
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
    if is_switch(sys.argv, '-x'): pull_xref(sol_cxn, switch_data(sys.argv, '-x'))

    sol_cxn.Close()

    print ('Completed in', time.time() - start_time, 'sec')


#def set_plexos_path(plexos_path):
#    pass


if __name__ == '__main__':

    
    main()
