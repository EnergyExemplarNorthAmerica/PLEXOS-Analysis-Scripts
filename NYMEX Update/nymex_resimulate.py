# for PLEXOS API
import os, sys, clr
import subprocess as sp

# get the PLEXOS Power BI Link
sys.path.append('../Power BI Link')
import power_bi_link as pbi

# for string parsing
import re

# for datetime manipulation
import calendar as cal
from datetime import date, timedelta

# for web scraping
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup

def add_month(dt):
    dt += timedelta(days = cal.monthrange(dt.year, dt.month)[1])
    return dt 

def setup_date_mapping():
    date_from = dict()
    #dt = DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)
    dt = DateTime(2024, DateTime.Today.Month, 1)
    date_from['HenryHub'] = dt
    date_from['Contract1'] = dt.AddMonths(1)
    date_from['Contract2'] = dt.AddMonths(2)
    date_from['Contract3'] = dt.AddMonths(3)
    date_from['Contract4'] = dt.AddMonths(4)
    return date_from

def pull_nymex_hh():
    url = 'https://www.eia.gov/dnav/ng/ng_pri_fut_s1_d.htm'
    req = Request(url, headers={'User-Agent' : "Magic Browser"}) 

    with urlopen( req ) as con:
        soup = BeautifulSoup(con, features='html.parser')

    gas_prices = dict()
    for row in soup.body.find('table', attrs={'class':'data1'}).find_all('tr', attrs={'class':'DataRow'}):
        try:
            header = re.sub(r'\s','',row.find('td',attrs={'class':'DataStub'}).text)
            data = re.sub(r'\s','',row.find('td', attrs={'class':'Current2'}).text)
            gas_prices[header] = float(data)
        except:
            continue

    return gas_prices

def parse_date(string):
    try:
        return DateTime.Parse(string)
    except:
        return None

def is_object(db, class_id, object_name):
    try:
        '''
        Int32 ObjectName2Id(
            ClassEnum ClassId,
            String Name
            )
        '''
        return db.ObjectName2Id(class_id, object_name) > 0
    except:
        return False

def create_object(db, class_id, object_name, strCategory = None, strDescription = None):
    '''
    Int32 AddObject(
        String strName,
        ClassEnum nClassId,
        Boolean bAddSystemMembership,
        String strCategory[ = None],
        String strDescription[ = None]
        )
    '''
    if not is_object(db, class_id, object_name):
        db.AddObject(object_name, class_id, True, strCategory, strDescription)
        return True
    else:
        return False

def is_membership(db, collection, parent, child):
    '''
    Int32 GetMembershipID(
        CollectionEnum nCollectionId,
        String strParent,
        String strChild
        )
    '''
    try:
        return db.GetMembershipID(collection, parent, child) > 0
    except:
        return False

def create_membership(db, collection, parent_class_id, parent, child_class_id, child, one_to_one = False):
    '''
    Int32 AddMembership(
        CollectionEnum nCollectionId,
        String strParent,
        String strChild
        )
    '''

    # convert the child to a list
    if type(child) is str:
        child = [child]

    # if the membership is 1-1, we need to remove any existing items
    if one_to_one:
        child = [child[0]]
        try:
            existing_list = db.GetChildMembers(collection, parent)
            existing_list = [] if existing_list is None else existing_list
            for c in existing_list:
                db.RemoveMembership(collection, parent, c)        
        except:
            pass

    # create the membership if required
    for c in child:
        if not is_membership(db, collection, parent, c):

            # make sure the objects exist
            create_object(db, parent_class_id, parent)
            create_object(db, child_class_id, c)

            # add the membership
            db.AddMembership(collection, parent, c)
        
    return True

def update_or_add_property(db, value, **params):

    # get membership id ... this is for a particular object
    mem_id = db.GetMembershipID(params['nCollectionId'], params['strParent'], params['strChild'])

    # add the scenario if it doesn't exist
    create_object(db, ClassEnum.Scenario, params['Scenario'])

    # delete the property row if it exists
    db.RemoveProperty(mem_id, params['EnumId'], params['BandId'], \
            params['DateFrom'], params['DateTo'], params['Variable'], \
            params['DataFile'], params['Pattern'], params['Scenario'], \
            params['Action'], PeriodEnum.Interval)

    # add a new property row
    db.AddProperty(mem_id, params['EnumId'], params['BandId'], \
            value, params['DateFrom'], params['DateTo'], params['Variable'], \
            params['DataFile'], params['Pattern'], params['Scenario'], \
            params['Action'], PeriodEnum.Interval)

def plexos_update_prices(db, gas_fuels, gas_scenario, prices = None):
    if prices is None: return False

    dates = setup_date_mapping()
    for k in prices:
        for f in gas_fuels:
            update_or_add_property(db, prices[k], \
                nCollectionId = CollectionEnum.SystemFuels, \
                strParent = 'System', strChild = f, \
                EnumId = SystemFuelsEnum.Price, \
                BandId = 1, \
                DateFrom = dates[k], \
                DateTo = None, \
                Variable = None, \
                DataFile = None, \
                Pattern = None, \
                Scenario = gas_scenario, \
                Action = '=')

    
    return True

def adjust_study_horizon(db, horizon_name, start_date, months):
    '''
    Update the Horizon object to run a number of days from the start_date
    '''
    db.UpdateAttribute(ClassEnum.Horizon, horizon_name, HorizonAttributeEnum.DateFrom, start_date.ToOADate())
    db.UpdateAttribute(ClassEnum.Horizon, horizon_name, HorizonAttributeEnum.StepType, 3.0)
    db.UpdateAttribute(ClassEnum.Horizon, horizon_name, HorizonAttributeEnum.StepCount, float(months))
    db.UpdateAttribute(ClassEnum.Horizon, horizon_name, HorizonAttributeEnum.ChronoDateFrom, start_date.ToOADate())
    db.UpdateAttribute(ClassEnum.Horizon, horizon_name, HorizonAttributeEnum.ChronoStepType, 2.0)
    db.UpdateAttribute(ClassEnum.Horizon, horizon_name, HorizonAttributeEnum.ChronoStepCount, start_date.AddMonths(months).ToOADate()-start_date.ToOADate())

def add_basic_outputs(db, report_object):

    # Check Daily and Monthly output
    db.UpdateAttribute(ClassEnum.Report, report_object, ReportAttributeEnum.OutputResultsbyDay, -1.0)
    db.UpdateAttribute(ClassEnum.Report, report_object, ReportAttributeEnum.OutputResultsbyMonth, -1.0)

    # a list of the report fields to put in output
    my_list = [\
        db.ReportPropertyName2PropertyId('System','Generator','Generators','Generation'),
        db.ReportPropertyName2PropertyId('System','Generator','Generators','Total Generation Cost'),
        db.ReportPropertyName2PropertyId('System','Generator','Generators','Fuel Offtake'),
        db.ReportPropertyName2PropertyId('System','Fuel','Fuels','Offtake'),
        db.ReportPropertyName2PropertyId('System','Fuel','Fuels','Price'),
        db.ReportPropertyName2PropertyId('System','Region','Regions','Price'),
        db.ReportPropertyName2PropertyId('System','Region','Regions','Total System Cost'),
        db.ReportPropertyName2PropertyId('System','Region','Regions','Load'),
        db.ReportPropertyName2PropertyId('System','Region','Regions','Generation')
        ]
    
    # enable summary reporting for all of the above
    for idx in my_list:
        db.AddReportProperty(report_object, idx, SimulationPhaseEnum.MTSchedule, False, True, False, False)

def plexos_update_project(db, project_name, base_scenarios, gas_scenario):
    model1 = project_name + '_Base'
    model2 = '{}_{}'.format(project_name, gas_scenario)

    # link up the project
    db.RemoveObject(gas_scenario, ClassEnum.Report)
    db.Close()
    db.Connection(db.DataSource)
    create_membership(db, CollectionEnum.ProjectModels, ClassEnum.Project, project_name, ClassEnum.Model, model1)
    create_membership(db, CollectionEnum.ProjectHorizon, ClassEnum.Project, project_name, ClassEnum.Horizon, gas_scenario, one_to_one=True)
    create_membership(db, CollectionEnum.ProjectReport, ClassEnum.Project, project_name, ClassEnum.Report, gas_scenario, one_to_one=True)

    # link up base case
    create_membership(db, CollectionEnum.ModelHorizon, ClassEnum.Model, model1, ClassEnum.Horizon, gas_scenario, one_to_one=True)
    create_membership(db, CollectionEnum.ModelReport, ClassEnum.Model, model1, ClassEnum.Report, gas_scenario, one_to_one=True)
    create_membership(db, CollectionEnum.ModelMTSchedule, ClassEnum.Model, model1, ClassEnum.MTSchedule, gas_scenario, one_to_one=True)
    create_membership(db, CollectionEnum.ModelScenarios, ClassEnum.Model, model1, ClassEnum.Scenario, base_scenarios)

    # link up update case
    create_membership(db, CollectionEnum.ProjectModels, ClassEnum.Project, project_name, ClassEnum.Model, model2)
    create_membership(db, CollectionEnum.ModelHorizon, ClassEnum.Model, model2, ClassEnum.Horizon, gas_scenario, one_to_one=True)
    create_membership(db, CollectionEnum.ModelReport, ClassEnum.Model, model2, ClassEnum.Report, gas_scenario, one_to_one=True)
    create_membership(db, CollectionEnum.ModelMTSchedule, ClassEnum.Model, model2, ClassEnum.MTSchedule, gas_scenario, one_to_one=True)
    create_membership(db, CollectionEnum.ModelScenarios, ClassEnum.Model, model2, ClassEnum.Scenario, base_scenarios + [gas_scenario])

    # setup the horizon
    dt = setup_date_mapping()['HenryHub']
    adjust_study_horizon(db, gas_scenario, dt, 5.0)

    # setup the report object
    add_basic_outputs(db, gas_scenario)

    return True

def plexos_update(plexos_db, fuel_objects, base_scenarios, gas_scenario, project_name, prices = None):

    # Connect to the PLEXOS input dataset
    cxn = DatabaseCore()
    cxn.DisplayAlerts = False
    cxn.Connection(plexos_db)
    cxn.DataSource = plexos_db

    # price updates
    if plexos_update_prices(cxn, fuel_objects, gas_scenario, prices):
        print('Updated PLEXOS gas prices')
    else:
        print('No update for PLEXOS gas prices')

    # update project and models
    if plexos_update_project(cxn, project_name, base_scenarios, gas_scenario):
        print('Updated PLEXOS project and models')
    else:
        print('No update to PLEXOS project and models')

    # close and save
    cxn.Close()

def plexos_launch(plexos_db, project_name):
    # launch the model on the local desktop
    # The \n argument is very important because it allows the PLEXOS
    # engine to terminate after completing the simulation
    sp.call([os.path.join(DatabaseCore().InstallPath, 'PLEXOS64.exe'), plexos_db, r'\n', r'\p', project_name])

def plexos_process(plexos_db, project_name):
    sol_file = os.path.join(os.path.dirname(plexos_db), 'Project {} Solution'.format(project_name), 'Project {} Solution.zip'.format(project_name))
    sol_cxn = Solution()
    sol_cxn.Connection(sol_file)
    sol_cxn.DisplayAlerts = False

    # pull data
    pbi.pull_data(sol_cxn, PeriodEnum.Month, ['', sol_file, '-m'], '-m', '_monthly.csv') # pull month results

def main(plexos_db, fuel_objects, base_scenarios, gas_scenario, project_name):

    plexos_update(plexos_db, fuel_objects, base_scenarios, gas_scenario, project_name, pull_nymex_hh())
    plexos_launch(plexos_db, project_name)
    plexos_process(plexos_db, project_name)

def parse_cli():
    cli = dict()
    last_switch = None
    for arg in sys.argv[1:]:
        if arg[0] == '-':
            cli[arg] = None
            last_switch = arg
        elif last_switch is not None:
            cli[last_switch] = arg
        else:
            cli['-'] = arg

    return cli

if __name__ == '__main__':
    
    plexos_path = 'C:/Program Files/Energy Exemplar/PLEXOS 8.2/'
    sys.path.append(plexos_path)
    clr.AddReference('PLEXOS7_NET.Core')
    clr.AddReference('EEUTILITY')

    from PLEXOS7_NET.Core import DatabaseCore, Solution, PLEXOSConnect
    from EEUTILITY.Enums import *
    from System import DateTime
    import power_bi_link as pbi
    pbi.set_plexos_path(plexos_path)

    cli = parse_cli()
    plexos_db = cli['-d'] if '-d' in cli else 'rtsDEMO/rts_PLEXOS.xml'
    gas_scenario = cli['-s'] if '-s' in cli else 'NYMEX'
    fuel_objects = cli['-g'].split(',') if '-g' in cli else ['NG/CC', 'NG/CT']
    project_name = cli['-p'] if '-p' in cli else '3Month'
    base_scenarios = cli['b'].split(',') if '-b' in cli else ["Gen Outages","Load: DA","RE: DA","Add Spin Up"]

    main(plexos_db, fuel_objects, base_scenarios, gas_scenario, project_name)
    