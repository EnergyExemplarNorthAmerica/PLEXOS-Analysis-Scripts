import glob
import sys
import pandas as pd
import os
import clr
import time

sys.path += ['C:/Program Files/Energy Exemplar/PLEXOS 8.2/']
clr.AddReference('PLEXOS7_NET.Core')
clr.AddReference('EEUTILITY')

# Import from .NET assemblies (both PLEXOS and system)
from PLEXOS7_NET.Core import *
from EEUTILITY.Enums import *
import System


def init_connection(db_file, solution_file):
    # initialize database connections
    db = DatabaseCore()
    db.Connection(db_file)
    print('Loaded PLEXOS input:', time.time() - start_time)

    sol = Solution()
    sol.Connection(solution_file)
    print('Loaded PLEXOS solution:', time.time() - start_time)
    return db, sol


def ptdf_generator(in_ptdf_file):
    print('Start PTDF:', time.time() - start_time)
    sol_ptdf = pd.read_csv(in_ptdf_file, sep='\t', low_memory=False)
    print('Finished Reading PTDF:', time.time() - start_time)
    sol_ptdf = sol_ptdf.T
    print('Finished Transpose:', time.time() - start_time)
    sol_ptdf = sol_ptdf.drop(sol_ptdf.index[[0]])
    sol_ptdf.reset_index(inplace=True)
    print('Completed PTDF:', time.time() - start_time)
    return sol_ptdf


def select_lines(db, sol, n_lines=5, db_file=None, sol_file=None):
    if (db is None) or (sol is None):
        db, sol = init_connection(db_file, sol_file)

    results = sol.Query(SimulationPhaseEnum.STSchedule, CollectionEnum.SystemLines, '', '', PeriodEnum.Interval,
                        SeriesTypeEnum.Periods, '1, 8, 10')  # Querying Flow, Shadow Price, and Congested Hour
    if results.EOF:
        print('No results')
        exit
    resultsRows = results.GetRows()
    sol_lines_data = pd.DataFrame(
        [[resultsRows[i, j] for i in range(resultsRows.GetLength(0))] for j in range(resultsRows.GetLength(1))],
        columns=[x.Name for x in results.Fields])
    sol_lines_data.drop(sol_lines_data.columns[[col for col in [*range(20)] if col not in (10, 12)]], axis=1, inplace=True)
    sol_lines_data = sol_lines_data.groupby("child_name").prod().abs()
    sol_lines_data['Max'], sol_lines_data['Max_Hour'] = sol_lines_data.max(axis=1), sol_lines_data.idxmax(axis=1)

    sol_lines_data.sort_values('Max', ascending=False, inplace=True)
    selected_lines = []
    for line in sol_lines_data.index.values:
        rt = db.GetPropertiesTable(CollectionEnum.SystemLines, "", line)
        resultsRows = rt.GetRows()
        property_table = pd.DataFrame(
            [[resultsRows[i, j] for i in range(resultsRows.GetLength(0))] for j in range(resultsRows.GetLength(1))],
            columns=[x.Name for x in rt.Fields])
        line_reactance = property_table[property_table['Property'] == 'Reactance']['Value']
        if not line_reactance.empty and float(line_reactance.iloc[0]) > 0:
            selected_lines.append(line)
        if len(selected_lines) >= n_lines:
            break

    selected_lines = pd.DataFrame(selected_lines, columns=['child_name']).merge(sol_lines_data['Max_Hour'], on='child_name')
    selected_lines['FromNode'] = selected_lines.apply(lambda row: db.GetChildMembers(CollectionEnum.LineNodeFrom, row['child_name'])[0], axis=1)
    selected_lines['ToNode'] = selected_lines.apply(lambda row: db.GetChildMembers(CollectionEnum.LineNodeTo, row['child_name'])[0], axis=1)
    selected_lines = selected_lines[['FromNode', 'ToNode', 'Max_Hour']]

    return list(selected_lines.values.flatten())


def congestion_analyzer(db, sol, FB, TB, TS, ptdf_data, db_file=None, sol_file=None):
    if (db is None) or (sol is None):
        db, sol = init_connection(db_file, sol_file)

    target_line = '{}_{}'.format(FB, TB)
    temp_file = '{}_sol.csv'.format(target_line)
    results_folder = './Results {}'.format(target_line)
    os.makedirs(results_folder, exist_ok=True)

    idxs = [idx for idx, from_bus, to_bus in
            zip(list(ptdf_data.columns), list(ptdf_data.iloc[0]), list(ptdf_data.iloc[1])) if
            (FB in from_bus and TB in to_bus) or (FB in to_bus and TB in from_bus)]
    print('Computed required PTDF columns:', time.time() - start_time)

    ptdf_df = ptdf_data.loc[:, [0] + idxs].iloc[2:].reset_index(drop=True)
    ptdf_df.rename(columns={0: 'child_name'}, inplace=True)
    print('Loaded PTDF data:', time.time() - start_time)

    sol.QueryToCSV(temp_file, False, SimulationPhaseEnum.STSchedule, CollectionEnum.SystemNodes, '', '',
                   PeriodEnum.Interval, SeriesTypeEnum.Periods, '22', System.DateTime.Parse(TS),
                   System.DateTime.Parse(TS).AddHours(1), '', '', '', 0, '', ',')
    injection_df = pd.read_csv(temp_file)
    os.remove(temp_file)

    injection_df.drop(injection_df.columns[[col for col in [*range(20)] if col not in (10, 12)]], axis=1, inplace=True)
    print('Queried Net Injections:', time.time() - start_time)

    # join injections to ptdfs
    temp_injection_df = pd.DataFrame(columns=['child_name', target_line])
    temp_injection_df['child_name'] = ptdf_df['child_name']
    temp_injection_df[target_line] = ptdf_df.iloc[:, 1]
    ptdf_df['child_name'] = ptdf_df.applymap(str)
    injection_df['child_name'] = injection_df.applymap(str)
    injection_df = injection_df.merge(temp_injection_df, on='child_name')

    # begin calculations of shift factors and impacts
    injection_df = injection_df.dropna()
    injection_df[target_line] = injection_df[target_line].astype(float)
    mult = injection_df.iloc[:, 2:-1].astype(float).multiply(injection_df[target_line], axis='index')
    mult.insert(0, target_line, injection_df.iloc[:, 0])
    mult['Shift Factor'] = injection_df[target_line]

    # drop columns that are not needed
    # we only need the columns for the transmission path, the timestamp we are querying, and the Shift Factor
    mult = mult[[target_line, TS, 'Shift Factor']]
    # we need another data frame for shift factor analysis: we will sort them differently
    sf = mult
    mult.to_csv(os.path.join(results_folder, 'NI x SF.csv'), index=False)
    mult = mult.sort_values([TS])  # sorting on injections
    sf = sf.sort_values(['Shift Factor'])  # sorting by shift factors

    # Part of our impact analysis is to focus on those buses / generators that are most impactful
    #   in relieving or causing congestion
    mult['impact'] = abs(mult[TS]) * abs(mult['Shift Factor'])
    merge = mult.nlargest(100, columns=['impact'])

    # this analysis will relate the generators to the buses in question
    df9 = injection_df
    df9.rename(columns={'child_name': 'Node'}, inplace=True)
    print('Load gen-bus mapping:', time.time() - start_time)
    mem = pd.DataFrame()

    try:
        for Node in merge[target_line]:
            node_generators = db.GetParentMembers(CollectionEnum.GeneratorNodes, Node)
            if node_generators is not None:
                for gen in node_generators:
                    mem = mem.append({"Node": Node, "Generator": gen}, ignore_index=True)
    except:
        pass

    mem = mem.merge(df9, on='Node', how='inner')
    mem = mem[['Node', TS, target_line, 'Generator']]

    # query genreation to a temporary file
    print('Loading generation:', time.time() - start_time)
    sol.QueryToCSV(temp_file, False, SimulationPhaseEnum.STSchedule, CollectionEnum.SystemGenerators, '', '',
                   PeriodEnum.Interval, SeriesTypeEnum.Periods, '1', System.DateTime.Parse(TS),
                   System.DateTime.Parse(TS).AddHours(1), '', '', '', 0, '', ',')
    print('Generation loaded:', time.time() - start_time)
    gen = pd.read_csv(temp_file)
    os.remove(temp_file)
    gen.drop(gen.columns[[col for col in [*range(20)] if col not in (10, 12)]], axis=1, inplace=True)
    gen = gen[['child_name', TS]]
    gen.rename(columns={"child_name": "Generator", TS: "Generation"}, inplace=True)

    try:
        lines = pd.DataFrame([line for line in db.GetObjects(ClassEnum.Line)], columns=['Lines'])
    except:
        pass
    line1 = lines[lines['Lines'].str.contains(target_line)]
    line1 = line1.values[0]
    print('Querying Line Shadow Price:', time.time() - start_time)
    sol.QueryToCSV(temp_file, False, SimulationPhaseEnum.STSchedule, CollectionEnum.SystemLines, '', '',
                   PeriodEnum.Interval, SeriesTypeEnum.Periods, '10', System.DateTime.Parse(TS),
                   System.DateTime.Parse(TS).AddHours(1), '', '', '', 0, '', ',')

    data = pd.read_csv(temp_file)
    os.remove(temp_file)
    sh = data[(data['child_name'] == line1[0])]
    mem = mem.merge(gen, on='Generator', how='outer')
    mem.rename(columns={target_line: 'Shift Factor'}, inplace=True)
    mem['Generation * SF'] = mem['Generation'] * mem['Shift Factor']
    mem = mem.sort_values(['Generation * SF'])
    mem['Line Shadow Price * SF'] = mem['Shift Factor'] * sh[TS].values[0]
    mem.rename(columns={TS: 'Net Injection'}, inplace=True)
    mem = mem.dropna()

    print('Preparing to write results:', time.time() - start_time)
    writer = pd.ExcelWriter(os.path.join(results_folder, '{}.xlsx.'.format(target_line)), engine='xlsxwriter')
    mult.to_excel(writer, sheet_name='NI x SF')
    sf.to_excel(writer, sheet_name='SF')
    mem.to_excel(writer, sheet_name='Flow Contribution')
    writer.save()
    # os.startfile(os.path.join(results_folder, '{}.xlsx'.format(target_line)))
    print("--- Completed at %s seconds ---" % (time.time() - start_time))


# Main Function for Congestion Analyzer
# Parameters: (1) input DB xml, (2) Solution Directory (3) Solution Zip File, rest should be in triplets (FB/TB/TS)
# PLEXOS Post-Simulation.bat should give:
# For (1):  %DATASET_PATH%
# For (2):  %SOLUTION_0%
# For (3):  Pushd %SOLUTION_0%
#           for %%i in (*.zip) do set SolutionName=%SOLUTION_0%%%~ni%%~xi
#           %SolutionName%
if __name__ == '__main__':
    if len(sys.argv) < 4:
        raise KeyError("Insufficient parameters, input database, study location, and solution file must be specified.")

    if not os.path.isfile(sys.argv[1]):
        raise FileNotFoundError("PLEXOS database not found.")
    plexos_db = sys.argv[1]

    if not os.path.isdir(sys.argv[2]):
        raise KeyError("Invalid parameter, solution location does not exist.")
    solution_folder = sys.argv[2]

    if not os.path.isfile((sys.argv[3])):
        raise FileNotFoundError("PLEXOS solution not found.")
    plexos_solution = sys.argv[3]

    raw_ptdf = glob.glob(solution_folder + '/*PTDF Diagnostics*.txt')
    if not raw_ptdf:
        raise FileNotFoundError("No PTDF diagnostic file found, ensure diagnostic object is enabled.")
    raw_ptdf = raw_ptdf[0]

    start_time = time.time()
    db_cxn, sol_cxn = init_connection(plexos_db, plexos_solution)  # Initialize the connection to PLEXOS

    if len(sys.argv) > 4:
        if (len(sys.argv) - 1) % 3 == 0:
            lines_to_analyze = sys.argv[4:]
        else:
            raise ValueError("User-specified parameters not in groups of 3s")
    else:
        lines_to_analyze = select_lines(db_cxn, sol_cxn, 3)  # Run Line Selector
    print('Selected Lines:', time.time() - start_time)

    processed_ptdf = ptdf_generator(raw_ptdf)  # Run PTDF Generator

    for from_bus, to_bus, timestamp in zip(lines_to_analyze[0::3], lines_to_analyze[1::3], lines_to_analyze[2::3]):
        print("--- Working on %s - %s - %s ---" % (from_bus, to_bus, timestamp))
        congestion_analyzer(db_cxn, sol_cxn, from_bus, to_bus, timestamp, processed_ptdf)  # Run Congestion Analyzer

    print("Exited main program")
