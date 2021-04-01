# standard Python/SciPy libraries
print("Loading Python Packages...")
import time

time.sleep(2)
print("Loading PLEXOS assemblies...")

import getpass, os, sys, clr
from os.path import dirname, join
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
import stdiomask
import tkinter as tk
from tkinter import filedialog
#
config_dataframe = pd.read_csv("run_config.csv")
plexos_path = "C:/Program Files/Energy Exemplar/PLEXOS " + str(config_dataframe["PLEXOS Version"][0])
sys.path.append(plexos_path)
clr.AddReference('PLEXOS7_NET.Core')
clr.AddReference('EEUTILITY')

from PLEXOS7_NET.Core import *
from EEUTILITY.Enums import *
from System import *

def connect_login():
    try:
        cxn.Connection('Data Source={};User Id={};Password={}'.format(server, username, password))
        # connection_status = cxn.IsConnectionValid()
        # time.sleep(4)
        print("\nYou have successfully logged in....")
        return
    except:
        print("\nFailed to authenticate the user [", username, "].")
        sys.exit()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    print("\n\n************* PLEXOS CONNECT LOGIN *************\n")
    # Login Credentials To Connect...
    server = str(config_dataframe["Server Address"][0])
    username = input('Username: ')
    password = stdiomask.getpass(prompt='Password: ', mask='*')

    cxn = PLEXOSConnect()

    # connect to the PLEXOS Connect server
    connect_login()
    is_end_to_end = config_dataframe["Uploading New Dataset?"][0]

    if (is_end_to_end):
        # Process of uploading database, running a simulation, uploading a solution view configuration file, exporting a solution view Excel Workbook.
        print("You selected an End-to-End Process!")

        dataset_name = str(config_dataframe["New Dataset Name"][0])

        addDatasetBool = cxn.AddDataset(None, dataset_name)
        if (addDatasetBool == False):
            print("Something went wrong with creating the dataset")
            sys.exit()

        # Database upload prompt.
        root = tk.Tk()
        root.withdraw()
        database_directory = filedialog.askopenfilename(initialdir="/", title="Select PLEXOS Database",
                                                        filetypes=(("XML File", "*.xml"), ("all files", "*.*")))

        databaseName = database_directory.split("/")[-1]
        sourceFolder = database_directory.replace(databaseName, "")

        print("The database is being uploaded .... ")
        try:
            cxn.UploadDataSet(sourceFolder, None, dataset_name, '1.0', 1, True)
            print("The database has been successfully uploaded to PLEXOS Connect!!")
        except:
            print("Something went wrong with uploading the dataset. The provided dataset name may already exist.")
            sys.exit()

        # set up the arguments for the job
        model_name = str(config_dataframe["Model Name"][0])
        args = ['"{}" -m "{}"'.format(databaseName, model_name)]
        run_id = cxn.AddRun('', dataset_name, '1.0', args, '', '', '', 0)

        # # track the progress of the run
        while not cxn.IsRunComplete(run_id):
            print(cxn.GetRunProgress(run_id))

        try:
            # Download the result
            cxn.DownloadSolution('.', run_id)
            print('Run', run_id, 'is complete.')
            print("\nThe solution file has been downloaded.\n")
        except:
            print("Something went wrong while downloading the solution file.")
            sys.exit()

        sol = Solution()
        try:
            solution_file = "Model " + model_name + " Solution.zip"
            sol.Connection(solution_file)
        except:
            print("Something went wrong while reading the solution file from the network\
             directory.")
            sys.exit()
        print("\n\n************* SOLUTION VIEW CONFIG FILE PARSING & OUTPUT REPORT SAVING *************\n")
        # Read and parse the configuration file.
        solutionConfiguration = str(config_dataframe["Solution View Config"][0])
        # wb = pd.ExcelWriter('Report1.xlsx')

        tree = ET.parse(solutionConfiguration)
        wb = pd.ExcelWriter(model_name + '.xlsx')
        tab_index = 1
        for item in tree.iter('SolutionHistoryItem'):
            properties = item.find('PropertyList')
            phaseEnum = item.find('Phases').find('int').text
            collectionName = 'System' + item.find('DisplayName').text.split('\\')[0]
            periodEnum = item.find('PeriodType').text
            seriesTypeEnum = item.find('SeriesType').text
            timeSliceList = item.find('TimesliceList').text
            sampleList = item.find('SampleList').text
            modelName = item.find('ModelList').text
            aggregationEnum = item.find('Aggregation').text
            categoryList = item.find('AggregationCategoryType').text

            # propertiesParsing = properties.split('Properties:')[1].split('Models')[0].split('\n')
            PropertiesList = []

            for property in properties:
                property = property.findall('string')[1].text.replace(' ', '').replace('&', '')
                propertyEnum = 'SystemOut' + item.find('DisplayName').text.split('\\')[0] + 'Enum.' + property
                PropertiesList.append(str(eval(propertyEnum)))

            PropertiesList = ','.join(PropertiesList)

            collectionEnum = 'CollectionEnum.' + collectionName
            collectionEnum = eval(collectionEnum)

            results = sol.Query(int(phaseEnum), \
                                collectionEnum, \
                                '', \
                                '', \
                                int(periodEnum), \
                                int(seriesTypeEnum), \
                                PropertiesList, \
                                None, \
                                None, \
                                '0', \
                                '', \
                                '', \
                                0, \
                                '', \
                                '')
            # Check to see if the query had results
            if results == None or results.EOF:
                print("No Results")
            else:
                # Create a DataFrame with a column for each column in the results
                cols = [x.Name for x in results.Fields]
                names = cols[cols.index('phase_name') + 1:]
                df = pd.DataFrame(columns=cols)

                resultsRows = results.GetRows()
                df = pd.DataFrame(
                    [[resultsRows[i, j] for i in range(resultsRows.GetLength(0))] for j in
                     range(resultsRows.GetLength(1))],
                    columns=[x.Name for x in results.Fields])
                tab_name = tree.find('Tabs')[tab_index - 1].find('Text').text

                df = df.drop(
                    columns=['parent_class_id', 'parent_id', 'collection_id', 'category_id', 'category_id', 'child_id',
                             'unit_id'])
                df.to_excel(wb, tab_name)  # 'Query' is the name of the worksheet

                tab_index += 1
            # List all datasets for the user to choose from.
            wb.save()

        print("\nThe solution views configuration based report has been saved.\n")
    else:
        # Uploading New Dataset? = False.
        # This is the process of running an existing dataset in the server.
        # set up the arguments for the job
        dataset_name = str(config_dataframe["Existing Dataset"][0])
        database_name = str(config_dataframe["Database Name"][0])
        model_name = str(config_dataframe["Model Name"][0])
        print("The model "+ model_name + " within the database "+ dataset_name + " under the dataset "+ dataset_name + " is being kicked off.")
        args = ['"{}" -m "{}"'.format(database_name, model_name)]
        run_id = cxn.AddRun('', dataset_name, '1.0', args, '', '', '', 0)

        # # track the progress of the run
        while not cxn.IsRunComplete(run_id):
            print(cxn.GetRunProgress(run_id))

        try:
            # Download the result
            cxn.DownloadSolution('.', run_id)
            print('Run', run_id, 'is complete.')
        except:
            print("Something went wrong while downloading the solution file.")
            sys.exit()

        sol = Solution()
        try:
            solution_file = "Model " + model_name + " Solution.zip"
            sol.Connection(solution_file)
        except:
            print("Something went wrong while reading the solution file from the network\
             directory.")
            sys.exit()

        print("\n\n************* SOLUTION VIEW CONFIG FILE PARSING & OUTPUT REPORT SAVING *************\n")
        # Read and parse the configuration file.
        solutionConfiguration = str(config_dataframe["Solution View Config"][0])
        # wb = pd.ExcelWriter('Report1.xlsx')

        tree = ET.parse(solutionConfiguration)
        wb = pd.ExcelWriter(model_name + '.xlsx')
        tab_index = 1
        for item in tree.iter('SolutionHistoryItem'):
            properties = item.find('PropertyList')
            phaseEnum = item.find('Phases').find('int').text
            collectionName = 'System' + item.find('DisplayName').text.split('\\')[0]
            periodEnum = item.find('PeriodType').text
            seriesTypeEnum = item.find('SeriesType').text
            timeSliceList = item.find('TimesliceList').text
            sampleList = item.find('SampleList').text
            modelName = item.find('ModelList').text
            aggregationEnum = item.find('Aggregation').text
            categoryList = item.find('AggregationCategoryType').text

            # propertiesParsing = properties.split('Properties:')[1].split('Models')[0].split('\n')
            PropertiesList = []

            for property in properties:
                property = property.findall('string')[1].text.replace(' ', '').replace('&', '')
                propertyEnum = 'SystemOut' + item.find('DisplayName').text.split('\\')[0] + 'Enum.' + property
                PropertiesList.append(str(eval(propertyEnum)))

            PropertiesList = ','.join(PropertiesList)

            collectionEnum = 'CollectionEnum.' + collectionName
            collectionEnum = eval(collectionEnum)

            results = sol.Query(int(phaseEnum), \
                                collectionEnum, \
                                '', \
                                '', \
                                int(periodEnum), \
                                int(seriesTypeEnum), \
                                PropertiesList, \
                                None, \
                                None, \
                                '0', \
                                '', \
                                '', \
                                0, \
                                '', \
                                '')
            # Check to see if the query had results
            if results == None or results.EOF:
                print("No Results")
            else:
                # Create a DataFrame with a column for each column in the results
                cols = [x.Name for x in results.Fields]
                names = cols[cols.index('phase_name') + 1:]
                df = pd.DataFrame(columns=cols)

                resultsRows = results.GetRows()
                df = pd.DataFrame(
                    [[resultsRows[i, j] for i in range(resultsRows.GetLength(0))] for j in
                     range(resultsRows.GetLength(1))],
                    columns=[x.Name for x in results.Fields])
                tab_name = tree.find('Tabs')[tab_index - 1].find('Text').text

                df = df.drop(
                    columns=['parent_class_id', 'parent_id', 'collection_id', 'category_id', 'category_id', 'child_id',
                             'unit_id'])
                df.to_excel(wb, tab_name)  # 'Query' is the name of the worksheet

                # Keep modelName, Collection, Child Name, Category, Fiscal Year, Properties Columns.
                # Save the workbook file.
                tab_index += 1
            # List all datasets for the user to choose from.
            wb.save()
        print("\nThe solution views configuration based report has been saved.\n")
