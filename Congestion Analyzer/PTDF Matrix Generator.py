"""
Author: Tarek Ibrahim

Acknowledgments: Energy Exemplar Solution Engineering Team
"""

"""
Runs once per model to post-process the PTDF matrix output from PLEXOS 
"""
import csv
import pandas as pd
import os
import time
import datetime
start_time = time.time()

Data = pd.read_csv("C:\TI\Congestion Analysis\PJM\PTDF.txt", sep = '\t')

Data = Data.T

Data = Data.drop(Data.index[[0]])
Data.reset_index(inplace = True) 

Data.to_csv(r'C:\TI\Congestion Analysis\PJM\Final PTDF.csv', index = False)
print("--- %s seconds ---" % (time.time() - start_time))
#print("Done")
