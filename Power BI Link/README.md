# Power BI Link
This module provides a way of quickly pushing PLEXOS results to a form that is consumable by Power BI. 
This format is structurally native to PLEXOS, with perhaps some minor accommodations for usability.

The Power BI Link pulls all PLEXOS output data of a specified time granularity.

# Power BI Link Command Line Interface
This script is setup to run in the following command line approach:
```
python power_bi_link.py <PLEXOS Solution File.zip> [-y [<annual output file>.csv]] 
                                                   [-q [<quarterly output file>.csv]]
                                                   [-m [<monthly output file>.csv]]
                                                   [-w [<weekly output file>.csv]]
                                                   [-d [<daily output file>.csv]]
                                                   [-h [<hourly output file>.csv]]
                                                   [-i [<interval output file>.csv]]
```
Only those output time granulaties that have been specified as options will be pushed to Power BI 
Link format. For example, the following will only produce annual and monthly output.
```
python power_bi_link.py "Model Year DA Solution.zip" -y -m
```
Additionally in the above, the annual and monthly data will be pushed to default .csv file names for 
those types. The default filenames replace .zip in the PLEXOS Solution File with "_<timestep>.csv"
where <timestep> is annual, monthly, etc.
