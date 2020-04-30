# Power BI Link
This module provides a way of quickly pushing PLEXOS results to a form that is 
consumable by Power BI. This format is structurally native to PLEXOS, with perhaps 
some minor accommodations for usability.

The Power BI Link pulls all PLEXOS output data of a specified time granularity.

# Power BI Link Command Line Interface
This script is setup to run in the following command line approach:
```
    python power_bi_link.py <solution_file> [-c [config_json]]
                                            [-y [yr_file]]
                                            [-q [qt_file]]
                                            [-m [mn_file]]
                                            [-w [wk_file]]
                                            [-d [dy_file]]
                                            [-h [hr_file]]
                                            [-i [in_file]]
                                            [-f [from_date]]
                                            [-t [to_date]]
```
Only those output time granulaties that have been specified as options will be pushed 
to Power BI Link format. For example, the following will only produce annual and 
monthly output.
```
python power_bi_link.py "Model Year DA Solution.zip" -y -m
```
Additionally in the above, the annual and monthly data will be pushed to default .csv
file names for those types. The default filenames replace .zip in the PLEXOS Solution 
File with "_<timestep>.csv" where <timestep> is annual, monthly, etc.

The ```annual.csv```, ```monthly.csv```, and ```hourly.csv``` files were produced
using the ```Model Year DA Solution.zip``` and the following two command lines.

```
python power_bi_link.py "Model Year DA Solution.zip" -y annual.csv -m monthly.csv
python power_bi_link.py "Model Year DA Solution.zip" -i hourly.csv -f 7/14/2024 -t 7/17/2024
```

## Configuration File Option
For users wanting to pull only specific outputs from PLEXOS into Power BI, the ```-c```
option provides a simple way to do this. To do so, one needs to create a configuration
file that indicates which outputs are desired. The structure of this file is shown in the
following example.

```
{
"queries": 
    [
        {
            "phase": "STSchedule",
            "parentclass": "System",
            "childclass": "Generator",
            "collection": "Generators",
            "properties": ["Generation", "Generation Cost", "SRMC"]
        },
        {
            "phase": "STSchedule",
            "parentclass": "System",
            "childclass": "Fuel",
            "collection": "Fuels"
        },
        {
            "phase": "STSchedule",
            "parentclass": "System",
            "childclass": "Region",
            "collection": "Regions",
            "properties": "Load"
        }
    ]
}
```

Each query indicates the simulation ```phase```: ```STSchedule```, ```MTSchedule```, ```PASA```, or ```LTPlan```.

The attributes ```parentclass``` and ```childclass``` indicate the those classes. Properties that come directly 
from a specific object generally have ```parentclass``` of ```System```. All three of the above are examples of this.

One might have ```parentclass``` of ```Generator``` and ```childclass``` of ```Fuel``` with ```collection```
of ```Generator.Fuels```. This would allow the user to determine how much of each fuel was used by each generator, 
or perhaps how much generation was producted by each generator-fuel pair.

The ```properties``` field may be missing, the name of a single property, or a list of properties. All of these are
demonstrated above.