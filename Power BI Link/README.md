# Power BI Link
This module provides a way of quickly pushing PLEXOS results to a form that is 
consumable by Power BI. This format is structurally native to PLEXOS, with perhaps 
some minor accommodations for usability.

The Power BI Link pulls all PLEXOS output data of a specified time granularity.

# Power BI Link Command Line Interface
This script is setup to run in the following command line approach:
```
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
option provides a simple way to do this. 

```
python power_bi_link.py "Model Year DA Solution.zip" -y annual.csv -c config.json
```

To do so, one needs to create a configuration file that indicates which outputs are desired. 
The structure of this file is shown in the following example. An example file is provided
in this project and is called ```config.json```. The configuation file maybe located wherever
the user prefers and named according to the user's preference as long as the path to the file
is indicated after the ```-c``` option.

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

## Cross Reference Table
It is often helpful to understand object linkages, say for example to know which 
generators are in a specific region. The ```-x``` option allows you to obtain this
metadata information in addition to numeric output results. For example,

```
python power_bi_link.py "Model Year DA Solution.zip" -x xref.csv
```

will produce a .csv file table that indicates all available object relationships in
the PLEXOS Solution file. 

The metadata is presented in a .csv tabular form similar to the following. For ease
of reading we have tab-separated instead of comma-separating, although the data appears
in .csv format.

```
     ParentClass ParentName ParentCategory       Collection ChildClass  ChildName  ChildCategory
1         System     System              -       Generators  Generator      101_3     Coal/Steam
2         System     System              -       Generators  Generator      101_4     Coal/Steam
3         System     System              -       Generators  Generator      101_5     Coal/Steam
4         System     System              -       Generators  Generator      102_3     Coal/Steam
5         System     System              -       Generators  Generator      102_4     Coal/Steam
6         System     System              -       Generators  Generator      115_3     Coal/Steam
7         System     System              -       Generators  Generator      116_1     Coal/Steam
8         System     System              -       Generators  Generator      123_2     Coal/Steam
9         System     System              -       Generators  Generator      123_3     Coal/Steam
10        System     System              -       Generators  Generator      201_3     Coal/Steam
11        System     System              -       Generators  Generator      202_3     Coal/Steam
```

The ```-x``` option can be used in combination with other option switches as needed.