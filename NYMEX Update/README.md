# NYMEX Update Example
This example shows how the end-to-end process of simulation can be
automated. In particular, in this case updates gas prices from
the internet and runs a simulation. The script does the following:

* pulls NYMEX gas prices from EIA's website
* updates a PLEXOS database
* creates and sets up new simulation models
* updates the horizon (to run five months beginning today in 2024)
* runs the simulation
* pulls results into a format accessible from Power BI

This example leverages the Power BI Link example in this repository.

# What is required?
Python 3.6+, Pythonnet library, a full download of this entire
repository, rtsDEMO unzipped in this folder, Power BI Desktop.

# Command Line Interface
```
python nymex_resimulate.py [-d <plexos_dataset>]
                           [-s <gas_scenario>]
                           [-g <gas_objects>]
                           [-p <project_name>]
                           [-b <base_scenarios>]
```
Each of these switches has a default value for this example. The default
value appears after the -->.
* <plexos_dataset> the PLEXOS input dataset --> rtsDEMO/rts_PLEXOS.xml
* <gas_scenario> the PLEXOS scenario for the gas price updates --> 'NYMEX'
* <gas_objects> a comma-seperated list of Gas objects in the PLEXOS data whose prices need updating from NYMEX --> "NG/CC,NG/CT"
* <project_name> the name of the PLEXOS project to configure and execute --> '3Month'
* <base_scenarios> a comma-separated list of the PLEXOS scenarios that should be included in the "base case" not updated with the NYMEX prices "Gen Outages,Load: DA,RE: DA,Add Spin Up"
