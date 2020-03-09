# Congestion Analysis Tool

A tool for congestion analysis for PLEXOS outputs. This script computes the shift factors and impact factors for a selected pathway.

# Command Line Interface

As written here, the script should be run from the command line as follows:

python congestion_analysis.py <plexos_db> <ptdf_file> <solution_file> <from_bus> <to_bus> <time_stamp> [<from_bus> <to_bus> <time_stamp> ...]

For example

"C:\\Users\\<username>\Miniconda3\python.exe" "c:\\Users\\<username>\Documents\GitHub\PLEXOS-Analysis-Scripts\Congestion Analyzer\congestion_analysis.py" "PLEXOS ERCOT Nodal 1.0.0.xml" "Final PTDF.csv" "Model ERCOT 2019 1 Day  Solution.zip" "8126_FORMOSA4A_138" "8140_JOSLIN4A_138" "6/14/2019 4:00 PM"

One could follow this with additional triples of from_bus to_bus time_stamp assuming they refer to the same input, output, and PTDF matrix.
