@echo off
REM 
REM Place a header in the console output
@echo *****************************In Post-simulation********************************
@echo DATASET_PATH=[%DATASET_PATH%]
@echo SOLUTION0=[%SOLUTION_0%]
@echo *******************************************************************************
REM
REM Setup a time stamp for the destination file
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "YY=%dt:~2,2%" & set "YYYY=%dt:~0,4%" & set "MM=%dt:~4,2%" & set "DD=%dt:~6,2%"
set "HH=%dt:~8,2%" & set "Min=%dt:~10,2%" & set "Sec=%dt:~12,2%"
set "datestamp=%YYYY%%MM%%DD%" & set "timestamp=%HH%%Min%%Sec%"
set "fullstamp=%YYYY%-%MM%-%DD%T%HH%-%Min%-%Sec%"
@echo %fullstamp%

for %%P in ("%DATASET_PATH%") do (
	pushd %%~dpP
	Call "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" "agg.cs"
	for /r %%f in (?T*.csv) do (
		pushd %%~dpP
		agg.exe "%SOLUTION_0%,%%~f,%fullstamp%"
		pushd %SOLUTION_0%
	) 
)