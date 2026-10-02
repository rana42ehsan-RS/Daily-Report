@echo off
REM ---------------------------------------------------------------
REM  Daily AQI Report - one-click run (edit the three paths below)
REM ---------------------------------------------------------------
cd /d "%~dp0"
set DATA=data
set AQMS_SHP=shp\56AQMS.shp

python daily_aqi_report_Punjab.py --csv "%DATA%" --shp "%AQMS_SHP%" --basemap osm --out output
pause
