# Load CSV to Configurations Script

## Description
Reads a CSV containing simple variations and writes it into Fusion 360 Configurations rows.

## Usage
Clone into your C:\\Users\\[YOUR_USERNAME]\\AppData\\Roaming\\Autodesk\\Autodesk Fusion 360\\API\\Scripts folder.

CSV must have headers. Headers must contain one "Part Number" column. This will set the row name and part number. See Sample Input file.

NOTE: Script will exit if a setting a configuration results in errors in the timeline. If this occurs, try rolling the timeline back to start and rerunning the script. 

## Todo
* Parameters Column parameter value seems weird. Like its being stored internally as cm rather than mm. Also column.parameter seems borked.
* Split Part Name and Part number up.
* Change to Addin rather than a script with header to column mapping and GUI mapping.
* Make Part Number header (and maybe column title header) optionally case insensitive.

## Done
* Add Insert column type support.