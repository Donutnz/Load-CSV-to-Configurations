# Load CSV to Configurations Script

## Description
Reads a CSV containing simple variations and writes it into Fusion 360 Configurations rows.

## Usage
Clone into your C:\\Users\\[YOUR_USERNAME]\\AppData\\Roaming\\Autodesk\\Autodesk Fusion 360\\API\\Scripts folder.

Recommend using https://github.com/Donutnz/Print-Configuration-Columns to get programmatic Configuration Table column titles as these can differ from what the GUI shows.

## Notes
* CSV must have headers. 

* "Part Number" must be the first column. This will set the row name and part number. See Sample Input file.

* Duplicate column titles (e.g. inserts) are handled as first come first served. I.e. the first column in the CSV with that title sets the first column in Configurations Table with same title.

* Script will exit if a setting a configuration results in errors in the timeline. If this occurs, try rolling the timeline back to start and rerunning the script. 

* Suppression Columns are handled as such: TRUE in the CSV = Not Suppressed

* Numerical (non-expression) length parameters are assumed to be in mm.

## Todo
* Validation of Thread Aspects
* Support Joint Snap Column type
* Split Part Name and Part number up.
* Change to Addin rather than a script with header to column mapping and GUI mapping.
* Make Part Number header (and maybe column title header) optionally case insensitive.

## Done
* Add Insert column type support.
* Handle duplicate columns (for Inserts)
* Parameters Column parameter value seems weird. Like its being stored internally as cm rather than mm. Also column.parameter seems borked.