#Author Josh TB
#Description Loads 

import adsk.core, adsk.fusion, adsk.cam, traceback
import csv

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui  = app.userInterface

        design=adsk.fusion.Design.cast(app.activeProduct)

        app.log("Starting CSV to Configurations Importer...")

        # Check if this is a configured design.
        if not design.isConfiguredDesign:
            ui.messageBox("Error: The current design is not configured! This script can only run on configured designs.")
            return

        sourceCSVFilePath=""

        csvFileDlg=ui.createFileDialog()
        csvFileDlg.title="Choose CSV file"
        csvFileDlg.filter='CSV Files (*.csv)'
        csvFileDlg.isMultiSelectEnabled=False
        
        dlgResult=csvFileDlg.showOpen()
        if dlgResult == adsk.core.DialogResults.DialogOK:
            sourceCSVFilePath=csvFileDlg.filename
        else:
            return
        
        topTable=design.configurationTopTable

        #TODO Add csv sniffer to check headers are there and correct.
        rowsAddedCnt=0
        rowsUpdatedCnt=0
        totalRowsCnt=0

        with open(sourceCSVFilePath) as csvSource:
            csvReader:csv.DictReader=csv.DictReader(csvSource)
            app.log("Headers: {}".format(csvReader.fieldnames))

            if "Part Number" not in csvReader.fieldnames:
                app.log("Header not found: Part Number")
                raise Exception("Header not found: Part Number")

            for csvRow in csvReader:
                app.log("Starting: {}".format(csvRow["Part Number"]))

                changesCnt=0 # Counts changes

                isUpdating=False # If a new config row is created or updating an existing one.

                confRow=None 

                for r in topTable.rows:
                    if r.name == csvRow["Part Number"]:
                        confRow=topTable.rows.itemByName(csvRow["Part Number"])
                        rowsUpdatedCnt+=1
                        app.log("Updating...")
                        isUpdating=True
                        break

                if confRow is None:
                    confRow=topTable.rows.add(csvRow["Part Number"])
                    rowsAddedCnt+=1
                    app.log("Adding...")

                for h,v in csvRow.items():
                    if h == "": # Skip headerless CSV columns
                        app.log("Skipping headerless column...")
                        continue

                    for tColumn in topTable.columns:
                        #app.log("Column title: {}".format(tColumn.title))
                        if h == tColumn.title:
                            cellToBeSet=confRow.getCellByColumnId(tColumn.id)

                            if cellToBeSet is None:
                                # If this fires, something is seriously weird.
                                raise Exception("No cell found in at row: {} column: {}. \n Something really weird has happened.".format(confRow.name, tColumn.title))
                            
                            if isinstance(tColumn, adsk.fusion.ConfigurationParameterColumn): # All parameters are numerical. I think.
                                try:
                                    if cellToBeSet.value != float(v): # Should be expression coz might be an expression in the CSV.
                                        app.log("Parameter: {}: E={}, V={} -> {}".format(tColumn.title, cellToBeSet.expression, cellToBeSet.value, v))
                                        #app.log("Test: {}".format(tColumn.parameter)) # Possibly a bug w column.parameter property. Have posted abt it.
                                        cellToBeSet.expression=v
                                        changesCnt+=1
                                except RuntimeError:
                                    app.log("Failed to set column {} to {}.".format(tColumn.title, v))
                                    raise

                            elif isinstance(tColumn, adsk.fusion.ConfigurationThemeColumn): # Mostly aimed at material and appearance for now. Can overcomplicate later.
                                themeTCell = adsk.fusion.ConfigurationThemeCell.cast(confRow.getCellByColumnId(tColumn.id))
                                for m in tColumn.referencedTable.rows:
                                    if v == m.name:
                                        if themeTCell.referencedTableRow != m:
                                            app.log("Theme table row: {}: {} -> {}".format(tColumn.title, themeTCell.referencedTableRow.name, m.name))
                                            themeTCell.referencedTableRow=m
                                            changesCnt+=1

                            elif isinstance(tColumn, adsk.fusion.ConfigurationPropertyColumn): # Description and part number
                                #app.log("Property {}".format(h))
                                if cellToBeSet.value != v:
                                    app.log("Property: {}: {} -> {}".format(tColumn.title, cellToBeSet.value, v))
                                    cellToBeSet.value=v

                                    changesCnt+=1

                            elif isinstance(tColumn, adsk.fusion.ConfigurationSuppressColumn):
                                if v.upper() == "TRUE":
                                    vAsBool=True
                                elif v.upper() == "FALSE":
                                    vAsBool=False
                                else:
                                    raise TypeError("Supression type column input data must be TRUE or FALSE. {} column in the input CSV is not bool.".format(tColumn.title))
                                    continue

                                vAsBool = not vAsBool # Ugh. "I want it active"=True -> isSuppressed=False

                                if cellToBeSet.isSuppressed != vAsBool:
                                    app.log("Suppress: {}: {} -> {}".format(tColumn.title, cellToBeSet.isSuppressed, vAsBool))
                                    cellToBeSet.isSuppressed = vAsBool
                                    changesCnt+=1

                            elif isinstance(tColumn, adsk.fusion.ConfigurationInsertColumn): # Configuration insert. Most sketchy bit and most time inefficent. Should probs get the occurance top table just once.
                                insTopTable:adsk.fusion.ConfigurationTopTable=tColumn.occurrence.configuredDataFile.configurationTable

                                for tRow in insTopTable.rows:
                                    if tRow.name == v:
                                        if cellToBeSet.row != tRow:
                                            app.log("Insert: {}: {} -> {}".format(tColumn.title, cellToBeSet.row.name, tRow.name))
                                            #app.log("Inserted row {} from table {}".format(insTopTable.name, tRow.name))
                                            cellToBeSet.row=tRow

                                            changesCnt+=1

                            else:
                                app.log("Unchanged: {}".format(tColumn.title))

                if isUpdating:
                    app.log("Changed {} columns".format(changesCnt))

                totalRowsCnt+=1

        ui.messageBox("Done! Added: {} Updated: {}".format(rowsAddedCnt, rowsUpdatedCnt))
        app.log("Done")

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
