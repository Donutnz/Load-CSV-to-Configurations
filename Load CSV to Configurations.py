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
                    for tColumn in topTable.columns:
                        app.log("Column title: {}".format(tColumn.title))
                        if h == tColumn.title:
                            cellToBeSet=confRow.getCellByColumnId(tColumn.id)

                            if cellToBeSet is None:
                                # If this fires, something is seriously weird.
                                raise Exception("No cell found in at row: {} column: {}. \n Something really weird has happened.".format(confRow.name, tColumn.title))
                            
                            if isinstance(tColumn, adsk.fusion.ConfigurationParameterColumn): # All parameters are numerical. I think.
                                try:
                                    cellToBeSet.expression=v
                                except RuntimeError:
                                    app.log("Failed to set column {} to {}.".format(tColumn.title, v))
                                    raise
                                
                                changesCnt+=1

                            elif isinstance(tColumn, adsk.fusion.ConfigurationThemeColumn): # Mostly aimed at material for now. Can overcomplicate later.
                                for m in tColumn.referencedTable.rows:
                                    if v == m.name:
                                        cellToBeSet.referencedTableRow=m
                                        app.log("Set theme table row: {}".format(m.name))
                                        changesCnt+=1

                            elif isinstance(tColumn, adsk.fusion.ConfigurationPropertyColumn): # Description and part number
                                #app.log("Property {}".format(h))
                                cellToBeSet.value=v

                                changesCnt+=1

                            elif isinstance(tColumn, adsk.fusion.ConfigurationSuppressColumn):
                                if v == "TRUE":
                                    cellToBeSet.isSuppressed=True
                                elif v == "FALSE":
                                    cellToBeSet.isSuppressed=False
                                else:
                                    raise TypeError("Supression type column input data must be TRUE or FALSE. {} column in the input CSV is not bool.".format(tColumn.title))
                                
                                changesCnt+=1

                            elif isinstance(tColumn, adsk.fusion.ConfigurationInsertColumn): # Configuration insert. Most sketchy bit and most time inefficent. Should probs get the occurance top table just once.
                                insTopTable:adsk.fusion.ConfigurationTopTable=tColumn.occurrence.configuredDataFile.configurationTable

                                for tRow in insTopTable.rows:
                                    if tRow.name == v:
                                        #app.log("Inserted row {} from table {}".format(insTopTable.name, tRow.name))
                                        cellToBeSet.row=tRow

                                changesCnt+=1
                if isUpdating:
                    app.log("Touched {} columns".format(changesCnt))

                totalRowsCnt+=1

        ui.messageBox("Done! Added: {} Updated: {}".format(rowsAddedCnt, rowsUpdatedCnt))

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
