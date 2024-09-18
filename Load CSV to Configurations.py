#Author Josh TB
#Description Loads 

import adsk.core, adsk.fusion, adsk.cam, traceback
import csv

def extractBool(strWBool:str):
    """Extracts bool from CSV string

    strWBool: str String containing some for of TRUE or FALSE 
    """
    if strWBool.upper() == "TRUE":
        retBool=True
    elif strWBool.upper() == "FALSE":
        retBool=False
    else:
        raise TypeError("Supression type column input data must be TRUE or FALSE. Column in the input CSV is not bool.")
    return retBool

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
            csvReader:csv.reader=csv.reader(csvSource) # Not DictReader coz need duplicate columns

            csvHeaderLine=next(csvReader) # First line in the CSV.

            if csvHeaderLine[0] != "Part Number":
                app.log("Header not found: Part Number")
                raise Exception("Header not found: Part Number")

            # Get Headers
            csvHeadersVsColumns=list() #List ordered by csv headers containing config column ID for that header
            unclaimedColumns=list(topTable.columns)
            for headerIndex in range(0, len(csvHeaderLine)):
                for t in unclaimedColumns:
                    if csvHeaderLine[headerIndex] == t.title:
                        csvHeadersVsColumns.append((headerIndex, t.id, t.title)) # Title just makes debugging easier.
                        unclaimedColumns.remove(t)

            app.log("Headers: {}".format(["{} -> {}".format(x[0], x[2]) for x in csvHeadersVsColumns]))

            for csvRow in csvReader:
                app.log("Starting: {}".format(csvRow[0]))

                if csvRow[0] == "":
                    app.log("Skipping CSV row with empty Part Number (aka column 0)")
                    continue

                changesCnt=0 # Counts changes

                isUpdating=False # If a new config row is created or updating an existing one.

                confRow=None 

                for r in topTable.rows:
                    if r.name == csvRow[0]:
                        confRow=topTable.rows.itemById(r.id)
                        rowsUpdatedCnt+=1
                        app.log("Updating...")
                        isUpdating=True
                        break

                if confRow is None:
                    confRow=topTable.rows.add(csvRow[0])
                    rowsAddedCnt+=1
                    app.log("Adding...")

                for csvRowIndex, colID, colTitle in csvHeadersVsColumns:
                    #app.log("Column: {}".format(colTitle))
                    rowCellValue=csvRow[csvRowIndex]
                    confColumn = topTable.columns.itemById(colID)

                    if topTable.getCell(confColumn.index, confRow.index) is None:
                        # If this fires, something is seriously weird.
                        raise Exception("No cell found in at row: {} column: {}. \n Something really weird has happened.".format(confRow.name, confColumn.title))
                    
                    if isinstance(confColumn, adsk.fusion.ConfigurationParameterColumn): # All parameters are numerical. I think.
                        paramCell:adsk.fusion.ConfigurationParameterCell=topTable.getCell(confColumn.index, confRow.index)
                        try:
                            cellAsFloat=float(rowCellValue)/10 # Convert mm to cm (Fusion's internal units is cm)
                            if paramCell.value != cellAsFloat:
                                app.log("Parameter: {}: E={}, V={} -> {}".format(confColumn.title, paramCell.expression, paramCell.value, cellAsFloat))
                                paramCell.value=cellAsFloat
                                changesCnt+=1
                        except ValueError:
                            app.log("Parameter {} is expression: {}".format(confColumn.title, rowCellValue))
                            try:
                                if paramCell.expression != rowCellValue: # Expression because CSV might be expression and for fractional inch stuff.
                                    app.log("Parameter: {}: E={}, V={} -> {}".format(confColumn.title, paramCell.expression, paramCell.value, rowCellValue))
                                    paramCell.expression=rowCellValue
                                    changesCnt+=1
                            except RuntimeError:
                                app.log("Failed to set column {} to {}.".format(confColumn.title, rowCellValue))
                                raise

                    elif isinstance(confColumn, adsk.fusion.ConfigurationThemeColumn): # Mostly aimed at material and appearance for now. Can overcomplicate later.
                        themeTCell = adsk.fusion.ConfigurationThemeCell.cast(topTable.getCell(confColumn.index, confRow.index))
                        for m in confColumn.referencedTable.rows:
                            if rowCellValue == m.name:
                                if themeTCell.referencedTableRow != m:
                                    app.log("Theme table row: {}: {} -> {}".format(confColumn.title, themeTCell.referencedTableRow.name, m.name))
                                    themeTCell.referencedTableRow=m
                                    changesCnt+=1

                    elif isinstance(confColumn, adsk.fusion.ConfigurationPropertyColumn): # Description and part number
                        #app.log("Property {}".format(h))
                        propCell:adsk.fusion.ConfigurationPropertyCell=topTable.getCell(confColumn.index, confRow.index)
                        if propCell.value != rowCellValue:
                            app.log("Property: {}: {} -> {}".format(confColumn.title, propCell.value, rowCellValue))
                            propCell.value=rowCellValue
                            changesCnt+=1

                    elif isinstance(confColumn, adsk.fusion.ConfigurationSuppressColumn):
                        vAsBool = extractBool(rowCellValue)

                        vAsBool = not vAsBool # Ugh. "I want it active"=True -> isSuppressed=False

                        suppressCell:adsk.fusion.ConfigurationSuppressCell=topTable.getCell(confColumn.index, confRow.index)

                        if suppressCell.isSuppressed != vAsBool:
                            app.log("Suppress: {}: {} -> {}".format(confColumn.title, suppressCell.isSuppressed, vAsBool))
                            suppressCell.isSuppressed = vAsBool
                            changesCnt+=1

                    elif isinstance(confColumn, adsk.fusion.ConfigurationInsertColumn): # Configuration insert. Most sketchy bit and most time inefficent. Should probs get the occurance top table just once.
                        insTopTable:adsk.fusion.ConfigurationTopTable=confColumn.occurrence.configuredDataFile.configurationTable

                        insCell:adsk.fusion.ConfigurationInsertCell=topTable.getCell(confColumn.index, confRow.index)
                        for tRow in insTopTable.rows:
                            if tRow.name == rowCellValue:
                                if insCell.row != tRow:
                                    app.log("Insert: {}: {} -> {}".format(confColumn.title, insCell.row.name, tRow.name))
                                    #app.log("Inserted row {} from table {}".format(insTopTable.name, tRow.name))
                                    insCell.row=tRow

                                    changesCnt+=1
                    elif isinstance(confColumn, adsk.fusion.ConfigurationJointSnapColumn): # TODO
                        app.log("Joint Snap column not supported yet")

                    elif isinstance(confColumn, adsk.fusion.ConfigurationFeatureAspectColumn): # WIP
                        #app.log("Feature Aspect Column: {} WIP!!!".format(confColumn.title))
                        if topTable.getCell(confColumn.index, confRow.index).objectType == adsk.fusion.ConfigurationFeatureAspectBooleanCell.classType():
                            app.log("Bool Aspect")
                            aspectBoolCell:adsk.fusion.ConfigurationFeatureAspectBooleanCell=topTable.getCell(confColumn.index, confRow.index)
                            isActive=extractBool(rowCellValue)

                            if aspectBoolCell.value != isActive:
                                app.log("Aspect (bool): {}: {} -> {}".format(confColumn.title, aspectBoolCell.value, rowCellValue))
                                aspectBoolCell.value=isActive
                                changesCnt+=1
                        elif topTable.getCell(confColumn.index, confRow.index).objectType == adsk.fusion.ConfigurationFeatureAspectStringCell.classType(): # confColumn.feature.objectType = adsk::fusion::ThreadFeature
                            strAspect:adsk.fusion.ConfigurationFeatureAspectStringCell=topTable.getCell(confColumn.index, confRow.index)
                            app.log("Aspect (String): {}: {} -> {}".format(confColumn.title, strAspect.value, rowCellValue))
                            strAspect.value=rowCellValue
                            changesCnt+=1
                    else:
                        app.log("Unchanged: {}".format(confColumn.title))

                if isUpdating:
                    app.log("Changed {} columns".format(changesCnt))

                totalRowsCnt+=1

        ui.messageBox("Done! Added: {} Updated: {}".format(rowsAddedCnt, rowsUpdatedCnt))
        app.log("Done")

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
