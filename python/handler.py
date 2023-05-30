import pandas as pd
import openpyxl
import docx


class Handler:
    def __init__(self):
        self.pathToOldWord_ = ""
        self.pathToNewWord_ = ""
        self.dataOldDoc_ = pd.DataFrame()
        self.dataNewDoc_ = pd.DataFrame()
        self.dataOldHandled_ = {}
        self.dataNewHandled_ = {}
        self.mistakedWithDiffRevInNew_ = {}
        self.mistakedWithDiffRevInOld_ = {}
        self.names_ = {"Not changed": [], "Mistaked": [], "Changed": [], "New:": [], "Deleted": []}
        self.pathToPDBExcel_ = ""
        self.pdbExcelData_ = {}
        self.inPDB_ = []
        self.notInPDB_ = []

    def setPathToOldWord(self, path: str) -> None:
        self.pathToOldWord_ = path

    def setPathToNewWord(self, path: str) -> None:
        self.pathToNewWord_ = path

    def parseOldWord(self) -> None:
        self.parseWord(self.pathToOldWord_, self.dataOldDoc_)
        self.handleMistakedInOneDoc(self.dataOldDoc_, self.mistakedWithDiffRevInOld_, self.dataOldHandled_)

    def exportOldWordData(self) -> None:
        pathToExtractedData = self.pathToOldWord_.replace(".docx", ".xlsx")
        self.exportWordData(self.dataOldDoc_, pathToExtractedData)

    def exportNewWordData(self) -> None:
        pathToExtractedData = self.pathToNewWord_.replace(".docx", ".xlsx")
        self.exportWordData(self.dataNewDoc_, pathToExtractedData)

    def parseWord(self, pathToWord: str, dataToFill: pd.DataFrame):
        document = docx.Document(pathToWord)
        tables = document.tables
        for table in tables:
            if self.__isTableIsValid(table):
                self.__getInfoFromTable(table, dataToFill)

    def __isTableIsValid(self, table):
        rows = table.rows
        first_row_cells = rows[0].cells

        lastCellText = first_row_cells[-1].text
        lastCellCond1 = "Weight" in lastCellText
        lastCellCond2 = "weight" in lastCellText
        lastCellCond3 = "(kg)" in lastCellText
        if not (lastCellCond1 or lastCellCond2 or lastCellCond3):
            return False

        secondCellText = first_row_cells[1].text

        secondCellContains1 = "Workpack" in secondCellText
        secondCellContains2 = "Block" in secondCellText
        secondCellContains3 = "Assembly" in secondCellText

        if secondCellContains1 or secondCellContains2 or secondCellContains3:
            return False

        secondCellCond1 = "Subassembly" in secondCellText
        secondCellCond2 = "node" in secondCellText
        secondCellCond3 = "Mark" in secondCellText
        secondCellCond4 = "DETAILS" in secondCellText
        secondCellCond5 = "Details" in secondCellText
        secondCellCond6 = "Single" in secondCellText
        secondCellCond7 = "part" in secondCellText
        secondCellComplexCond = secondCellCond1 or secondCellCond2 or secondCellCond3 or \
                                secondCellCond4 or secondCellCond5 or secondCellCond6 or secondCellCond7

        if not secondCellComplexCond:
            return False

        return True

    def __getInfoFromTable(self, table, dataToFill):
        rows = table.rows
        for rowNum in range(1, len(rows)):
            row = rows[rowNum]
            cells = row.cells
            if len(cells) < 3:
                continue

            thirdCell = cells[2]
            nomination = thirdCell.text

            fourthCell = cells[3]
            revisionNumber = fourthCell.text

            if nomination not in dataToFill:
                dataToFill[nomination] = []

            dataToFill[nomination].append(revisionNumber)

    def __handleMistakesInOneDoc(self, dataDoc, withDiffRevInDoc, dataHandled):
        for nomination in dataDoc.keys():
            revisions = dataDoc[nomination]
            if self.__nominationIsValidInDoc(revisions):
                if nomination not in dataHandled:
                    dataHandled[nomination] = revisions[0]
            else:
                for uniqueRev in set(revisions):
                    if nomination in withDiffRevInDoc:
                        withDiffRevInDoc[nomination].append(uniqueRev)
                    else:
                        withDiffRevInDoc[nomination] = [uniqueRev]

    def __nominationIsValidInDoc(self, revisions):
        if len(revisions) == 1:
            return True
        for i in range(1, len(revisions)):
            if revisions[i - 1] != revisions[i]:
                return False
        return True

    def compareHashTables(self):
        newDataKeys = self.dataNewHandled.keys()
        for newDataKey in newDataKeys:
            if newDataKey in self.dataOldDoc:
                newRevisionValue = self.dataNewHandled[newDataKey]
                oldRevisionValue = self.dataOldHandled[newDataKey]
                if newRevisionValue == oldRevisionValue:
                    self.names["notChanged"].append(newDataKey)
                else:
                    oldRevVal = int(oldRevisionValue)
                    newRevVal = int(newRevisionValue)
                    if newRevVal == oldRevVal + 1:
                        self.names["changed"].append(newDataKey)
                    else:
                        self.names["mistaked"].append(newDataKey)
            else:
                self.names["new"].append(newDataKey)
        self.__findDeletedNominations()

    def __findDeletedNominations(self):
        oldNominations = self.dataOldHandled.keys()
        for oldKey in oldNominations:
            if oldKey not in self.dataNewHandled and oldKey not in self.mistakedWithDiffRevInNew:
                self.names["deleted"].append(oldKey)

    def exportDataToExcel(self, path):
        wb = openpyxl.Workbook()
        workSheet = wb.active
        workSheet.title = "Data"

        self.__exportNotChanged(workSheet)
        self.__exportChanged(workSheet)
        self.__exportNew(workSheet)
        self.__exportDeleted(workSheet)
        self.__exportMistaked(workSheet)
        self.__exportMistakedInNewDoc(workSheet)

        wb.save(path)

    def __exportChanged(self, workSheet):
        rowNum = 1
        workSheet.Cells(rowNum, 4).Value = "Changed"
        workSheet.Cells(rowNum, 5).Value = "NEW REV."
        rowNum += 1
        for nomination in self.names_("Changed"):
            obj = nomination
            workSheet.Cells(rowNum, 4).Value = obj

            newRev = self.dataNewHandled_(nomination)
            workSheet.Cells(rowNum, 5).Value = newRev

            rowNum += 1

    def __exportNew(self, workSheet):
        rowNum = 1
        workSheet.Cells(rowNum, 7).Value = "New"
        workSheet.Cells(rowNum, 8).Value = "NEW REV."
        rowNum += 1
        for nomination in self.names_("New"):
            obj = nomination
            workSheet.Cells(rowNum, 7).Value = obj

            newRev = self.dataNewHandled_(nomination)
            workSheet.Cells(rowNum, 8).Value = newRev

            rowNum += 1

    def __exportDeleted(self, workSheet):
        rowNum = 1
        workSheet.Cells(rowNum, 10).Value = "Deleted"
        workSheet.Cells(rowNum, 11).Value = "OLD REV."
        rowNum += 1
        for nomination in self.names_("Deleted"):
            obj = nomination
            workSheet.Cells(rowNum, 10).Value = obj

            newRev = self.dataOldHandled_(nomination)
            workSheet.Cells(rowNum, 11).Value = newRev

            rowNum += 1

    def __exportMistaked(self, workSheet):
        rowNum = 1
        workSheet.Cells(rowNum, 13).Value = "Mistaked"
        workSheet.Cells(rowNum, 14).Value = "OLD REV."
        workSheet.Cells(rowNum, 15).Value = "NEW REV."

        rowNum += 1
        for nomination in self.names_("Mistaked"):
            workSheet.Cells(rowNum, 13).Value = nomination
            workSheet.Cells(rowNum, 14).Value = self.dataOldHandled_(nomination)
            workSheet.Cells(rowNum, 15).Value = self.dataNewHandled_(nomination)

            rowNum += 1

    def __exportMistakedInNewDoc(self, workSheet):
        rowNum = 1
        workSheet.Cells(rowNum, 17).Value = "Mistaked with diff REV."
        workSheet.Cells(rowNum, 18).Value = "REV."

        rowNum += 1
        for item in self.mistakedWithDiffRevInNew_:
            revs = item.Value()
            for rev in revs:
                workSheet.Cells(rowNum, 17).Value = item.Key()
                workSheet.Cells(rowNum, 18).Value = rev
                rowNum += 1

    def setPathToPDBExcel(self, path):
        self.pathToPDBExcel_ = path

    def parsePDBExcel(self):
        # Open Workbook
        workbook = openpyxl.load_workbook(self.pathToPDBExcel_)
        worksheet = workbook.active

        # Read nominations and its revisions
        nominations = []
        revisions = []
        for row in worksheet.iter_rows(values_only=True):
            nominations.append(row[14])
            revisions.append(row[21])

        # Add data into the pdbExcelData_
        for nom, rev in zip(nominations, revisions):
            if nom in self.pdbExcelData_:
                raise Exception("В Excel-файле в столбце O есть дубликаты! Исправьте и перезагрузите")
            else:
                self.pdbExcelData_[nom] = rev

    def compareHandledDataWithPDB(self):
        consideredNames = ["Mistaked", "Changed", "New"]
        for key in consideredNames:
            tagNominations = self.names_[key]

            for nom in tagNominations:
                if nom in self.pdbExcelData_:
                    self.inPDB_.append(nom)
                else:
                    self.notInPDB_.append(nom)

    def exportComparisonPDBToExcel(self, path: str):
        wb = openpyxl.Workbook()
        workSheet = wb.active
        workSheet.title = "Summary"

        workSheet.Range("A1:D1").Merge()
        workSheet.Range("A1").Value = "In PDB"
        workSheet.Cells(2, 1).Value = "NAME"
        workSheet.Cells(2, 2).Value = "OLD REV."
        workSheet.Cells(2, 3).Value = "NEW REV."
        workSheet.Cells(2, 4).Value = "PDB REV."

        workSheet.Range("F1:I1").Merge()
        workSheet.Range("F1").Value = "Not in PDB"
        workSheet.Cells(2, 6).Value = "NAME"
        workSheet.Cells(2, 7).Value = "OLD REV."
        workSheet.Cells(2, 8).Value = "NEW REV."
        workSheet.Cells(2, 9).Value = "PDB REV."

        rowNum = 3
        for nom in self.inPDB_:
            workSheet.Cells(rowNum, 1).Value = nom
            if nom in self.dataOldHandled_:
                workSheet.Cells(rowNum, 2).Value = self.dataOldHandled_[nom]
            else:
                workSheet.Cells(rowNum, 2).Value = "-"
            workSheet.Cells(rowNum, 3).Value = self.dataNewHandled_[nom]
            workSheet.Cells(rowNum, 4).Value = self.pdbExcelData_[nom]
            rowNum += 1

        rowNum = 3
        for nom in self.notInPDB_:
            workSheet.Cells(rowNum, 6).Value = nom
            workSheet.Cells(rowNum, 7).Value = self.dataOldHandled_[nom]
            workSheet.Cells(rowNum, 8).Value = self.dataNewHandled_[nom]
            workSheet.Cells(rowNum, 9).Value = "-"
            rowNum += 1

        wb.save(path)
