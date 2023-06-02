import pandas as pd
import openpyxl
import docx
from collections import defaultdict


class Handler:
    def __init__(self):
        self.pathToOldWord_ = ""
        self.pathToNewWord_ = ""
        self.dataOldDoc_ = defaultdict(list)
        self.dataNewDoc_ = defaultdict(list)
        self.dataOldHandled_ = {}
        self.dataNewHandled_ = {}
        self.mistakedWithDiffRevInNew_ = {}
        self.mistakedWithDiffRevInOld_ = {}
        self.names_ = {"Not changed": [], "Mistaked": [], "Changed": [], "New:": [], "Deleted": []}

    def setPathToOldWord(self, path: str) -> None:
        self.pathToOldWord_ = path

    def setPathToNewWord(self, path: str) -> None:
        self.pathToNewWord_ = path

    def parseOldWord(self) -> None:
        self.__parseWord(self.pathToOldWord_, self.dataOldDoc_)
        self.__handleMistakedInOneDoc(self.dataOldDoc_, self.mistakedWithDiffRevInOld_, self.dataOldHandled_)

    def parseNewWord(self) -> None:
        self.__parseWord(self.pathToNewWord_, self.dataNewDoc_)
        self.__handleMistakedInOneDoc(self.dataNewDoc_, self.mistakedWithDiffRevInNew_, self.dataNewHandled_)

    def exportOldWordData(self) -> None:
        pathToExtractedData = self.pathToOldWord_.replace(".docx", ".xlsx")
        self.exportWordData(self.dataOldDoc_, pathToExtractedData)

    def exportNewWordData(self) -> None:
        pathToExtractedData = self.pathToNewWord_.replace(".docx", ".xlsx")
        self.exportWordData(self.dataNewDoc_, pathToExtractedData)

    def __parseWord(self, pathToWord: str, dataToFill: pd.DataFrame) -> None:
        document = docx.Document(pathToWord)
        tables = document.tables
        for table in tables:
            if self.__isTableIsValid(table):
                self.__getInfoFromTable(table, dataToFill)

    def __isTableIsValid(self, table):
        rows = table.rows
        firstRowCells = rows[0].cells

        lastCellText = firstRowCells[-1].text.lower()
        neededWords = ["weight", "(kg)"]
        if not any(word in lastCellText for word in neededWords):
            return False

        secondCellText = firstRowCells[1].text
        secondCellBadWords = {"Workpack", "Block", "Assembly"}
        secondCellText = firstRowCells[1].text
        if any(word in secondCellText for word in secondCellBadWords):
            return False

        secondCellConditions = {"Subassembly", "node", "Mark", "DETAILS", "Details", "Single", "part"}
        if not any(cond in secondCellText for cond in secondCellConditions):
            return False

        return True

    def __getInfoFromTable(self, table, dataToFill):
        for i, cells in enumerate([row.cells for row in table.rows[1:] if len(row.cells) >= 3], start=1):
            nomination = cells[2].text
            revisionNumber = cells[3].text

            dataToFill[nomination].append(revisionNumber)

    def __handleMistakedInOneDoc(self, dataDoc, withDiffRevInDoc, dataHandled):
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

    def compareHashTables(self) -> None:
        newDataKeys = self.dataNewHandled_.keys()
        for newDataKey in newDataKeys:
            if newDataKey in self.dataOldDoc_:
                newRevisionValue = self.dataNewHandled_[newDataKey]
                oldRevisionValue = self.dataNewHandled_[newDataKey]
                if newRevisionValue == oldRevisionValue:
                    self.names_["Not changed"].append(newDataKey)
                else:
                    oldRevVal = int(oldRevisionValue)
                    newRevVal = int(newRevisionValue)
                    if newRevVal == oldRevVal + 1:
                        self.names_["Changed"].append(newDataKey)
                    else:
                        self.names_["Mistaked"].append(newDataKey)
            else:
                self.names_["New"].append(newDataKey)
        self.__findDeletedNominations()

    def __findDeletedNominations(self):
        oldNominations = self.dataNewHandled_.keys()
        for oldKey in oldNominations:
            if oldKey not in self.dataNewHandled_ and oldKey not in self.mistakedWithDiffRevInNew_:
                self.names_["Deleted"].append(oldKey)

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

    def __exportNotChanged(self, workSheet) -> None:
        rowNum = 1
        workSheet.Cells(rowNum, 1).Value = "Not changed"
        workSheet.Cells(rowNum, 2).Value = "NEW REV."
        rowNum += 1
        for nomination in self.names_("Not changed"):
            obj = nomination
            workSheet.Cells(rowNum, 4).Value = obj

            newRev = self.dataNewHandled_(nomination)
            workSheet.Cells(rowNum, 5).Value = newRev

            rowNum += 1

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
