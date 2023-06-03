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
        self.names_ = {"Not changed": [], "Mistaked": [], "Changed": [], "New": [], "Deleted": []}

    def clearData(self):
        self.pathToOldWord_ = ""
        self.pathToNewWord_ = ""
        self.dataOldDoc_.clear()
        self.dataNewDoc_.clear()
        self.dataOldHandled_.clear()
        self.dataNewHandled_.clear()
        self.mistakedWithDiffRevInNew_.clear()
        self.mistakedWithDiffRevInOld_.clear()
        self.names_.clear()
        self.names_ = {"Not changed": [], "Mistaked": [], "Changed": [], "New": [], "Deleted": []}

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
        self.__exportWordData(self.dataOldDoc_, pathToExtractedData)

    def exportNewWordData(self) -> None:
        pathToExtractedData = self.pathToNewWord_.replace(".docx", ".xlsx")
        self.__exportWordData(self.dataNewDoc_, pathToExtractedData)

    def __parseWord(self, pathToWord: str, dataToFill: pd.DataFrame) -> None:
        document = docx.Document(pathToWord)
        tables = document.tables
        for table in tables:
            if self.__isTableIsValid(table):
                self.__getInfoFromTable(table, dataToFill)

    def __isTableIsValid(self, table) -> bool:
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

    def __getInfoFromTable(self, table, dataToFill) -> None:
        for row in table.rows[1:-1]:
            nomination = row.cells[2].text
            revisionNumber = row.cells[3].text

            dataToFill[nomination].append(revisionNumber)

    def __handleMistakedInOneDoc(self, dataDoc, withDiffRevInDoc, dataHandled) -> None:
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

    def __nominationIsValidInDoc(self, revisions) -> bool:
        if len(revisions) == 1:
            return True
        for i in range(1, len(revisions)):
            if revisions[i - 1] != revisions[i]:
                return False
        return True

    def compareHashTables(self) -> None:
        for newDataKey in self.dataNewHandled_.keys():
            if newDataKey in self.dataOldHandled_.keys():
                newRevisionValue = self.dataNewHandled_[newDataKey]
                oldRevisionValue = self.dataOldHandled_[newDataKey]
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

    def __findDeletedNominations(self) -> None:
        oldNominations = self.dataNewHandled_.keys()
        for oldKey in oldNominations:
            if oldKey not in self.dataNewHandled_ and oldKey not in self.mistakedWithDiffRevInNew_:
                self.names_["Deleted"].append(oldKey)

    def exportDataToExcel(self, path) -> None:
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
        workSheet['A1'] = "Not changed"
        workSheet['B1'] = "REV."
        rowNum += 1
        for nomination in self.names_["Not changed"]:
            obj = nomination
            workSheet[f'A{str(rowNum)}'] = obj

            newRev = self.dataNewHandled_[nomination]
            workSheet[f'B{str(rowNum)}'] = newRev

            rowNum += 1

    def __exportChanged(self, workSheet) -> None:
        rowNum = 1
        workSheet['D1'] = "Changed"
        workSheet['E1'] = "NEW changed REV."
        rowNum += 1
        for nomination in self.names_["Changed"]:
            obj = nomination
            workSheet[f'D{str(rowNum)}'] = obj

            newRev = self.dataNewHandled_[nomination]
            workSheet[f'E{str(rowNum)}'] = newRev

            rowNum += 1

    def __exportNew(self, workSheet) -> None:
        rowNum = 1
        workSheet['G1'] = "New"
        workSheet['H1'] = "NEW REV."
        rowNum += 1
        for nomination in self.names_["New"]:
            obj = nomination
            workSheet[f'G{str(rowNum)}'] = obj

            newRev = self.dataNewHandled_[nomination]
            workSheet[f'H{str(rowNum)}'] = newRev

            rowNum += 1

    def __exportDeleted(self, workSheet) -> None:
        rowNum = 1
        workSheet['J1'] = "Deleted"
        workSheet['K1'] = "OLD REV."
        rowNum += 1
        for nomination in self.names_["Deleted"]:
            obj = nomination
            workSheet[f'J{str(rowNum)}'] = obj

            newRev = self.dataOldHandled_[nomination]
            workSheet[f'K{str(rowNum)}'] = newRev

            rowNum += 1

    def __exportMistaked(self, workSheet) -> None:
        rowNum = 1
        workSheet['M1'] = "Mistaked"
        workSheet['N1'] = "OLD REV."
        workSheet['O1'] = "NEW REV."
        rowNum += 1
        for nomination in self.names_["Mistaked"]:
            workSheet[f'M{str(rowNum)}'] = nomination
            workSheet[f'N{str(rowNum)}'] = self.dataOldHandled_[nomination]
            workSheet[f'O{str(rowNum)}'] = self.dataNewHandled_[nomination]

            rowNum += 1

    def __exportMistakedInNewDoc(self, workSheet) -> None:
        rowNum = 1
        workSheet['Q1'] = "Mistaked with diff REV."
        workSheet['R1'] = "REV."

        rowNum += 1
        for nom, revs in self.mistakedWithDiffRevInNew_.items():
            for rev in revs:
                workSheet[f'Q{str(rowNum)}'] = nom
                workSheet[f'R{str(rowNum)}'] = rev
                rowNum += 1

    def __exportWordData(self, data, pathToExtractedData) -> None:
        wb = openpyxl.Workbook()
        workSheet = wb.active
        workSheet.title = "Extracted data"

        rowNum = 1
        workSheet['A1'] = "NAME"
        workSheet['B1'] = "REV."
        rowNum += 1
        for key, values in data.items():
            for value in values:
                workSheet[f'A{rowNum}'] = key
                workSheet[f'B{rowNum}'] = value
                rowNum += 1

        wb.save(pathToExtractedData)
