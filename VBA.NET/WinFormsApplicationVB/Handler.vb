Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Imports System.IO
Imports WordTable = DocumentFormat.OpenXml.Wordprocessing.Table

Public Class Handler
    Private pathToOldWord_ As String
    Private pathToNewWord_ As String

    Private dataOldDoc_ = New Dictionary(Of String, List(Of String))()
    Private dataNewDoc_ = New Dictionary(Of String, List(Of String))()

    Private dataOldHandled_ = New Hashtable()
    Private dataNewHandled_ = New Hashtable()

    Private mistakedWithDiffRevInNew_ = New Dictionary(Of String, List(Of String))()
    Private mistakedWithDiffRevInOld_ = New Dictionary(Of String, List(Of String))()

    Private names_ = New Dictionary(Of String, List(Of String))() From {
            {"Not changed", New List(Of String)()}, {"Mistaked", New List(Of String)()}, {"Changed", New List(Of String)()},
            {"New", New List(Of String)()}, {"Deleted", New List(Of String)}
        }

    Private pathToPDBExcel_ As String
    Private pdbExcelData_ = New Dictionary(Of String, String)()

    Private inPDB_ = New List(Of String)()
    Private notInPDB_ = New List(Of String)()

    Public Sub setPathToOldWord(path As String)
        pathToOldWord_ = path
    End Sub

    Public Sub setPathToNewWord(path As String)
        pathToNewWord_ = path
    End Sub

    Public Sub parseOldWord()
        acceptAllRevsInDoc(pathToOldWord_)
        parseWord(pathToOldWord_, dataOldDoc_)
        handleMistakedInOneDoc(dataOldDoc_, mistakedWithDiffRevInOld_, dataOldHandled_)
    End Sub

    Public Sub exportOldWordData()
        exportWordData(dataOldDoc_, pathToOldWord_.Replace(".docx", ".xlsx"))
    End Sub

    Public Sub exportNewWordData()
        exportWordData(dataNewDoc_, pathToNewWord_.Replace(".docx", ".xlsx"))
    End Sub

    Public Sub acceptAllRevs(pathToWordsFodler As String)
        Dim files = Directory.GetFiles(pathToWordsFodler)

        For Each file In files
            acceptAllRevsInDoc(file)
        Next
    End Sub

    Public Sub acceptAllRevsInDoc(pathToDoc As String)
        Dim wordApp As New Microsoft.Office.Interop.Word.Application()
        Dim doc As Microsoft.Office.Interop.Word.Document = wordApp.Documents.Open(pathToDoc)
        doc.AcceptAllRevisions()
        doc.Save()
        doc.Close()
        wordApp.Quit()
    End Sub

    Public Sub makeExcelDBFromWords(pathToWordsFodler As String)
        Dim dataDocs = New Dictionary(Of String, Dictionary(Of String, List(Of String)))

        Dim files = Directory.GetFiles(pathToWordsFodler)

        For Each file In files
            Dim dataDoc = New Dictionary(Of String, List(Of String))()
            parseWord(file, dataDoc)

            Dim fileName = Path.GetFileName(file)
            Dim start = fileName.IndexOf("DP-")
            Dim finish = fileName.LastIndexOf("_")
            Dim identificator = fileName.Substring(start, finish - start)
            dataDocs.Add(identificator, dataDoc)
        Next

        makeExcelDB(pathToWordsFodler + "/DB.xlsx", dataDocs)
    End Sub

    Private Function makeExcelDB(path As String, data As Dictionary(Of String, Dictionary(Of String, List(Of String))))
        Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
        Dim workBook = excelApp.Workbooks.Add()
        Dim workSheet = workBook.Worksheets(1)

        workSheet.Cells(1, 1).Value = "NAME"
        workSheet.Cells(1, 2).Value = "REV."
        workSheet.Cells(1, 3).Value = "WORD FILE"

        Dim rowNum = 2
        For Each item In data
            Dim fileName = item.Key
            Dim fileContent = item.Value

            For Each innerItem In fileContent
                Dim nomination = innerItem.Key
                Dim revisions = innerItem.Value
                For Each rev In revisions
                    workSheet.Cells(rowNum, 1).Value = nomination
                    workSheet.Cells(rowNum, 2).Value = rev
                    workSheet.Cells(rowNum, 3).Value = fileName
                    rowNum += 1
                Next
            Next
        Next

        workSheet.Name = "DB"
        excelApp.ActiveWorkbook.SaveAs(path)
        workBook.Close()
        excelApp.Quit()
    End Function

    Private Function exportWordData(data As Dictionary(Of String, List(Of String)), path As String)
        Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
        Dim workBook = excelApp.Workbooks.Add()
        Dim workSheet = workBook.Worksheets(1)

        workSheet.Cells(1, 1).Value = "NAME"
        workSheet.Cells(1, 2).Value = "REV."

        Dim rowNum = 2
        For Each item In data
            Dim name = item.Key
            Dim revs = item.Value

            For Each rev In revs
                workSheet.Cells(rowNum, 1).Value = name
                workSheet.Cells(rowNum, 2).Value = rev
                rowNum += 1
            Next rev
        Next item

        workSheet.Name = "DataOldDoc"
        excelApp.ActiveWorkbook.SaveAs(path)
        workBook.Close()
        excelApp.Quit()
    End Function

    Public Sub parseNewWord()
        acceptAllRevsInDoc(pathToNewWord_)
        parseWord(pathToNewWord_, dataNewDoc_)
        handleMistakedInOneDoc(dataNewDoc_, mistakedWithDiffRevInNew_, dataNewHandled_)
    End Sub

    Private Sub parseWord(pathToWord As String, dataToFill As Dictionary(Of String, List(Of String)))
        Using doc As WordprocessingDocument = WordprocessingDocument.Open(pathToWord, False)
            Dim tables As IEnumerable(Of WordTable) = doc.MainDocumentPart.Document.Body.Elements(Of WordTable)()
            For Each table As WordTable In tables
                If isTableIsValid(table) Then
                    getInfoFromTable(table, dataToFill)
                End If
            Next
        End Using
    End Sub

    Private Function isTableIsValid(table As WordTable) As Boolean
        Dim rows As IEnumerable(Of TableRow) = table.Elements(Of TableRow)()
        Dim firstRow As TableRow = rows.First()
        Dim firstRowCells As IEnumerable(Of TableCell) = firstRow.Elements(Of TableCell)

        Dim lastCellText As String = firstRowCells.Last().InnerText
        Dim lastCellCond As Boolean = lastCellText.ToLower().Contains("weight") Or lastCellText.Contains("(kg)")
        If Not lastCellCond Then
            Return False
        End If

        Dim secondCellText As String = firstRowCells.ElementAt(1).InnerText.ToLower()

        If secondCellText.Contains("workpack") Or secondCellText.Contains("block") Then
            Return False
        End If

        If secondCellText.Contains("assembly") Then
            Dim fifthCellText As String = firstRowCells.ElementAt(4).InnerText.ToLower()
            If fifthCellText.Contains("DP") Or fifthCellText.Contains("dp") Then
                Return False
            End If
        End If

        Dim keyWords() As String = {"Subassembly", "subassembly", "node", "mark", "details", "single", "part"}
        For Each keyWord As String In keyWords
            If secondCellText.Contains(keyWord) Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Function getInfoFromTable(table As WordTable, dataToFill As Dictionary(Of String, List(Of String)))
        Dim rows As IEnumerable(Of TableRow) = table.Elements(Of TableRow)
        For rowNum = 1 To rows.Count() - 1
            Dim row = rows.ElementAt(rowNum)
            Dim cells As IEnumerable(Of TableCell) = row.Elements(Of TableCell)()
            If (cells.Count() < 3) Then
                Continue For
            End If

            Dim thirdCell As TableCell = cells.ElementAt(2)
            Dim nomination As String = thirdCell.InnerText

            Dim fourthCell As TableCell = cells.ElementAt(3)
            Dim revisionNumber As String = fourthCell.InnerText

            If Not dataToFill.ContainsKey(nomination) Then
                dataToFill(nomination) = New List(Of String)()
            End If
            dataToFill(nomination).Add(revisionNumber)
        Next
    End Function

    Private Sub handleMistakedInOneDoc(dataDoc As Dictionary(Of String, List(Of String)), withDiffRevInDoc As Dictionary(Of String, List(Of String)), dataHandled As Hashtable)
        For Each nomination In dataDoc.Keys()
            Dim revisions = dataDoc(nomination)
            If nominationIsValidInDoc(revisions) Then
                If Not dataHandled.ContainsKey(nomination) Then
                    dataHandled.Add(nomination, revisions.First())
                End If

            Else
                For Each uniqueRev In revisions.ToHashSet()
                    If (withDiffRevInDoc.ContainsKey(nomination)) Then
                        withDiffRevInDoc(nomination).Add(uniqueRev)
                    Else
                        withDiffRevInDoc.Add(nomination, New List(Of String) From {uniqueRev})
                    End If

                Next uniqueRev
            End If
        Next nomination
    End Sub

    Private Function nominationIsValidInDoc(revisions As List(Of String)) As Boolean
        If Equals(1, revisions.Count()) Then
            Return True
        End If

        For i = 1 To revisions.Count() - 1
            If Not Equals(revisions(i - 1), revisions(i)) Then
                Return False
            End If
        Next i
        Return True
    End Function

    Public Sub compareHashTables()
        Dim newDataKeys = dataNewHandled_.Keys()
        For Each newDataKey In newDataKeys
            'Check if key exists already in the old document
            If dataOldDoc_.ContainsKey(newDataKey) Then
                Dim newRevisionValue As String = dataNewHandled_(newDataKey)
                Dim oldRevisionValue As String = dataOldHandled_(newDataKey)

                ' Check if there is no difference between old and new revision values
                If (Equals(newRevisionValue, oldRevisionValue)) Then
                    Dim notChanged = names_("Not changed")
                    notChanged.Add(newDataKey)
                Else
                    ' Check if there is difference (should be only so: newRevisionValue = oldRevisionValue + 1)
                    Dim oldRevVal As Integer = Val(oldRevisionValue)
                    Dim newRevVal As Integer = Val(newRevisionValue)
                    If Equals(newRevVal, oldRevVal + 1) Then
                        names_("Changed").Add(newDataKey)
                    Else
                        names_("Mistaked").Add(newDataKey)
                    End If
                End If
            Else
                names_("New").Add(newDataKey)
            End If
        Next

        findDeletedNominations()
    End Sub

    Private Sub findDeletedNominations()
        Dim oldNominations = dataOldHandled_.Keys()
        For Each oldKey In oldNominations
            If Not dataNewHandled_.ContainsKey(oldKey) And Not mistakedWithDiffRevInNew_.ContainsKey(oldKey) Then
                names_("Deleted").Add(oldKey)
            End If
        Next
    End Sub

    Public Sub exportDataToExcel(path As String)
        Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
        Dim workBook = excelApp.Workbooks.Add()
        Dim workSheet = workBook.Worksheets(1)

        exportNotChanged(workSheet)
        exportChanged(workSheet)
        exportNew(workSheet)
        exportDeleted(workSheet)
        exportMistaked(workSheet)
        exportMistakedInNewDoc(workSheet)

        workSheet.Name = "Data"
        excelApp.ActiveWorkbook.SaveAs(path)
        workBook.Close()
        excelApp.Quit()
    End Sub

    Private Sub exportNotChanged(workSheet As Object)
        Dim rowNum As Integer = 1
        workSheet.Cells(rowNum, 1).Value = "Not changed"
        workSheet.Cells(rowNum, 2).Value = "REV."
        rowNum += 1
        For Each nomination In names_("Not changed")
            Dim obj As Object = nomination
            workSheet.Cells(rowNum, 1).Value = obj

            Dim rev As Object = dataNewHandled_(nomination)
            workSheet.Cells(rowNum, 2).Value = rev

            rowNum += 1

        Next nomination
    End Sub

    Private Sub exportChanged(workSheet As Object)
        Dim rowNum As Integer = 1
        workSheet.Cells(rowNum, 4).Value = "Changed"
        workSheet.Cells(rowNum, 5).Value = "NEW REV."
        rowNum += 1
        For Each nomination In names_("Changed")
            Dim obj As Object = nomination
            workSheet.Cells(rowNum, 4).Value = obj

            Dim newRev As Object = dataNewHandled_(nomination)
            workSheet.Cells(rowNum, 5).Value = newRev

            rowNum += 1

        Next nomination
    End Sub

    Private Sub exportNew(workSheet As Object)
        Dim rowNum As Integer = 1
        workSheet.Cells(rowNum, 7).Value = "New"
        workSheet.Cells(rowNum, 8).Value = "NEW REV."
        rowNum += 1
        For Each nomination In names_("New")
            Dim obj As Object = nomination
            workSheet.Cells(rowNum, 7).Value = obj

            Dim newRev As Object = dataNewHandled_(nomination)
            workSheet.Cells(rowNum, 8).Value = newRev

            rowNum += 1

        Next nomination
    End Sub

    Private Sub exportDeleted(workSheet As Object)
        Dim rowNum As Integer = 1
        workSheet.Cells(rowNum, 10).Value = "Deleted"
        workSheet.Cells(rowNum, 11).Value = "OLD REV."
        rowNum += 1
        For Each nomination In names_("Deleted")
            Dim obj As Object = nomination
            workSheet.Cells(rowNum, 10).Value = obj

            Dim newRev As Object = dataOldHandled_(nomination)
            workSheet.Cells(rowNum, 11).Value = newRev

            rowNum += 1

        Next nomination
    End Sub

    Private Sub exportMistaked(workSheet As Object)
        Dim rowNum As Integer = 1
        workSheet.Cells(rowNum, 13).Value = "Mistaked"
        workSheet.Cells(rowNum, 14).Value = "OLD REV."
        workSheet.Cells(rowNum, 15).Value = "NEW REV."

        rowNum += 1
        For Each nomination In names_("Mistaked")
            workSheet.Cells(rowNum, 13).Value = nomination
            workSheet.Cells(rowNum, 14).Value = dataOldHandled_(nomination)
            workSheet.Cells(rowNum, 15).Value = dataNewHandled_(nomination)

            rowNum += 1

        Next nomination
    End Sub

    Private Sub exportMistakedInNewDoc(workSheet As Object)
        Dim rowNum As Integer = 1
        workSheet.Cells(rowNum, 17).Value = "Mistaked with diff REV."
        workSheet.Cells(rowNum, 18).Value = "REV."

        rowNum += 1
        For Each item In mistakedWithDiffRevInNew_
            Dim revs As List(Of String) = item.Value()
            For Each rev In revs
                workSheet.Cells(rowNum, 17).Value = item.Key()
                workSheet.Cells(rowNum, 18).Value = rev
                rowNum += 1
            Next
        Next
    End Sub

    Public Sub setPathToPDBExcel(path As String)
        pathToPDBExcel_ = path
    End Sub

    Public Sub parsePDBExcel()
        Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
        Dim workbook As Microsoft.Office.Interop.Excel.Workbook = excelApp.Workbooks.Open(pathToPDBExcel_)
        Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = workbook.Sheets(1)

        Dim nominationsRange As Microsoft.Office.Interop.Excel.Range = worksheet.Range("O:O")
        Dim nominations = New List(Of String)

        For Each nomination In nominationsRange.Value
            If IsNothing(nomination) Then
                Exit For
            End If
            nominations.Add(nomination.ToString())
        Next nomination

        Dim revisionsRange As Microsoft.Office.Interop.Excel.Range = worksheet.Range("V:V")
        Dim revisions = New List(Of String)

        For Each revision In revisionsRange.Value
            If IsNothing(revision) Then
                Exit For
            End If
            revisions.Add(revision.ToString())
        Next

        For i = 1 To nominations.Count() - 1
            If pdbExcelData_.ContainsKey(nominations(i)) Then
                workbook.Close()
                Throw New Exception("В Excel-файле в столбце O есть дубликаты! Исправьте и перезагрузите")
            Else
                pdbExcelData_.Add(nominations(i), revisions(i))
            End If
        Next
    End Sub

    Public Sub compareHandledDataWithPDB()
        ' Firstly - compare names
        Dim consideredNames = New List(Of String) From {"Mistaked", "Changed", "New"}
        For Each key In consideredNames
            Dim tagNominations = names_(key)

            For Each nom In tagNominations
                If pdbExcelData_.ContainsKey(nom) Then
                    inPDB_.Add(nom)
                Else
                    notInPDB_.Add(nom)
                End If
            Next
        Next
    End Sub

    Public Sub exportComparisonPDBToExcel(path As String)
        Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
        Dim workBook = excelApp.Workbooks.Add()
        Dim workSheet = workBook.Worksheets(1)

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

        Dim rowNum As Integer = 3
        For Each nom In inPDB_
            workSheet.Cells(rowNum, 1).Value = nom
            If dataOldHandled_.ContainsKey(nom) Then
                workSheet.Cells(rowNum, 2).Value = dataOldHandled_(nom)
            Else
                workSheet.Cells(rowNum, 2).Value = "-"
            End If
            workSheet.Cells(rowNum, 3).Value = dataNewHandled_(nom)
            workSheet.Cells(rowNum, 4).Value = pdbExcelData_(nom)
            rowNum += 1
        Next nom

        rowNum = 3
        For Each nom In notInPDB_
            workSheet.Cells(rowNum, 6).Value = nom
            workSheet.Cells(rowNum, 7).Value = dataOldHandled_(nom)
            workSheet.Cells(rowNum, 8).Value = dataNewHandled_(nom)
            workSheet.Cells(rowNum, 9).Value = "-"
            rowNum += 1
        Next nom

        workSheet.Name = "Summary"
        excelApp.ActiveWorkbook.SaveAs(path)
        workBook.Close()
        excelApp.Quit()
    End Sub

    Public Sub clearData()
        pathToOldWord_ = ""
        pathToNewWord_ = ""
        pathToPDBExcel_ = ""

        dataOldDoc_.Clear()
        dataOldDoc_ = New Dictionary(Of String, List(Of String))()

        dataNewDoc_.Clear()
        dataNewDoc_ = New Dictionary(Of String, List(Of String))()

        dataOldHandled_.Clear()
        dataOldHandled_ = New Hashtable()

        dataNewHandled_.Clear()
        dataNewHandled_ = New Hashtable()

        mistakedWithDiffRevInNew_.Clear()
        mistakedWithDiffRevInNew_ = New Dictionary(Of String, List(Of String))()

        mistakedWithDiffRevInOld_.Clear()
        mistakedWithDiffRevInOld_ = New Dictionary(Of String, List(Of String))()

        names_.Clear()
        names_ = New Dictionary(Of String, List(Of String)) From {
            {"Not changed", New List(Of String)()}, {"Mistaked", New List(Of String)()}, {"Changed", New List(Of String)()},
            {"New", New List(Of String)()}, {"Deleted", New List(Of String)}
        }

        pdbExcelData_ = New Dictionary(Of String, String)()
    End Sub
End Class
