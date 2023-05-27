Public Class ConverterWindow
    Private handler_ As Handler
    Public Sub New()
        InitializeComponent()
        handler_ = New Handler()
    End Sub

    Private Sub onLoadOldWord(sender As Object, e As EventArgs) Handles pbLoadOldWord.Click
        Dim pathToOldWord = getOpenFileName("docx")
        If Equals(pathToOldWord, "") Then
            Return
        End If

        handler_.setPathToOldWord(pathToOldWord)
        Try
            handler_.parseOldWord()
            If chbxExportOld2Excel.Checked Then
                handler_.exportOldWordData()
            End If
            lblOldWord.Text = " : " + pathToOldWord
        Catch exc As Exception
            MsgBox(exc.Message)
        End Try
    End Sub

    Private Sub onLoadNewWord(sender As Object, e As EventArgs) Handles pbLoadNewWord.Click
        Dim pathToNewWord = getOpenFileName("docx")
        If Equals(pathToNewWord, "") Then
            Return
        End If

        handler_.setPathToNewWord(pathToNewWord)
        Try
            handler_.parseNewWord()
            If chbxExportNew2Excel.Checked Then
                handler_.exportNewWordData()
            End If
            lblNewWord.Text = " : " + pathToNewWord
        Catch exc As Exception
            MsgBox(exc.Message)
        End Try
    End Sub

    Private Sub onCompareAndExport(sender As Object, e As EventArgs) Handles pbCompareAndExport.Click
        Try
            handler_.compareHashTables()
            exportToExcel()
        Catch exc As Exception
            MsgBox(exc.Message)
        End Try

    End Sub

    Private Sub exportToExcel()
        Dim pathToExcel = getSaveFileName("xlsx")
        If Equals(pathToExcel, "") Then
            Return
        End If

        Try
            handler_.exportDataToExcel(pathToExcel)
        Catch exc As Exception
            MsgBox(exc.Message)
        End Try
    End Sub

    Private Sub onClean(sender As Object, e As EventArgs) Handles pbClean.Click
        handler_ = New Handler()

        ' Return labels to previous state
        lblOldWord.Text = " : "
        lblNewWord.Text = " : "
        lblLoadPDBExcel.Text = " : "
        pbCompareWithPDBAndExport.Enabled = False
    End Sub

    Private Sub onLoadPDBExcel(sender As Object, e As EventArgs) Handles pbLoadPDBExcel.Click
        Dim pathToPDBExcel = getOpenFileName("xlsx")
        If Equals(pathToPDBExcel, "") Then
            Return
        End If

        handler_.setPathToPDBExcel(pathToPDBExcel)
        Try
            handler_.parsePDBExcel()
            lblLoadPDBExcel.Text = " : " + pathToPDBExcel
            pbCompareWithPDBAndExport.Enabled = True
        Catch exc As Exception
            MsgBox(exc.Message)
        End Try
    End Sub

    Private Sub onCompareWithPDBAndExport(sender As Object, e As EventArgs) Handles pbCompareWithPDBAndExport.Click
        Try
            handler_.compareHandledDataWithPDB()
            Dim pathToResultExcel = getSaveFileName("xlsx")
            If Equals(pathToResultExcel, "") Then
                Return
            End If
            handler_.exportComparisonPDBToExcel(pathToResultExcel)
        Catch exc As Exception
            MsgBox(exc.Message)
        End Try
    End Sub

    Private Sub tryToEnableComparePDBButton()
        If (Equals(lblOldWord.Text, " : ")) Then
            Return
        End If

        If (Equals(lblNewWord.Text, " : ")) Then
            Return
        End If

        pbLoadPDBExcel.Enabled = True
    End Sub

    Private Function getOpenFileName(extension As String) As String
        Dim dialog As New OpenFileDialog
        dialog.Filter = "Файлы " + "(*." + extension + ")|*." + extension
        dialog.ShowDialog()
        Return dialog.FileName
    End Function

    Private Function getSaveFileName(extension As String) As String
        Dim dialog As New SaveFileDialog
        dialog.Filter = "Файлы " + "(*." + extension + ")|*." + extension
        dialog.ShowDialog()
        Return dialog.FileName
    End Function

End Class
