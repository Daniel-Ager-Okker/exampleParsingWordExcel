Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic.DateAndTime

Public Class Logger
    Private path2LogFile_ As String
    Private file_ As FileStream

    Public Sub setPath2LogFile(path As String)
        path2LogFile_ = path
    End Sub

    Public Sub startLogging()
        file_ = File.Create(path2LogFile_)
    End Sub

    Public Sub endLogging()
        file_.Close()
    End Sub

    Public Sub setLogMessage(msg As String)
        Dim time = DateAndTime.Now.ToString
        Dim message = time + ":" + msg + vbNewLine
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(message)
        file_.Write(info, 0, info.Length)
    End Sub

End Class