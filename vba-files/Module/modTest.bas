Attribute VB_Name = "modTest"
Option Explicit

Public Sub test()
    Dim cls         As clsToolsVBACodeStatistics
    Set cls = New clsToolsVBACodeStatistics
    Dim JSON        As String
    Dim sPath       As String

    With cls
        JSON = .getJSONCodeBase(ActiveWorkbook)
        sPath = ActiveWorkbook.Path & Application.PathSeparator & "ts.json"
        If FileHave(sPath, vbNormal) Then Call Kill(sPath)
        Call TXTAddIntoTXTFile(sPath, JSON, True)
        Debug.Print JSON
    End With
End Sub


Public Function TXTAddIntoTXTFile(ByVal FileName As String, ByVal txt As String, Optional AddFile As Boolean = True) As Boolean
    'TXTAddIntoTXTFile - logical variable, True - addition succeeded, False - failed
    'FileName - string variable, full file path
    'txt - text to be added to the file
    'AddFile - logical variable, default True, creates the file if it doesn't exist

    Dim fso         As Object
    Dim ts          As Object
    On Error Resume Next: Err.Clear
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.OpenTextFile(FileName, 8, AddFile): ts.Write txt: ts.Close
    TXTAddIntoTXTFile = Err = 0
    Set ts = Nothing: Set fso = Nothing
End Function