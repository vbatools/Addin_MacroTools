Attribute VB_Name = "modAddinInstall"
Option Explicit
Option Private Module

Private Type TB
    tbNameTB        As String
    tbArrBody       As Variant
End Type

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : InstallationAddMacro - add-in installation procedure
'* Created    : 22-03-2023 15:14
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub InstallationAddinMacroTools()
    Dim addFolder   As String
    Dim sFullName   As String
    Dim existingAddIn As AddIn

    On Error GoTo InstallationAdd_Err
    addFolder = VBA.Replace(Application.UserLibraryPath & Application.PathSeparator, _
            Application.PathSeparator & Application.PathSeparator, _
            Application.PathSeparator)

    If Dir(addFolder, vbDirectory) = vbNullString Then
        MsgBox "Unfortunately, the program cannot install the add-in on this computer." & vbCrLf & _
                "The add-ins directory is missing." & vbCrLf & _
                "Please contact the program developer.", vbCritical, "Add-in Installation Failed"
        Exit Sub
    End If

    sFullName = addFolder & modAddinConst.NAME_ADDIN & ".xlam"
    On Error Resume Next
    Set existingAddIn = Application.AddIns(modAddinConst.NAME_ADDIN)
    On Error GoTo InstallationAdd_Err

    Dim tableNames As Variant
    Dim arrTB() As TB
    tableNames = Array(TB_COMMENTS)
    On Error Resume Next
    arrTB = ReadTableDataIntoTBArray(Workbooks(modAddinConst.NAME_ADDIN & ".xlam"), shSettings.Name, tableNames)
    On Error GoTo 0

    If (Not (Not (arrTB))) = 0 Then
        Debug.Print ">> Warning: Failed to load table data. Update aborted."
    Else
        Call UpdateTablesFromTBArray(ThisWorkbook, shSettings.Name, arrTB)
    End If

    If Not existingAddIn Is Nothing Then
        If existingAddIn.Installed Then
            existingAddIn.Installed = False
        End If
    End If

    If WorkbookIsOpen(modAddinConst.NAME_ADDIN & ".xlam") Then
        MsgBox "The add-in file is already open." & vbCrLf & _
                "It may have been installed previously.", vbCritical, "Add-in Installation Failed"
        Exit Sub
    End If

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ThisWorkbook.SaveAs FileName:=sFullName, FileFormat:=xlOpenXMLAddIn

    Call AddIns.Add(FileName:=sFullName)
    AddIns(modAddinConst.NAME_ADDIN).Installed = True

    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "The program has been successfully installed!" & vbCrLf & _
            "Please open or create a new document.", vbInformation, _
            "Add-in Installation:" & modAddinConst.NAME_ADDIN

    ThisWorkbook.Close False
    Exit Sub

InstallationAdd_Err:
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    If Err.Number = 1004 Then
        MsgBox "To install the add-in, please close this file and run it again.", _
                vbCritical, "Installation:"
    Else
        Call WriteErrorLog("InstallationAddinMacroTools", True)
    End If
End Sub

Private Function ReadTableDataIntoTBArray(ByRef targetWorkbook As Workbook, _
                            ByRef shName As String, _
                            ByRef tableNames As Variant) As TB()

    If Not IsArray(tableNames) Then
        Err.Raise vbObjectError + 1, "GetTBLists", "An array of table names is expected. Received:" & TypeName(tableNames)
    End If

    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = targetWorkbook.Worksheets(shName)
    On Error GoTo 0

    If targetSheet Is Nothing Then
        Err.Raise vbObjectError + 2, "GetTBLists", "Sheet '" & shName & "' not found in workbook '" & targetWorkbook.Name & "'."
    End If

    Dim tableCount As Long
    tableCount = UBound(tableNames)

    Dim arrTB() As TB
    ReDim arrTB(0 To tableCount)

    Dim i As Long
    Dim tableName As String
    Dim tbl As ListObject

    For i = 0 To tableCount
        tableName = tableNames(i)
        With arrTB(i)
            .tbNameTB = tableName
            If IsTableExists(targetSheet, tableName) Then
                Set tbl = targetSheet.ListObjects(tableName)
                If Not tbl.DataBodyRange Is Nothing Then
                    .tbArrBody = tbl.DataBodyRange.Value2
                Else
                    .tbArrBody = Empty
                End If
            Else
                Debug.Print ">> Warning: Table '" & tableName & "' not found in sheet '" & targetSheet.Name & "' in workbook '" & targetWorkbook.Name & "'."
                .tbArrBody = Empty
            End If
        End With
    Next i
    ReadTableDataIntoTBArray = arrTB
End Function

Private Sub UpdateTablesFromTBArray(ByRef targetWorkbook As Workbook, _
                               ByRef shName As String, _
                               ByRef arrTB() As TB)

    Dim targetSheet As Worksheet
    On Error Resume Next
    Set targetSheet = targetWorkbook.Worksheets(shName)
    On Error GoTo 0

    If targetSheet Is Nothing Then
        Err.Raise vbObjectError + 2, "loadTBListsInTable", "Sheet '" & shName & "' not found in workbook '" & targetWorkbook.Name & "'."
    End If

    Dim i As Long
    Dim tableName As String
    Dim tbl As ListObject
    Dim data As Variant

    For i = LBound(arrTB) To UBound(arrTB)
        tableName = arrTB(i).tbNameTB
        data = arrTB(i).tbArrBody
        If IsTableExists(targetSheet, tableName) Then
            Set tbl = targetSheet.ListObjects(tableName)
            If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
            tbl.Range(2, 1).Resize(UBound(data, 1), UBound(data, 2)).Value2 = data
            Debug.Print ">> Table: " & tableName & "— updated"
        Else
            Debug.Print ">> Warning: Table '" & tableName & "' not found on sheet '" & targetSheet.Name & "'."
        End If
    Next i
End Sub
