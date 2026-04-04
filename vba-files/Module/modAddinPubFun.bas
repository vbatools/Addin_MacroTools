Attribute VB_Name = "modAddinPubFun"
Option Explicit
Option Private Module

#If VBA7 Then
    ' For 64-bit systems
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Enum enumParametrVersion
    enName = 1
    enAuthor
    enVersion
    enLicense
    enDateOfCreation
    enDateOfUpdate
    enDescription
    enAll
    [_First] = enName
    [_Last] = enAll
End Enum

Public Function Version(ByVal Parametr As enumParametrVersion) As String
    Dim sRes        As String
    Dim arr         As Variant
    arr = shSettings.ListObjects(TB_ABOUT).DataBodyRange.Value2
    Select Case Parametr
        Case enumParametrVersion.enAll:
            Dim i   As Byte
            For i = enumParametrVersion.[_First] To enumParametrVersion.[_Last] - 1
                If sRes <> vbNullString Then sRes = sRes & vbNewLine
                If arr(i, 3) = 1 Then arr(i, 2) = VBA.Format$(arr(i, 2), FORMAT_DATE)
                sRes = sRes & arr(i, 1) & ": " & arr(i, 2)
            Next i
        Case Else:
            If arr(Parametr, 3) = 1 Then arr(i, 2) = VBA.Format$(arr(Parametr, 2), FORMAT_DATE)
            sRes = arr(Parametr, 1) & ": " & arr(Parametr, 2)
    End Select
    Version = sRes
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : modPublicFunctions - global public functions
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Public Function getCommentVBATools() As String
    Const STR_1     As String = "' __      ______       _______         ,_ "
    Const STR_2     As String = "' \ \    / /  _ \   /\|__   __|        | |"
    Const STR_3     As String = "'  \ \  / /| |_) | /  \  | | ___   ___ | |___"
    Const STR_4     As String = "'   \ \/ / |  _ < / /\ \ | |/ _ \ / _ \| / __|"
    Const STR_5     As String = "'    \  /  | |_) / ____ \| | (_) | (_) | \__ \"
    Const STR_6     As String = "'     \/   |____/_/    \_\_|\___/ \___/|_|___/"
    getCommentVBATools = STR_1 & vbNewLine & STR_2 & vbNewLine & STR_3 & vbNewLine & STR_4 & vbNewLine & STR_5 & vbNewLine & STR_6 & vbNewLine
End Function

Public Sub URLLinks(ByVal url_str As String)
    If VBA.Len(url_str) = 0 Then Exit Sub
    On Error GoTo ErrorHandler

    Dim appEX       As Object
    Set appEX = CreateObject("Wscript.Shell")
    appEX.Run url_str
    Set appEX = Nothing
    Exit Sub
ErrorHandler:
    Select Case Err
        Case Else:
            Call MsgBox("An error occurred in URLLinks" & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly + vbCritical, "Error in URLLinks")
    End Select
    Set appEX = Nothing
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : FileHave - checks if a file exists
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* sPath As String                                : - string, path to file or folder
'* Optional Atributes As FileAttribute = vbNormal : - check type, file or folder
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function FileHave(ByVal Path As String, ByVal fileAttribute As VbFileAttribute) As Boolean
    Dim fso         As Object

    ' Check for empty
    If Path = vbNullString Then Exit Function
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Depending on the IsFolder parameter, choose the check method
    Select Case fileAttribute
             Case VbFileAttribute.vbDirectory
            ' Look for folder
            FileHave = fso.FolderExists(Path)
        Case VbFileAttribute.vbNormal
            ' Look for file
            FileHave = fso.FileExists(Path)
    End Select
    ' Free memory
    Set fso = Nothing
End Function

Public Function sGetBaseName(ByVal sPathFile As String) As String
    'sPathFile - string, path.
    'Returns the name (without extension) of the last component in the specified path.
    Dim fso         As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    sGetBaseName = fso.GetBaseName(sPathFile)
    Set fso = Nothing
End Function

Public Function sGetExtensionName(ByVal sPathFile As String) As String
    'sPathFile - string, path.
    'Returns the extension of the last component in the specified path.
    Dim fso         As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    sGetExtensionName = fso.GetExtensionName(sPathFile)
    Set fso = Nothing
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetFileName - returns the name (with extension) of the last component in the specified path.
'* Created    : 22-03-2023 14:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByVal sPathFile As String : - string, path to file
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function sGetFileName(ByVal sPathFile As String) As String
    Dim fso         As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    sGetFileName = fso.getFileName(sPathFile)
    Set fso = Nothing
End Function

Public Function MoveFile(OldFile As String, NewPathFile As String) As Boolean
    'Move file
    Dim objFso As Object, objFile As Object
    If Dir(OldFile, 16) = vbNullString Then Exit Function
    'move the file
    Set objFso = CreateObject("Scripting.FileSystemObject"): Set objFile = objFso.GetFile(OldFile)
    objFile.Copy NewPathFile
    Set objFile = Nothing: Set objFso = Nothing
    MoveFile = True
End Function

Public Function sGetParentFolderName(ByVal sPathFile As String) As String
    'sPathFile - string, path.
    'Returns the path to the last component in the specified path (its directory).
    Dim fso         As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    sGetParentFolderName = fso.GetParentFolderName(sPathFile)
    Set fso = Nothing
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WorkbookIsOpen - Returns TRUE if the workbook named wname is open
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByRef WBName As String : Workbook name
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function WorkbookIsOpen(ByRef WBName As String) As Boolean
    Dim wb          As Workbook
    On Error Resume Next
    Set wb = Workbooks(WBName)
    WorkbookIsOpen = Err.Number = 0
    On Error GoTo 0
End Function

Public Function IsTableExists(ByVal sh As Worksheet, ByVal sTBName As String) As Boolean
    On Error Resume Next
    IsTableExists = Not sh.ListObjects(sTBName) Is Nothing
    On Error GoTo 0
End Function

Public Sub base64ToFile(ByVal sHashBase64 As String, ByVal sFilePath As String)
    Dim byteArr()   As Byte
    Dim oBase       As Object
    Set oBase = CreateObject("MSXML2.DOMDocument").createElement("b64")
    With oBase
        .DataType = "bin.base64"
        .text = sHashBase64
        byteArr = .nodeTypedValue
    End With
    Open sFilePath For Binary Access Write As #1
    Put #1, 1, byteArr
    Close #1
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function     :   addTabelFromArray - convert array to markdown table
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   22-01-2026 10:24
'* Argument(s)  :               Description
'*
'* ByRef arr As Variant       :
'* ByVal sDelimetr As String  :
'* ByVal haveHeder As Boolean :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function addTabelFromArray(ByRef arr As Variant, ByVal sDelimetr As String, ByVal haveHeder As Boolean, ByVal iEndCol As Byte) As String
    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim k           As Long
    Dim iMaxLen     As Long
    Dim Shift       As Byte
    Dim shift_2     As Byte

    iCount = UBound(arr, 1)
    If LBound(arr, 1) = 1 Then
        Shift = 1
    Else
        shift_2 = 1
    End If

    If haveHeder Then
        ReDim arrCol(0 To iCount + shift_2) As Variant
        ReDim arrLen(0 To iCount + Shift + shift_2) As Integer
    Else
        ReDim arrCol(0 To iCount - Shift) As Variant
        ReDim arrLen(0 To iCount + shift_2) As Integer
    End If

    If iEndCol = 0 Then iEndCol = UBound(arr, 2)
    iEndCol = iEndCol - shift_2

    For j = LBound(arr, 2) To iEndCol
        For i = LBound(arr, 1) To iCount
            k = k + 1
            If j > Shift Then arrCol(k - 1) = arrCol(k - 1) & VBA.String$(arrLen(iCount + shift_2) - arrLen(i - Shift), " ") & " " & sDelimetr & " "
            If j = Shift Then arrCol(k - 1) = " " & sDelimetr & " " & arrCol(k - 1)

            arrCol(k - 1) = arrCol(k - 1) & arr(i, j)
            arrLen(i - Shift) = VBA.Len(arr(i, j))
            If arrLen(i - Shift) > iMaxLen Then iMaxLen = arrLen(i - Shift)
            If haveHeder And i = Shift Then
                k = k + 1
                arrCol(k - 1) = arrCol(k - 1) & VBA.String$(arrLen(iCount + shift_2), "-") & " " & sDelimetr & " "
            End If
        Next i
        arrLen(iCount + shift_2) = iMaxLen
        iMaxLen = 0: k = 0
    Next j
    For i = LBound(arr, 1) To iCount
        k = k + 1
        arrCol(k - 1) = arrCol(k - 1) & VBA.String$(arrLen(iCount + shift_2) - arrLen(i - Shift), " ") & " " & sDelimetr & " "
        If haveHeder And i = Shift Then
            k = k + 1
            arrCol(k - 1) = arrCol(k - 1) & VBA.String$(arrLen(iCount + shift_2), "-") & " " & sDelimetr & " "
        End If
    Next i
    addTabelFromArray = VBA.Join(arrCol, vbNewLine)
End Function

Public Function fileDialogFun(ByVal sPath As String, _
        ByRef bMultiSelect As Boolean, _
        Optional sExpansion As String = "*.xlsm;*.xlsb;*.xlsx") As String()

    If sPath = vbNullString Or Not (Dir(sPath, vbDirectory) <> vbNullString) Then sPath = ThisWorkbook.Path

    Dim oFd         As FileDialog
    Set oFd = Application.FileDialog(msoFileDialogFilePicker)
    With oFd
        .AllowMultiSelect = bMultiSelect
        'dialog window title
        .Title = "Select Files:"
        'clear previously set file types
        .Filters.Clear
        'set the ability to select only Excel files
        .Filters.Add "Microsoft Excel Files", sExpansion, 1
        'assign the display folder and default file name
        .InitialFileName = sPath
        'dialog window view (9 variants available)
        .InitialView = msoFileDialogViewDetails
        If .Show = 0 Then
            Call MsgBox("No files selected!", vbCritical, "Select Files:")
            Exit Function
        End If
        Dim iCount  As Integer
        Dim i       As Integer
        iCount = .SelectedItems.Count
        'ReDim arr(1 To iCount) As String
        ReDim arr(1 To iCount, 1 To 1) As String
        For i = 1 To iCount
            'arr(i) = VBA.CStr(.SelectedItems.Item(i))
            arr(i, 1) = VBA.CStr(.SelectedItems.item(i))
        Next
    End With
    fileDialogFun = arr

    'For procedure
    'If (Not (Not (v))) = 0 Then Exit Sub
End Function

' Main function: Returns a two-dimensional array with file information
Function GetFilesTable(ByVal folderPath As String) As Variant
    Dim fso         As Object
    Dim folder      As Object
    Dim fileCount   As Long
    Dim varResult   As Variant
    Dim rowIndex    As Long

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 1. Check if folder exists
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Specified folder does not exist:" & folderPath, vbExclamation, "Error"
        GetFilesTable = Array()    ' Return empty array
        Exit Function
    End If

    Set folder = fso.GetFolder(folderPath)

    ' 2. First pass: Count files
    fileCount = CountFilesRecursive(folder)

    ' If no files found
    If fileCount = 0 Then
        GetFilesTable = Array()
        Exit Function
    End If

    ' 3. Initialize two-dimensional array
    ' Rows: from 1 to fileCount
    ' Columns: from 1 to 4 (Path, Name, Size, Date)
    ReDim varResult(1 To fileCount, 1 To 4)

    ' 4. Second pass: Fill array
    rowIndex = 1    ' Start filling from the first row
    Call FillFilesTableRecursive(folder, varResult, rowIndex)

    ' Return filled table
    GetFilesTable = varResult

    ' Free memory
    Set folder = Nothing
    Set fso = Nothing
End Function

' Helper function: Recursive file count
Private Function CountFilesRecursive(ByVal currentFolder As Object) As Long
    Dim subFolder   As Object
    Dim lCount       As Long

    lCount = currentFolder.Files.Count

    For Each subFolder In currentFolder.SubFolders
        lCount = lCount + CountFilesRecursive(subFolder)
    Next subFolder

    CountFilesRecursive = lCount
End Function

' Helper procedure: Recursive array filling
Private Sub FillFilesTableRecursive(ByVal currentFolder As Object, ByRef tableArray As Variant, ByRef currentRow As Long)
    Dim subFolder   As Object
    Dim file        As Object

    ' Fill data for each file in the current folder
    For Each file In currentFolder.Files
        tableArray(currentRow, 1) = file.Name                ' File name
        tableArray(currentRow, 2) = file.Path                ' Full path
        tableArray(currentRow, 3) = file.Size                ' Size (bytes)
        tableArray(currentRow, 4) = file.DateLastModified    ' Modification date

        currentRow = currentRow + 1
    Next file

    ' Recursive call for subfolders
    For Each subFolder In currentFolder.SubFolders
        Call FillFilesTableRecursive(subFolder, tableArray, currentRow)
    Next subFolder
End Sub

Public Function GetArrayFromDictionary(ByRef oDic As Dictionary) As Variant
    If oDic Is Nothing Then Exit Function
    If oDic.Count = 0 Then Exit Function

    Dim lRowCount   As Long
    Dim lColCount   As Long
    Dim arrRes()    As String
    Dim arrItems    As Variant
    Dim arrItem     As Variant
    Dim i As Long, j As Long

    lRowCount = oDic.Count
    arrItems = oDic.Items

    On Error Resume Next
    lColCount = UBound(arrItems(0), 2)
    On Error GoTo 0

    If lColCount = 0 Then Exit Function
    ReDim arrRes(1 To lRowCount, 1 To lColCount)
    For i = 0 To lRowCount - 1
        arrItem = arrItems(i)

        For j = 1 To lColCount
            arrRes(i + 1, j) = arrItem(1, j)
        Next j
    Next i
    GetArrayFromDictionary = arrRes
End Function

' Procedure for outputting data to a sheet
Public Sub OutputResults(ByRef wb As Workbook, ByRef sSheetName As String, ByRef sNameVBAProj As String, ByRef arrOutput As Variant)
    If IsEmpty(arrOutput) Then Exit Sub
    Dim wsTarget    As Worksheet

    On Error Resume Next
    Set wsTarget = wb.Worksheets(sSheetName)
    On Error GoTo 0

    If wsTarget Is Nothing Then
        Set wsTarget = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsTarget.Name = sSheetName
    Else
        wsTarget.Cells.ClearContents
    End If

    Application.ScreenUpdating = False
    Dim iCol        As Integer
    Dim iRow        As Integer
    Dim i           As Integer

    On Error Resume Next
    iCol = UBound(arrOutput, 2)
    If LBound(arrOutput, 2) = 0 Then iCol = iCol + 1
    If Err.Number <> 0 Then iCol = 1
    On Error GoTo 0

    iRow = UBound(arrOutput, 1)
    If LBound(arrOutput, 1) = 0 Then iRow = iRow + 1
    With wsTarget
        If IsArray(arrOutput) Then
            .Cells(3, 1).Resize(iRow, iCol).Value2 = arrOutput
            Rows("3:" & iRow + 3).RowHeight = 15
        End If
        For i = 1 To iCol
            With .Cells(1, i)
                .EntireColumn.AutoFit
                If .ColumnWidth > 50 Then .ColumnWidth = 50
            End With
            .Cells(2, i).Value2 = i
        Next i
        .Cells(1, 1).value = sNameVBAProj

        .Range(.Cells(2, 1), .Cells(2, iCol)).AutoFilter
        .Activate
    End With
End Sub

Public Sub WriteErrorLog(ByVal sNameFunc As String, ByVal bMsgBox As Boolean)

    If bMsgBox Then
        Call MsgBox("Error! In " & sNameFunc & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line" & Erl, vbOKOnly + vbExclamation, "Critical Error:")
    Else
        Debug.Print ">> Error! In " & sNameFunc & vbCrLf & Err.Number & vbCrLf & Err.Description & vbCrLf & "at line" & Erl
    End If

    Dim clsLoger    As clsLogging
    Set clsLoger = New clsLogging
    Call clsLoger.LogError(sNameFunc)
    Set clsLoger = Nothing
End Sub

Public Function FilterArrayByText(ByRef vSourceArray As Variant, ByVal iColSerch As Byte, ByVal sSearch As String) As Variant
    Dim i           As Long
    Dim lCount      As Long
    Dim lRowsStart       As Long
    Dim lRowsEnd       As Long
    Dim lColsStart       As Long
    Dim lColsEnd       As Long
    Dim arrResult() As Variant

    sSearch = VBA.UCase$(sSearch)
    lRowsStart = LBound(vSourceArray, 1)
    lRowsEnd = UBound(vSourceArray, 1)
    lColsStart = LBound(vSourceArray, 2)
    lColsEnd = UBound(vSourceArray, 2)

    ' 1. Count matches for accurate memory allocation
    For i = lRowsStart To lRowsEnd
        If VBA.InStr(1, VBA.UCase$(CStr(vSourceArray(i, iColSerch))), sSearch) > 0 Then
            lCount = lCount + 1
        End If
    Next i

    ' If nothing found, return Empty
    If lCount = 0 Then
        FilterArrayByText = Empty
        Exit Function
    End If

    ' 2. Fill result array
    If lRowsStart = 0 Then lCount = lCount - 1
    ReDim arrResult(lRowsStart To lCount, lColsStart To lColsEnd)
    lCount = 0

    Dim j           As Long
    lCount = lRowsStart
    For i = lRowsStart To lRowsEnd
        If VBA.InStr(1, VBA.UCase$(CStr(vSourceArray(i, iColSerch))), sSearch) > 0 Then
            For j = lColsStart To lColsEnd
                arrResult(lCount, j) = vSourceArray(i, j)
            Next j
            lCount = lCount + 1
        End If
    Next i

    FilterArrayByText = arrResult
End Function

Public Function GetTargetWorkbook(ByRef wb As Workbook, ByVal sCaptionForm As String, ByVal sCaptionBtnOk As String) As Boolean
    With frmListWBOpen
        .Caption = sCaptionForm
        .lbOk.Caption = sCaptionBtnOk
        .Show

        If .cmbMain.value = vbNullString Or .lbRes.Caption = "0" Then Exit Function
        On Error Resume Next
        Set wb = Workbooks(.cmbMain.value)
        If Err.Number <> 0 Then Exit Function
        On Error GoTo 0
    End With
    GetTargetWorkbook = True
End Function