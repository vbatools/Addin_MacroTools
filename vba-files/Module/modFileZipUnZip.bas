Attribute VB_Name = "modFileZipUnZip"
Option Explicit
Option Private Module

Private Const TYPE_FILES    As String = "*.xlsm;*.xlsb;*.xlam;*.xlsx;*.docm;*.dotm;*.dotx;*.docx;*.pptx;*.pptm;*.potx;*.potm"
Private Const EXT_ZIP As String = ".zip"

Public Sub UnZipFile()
    Dim arrFiles()  As String
    arrFiles = fileDialogFun(ActiveWorkbook.Path, False, TYPE_FILES)
    If (Not (Not (arrFiles))) = 0 Then Exit Sub
    Dim clsZIP      As clsOfficeArchiveManager
    Set clsZIP = New clsOfficeArchiveManager
    With clsZIP
        If .Initialize(arrFiles(1, 1), True) Then
            If .UnZipFile Then Call MsgBox("File unpacked!", vbInformation)
        End If
    End With
    Set clsZIP = Nothing
End Sub

Public Sub ZipFile()
    Dim arrFiles()  As String
    arrFiles = fileDialogFun(ActiveWorkbook.Path, False, TYPE_FILES)
    If (Not (Not (arrFiles))) = 0 Then Exit Sub
    Dim clsZIP      As clsOfficeArchiveManager
    Set clsZIP = New clsOfficeArchiveManager
    With clsZIP
        If .Initialize(arrFiles(1, 1), True) Then
            If .ZipFilesInFolder Then Call MsgBox("File packed!", vbInformation)
        End If
    End With
    Set clsZIP = Nothing
End Sub

Public Sub addListInFileFiles()
    Dim arrFiles()  As String
    arrFiles = fileDialogFun(ActiveWorkbook.Path, False, TYPE_FILES)
    If (Not (Not (arrFiles))) = 0 Then Exit Sub
    Dim clsZIP      As clsOfficeArchiveManager
    Set clsZIP = New clsOfficeArchiveManager
    With clsZIP
        If .Initialize(arrFiles(1, 1), True) Then
            If .UnZipFile Then
                Dim arrFilesTable As Variant
                arrFilesTable = GetFilesTable(.GetSettings(FolderUnzipped))
                If Not IsEmpty(arrFilesTable) Then
                    Dim iCount As Long
                    iCount = UBound(arrFilesTable, 1)
                    If iCount > 0 Then
                        Dim bReferenceStyle As Boolean
                        Dim sh As Worksheet: Set sh = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
                        ' create table headers
                        With sh
                            With .Cells(2, 1).Resize(1, 4)
                                .value = Array("FILE NAME", "FULL PATH", "SIZE (BYTES)", "MODIFICATION DATE")
                                .Font.Bold = True: .Interior.ColorIndex = 17
                            End With
                            .Cells(3, 1).Resize(iCount, 4).value = arrFilesTable
                            If Application.ReferenceStyle = xlR1C1 Then
                                Application.ReferenceStyle = xlA1
                                bReferenceStyle = True
                            End If
                            .Range("A:D").EntireColumn.AutoFit
                            .Cells(1, 1).Value2 = "FILE:"
                            .Cells(1, 2).Value2 = arrFiles(1, 1)
                            If bReferenceStyle Then Application.ReferenceStyle = xlR1C1
                        End With
                    End If
                End If
                If .ZipFilesInFolder Then
                End If
            End If
        End If
    End With
    Set clsZIP = Nothing
End Sub

Public Function ZipAllFilesInFolder(ByRef mFSO As Object, ByVal mUnzippedFolderPath As String, ByVal mFileFullName As String) As Boolean
    
    On Error GoTo ErrHandler
    
    Dim sDestZip    As String
    sDestZip = mFileFullName & EXT_ZIP

    ' 1. Prepare an empty ZIP container next to the original
    If mFSO.FileExists(sDestZip) Then mFSO.DeleteFile sDestZip, True
    Call CreateEmptyZipFile(mFSO, sDestZip)

    ' 2. Copy folder contents into ZIP
    If CopyItemsShell(mUnzippedFolderPath, sDestZip) Then
        ' 3. Replace the original
        Call DeleteFolderSafe(mFSO, mUnzippedFolderPath)
        mFSO.DeleteFile mFileFullName, True
        Name sDestZip As mFileFullName
        ZipAllFilesInFolder = True
    End If

    Exit Function
ErrHandler:
    Debug.Print ">> Error ZipAllFilesInFolder: " & Err.Number & " " & Err.Description
End Function

' Extract file to a folder
Public Function FileUnZip(ByRef mFSO As Object, ByVal mUnzippedFolderPath As String, ByVal mFileFullName As String) As Boolean
    
    On Error GoTo ErrHandler
    
    If mFSO Is Nothing Then Set mFSO = CreateObject("Scripting.FileSystemObject")
    Dim sTempZipFile As String
    sTempZipFile = mUnzippedFolderPath & EXT_ZIP

    ' 1. Cleanup
    If mFSO.FolderExists(mUnzippedFolderPath) Then Call DeleteFolderSafe(mFSO, mUnzippedFolderPath)
    If mFSO.FileExists(sTempZipFile) Then mFSO.DeleteFile sTempZipFile, True

    ' 2. Preparation: copy original and rename to .zip
    FileCopy mFileFullName, mUnzippedFolderPath    ' Create a copy with the folder name
    Name mUnzippedFolderPath As sTempZipFile      ' Rename the copy to .zip

    ' 3. Create folder and extract
    MkDir mUnzippedFolderPath

    If CopyItemsShell(sTempZipFile, mUnzippedFolderPath) Then
        FileUnZip = True
        ' Delete temporary zip
        If mFSO.FileExists(sTempZipFile) Then mFSO.DeleteFile sTempZipFile, True
    End If
    
    Exit Function
ErrHandler:
    Debug.Print ">> Error ZipAllFilesInFolder: " & Err.Number & " " & Err.Description
End Function

' Create a ZIP file stub
Private Sub CreateEmptyZipFile(ByRef mFSO As Object, ByVal zipPath As String)
    Dim ts          As Object
    Set ts = mFSO.CreateTextFile(zipPath, True)
    ts.Write "PK" & Chr(5) & Chr(6) & String(18, 0)
    ts.Close
    Set ts = Nothing
End Sub
' Universal copying via Shell (Source -> Dest)
' Works for both Zip->Folder and Folder->Zip
Private Function CopyItemsShell(ByVal sSourcePath As String, ByVal sDestPath As String) As Boolean
    On Error GoTo ErrHandler
    Dim objShell    As Object
    Dim itemsCount  As Long
    Dim k           As Long

    Set objShell = CreateObject("Shell.Application")

    itemsCount = objShell.Namespace(VBA.CVar(sSourcePath)).Items.Count

    If itemsCount > 0 Then
        ' Flag 20: Overwrite without dialogs
        objShell.Namespace(VBA.CVar(sDestPath)).CopyHere objShell.Namespace(VBA.CVar(sSourcePath)).Items, 20

        ' Wait for completion (check file count)
        Do Until objShell.Namespace(VBA.CVar(sDestPath)).Items.Count = itemsCount
            Call Sleep(200)
            k = k + 1
            If k > 60 Then    ' Timeout 1 minute
                Debug.Print ">> Timeout copying items: " & sSourcePath
                Exit Function
            End If
        Loop
        CopyItemsShell = True
    Else
        ' If source is empty, consider it a success
        CopyItemsShell = True
    End If

    Set objShell = Nothing
    Exit Function
ErrHandler:
    Debug.Print ">> Error CopyItemsShell: " & Err.Description
    CopyItemsShell = False
End Function

' Safe folder deletion
Private Sub DeleteFolderSafe(ByRef mFSO As Object, ByVal sFolderPath As String)
    On Error Resume Next
    If mFSO.FolderExists(sFolderPath) Then
        mFSO.DeleteFolder sFolderPath, True
        ' FSO sometimes doesn't delete immediately, check and force clean if still there
        If mFSO.FolderExists(sFolderPath) Then
            ' Recursive deletion via script (more reliable than Kill)
            Dim scriptObj As Object
            Set scriptObj = CreateObject("WScript.Shell")
            scriptObj.Run "cmd /c rd /s /q """ & sFolderPath & """", 0, True
        End If
    End If
    On Error GoTo 0
End Sub