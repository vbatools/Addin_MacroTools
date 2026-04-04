VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOfficeArchiveManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class        :   clsOfficeArchiveManager - Class for safe management of Office Open XML file structure.
'*                                            Provides unpacking, access to internal components (xl, word, ppt),
'*                                            reverse archiving and backup management with file lock control.
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   28-01-2026 09:54
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

' Enum for path settings
Public Enum SettingsValue
    FileFolder
    FileFullName
    FileName
    FileExtension
    FileType
    FileRelsPath
    FileFullNameBackup
    FileVBAProject
    FilePrinterSettings
    FileCustomUI
    FileCustomUI14
    
    FolderPrinterSettings
    FolderUnzipped    ' Folder with unpacked files
    FolderZip    ' Temporary archive file
    FolderRels
    FolderXl
    FolderCustomUI
    FolderDocProps
    'for excel
    ExlFolderWorkSheets
    ExlFileWorkBook
    ExlFileWorkBookRels

    [_First] = FileFolder
    [_Last] = ExlFileWorkBook
End Enum

Public Enum StreamTypeEnum
    adTypeBinary = 1
    adTypeText = 2
End Enum

Private Const PREFIX_UNZIPPED As String = "UnZipped_"
Private Const EXT_ZIP As String = ".zip"
Private Const SUFFIX_BACKUP As String = "(backup)"

' Class variables
Private mFileFullName As String
Private mFileFullNameBackup As String
Private mFileName   As String
Private mFileExtension As String
Private mFileFolder As String
Private mFileType   As String
Private mFileRelsPath As String
Private mUnzippedFolderPath As String    ' Folder for unpacking
Private mAddBackup  As Boolean
Private mIsInitialized As Boolean
Private mFSO        As Object

' Getter for settings
Public Function GetSettings(ByVal enValue As SettingsValue) As String
    If Not mIsInitialized Then Err.Raise vbObjectError + 101, "ZipManager", "Class not initialized"

    Dim sBasePath   As String
    ' Base path to unpacked folder always ends with separator
    sBasePath = mUnzippedFolderPath & Application.PathSeparator

    Select Case enValue
             Case SettingsValue.FolderUnzipped: GetSettings = sBasePath
        Case SettingsValue.FolderZip: GetSettings = mUnzippedFolderPath & EXT_ZIP    ' Temporary zip file
        Case SettingsValue.FolderCustomUI: GetSettings = sBasePath & "customUI" & Application.PathSeparator
        Case SettingsValue.FolderDocProps: GetSettings = sBasePath & "docProps" & Application.PathSeparator
        Case SettingsValue.FolderRels: GetSettings = sBasePath & "_rels" & Application.PathSeparator
        Case SettingsValue.FolderXl: GetSettings = sBasePath & mFileType & Application.PathSeparator

        Case SettingsValue.FileCustomUI: GetSettings = sBasePath & "customUI" & Application.PathSeparator & "customUI.xml"
        Case SettingsValue.FileCustomUI14: GetSettings = sBasePath & "customUI" & Application.PathSeparator & "customUI14.xml"

        Case SettingsValue.FileFullName: GetSettings = mFileFullName
        Case SettingsValue.FileFullNameBackup: GetSettings = mFileFullNameBackup
        Case SettingsValue.FileName: GetSettings = mFileName
        Case SettingsValue.FileExtension: GetSettings = mFileExtension
        Case SettingsValue.FileFolder: GetSettings = mFileFolder
        Case SettingsValue.FileType: GetSettings = mFileType
        Case SettingsValue.FileRelsPath: GetSettings = mFileRelsPath
        Case SettingsValue.FileVBAProject
            If mFileType = "xl" Then GetSettings = sBasePath & mFileType & Application.PathSeparator & "vbaProject.bin"
            'for excel
        Case SettingsValue.FolderPrinterSettings
            If mFileType = "xl" Then GetSettings = sBasePath & mFileType & Application.PathSeparator & "printerSettings"
        Case SettingsValue.FilePrinterSettings
            If mFileType = "xl" Then GetSettings = sBasePath & mFileType & Application.PathSeparator & "printerSettings.bin"
            'for excel
        Case SettingsValue.ExlFileWorkBookRels:
            If mFileType = "xl" Then GetSettings = sBasePath & mFileType & Application.PathSeparator & "_rels\workbook.xml.rels"
        Case SettingsValue.ExlFolderWorkSheets:
            If mFileType = "xl" Then GetSettings = sBasePath & mFileType & Application.PathSeparator & "worksheets" & Application.PathSeparator
        Case SettingsValue.ExlFileWorkBook:
            If mFileType = "xl" Then GetSettings = sBasePath & mFileType & Application.PathSeparator & "workbook.xml"
    End Select
End Function

Public Function addPathToSheet(ByVal idSheet As Long) As String
    addPathToSheet = GetSettings(ExlFolderWorkSheets) & "sheet" & idSheet & ".xml"
End Function

' Initialize class
Public Function Initialize(ByVal sPathFile As String, ByVal addBackup As Boolean) As Boolean
    If Not FileHave(sPathFile, vbNormal) Then
        Call MsgBox("File not found: " & sPathFile, vbExclamation)
        Exit Function
    End If

    Set mFSO = CreateObject("Scripting.FileSystemObject")
    If IsFileLocked(sPathFile) Then
        Call MsgBox("File is open." & vbNewLine & "Close the file to continue: " & sPathFile, vbCritical)
        Set mFSO = Nothing
        Exit Function
    End If

    mFileFullName = sPathFile
    mAddBackup = addBackup

    With mFSO
        mFileFolder = .GetParentFolderName(sPathFile)
        mFileExtension = .GetExtensionName(sPathFile)
        mFileName = .getFileName(sPathFile)
        If mFileName Like "*—*" Then
            mFileName = VBA.Replace(mFileName, "—", "-")
            Name mFileFullName As mFileFolder & Application.PathSeparator & mFileName
            mFileFullName = mFileFolder & Application.PathSeparator & mFileName
        End If
        ' Build path to folder where files will be unpacked (e.g., C:\Path\UnZipped_File.xlsx)
        mUnzippedFolderPath = mFileFolder & Application.PathSeparator & PREFIX_UNZIPPED & mFileName
    End With

    ' Determine office file type
    Select Case VBA.LCase(mFileExtension)
             Case "docm", "dotm", "dotx", "docx", "doc", "dot": mFileType = "word": mFileRelsPath = "_rels" & Application.PathSeparator & "document.xml.rels"
        Case "xlsm", "xlsb", "xlam", "xla", "xlsx", "xls": mFileType = "xl": mFileRelsPath = "_rels" & Application.PathSeparator & "workbook.xml.rels"
        Case "pptx", "pptm", "ppt", "potx", "potm", "pot": mFileType = "ppt": mFileRelsPath = "_rels" & Application.PathSeparator & "presentation.xml.rels"
    End Select

    If mAddBackup Then mAddBackup = AddBackupFile()

    mIsInitialized = True
    Initialize = True
End Function

' Create a backup copy
Private Function AddBackupFile() As Boolean
    mFileFullNameBackup = mFileFullName & SUFFIX_BACKUP
    If FileHave(mFileFullNameBackup, vbNormal) Then Kill mFileFullNameBackup
    FileCopy mFileFullName, mFileFullNameBackup
    AddBackupFile = True
End Function

' Unpack file to folder
Public Function UnZipFile() As Boolean
    If Not mIsInitialized Then Err.Raise vbObjectError + 101, "ZipManager", "Class not initialized"
    On Error GoTo ErrHandler

    Dim sTempZipFile As String
    sTempZipFile = mUnzippedFolderPath & EXT_ZIP
    
    UnZipFile = FileUnZip(mFSO, mUnzippedFolderPath, mFileFullName)


    Exit Function
ErrHandler:
    Debug.Print ">> Error UnZipFile: " & Err.Number & " " & Err.Description
End Function

' Archive folder back to file
Public Function ZipFilesInFolder() As Boolean
    If Not mIsInitialized Then Err.Raise vbObjectError + 101, "ZipManager", "Class not initialized"
    On Error GoTo ErrHandler

    If Not mFSO.FolderExists(mUnzippedFolderPath) Then
        Call MsgBox("Archive folder not found: " & mUnzippedFolderPath, vbCritical)
        Exit Function
    End If

    If IsFileLocked(mFileFullName) Then
        Call MsgBox("File is open. Close the file to continue: " & mFileFullName, vbCritical)
        Exit Function
    End If

    Dim sDestZip    As String
    sDestZip = mFileFullName & EXT_ZIP
    
    ZipFilesInFolder = ZipAllFilesInFolder(mFSO, mUnzippedFolderPath, mFileFullName)
    

    Exit Function
ErrHandler:
    Debug.Print ">> Error ZipFilesInFolder: " & Err.Number & " " & Err.Description
End Function

' Check if file is locked (by rename attempt)
Private Function IsFileLocked(ByVal sFileFullName As String) As Boolean
    If Not mFSO.FileExists(sFileFullName) Then Exit Function

    Const TMP_EXT   As String = ".tmp"
    Dim sTempName   As String
    sTempName = sFileFullName & TMP_EXT

    On Error Resume Next
    Name sFileFullName As sTempName

    Select Case Err.Number
             Case 58, 75: IsFileLocked = True
        Case 0: Name sTempName As sFileFullName
        Case Else: Debug.Print ">> Unexpected lock error: " & Err.Number
    End Select

    On Error GoTo 0
End Function

Private Sub Class_Terminate()
    Set mFSO = Nothing
End Sub

'* XML Block
Public Function getXMLDOC(ByVal sPathXML As String) As MSXML2.DOMDocument
    ' Check if file exists
    If Not FileHave(sPathXML, vbNormal) Then Exit Function

    Dim oXMLDoc     As MSXML2.DOMDocument
    Set oXMLDoc = New MSXML2.DOMDocument
    With oXMLDoc
        .async = False
        .validateOnParse = False
        If Not .Load(sPathXML) Then Exit Function
        If .parseError.ErrorCode <> 0 Then Exit Function
        Set getXMLDOC = oXMLDoc
    End With
End Function

Public Function readXMLFromFile(ByVal sPathXML As String) As String
    Dim oXMLDoc     As MSXML2.DOMDocument
    Set oXMLDoc = getXMLDOC(sPathXML)
    If oXMLDoc Is Nothing Then Exit Function
    readXMLFromFile = oXMLDoc.XML
End Function

Public Sub writeXMLToFile(ByVal sPathXML As String, sXML As String)
    Dim oXMLDoc     As MSXML2.DOMDocument
    Set oXMLDoc = getXMLDOC(sPathXML)
    If oXMLDoc Is Nothing Then Exit Sub
    Call oXMLDoc.loadXML(sXML)
    Call oXMLDoc.Save(sPathXML)
End Sub

Public Function getArraySheetsName() As Variant
    Dim sPathXML    As String
    Dim oXMLDoc     As MSXML2.DOMDocument
    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList

    ' Get path to settings file
    sPathXML = GetSettings(ExlFileWorkBook)

    ' Initialize XML parser
    Set oXMLDoc = getXMLDOC(sPathXML)
    If oXMLDoc Is Nothing Then Exit Function
    Set oXMLNodeList = oXMLDoc.SelectNodes("/workbook/sheets/sheet")

    ' Check that node is found and not empty
    If oXMLNodeList Is Nothing Then Exit Function
    Dim iCount      As Long
    iCount = oXMLNodeList.Length
    If iCount = 0 Then Exit Function

    Dim i           As Long
    ReDim arr(1 To iCount, 1 To 3) As String

    ' Iterate through only real sheet nodes
    For i = 0 To iCount - 1
        With oXMLNodeList(i).Attributes
            arr(i + 1, 1) = .getNamedItem("name").text
            arr(i + 1, 2) = .getNamedItem("sheetId").text
            arr(i + 1, 3) = .getNamedItem("r:id").text
        End With
    Next i
    getArraySheetsName = arr

    Set oXMLNodeList = Nothing
    Set oXMLDoc = Nothing
End Function

Public Function delPasswordSheet(ByVal idSheet As Long) As Boolean
    Dim sPathXML    As String
    ' Get path to sheet settings file
    sPathXML = addPathToSheet(idSheet)
    delPasswordSheet = delXMLNode(sPathXML, "/worksheet/sheetProtection")
End Function

Public Function delPasswordWBook() As Boolean
    Dim sPathXML    As String
    ' Get path to sheet settings file
    sPathXML = GetSettings(ExlFileWorkBook)
    delPasswordWBook = delXMLNode(sPathXML, "/workbook/workbookProtection")
End Function

Private Function delXMLNode(ByVal sPathXML As String, ByVal sXPath As String) As Boolean
    Dim oXMLDoc     As MSXML2.DOMDocument
    Set oXMLDoc = getXMLDOC(sPathXML)
    If oXMLDoc Is Nothing Then Exit Function
    Dim oNode       As MSXML2.IXMLDOMNode

    ' Find the protection node (use SelectSingleNode to search for a single element)
    Set oNode = oXMLDoc.SelectSingleNode(sXPath)
    If Not oNode Is Nothing Then
        Call oNode.ParentNode.RemoveChild(oNode)
        Call oXMLDoc.Save(sPathXML)
        delXMLNode = True
    End If

    Set oNode = Nothing
    Set oXMLDoc = Nothing
End Function

Public Function getBinaryArrayVBAProject(ByVal StreamType As StreamTypeEnum) As Variant
    Dim sPathVBAProject As String
    Dim stream      As Object
    sPathVBAProject = getPathVBAProject()
    If Not FileHave(sPathVBAProject, vbNormal) Then Exit Function
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = StreamType
        If StreamType = adTypeText Then .Charset = "windows-1251"
        .Open
        Call .LoadFromFile(sPathVBAProject)
        If StreamType = adTypeText Then
            getBinaryArrayVBAProject = .ReadText
        Else
            getBinaryArrayVBAProject = .Read
        End If
        .Close
        Set stream = Nothing
    End With
End Function

Public Function putBinaryArrayVBAProject(ByRef arr As Variant, ByVal StreamType As StreamTypeEnum) As Boolean
    Dim sPathVBAProject As String
    Dim stream      As Object
    sPathVBAProject = getPathVBAProject()
    If Not FileHave(sPathVBAProject, vbNormal) Then Exit Function
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = StreamType
        If StreamType = adTypeText Then .Charset = "windows-1251"
        .Open
        If StreamType = adTypeText Then
            Call .WriteText(arr)
        Else
            Call .Write(arr)
        End If
        Call .SaveToFile(sPathVBAProject, 2)
        .Close
        Set stream = Nothing
        putBinaryArrayVBAProject = True
    End With
End Function

Public Function getPathVBAProject() As String
    Dim oXMLDoc     As MSXML2.DOMDocument
    Set oXMLDoc = getXMLDOC(GetSettings(ExlFileWorkBookRels))

    ' Check successful XML object initialization
    If oXMLDoc Is Nothing Then Exit Function

    Dim oRootNode   As MSXML2.IXMLDOMNode
    Set oRootNode = oXMLDoc.SelectSingleNode("Relationships")

    ' Check for root node existence
    If oRootNode Is Nothing Then Exit Function

    Dim oNodeItem   As MSXML2.IXMLDOMNode
    Dim attrType    As MSXML2.IXMLDOMNode
    Dim attrTarget  As MSXML2.IXMLDOMNode

    For Each oNodeItem In oRootNode.ChildNodes
        ' Explicitly get attributes into temporary variables
        Set attrType = oNodeItem.Attributes.getNamedItem("Type")
        Set attrTarget = oNodeItem.Attributes.getNamedItem("Target")

        ' Check for Nothing prevents error 91 if attribute is missing
        If Not attrType Is Nothing And Not attrTarget Is Nothing Then
            If attrType.text Like "*/vbaProject" Then
                getPathVBAProject = GetSettings(FolderXl) & attrTarget.text
                Exit Function
            End If
        End If
    Next oNodeItem
End Function