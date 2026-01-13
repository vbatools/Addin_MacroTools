Attribute VB_Name = "modAddinPubFun"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : modPublicFunctions - global public functions
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : FileHave - Checks if a file exists
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* sPath As String                                : - Path to the file to check
'* Optional Atributes As FileAttribute = vbNormal : - File attributes to check
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function FileHave(ByVal Patch As String) As Boolean
    FileHave = (Dir(Patch, vbDirectory) <> vbNullString)
    If Patch = vbNullString Then FileHave = False
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetFileName - Returns the file name (with extension) from a full path
'* Created    : 22-03-2023 14:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByVal sPathFile As String : - Path to the file
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function sGetFileName(ByVal sPathFile As String) As String
    Dim fso         As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    sGetFileName = fso.GetFileName(sPathFile)
    Set fso = Nothing
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WorkbookIsOpen - Checks if a workbook with the specified name is open
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByRef WBName As String : Name of the workbook
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function WorkbookIsOpen(ByRef WBName As String) As Boolean
    Dim wb          As Workbook
    On Error Resume Next
    Set wb = Workbooks(WBName)
    WorkbookIsOpen = Err.Number = 0
End Function
