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
'* Function   : FileHave - проверка существования файла
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* sPath As String                                : - строка, путь к файлу или папке
'* Optional Atributes As FileAttribute = vbNormal : - тип проверки, файл или папка
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function FileHave(ByVal Patch As String) As Boolean
    FileHave = (Dir(Patch, vbDirectory) <> vbNullString)
    If Patch = vbNullString Then FileHave = False
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetFileName - возвращает имя (с расширением) последнего компонента в заданном пути.
'* Created    : 22-03-2023 14:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                 Description
'*
'* ByVal sPathFile As String : - строка, путь к файлу
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function sGetFileName(ByVal sPathFile As String) As String
    Dim fso         As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    sGetFileName = fso.GetFileName(sPathFile)
    Set fso = Nothing
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WorkbookIsOpen - Возвращает ИСТИНА если открыта книга под названием wname
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByRef WBName As String : Имя книги
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function WorkbookIsOpen(ByRef WBName As String) As Boolean
    Dim wb          As Workbook
    On Error Resume Next
    Set wb = Workbooks(WBName)
    WorkbookIsOpen = Err.Number = 0
End Function

