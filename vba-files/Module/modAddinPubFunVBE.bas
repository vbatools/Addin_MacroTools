Attribute VB_Name = "modAddinPubFunVBE"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetSelectControl - получение выделленого контрола на форме в конструкторе VBE
'* Created    : 22-03-2023 16:01
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* Optional bUserForm As Boolean = False :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetSelectControl(Optional bUserForm As Boolean = False) As Object
    On Error GoTo ErrorHandler

    If bUserForm Then
        Dim Form    As UserForm
        Set Form = Application.VBE.SelectedVBComponent.Designer
        If Not Form Is Nothing Then
            Set GetSelectControl = Form
            Exit Function
        End If
    Else
        If Application.VBE.ActiveWindow.Type = vbext_wt_Designer Then
            Dim objActiveModule As VBComponent
            Set objActiveModule = getActiveModule()
            If Not objActiveModule Is Nothing Then
                Dim collControls As Collection
                Set collControls = getSelectedControlsCollection
                If collControls.Count = 1 Then
                    Set GetSelectControl = collControls.Item(1)
                    Set collControls = Nothing
                    Exit Function
                End If
            End If
        End If
    End If

    Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 9:
            Debug.Print "Для работы инструмента, откройте окно View -> Properties Window"
        Case Else:
            Debug.Print "Ошибка! в GetSelectControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
            'Call WriteErrorLog("GetSelectControl")
    End Select
    Err.Clear
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : getSelectedControlsCollection - получение коллекции выделеных контролов на форме в конструкторе VBE
'* Created    : 22-03-2023 16:00
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function getSelectedControlsCollection() As Collection
    Dim ctl         As Control
    Dim coll        As New Collection
    Dim actModule   As VBComponent
    Set actModule = getActiveModule
    If actModule Is Nothing Then Exit Function
    For Each ctl In actModule.Designer.Selected
        Call coll.Add(ctl)
    Next ctl
    Set getSelectedControlsCollection = coll
    Set coll = Nothing
End Function


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : getActiveModule - получение активного модуля VBA
'* Created    : 22-03-2023 16:00
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function getActiveModule() As VBComponent
    Set getActiveModule = Application.VBE.SelectedVBComponent
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetCodeFromModule - получить код из модуля в строковую переменную
'* Created    : 20-04-2020 18:20
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : модуль VBA
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetCodeFromModule(ByRef moCM As VBIDE.CodeModule) As String
    With moCM
        If .CountOfLines > 0 Then GetCodeFromModule = .Lines(1, .CountOfLines)
    End With
End Function
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : SetCodeInModule - загрузить код из строковой переменой в модуль
'* Created    : 20-04-2020 18:21
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByRef objVBComp As VBIDE.VBComponent : модуль VBA
'* ByVal SCode As String                : строковая переменная
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SetCodeInModule(ByRef moCM As VBIDE.CodeModule, ByVal sCode As String)
    With moCM
        If .CountOfLines > 0 Then Call .DeleteLines(1, .CountOfLines)
        Call .InsertLines(1, VBA.Trim$(sCode))
    End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WhatIsTextInComboBoxHave - получение текущего значения в CommboBox понели инструментов редактора VBE
'* Created    : 22-03-2023 14:34
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function WhatIsTextInComboBoxHave(ByVal sTagCombobox As String) As String
    Dim myCommandBar As CommandBar
    Dim cntrl       As CommandBarControl

    Set myCommandBar = Application.VBE.CommandBars(modAddinConst.MENU_TOOLS)
    For Each cntrl In myCommandBar.Controls
        If cntrl.Tag = sTagCombobox Then
            WhatIsTextInComboBoxHave = cntrl.Text
            Exit Function
        End If
    Next cntrl
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : TrimLinesTabAndSpase - удаляем крайние табуляции и пробелы (все строки прижимаются к левому краю):
'* Created    : 22-03-2023 16:23
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByRef CurCodeModule As VBIDE.CodeModule :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub TrimLinesTabAndSpase(ByRef CurCodeModule As VBIDE.CodeModule)
    Dim sLines      As String

    sLines = GetCodeFromModule(CurCodeModule)
    If sLines = vbNullString Then Exit Sub
    Call SetCodeInModule(CurCodeModule, fnTrimLinesTabAndSpase(sLines))
End Sub

Public Function fnTrimLinesTabAndSpase(ByVal strLine As String) As String
    If strLine = vbNullString Then Exit Function
    Dim j           As Long
    Dim arr         As Variant
    Dim sResult     As String
    arr = VBA.Split(strLine, vbNewLine)
    For j = 0 To UBound(arr, 1)
        If sResult <> vbNullString Then sResult = sResult & vbNewLine
        sResult = sResult & VBA.Trim$(arr(j))
    Next j
    fnTrimLinesTabAndSpase = VBA.Trim$(sResult)
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : VBAIsTrusted - проверка доступа к объектной моделе VBA
'* Created    : 22-03-2023 14:33
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function VBAIsTrusted() As Boolean
    On Error GoTo ErrorHandler
    Dim sTxt        As String
    sTxt = Application.VBE.Version
    VBAIsTrusted = True
    Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 1004:
            'If ThisWorkbook.Name = C_Const.NAME_ADDIN & ".xlam" Then
            Call MsgBox("Предупреждение! " & modAddinConst.NAME_ADDIN & vbLf & vbNewLine & _
                    "Отключено: [Доверять доступ к объектной модели VBE]" & vbLf & _
                    "Для включения перейдите: Файл->Параметры->Центр управления безопасностью->Параметры макросов" & _
                    vbLf & vbNewLine & "И перезапустите Excel", vbCritical, "Предупреждение:")
        Case Else:
            Debug.Print "Ошибка! в VBAIsTrusted" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
            'Call WriteErrorLog("VBAIsTrusted")
    End Select
    Err.Clear
    VBAIsTrusted = False
End Function

