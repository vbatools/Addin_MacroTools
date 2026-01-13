Attribute VB_Name = "modUFControlsReName"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1




'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RenameControl - Renames the selected control
'* Created    : 08-10-2020 14:11
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub RenameControl()
    Dim cnt         As Control
    Dim sNewName    As String
    Dim sOldName    As String
    Dim strVar      As String
    Dim moCM        As CodeModule

    On Error GoTo ErrorHandler

    Set cnt = GetSelectControl()
    If cnt Is Nothing Then Exit Sub
    sOldName = cnt.Name
    sNewName = InputBox("Enter new Control name", "Rename Control:", sOldName)
    If sNewName = vbNullString Or sNewName = sOldName Then Exit Sub

    cnt.Name = sNewName
    Set moCM = Application.VBE.SelectedVBComponent.CodeModule
    strVar = GetCodeFromModule(moCM)
    If strVar = vbNullString Then Exit Sub
    Call SetCodeInModule(moCM, ReplceCode(strVar, sOldName, sNewName))

    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 40044:
            Call MsgBox("Error! Cannot rename Control [ " & sNewName & " ], name already exists!", vbCritical, "Cannot rename Control:")
            Exit Sub
        Case -2147319764:
            Call MsgBox("Control name already exists [" & sNewName & " ], name already exists!", vbCritical, "Name already exists:")
            Exit Sub
        Case Else:
            Debug.Print "Error in RenameControl" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
            'Call WriteErrorLog("RenameControl")
    End Select
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ReplceCode - Replaces old control name with new name in code
'* Created    : 26-03-2020 13:11
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByVal sInCode As String : Input code
'* sOldName As String      : Old name
'* sNewName As String      : New name
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function ReplceCode(ByVal sInCode As String, sOldName As String, sNewName As String) As String
    If sInCode = vbNullString Then Exit Function
    Dim sCode       As String
    sCode = sInCode
    sCode = VBA.Replace(sCode, " " & sOldName & ".", " " & sNewName & ".", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, " " & sOldName & ",", " " & sNewName & ",", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, " " & sOldName & ")", " " & sNewName & ")", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, "(" & sOldName & ".", "(" & sNewName & ".", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, "(" & sOldName & ",", "(" & sNewName & ",", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, "=" & sOldName & ".", "=" & sNewName & ".", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, "=" & sOldName & vbNewLine, "=" & sNewName & vbNewLine, , , vbTextCompare)
    sCode = VBA.Replace(sCode, "(" & sOldName & " ", "(" & sNewName & " ", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, "(" & sOldName & ")", "(" & sNewName & ")", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, "." & sOldName & ".", "." & sNewName & ".", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, "." & sOldName & vbNewLine, "." & sNewName & vbNewLine, , , vbTextCompare)
    sCode = VBA.Replace(sCode, " " & sOldName & "_", " " & sNewName & "_", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, vbNewLine & sOldName & "_", vbNewLine & sNewName & "_", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, """ & sOldName & """, """ & sNewName & """, 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, " " & sOldName & " ", " " & sNewName & " ", 1, -1, vbTextCompare)
    sCode = VBA.Replace(sCode, " " & sOldName & vbNewLine, " " & sNewName & vbNewLine, 1, -1, vbTextCompare)
    ReplceCode = sCode
End Function
