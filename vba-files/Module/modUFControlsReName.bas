Attribute VB_Name = "modUFControlsReName"
Option Explicit
Option Private Module

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub          :   RenameControl - Renames a control on a UserForm and updates all references in the VBA project.
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   18-03-2026 17:53
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub RenameControl()
    On Error GoTo ErrorHandler

    Dim selectedControls As Object
    Dim selectedControl As control
    Dim newName     As String
    Dim regExp      As Object

    ' Get selected control
    Set selectedControls = GetSelectControl()
    If selectedControls Is Nothing Then Exit Sub
    Select Case TypeName(selectedControls)
             Case "Controls"
            Set selectedControl = selectedControls.item(0)
        Case Else
            Set selectedControl = selectedControls
    End Select
    Set selectedControls = Nothing

    ' Initialize RegExp engine
    Set regExp = CreateObject("VBScript.RegExp")
    With regExp
        .Global = True
        .IgnoreCase = True
    End With

    newName = InputBox("Enter new Control name", "Renaming Control:", selectedControl.Name)
    Call RenameControlUserForm(regExp, Application.VBE.SelectedVBComponent, selectedControl, newName)
    Set regExp = Nothing

    Exit Sub
ErrorHandler:
    Select Case Err.Number
             Case 40044
            MsgBox "Error! Invalid Control name [" & newName & "], please enter a different name!", vbCritical, "Invalid Control name entered:"
        Case -2147319764
            MsgBox "This Control name is already in use [" & newName & "], please enter a different name!", vbCritical, "Ambiguous name:"
        Case Else
            WriteErrorLog "RenameControl", True
            MsgBox "An unexpected error occurred. See log file for details.", vbExclamation, "Error:"
    End Select
    Err.Clear
    Set regExp = Nothing
End Sub

Public Sub RenameControlUserForm(ByRef regExp As Object, ByRef VBComp As vbComponent, ByRef selectedControl As control, ByRef newName As String)
    Dim oldName     As String
    Dim objCodeModule As codeModule
    Dim codeContent As String

    oldName = selectedControl.Name
    ' Validate input
    If newName = vbNullString Or newName = oldName Then Exit Sub

    ' Rename control
    selectedControl.Name = newName

    ' Get active code module
    Set objCodeModule = VBComp.codeModule
    codeContent = GetCodeFromModule(objCodeModule)
    'If codeContent = vbNullString Then Exit Sub

    ' Replace direct references to the old control name
    codeContent = ReplaceControlReferences(regExp, codeContent, oldName, newName)

    ' Update code module
    UpdateCodeModule objCodeModule, codeContent

    ' Search and replace references in all other components
    Call UpdateAllProjectReferences(regExp, objCodeModule.Parent.Collection.Parent, objCodeModule.Parent.Name, oldName, newName)
End Sub

'----------------------------------------------------------------------------------
' Replaces all occurrences of oldName with newName in the given code content.
' Handles word boundaries to avoid partial matches.
'----------------------------------------------------------------------------------
Public Function ReplaceControlReferences(ByRef regExp As Object, ByVal code As String, ByVal oldName As String, ByVal newName As String) As String
    Dim pattern     As String
    Dim result      As String

    ' Match: [space|(|=|.|\r\n|&"]oldName[space|)|,|.|_|\r\n|&"]
    pattern = "([ \(\=\.\r\n]|"" & )" & oldName & "([ \)\,\.\_\r\n]| & "")"
    result = ReplaceCode(regExp, code, pattern, "$1" & newName & "$2")

    ReplaceControlReferences = result
End Function

'----------------------------------------------------------------------------------
' Updates the entire code module with new content.
'----------------------------------------------------------------------------------
Private Sub UpdateCodeModule(ByRef objCodeModule As codeModule, ByVal newContent As String)
    Dim lineCount   As Long
    lineCount = objCodeModule.CountOfLines

    If lineCount > 0 Then
        objCodeModule.DeleteLines 1, lineCount
        objCodeModule.InsertLines 1, newContent
    End If
End Sub

'----------------------------------------------------------------------------------
' Searches and replaces control references across all components in the VBA project.
' Handles:
'   - FormName.OldControlName
'   - .OldControlName inside With FormName blocks
'   - AliasName.OldControlName (if alias is declared as FormName)
'   - .OldControlName inside With AliasName blocks
'----------------------------------------------------------------------------------
Private Sub UpdateAllProjectReferences(ByRef regExp As Object, ByRef objVBProject As vbProject, ByVal sUserFormName As String, _
        ByVal sOldName As String, ByVal sNewName As String)
    Dim objVBComponent As vbComponent
    Dim objCodeModule As codeModule
    Dim sCodeContent As String
    Dim hasChanges  As Boolean
    Dim aliasName   As String
    Dim pattern     As String

    For Each objVBComponent In objVBProject.VBComponents
        ' Skip the form being renamed
        If objVBComponent.Name = sUserFormName Then GoTo NextComponent

        Set objCodeModule = objVBComponent.codeModule
        If objCodeModule.CountOfLines = 0 Then GoTo NextComponent

        sCodeContent = objCodeModule.Lines(1, objCodeModule.CountOfLines)
        hasChanges = False
        sCodeContent = ReplaceControlName(regExp, sCodeContent, sUserFormName, sOldName, sNewName, hasChanges)
        If hasChanges Then UpdateCodeModule objCodeModule, sCodeContent
NextComponent:
    Next objVBComponent
End Sub

Public Function ReplaceControlName(ByRef regExp As Object, ByVal sCodeContent As String, ByVal sUserFormName As String, _
        ByVal sOldName As String, ByVal sNewName As String, ByRef hasChanges As Boolean) As String

    Dim sPattern As String
    Dim sAliasName As String

    ' ---------------------------------------------------------
    ' 1. Replace sUserFormName.OldControlName > sUserFormName.NewControlName
    ' Use context capture ($1, $2) to preserve delimiters
    ' ---------------------------------------------------------
    sPattern = "(^|\W)(" & sUserFormName & "\." & sOldName & ")(\W|$)"
    If FindMatchCount(regExp, sCodeContent, sPattern) > 0 Then
        sCodeContent = ReplaceCode(regExp, sCodeContent, sPattern, "$1" & sUserFormName & "." & sNewName & "$3")
        hasChanges = True
    End If

    ' ---------------------------------------------------------
    ' 2. Replace .OldControlName inside With sUserFormName blocks
    ' ---------------------------------------------------------
    If FindMatchCount(regExp, sCodeContent, "\bWith\s+" & sUserFormName & "\b") > 0 Then
        ' Look for dot before name, preserving the character before it ($1)
        sPattern = "([^a-zA-Z0-9_]|^)\." & sOldName & "\b"
        If FindMatchCount(regExp, sCodeContent, sPattern) > 0 Then
            sCodeContent = ReplaceCode(regExp, sCodeContent, sPattern, "$1." & sNewName)
            hasChanges = True
        End If
        sPattern = "\." & sOldName & "\."
        If FindMatchCount(regExp, sCodeContent, sPattern) > 0 Then
            sCodeContent = ReplaceCode(regExp, sCodeContent, sPattern, "." & sNewName & ".")
            hasChanges = True
        End If
    End If

    ' ---------------------------------------------------------
    ' 3. Handle alias variables declared as sUserFormName
    ' ---------------------------------------------------------
    sAliasName = FindSubMatch(regExp, sCodeContent, "\b(\w+)\s+As\s+" & sUserFormName & "\b")
    If sAliasName <> vbNullString Then

        ' 3a. Direct reference: sAliasName.OldControlName
        ' Use reliable pattern with context capture
        sPattern = "(^|\W)(" & sAliasName & "\." & sOldName & ")(\W|$)"
        If FindMatchCount(regExp, sCodeContent, sPattern) > 0 Then
            sCodeContent = ReplaceCode(regExp, sCodeContent, sPattern, "$1" & sAliasName & "." & sNewName & "$3")
            hasChanges = True
        End If

        ' 3b. Dot reference inside With sAliasName block
        ' WARNING: This replacement applies to the entire module if a With block is found.
        ' This may lead to false positives if sOldName is used in other contexts.
        ' Uncomment only if you are confident in the uniqueness of names.

        If FindMatchCount(regExp, sCodeContent, "\bWith\s+" & sAliasName & "\b") > 0 Then
            ' Stricter pattern: look for .sOldName (with a dot) to avoid replacing just a variable
            sPattern = "(^|\W)\." & sOldName & "(\W|$)"
            If FindMatchCount(regExp, sCodeContent, sPattern) > 0 Then
                sCodeContent = ReplaceCode(regExp, sCodeContent, sPattern, "$1." & sNewName & "$2")
                hasChanges = True
            End If
            sPattern = "\." & sOldName & "\."
            If FindMatchCount(regExp, sCodeContent, sPattern) > 0 Then
                sCodeContent = ReplaceCode(regExp, sCodeContent, sPattern, "." & sNewName & ".")
                hasChanges = True
            End If
        End If
    End If
    ReplaceControlName = sCodeContent
End Function

'----------------------------------------------------------------------------------
' Returns number of matches for a given pattern.
'----------------------------------------------------------------------------------
Private Function FindMatchCount(ByRef regExp As Object, ByVal text As String, ByVal pattern As String) As Long
    With regExp
        .pattern = pattern
        FindMatchCount = .Execute(text).Count
    End With
End Function

'----------------------------------------------------------------------------------
' Extracts first submatch from a pattern (e.g., alias name from "Dim x As FormName").
'----------------------------------------------------------------------------------
Private Function FindSubMatch(ByRef regExp As Object, ByVal text As String, ByVal pattern As String) As String
    Dim matches     As Object

    With regExp
        .pattern = pattern
        Set matches = .Execute(text)
        If matches.Count > 0 Then
            If matches(0).SubMatches.Count > 0 Then
                FindSubMatch = matches(0).SubMatches(0)
            End If
        End If
    End With
End Function

'----------------------------------------------------------------------------------
' Replaces all occurrences of pattern with replacement using RegExp.
'----------------------------------------------------------------------------------
Private Function ReplaceCode(ByRef regExp As Object, ByVal inputCode As String, ByVal pattern As String, ByVal replacement As String) As String
    If Len(inputCode) = 0 Or Len(pattern) = 0 Then
        ReplaceCode = inputCode
        Exit Function
    End If

    With regExp
        .pattern = pattern
        ReplaceCode = .Replace(inputCode, replacement)
    End With
End Function