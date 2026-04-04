Attribute VB_Name = "modAddinPubFunVBEModule"
Option Explicit
Option Private Module

Public Function exportModuleToFile(ByRef VBComp As VBIDE.vbComponent, ByVal sPath As String) As Boolean
      On Error GoTo Error_Handler
      With VBComp
          sPath = sPath & .Name & TypeExtensionModule(.Type)
          Call .Export(sPath)
          exportModuleToFile = True
        Debug.Print ">> Module: [" & .Name & "] was exported to folder [" & sPath & "]"
    End With
    Exit Function
Error_Handler:
    Debug.Print ">> Error in AddModuleToProject " & vbCrLf & Err.Number & vbCrLf & Err.Description & vbCrLf & "at line" & Erl
End Function

Public Sub CopyModyleVBE()
    Call CopyModuleToProject(Application.VBE.ActiveVBProject, Application.VBE.SelectedVBComponent)
End Sub

Public Function AddModuleToProject(ByRef vbProj As VBIDE.vbProject, _
        ByVal sNameModule As String, _
        ByVal TypeModule As vbext_ComponentType, _
        ByVal sCodeVBA As String, _
        ByVal bAddUniqueName As Boolean) As VBIDE.vbComponent

    If TypeModule = vbext_ct_Document Then
        Debug.Print ">> Module type Sheet or Workbook cannot be created"
        Exit Function
    End If

    Dim VBComp      As VBIDE.vbComponent
    Dim vbCodeModule As VBIDE.codeModule
    Dim sNameModuleCheck As String
    On Error GoTo errMsg
    Set VBComp = vbProj.VBComponents.Add(TypeModule)
    sNameModuleCheck = sNameModule
    If bAddUniqueName Then sNameModuleCheck = AddModuleUniqueName(vbProj, sNameModule)

    VBComp.Name = sNameModuleCheck

    If sCodeVBA <> vbNullString Then
        Set vbCodeModule = VBComp.codeModule
        With vbCodeModule
            Call .InsertLines(.CountOfLines + 1, sCodeVBA)
        End With
    End If

    Debug.Print ">> Module: [" & sNameModuleCheck & "] added to workbook [" & getFileNameOnVBProject(vbProj) & "]"
    Set AddModuleToProject = VBComp
    Exit Function
errMsg:
    Select Case Err.Number
        Case 32813:
            Debug.Print ">> Module: [" & sNameModuleCheck & "] was already added to workbook [" & getFileNameOnVBProject(vbProj) & "]"
            vbProj.VBComponents.Remove VBComp
        Case 76:
            Debug.Print ">> Module: [" & sNameModuleCheck & "] added to workbook:" & ActiveWorkbook.Name & vbCrLf & "File not saved!"
        Case Else:
            Call WriteErrorLog("AddModuleToProject", False)
    End Select
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CopyModyleVBE - copy VBA module to Project Explorer
'* Created    : 22-03-2023 15:09
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CopyModuleToProject(ByRef vbProjPaste As VBIDE.vbProject, ByRef VBCompCopy As VBIDE.vbComponent) As VBIDE.vbComponent
    If VBCompCopy Is Nothing Then Exit Function
    Dim sCodeVBA    As String
    With VBCompCopy
        Select Case VBCompCopy.Type
            Case vbext_ct_MSForm:
                Set CopyModuleToProject = CopyModuleTypeForm(vbProjPaste, VBCompCopy)
            Case vbext_ct_Document:
                If MsgBox("Workbook or Sheet module cannot be copied. Copy to a standard module?", _
                        vbYesNo + vbQuestion, "Copy Module:") = vbNo Then Exit Function
                If .codeModule.CountOfLines <> 0 Then
                    sCodeVBA = .codeModule.Lines(1, .codeModule.CountOfLines)
                    sCodeVBA = VBA.Replace(sCodeVBA, "Option Explicit", vbNullString)
                End If
                Set CopyModuleToProject = AddModuleToProject(vbProjPaste, .Name, vbext_ct_StdModule, sCodeVBA, True)
            Case Else:
                If .codeModule.CountOfLines <> 0 Then
                    sCodeVBA = .codeModule.Lines(1, .codeModule.CountOfLines)
                    sCodeVBA = VBA.Replace(sCodeVBA, "Option Explicit", vbNullString)
                End If
                Set CopyModuleToProject = AddModuleToProject(vbProjPaste, .Name, VBCompCopy.Type, sCodeVBA, True)
        End Select
    End With
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : CopyModuleForm - copy VBA module
'* Created    : 22-03-2023 15:02
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByRef vbProj As VBIDE.VBProject      : destination project
'* ByRef vbComCopy As VBIDE.VBComponent : module to copy
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function CopyModuleTypeForm(ByRef vbProj As VBIDE.vbProject, ByRef VBComp As VBIDE.vbComponent) As VBIDE.vbComponent
    Dim sFullFileName As String
    Dim sNameFile   As String
    Dim sNameMod    As String
    Dim sNameModNew As String
    Dim sNamePath   As String
    Dim sFullFileNameNew As String

    On Error GoTo ErrorHandler

    sFullFileName = vbProj.FileName
    sNameFile = sGetFileName(sFullFileName)
    sNamePath = sGetParentFolderName(sFullFileName)

    sNameMod = VBComp.Name
    sNameModNew = AddModuleUniqueName(vbProj, sNameMod)
    sFullFileNameNew = sNamePath & sNameModNew & ".bas"

    VBComp.Name = sNameModNew
    Call VBComp.Export(FileName:=sFullFileNameNew)
    VBComp.Name = sNameMod
    Call vbProj.VBComponents.Import(FileName:=sFullFileNameNew)
    Call Kill(sFullFileNameNew)
    Call Kill(sNamePath & sNameModNew & ".frx")

    Set CopyModuleTypeForm = VBComp
    Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 76:
            Call MsgBox("File not saved, please save the file!", vbCritical, "Error:")
        Case Else:
            Call WriteErrorLog("CopyModuleTypeForm", False)
    End Select
    Err.Clear
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : AddModuleName - create module name
'* Created    : 22-03-2023 15:01
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                     Description
'*
'* ByRef vbProj As VBIDE.VBProject : VBA project
'* ByVal NameModule As String      : new module name
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function AddModuleUniqueName(ByRef vbProj As VBIDE.vbProject, ByVal NameModule As String) As String
    Dim objCol      As Collection
    Dim VBComp      As VBIDE.vbComponent
    Dim i           As Integer
    Dim bFlag       As Boolean
    Set objCol = New Collection

    For Each VBComp In vbProj.VBComponents
        objCol.Add VBComp.Name, VBComp.Name
    Next VBComp

    On Error Resume Next
    objCol.Add NameModule, NameModule
    If Err = 0 Then
        AddModuleUniqueName = NameModule
    Else
        bFlag = True
        Do While bFlag
            Err.Clear
            i = i + 1
            objCol.Add NameModule & "_" & i, NameModule & "_" & i
            If Err.Number = 0 Then
                bFlag = False
                AddModuleUniqueName = NameModule & "_" & i
            End If
        Loop
    End If
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : DeleteModuleToProject - delete module from VBA project
'* Created    : 22-03-2023 15:06
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByVal VBName As String : module name
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function DeleteModuleToProject(ByRef vbProj As VBIDE.vbProject, ByVal VBName As String) As Boolean
    Dim VBComp      As VBIDE.vbComponent
    On Error GoTo ErrorHandler
    Set VBComp = vbProj.VBComponents(VBName)
    Dim sMsg As String
    If VBComp.Type = vbext_ct_Document Then
        With VBComp.codeModule
            If .CountOfLines > 1 Then
                .DeleteLines 1, .CountOfLines
                .InsertLines 1, "Option Explicit"
            End If
        End With
        sMsg = "] cannot be deleted, VBA code removed from workbook ["
    Else
        Call vbProj.VBComponents.Remove(VBComp)
        sMsg = "] was deleted from workbook ["
    End If
    Debug.Print ">> Module: [" & VBName & sMsg & getFileNameOnVBProject(vbProj) & "]"

    DeleteModuleToProject = True
    Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 9:
            Debug.Print ">> Module: [ " & VBName & " ] not found in workbook [" & getFileNameOnVBProject(vbProj) & "]"
        Case 76:
            Debug.Print ">> Module: [" & VBName & "] was deleted from workbook:" & ActiveWorkbook.Name & vbCrLf & "File not saved!"
        Case Else:
            Call WriteErrorLog("DeleteModuleToProject", False)
    End Select
    Err.Clear
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : getFileName - get filename of active VBA project
'* Created    : 22-03-2023 15:07
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function getFileNameOnVBProject(Optional vbProj As VBIDE.vbProject = Nothing) As String
    On Error GoTo ErrorHandler
    If vbProj Is Nothing Then
        getFileNameOnVBProject = sGetFileName(Application.VBE.ActiveVBProject.FileName)
    Else
        getFileNameOnVBProject = sGetFileName(vbProj.FileName)
    End If
    Exit Function
ErrorHandler:
    Select Case Err.Number
        Case 76:
            getFileNameOnVBProject = ActiveWorkbook.Name
        Case Else:
            Call WriteErrorLog("getFileNameOnVBProject", False)
    End Select
    Err.Clear
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : ModuleLineCount - count number of code lines in VBA module
'* Created    : 23-03-2023 09:59
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                     Description
'*
'* oVBComp As VBIDE.VBComponent : VBA module
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function moduleLineCount(oVBComp As VBIDE.vbComponent) As Long
    Dim sLine       As String
    Dim j           As Long
    Dim i           As Long
    With oVBComp.codeModule
        If .CountOfLines > 0 Then
            For i = 1 To .CountOfLines
                sLine = Trim(.Lines(i, 1))
                If Left(sLine, 11) = "Option Base" Or Left(sLine, 14) = "Option Compare" Or Left(sLine, 15) = "Option Explicit" Or Left(sLine, 21) = "Option Private Module" Then sLine = vbNullString
                If sLine <> vbNullString Then j = j + 1
            Next i
        End If
    End With
    moduleLineCount = j
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : TypeProcedyre - extract procedure type
'* Created    : 22-03-2023 15:34
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                             Description
'*
'* ByRef StrDeclarationProcedure As String : declaration line
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function TypeProcedyre(ByRef StrDeclarationProcedure As String) As String
    If StrDeclarationProcedure Like "*Sub *" Then
        TypeProcedyre = "Sub"
    ElseIf StrDeclarationProcedure Like "*Function *" Then
        TypeProcedyre = "Function"
    ElseIf StrDeclarationProcedure Like "*Property Set *" Then
        TypeProcedyre = "Property Set"
    ElseIf StrDeclarationProcedure Like "*Property Get *" Then
        TypeProcedyre = "Property Get"
    ElseIf StrDeclarationProcedure Like "*Property Let *" Then
        TypeProcedyre = "Property Let"
    Else
        TypeProcedyre = "Unknown Type"
    End If
End Function

Public Function TypeProcedyreModifier(ByRef StrDeclarationProcedure As String) As String
    If StrDeclarationProcedure Like "Public *" Then
        TypeProcedyreModifier = "Public"
    ElseIf StrDeclarationProcedure Like "Private *" Then
        TypeProcedyreModifier = "Private"
    Else
        TypeProcedyreModifier = "Public"
    End If
End Function
Public Function TypeExtensionModule(ByRef ComponentType As vbext_ComponentType) As String
    Select Case ComponentType
             Case vbext_ct_ClassModule
            TypeExtensionModule = ".cls"
        Case vbext_ct_Document
            TypeExtensionModule = ".cls"
        Case vbext_ct_StdModule
            TypeExtensionModule = ".bas"
        Case vbext_ct_MSForm
            TypeExtensionModule = ".frm"
        Case Else
            TypeExtensionModule = ".txt"
    End Select
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : moduleType - module type
'* Created    : 22-03-2023 15:35
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                         Description
'*
'* ByRef ComponentType As VBIDE.vbext_ComponentType : VBA module
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function moduleTypeName(ByRef ComponentType As VBIDE.vbext_ComponentType) As String
    Select Case ComponentType
             Case vbext_ct_ActiveXDesigner
            moduleTypeName = "ActiveX Designer"
        Case vbext_ct_ClassModule
            moduleTypeName = "Class Module"
        Case vbext_ct_Document
            moduleTypeName = "Document Module"
        Case vbext_ct_MSForm
            moduleTypeName = "UserForm"
        Case vbext_ct_StdModule
            moduleTypeName = "Code Module"
        Case Else
            moduleTypeName = "Unknown Type: " & CStr(ComponentType)
    End Select
End Function

Public Function getVBModuleByName(ByRef vbProj As VBIDE.vbProject, ByVal sNameModule As String) As VBIDE.vbComponent
    On Error Resume Next
    Set getVBModuleByName = vbProj.VBComponents(sNameModule)
    On Error GoTo 0
End Function