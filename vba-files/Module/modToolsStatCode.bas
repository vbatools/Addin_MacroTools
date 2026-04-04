Attribute VB_Name = "modToolsStatCode"
Option Explicit
Option Private Module

' Enumeration for determining the type of statistics to collect
Private Enum StatMode
    msAll = 0
    msModules = 1
    msProcedures = 2
    msUserForms = 3
    msDeclarations = 4
End Enum

Public Sub addListVariableProjectOfuscation()
    Call RunStatCollection(msAll, "Collecting Data for Obfuscation:", ms_VARIABLE_SHEET, False, False, False)
End Sub

Public Sub addStatAll()
    Call RunStatCollection(msAll, "Collecting Statistics:", "ALL", False, False, False)
End Sub

Public Sub addStatModules()
    Call RunStatCollection(msModules, "Collecting Statistics - Module:", "MODULES", False, False, False)
End Sub

Public Sub addStatModuleProcedures()
    Call RunStatCollection(msProcedures, "Collecting Statistics - Procedures:", "PROCEDURES", False, False, False)
End Sub

Public Sub addStatUserFormsControl()
    Call RunStatCollection(msUserForms, "Collecting Statistics - UserForms:", "FORMS", False, False, False)
End Sub

Public Sub addStatDeclaration()
    Call RunStatCollection(msDeclarations, "Collecting Statistics - Declaration:", "DECLARATION", False, False, False)
End Sub

' Main private procedure containing all business logic
Private Sub RunStatCollection(ByVal mode As StatMode, ByVal formCaption As String, ByVal sheetName As String, _
        ByVal bGetCodeModule As Boolean, ByVal bGetCodeProcs As Boolean, ByVal bGetCodeDeclarations As Boolean)

    Dim wb          As Workbook
    ' 1. Initialize and display the form
    If Not GetTargetWorkbook(wb, formCaption, "COLLECT") Then Exit Sub

    Dim cls         As clsToolsVBACodeStatistics
    Dim dtStart     As Date

    Application.ScreenUpdating = False
    If wb Is Nothing Then
        MsgBox "Could not access the selected workbook.", vbExclamation
        Exit Sub
    End If

    ' 2. Check VBA project protection
    If wb.vbProject.Protection = vbext_pp_locked Then
        MsgBox "Project is password protected!", vbCritical
        Exit Sub
    End If

    ' 3. Collect data
    Set cls = New clsToolsVBACodeStatistics
    dtStart = VBA.Now()
    With cls
        Select Case mode
                Case msAll
                Debug.Print ">> " & Format(VBA.Now() - dtStart, FORMAT_TIME) & " Code Statistics Start"
                Call .addListModules(wb, bGetCodeModule)
                Debug.Print ">> " & Format(VBA.Now() - dtStart, FORMAT_TIME) & " Modules"

                Call .addListDeclarations(wb, bGetCodeDeclarations)
                Debug.Print ">> " & Format(VBA.Now() - dtStart, FORMAT_TIME) & " Declarations"

                Call .addListProcs(wb, bGetCodeProcs)
                Debug.Print ">> " & Format(VBA.Now() - dtStart, FORMAT_TIME) & " Procedures"

                Call .addListControlsUserForms(wb)
                Debug.Print ">> " & Format(VBA.Now() - dtStart, FORMAT_TIME) & " UserForm Controls"
            Case msModules
                Call .addListModules(wb, bGetCodeModule)
            Case msProcedures
                Call .addListProcs(wb, bGetCodeProcs)
            Case msUserForms
                Call .addListControlsUserForms(wb)
            Case msDeclarations
                Call .addListDeclarations(wb, bGetCodeDeclarations)
        End Select

        Call OutputResults(ActiveWorkbook, sheetName, wb.FullName, .getArrayCodeBase)
    End With
    Set cls = Nothing
End Sub