Attribute VB_Name = "modToolsObfuscation"
Option Explicit
Option Private Module

Public Const ms_VARIABLE_SHEET As String = "OBFUSCATION_VARIABLE"

Public Sub ObfuscationVBAProject()

    ' 1. Get target workbook
    Dim wb          As Workbook
    If Not GetTargetWorkbook(wb, "Obfuscate VBA Project:", "OBFUSCATE") Then Exit Sub
    
    If wb.vbProject.Protection = vbext_pp_locked Then
        MsgBox "VBA project is password protected!", vbCritical
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' 2. Initialize obfuscator
    Dim oObfuscator As clsObfuscator
    Set oObfuscator = New clsObfuscator

    ' 3. Perform obfuscation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If oObfuscator.Execute(wb) Then
        Application.ScreenUpdating = True
        Call MsgBox("Obfuscation completed successfully!", vbInformation, "Result")
    Else
        Application.ScreenUpdating = True
        Call MsgBox("Obfuscation completed with errors.", vbExclamation, "Result")
    End If

CleanUp:
    ' 4. Cleanup
    Set oObfuscator = Nothing

    Exit Sub
ErrorHandler:
    Call WriteErrorLog("ObfuscationVBAProject", True)
    Resume CleanUp
End Sub