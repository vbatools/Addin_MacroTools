Attribute VB_Name = "modAddinInstall"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : InstallationAddMacro - Installs the add-in
'* Created    : 22-03-2023 15:14
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub InstallationAddinMacroTools()
    Dim AddFolder   As String
    On Error GoTo InstallationAdd_Err
    'Get the folder path for add-ins
    AddFolder = VBA.Replace(Application.UserLibraryPath & Application.PathSeparator, Application.PathSeparator & Application.PathSeparator, Application.PathSeparator)
    'Check if the folder exists
    If Dir(AddFolder, vbDirectory) = vbNullString Then
        Call MsgBox("Unable to find the folder for installing the add-in." & vbCrLf & _
                "Please check the installation path." & vbCrLf & _
                "Contact support for assistance.", vbCritical, _
                "Installation Error:")
        Exit Sub
    End If
    Dim sFullName   As String
    sFullName = AddFolder & modAddinConst.NAME_ADDIN & ".xlam"
    'Uninstall previous version if exists
    If FileHave(sFullName) Then AddIns(modAddinConst.NAME_ADDIN).Installed = False
    'Check if the file is already open
    If WorkbookIsOpen(modAddinConst.NAME_ADDIN & ".xlam") Then
        Call MsgBox("The file is currently open." & vbCrLf & _
                "Please close the file before continuing.", vbCritical, _
                "Installation Error:")
        Exit Sub
    End If
    'Install the add-in
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    If Workbooks.Count = 0 Then Workbooks.Add
    Call ThisWorkbook.SaveAs(Filename:=sFullName, FileFormat:=xlOpenXMLAddIn)
    Call AddIns.Add(Filename:=sFullName)
    AddIns(modAddinConst.NAME_ADDIN).Installed = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Call MsgBox("Installation completed successfully! " & vbCrLf & _
            "The add-in is now ready to use.", vbInformation, _
            "Installation Complete: " & modAddinConst.NAME_ADDIN)
    Call ThisWorkbook.Close(False)
    Exit Sub
InstallationAdd_Err:
    If Err.Number = 1004 Then
        Call MsgBox("During installation, an error occurred related to permissions or file access.", _
                vbCritical, "Installation Error:")
    Else
        Call MsgBox(Err.Description & vbCrLf & "in F_AddInInstall.InstallationAddinMacroTools " & vbCrLf & "at line " & Erl, _
                vbExclamation + vbOKOnly, "Error:")
        'Call WriteErrorLog("modInstallAddin.InstallationAddinMacroTools")
    End If
End Sub
