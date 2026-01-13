Attribute VB_Name = "modAddinInstall"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : InstallationAddMacro - процедура установка надстройки
'* Created    : 22-03-2023 15:14
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub InstallationAddinMacroTools()
    Dim AddFolder   As String
    On Error GoTo InstallationAdd_Err
    'Проверяем имеется ли данная директория
    AddFolder = VBA.Replace(Application.UserLibraryPath & Application.PathSeparator, Application.PathSeparator & Application.PathSeparator, Application.PathSeparator)
    'Проверка на наличие дириктории
    If Dir(AddFolder, vbDirectory) = vbNullString Then
        Call MsgBox("К сожалению, программа не может выполнить установку надстройки на данном компьютере." & vbCrLf & _
                "Отсутствует директория с надстройками." & vbCrLf & _
                "Обратитесь к разработчику программы.", vbCritical, _
                "Сбой установки надстройки:")
        Exit Sub
    End If
    Dim sFullName   As String
    sFullName = AddFolder & modAddinConst.NAME_ADDIN & ".xlam"
    'Отключаем ранее установленую надстройку
    If FileHave(sFullName) Then AddIns(modAddinConst.NAME_ADDIN).Installed = False
    'Проверяем открыта ли надстройка
    If WorkbookIsOpen(modAddinConst.NAME_ADDIN & ".xlam") Then
        Call MsgBox("Файл с надстройкой уже открыт." & vbCrLf & _
                "Возможно она уже была установлена ранее.", vbCritical, _
                "Сбой установки надстройки:")
        Exit Sub
    End If
    'Сохраняем как
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    If Workbooks.Count = 0 Then Workbooks.Add
    Call ThisWorkbook.SaveAs(Filename:=sFullName, FileFormat:=xlOpenXMLAddIn)
    Call AddIns.Add(Filename:=sFullName)
    AddIns(modAddinConst.NAME_ADDIN).Installed = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Call MsgBox("Программа успешно установлена! " & vbCrLf & _
            "Просто откройте или создайте новый документ.", vbInformation, _
            "Установка надстройки: " & modAddinConst.NAME_ADDIN)
    Call ThisWorkbook.Close(False)
    Exit Sub
InstallationAdd_Err:
    If Err.Number = 1004 Then
        Call MsgBox("Для установки надстройки, пожалуйста закройте данный файл и запустите его еще раз.", _
                vbCritical, "Установка:")
    Else
        Call MsgBox(Err.Description & vbCrLf & "в F_AddInInstall.InstallationAddinMacroTools " & vbCrLf & "в строке " & Erl, _
                vbExclamation + vbOKOnly, "Ошибка:")
        'Call WriteErrorLog("modInstallAddin.InstallationAddinMacroTools")
    End If
End Sub
