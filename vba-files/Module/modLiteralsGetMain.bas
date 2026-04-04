Attribute VB_Name = "modLiteralsGetMain"
Option Explicit
Option Private Module

Public Const STR_UF As String = "STR_UF"
Public Const STR_CODE As String = "STR_CODE"
Public Const STR_UI As String = "STR_UI"

Public Sub getAllLiteralsFile()

    Dim wb          As Workbook
    If Not GetTargetWorkbook(wb, "Collecting String Literals:", "COLLECT") Then Exit Sub

    Application.ScreenUpdating = False
    If wb.vbProject.Protection = vbext_pp_locked Then
        Call MsgBox("VBA project is password protected!", vbCritical)
        Exit Sub
    End If

    Dim sFullNameFile As String
    sFullNameFile = wb.FullName

    Dim arr         As Variant
    Dim dtStart     As Date
    dtStart = VBA.Now()
    Debug.Print ">> Data collection: " & sFullNameFile

    arr = GetArrayFromDictionary(parserLiteralsFormControls(wb))
    Call OutputResults(ActiveWorkbook, STR_UF, sFullNameFile, arr)
    Debug.Print vbTab & ">> " & VBA.Format$(VBA.Now() - dtStart, FORMAT_TIME) & " User Forms"

    arr = GetArrayFromDictionary(parserLiteralsFormCode(wb, False))
    Call OutputResults(ActiveWorkbook, STR_CODE, sFullNameFile, arr)
    Debug.Print vbTab & ">> " & VBA.Format$(VBA.Now() - dtStart, FORMAT_TIME) & " Code VBA"

    arr = GetArrayFromDictionary(parserLiteralsFormUI(wb, Array("label", "supertip", "screentip", "title", "description"), False, True))
    Call OutputResults(ActiveWorkbook, STR_UI, sFullNameFile, arr)
    Debug.Print vbTab & ">> " & VBA.Format$(VBA.Now() - dtStart, FORMAT_TIME) & " Ribbon Control UI"
End Sub