Attribute VB_Name = "modAddinRibbonCallbacks"
Option Explicit
Option Private Module

Private Sub btnRefresh(control As IRibbonControl)
      If VBAIsTrusted Then Call RefreshMenu
End Sub

Private Sub btnInfoFile(control As IRibbonControl)
      Call frmInfoFile.Show
End Sub

Private Sub btnOpenFile(control As IRibbonControl)
    Call UnZipFile
End Sub

Private Sub btnCloseFile(control As IRibbonControl)
    Call ZipFile
End Sub

Private Sub btnInToFile(control As IRibbonControl)
    Call addListInFileFiles
End Sub

Private Sub btnUnProtectVBA(control As IRibbonControl)
    Call unProtectVBA
End Sub

Private Sub btnExportVBA(control As IRibbonControl)
    Call frmMendgerVBAModules.Show
End Sub

Private Sub btnUnProtectVBAUnivable(control As IRibbonControl)
    Call delProtectVBAUnviewable
End Sub

Private Sub btnProtectVBAUnivable(control As IRibbonControl)
    Call setProtectVBAUnviewable
End Sub

Private Sub btnHiddenModule(control As IRibbonControl)
    Call frmHideModule.Show
End Sub

Private Sub btnUnProtectSheetsXML(control As IRibbonControl)
    Call delPasswordWBook
End Sub

Private Sub btnObfuscator(control As IRibbonControl)
    Call ObfuscationVBAProject
End Sub

Private Sub btnObfuscatorVariable(control As IRibbonControl)
    Call addListVariableProjectOfuscation
End Sub

Private Sub btnSerchVariableUnUsed(control As IRibbonControl)
    Call showFormUnUsedVariable
End Sub

Private Sub btnAddStatisticAll(control As IRibbonControl)
    Call addStatAll
End Sub

Private Sub btnAddStatisticForms(control As IRibbonControl)
    Call addStatUserFormsControl
End Sub

Private Sub btnAddStatisticModules(control As IRibbonControl)
    Call addStatModules
End Sub

Private Sub btnAddStatisticDeclaretions(control As IRibbonControl)
    Call addStatDeclaration
End Sub

Private Sub btnAddStatisticProcedures(control As IRibbonControl)
    Call addStatModuleProcedures
End Sub

Private Sub btnAddStatisticShape(control As IRibbonControl)
    Call addShapeStatistic
End Sub

Private Sub btnParserLiterals(control As IRibbonControl)
    Call getAllLiteralsFile
End Sub

Private Sub btnReNameLiterals(control As IRibbonControl)
    Call ReNameLiteralsFile
End Sub

Private Sub btnToolCharMonitor(control As IRibbonControl)
    Call frmCharsMonitor.Show
End Sub

Private Sub btnRegExpr(control As IRibbonControl)
    Call AddSheetTestRegExp
End Sub

Private Sub btnDeleteExternalLinks(control As IRibbonControl)
    Call ExternalLinkUtility
End Sub

Private Sub btnReferenceStyle(control As IRibbonControl)
    With Application
        If .ReferenceStyle = xlR1C1 Then
            .ReferenceStyle = xlA1
        Else
            .ReferenceStyle = xlR1C1
        End If
    End With
End Sub

Private Sub btnAddIn(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.Dialogs(xlDialogAddinManager).Show
    Exit Sub
ErrorHandler:
    Err.Clear
    Call MsgBox("No open Excel files" & Chr(34) & "Files Excel" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
End Sub

Private Sub btnVBAWindowOpen(control As IRibbonControl)
    Call Application.SendKeys("%{F11}")
End Sub

Private Sub btnOptionsStyle(control As IRibbonControl)
    frmSettingsIndent.Show
End Sub

Private Sub btnOptionsComment(control As IRibbonControl)
    Call frmSettingsKomments.Show
End Sub

Private Sub btnBlackTheme(control As IRibbonControl)
    Call changeColorDarkTheme
End Sub

Private Sub btnWhiteTheme(control As IRibbonControl)
    Call changeColorWhiteTheme
End Sub

Private Sub btnOpenLogFile(control As IRibbonControl)
    Dim clsLoger          As clsLogging
    Set clsLoger = New clsLogging
    Call clsLoger.ShowLog
    Set clsLoger = Nothing
End Sub

Private Sub btnDeleteLogFile(control As IRibbonControl)
    Dim clsLoger          As clsLogging
    Set clsLoger = New clsLogging
    Call clsLoger.ResetLogs
    Set clsLoger = Nothing
End Sub

Private Sub btnAbout(control As IRibbonControl)
    Call frmAboutInfo.Show
End Sub