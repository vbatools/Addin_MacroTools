VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMendgerVBAModules 
   Caption         =   "Password Removal:"
   ClientHeight    =   8535.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14220
   OleObjectBlob   =   "frmMendgerVBAModules.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMendgerVBAModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim arrList As Variant
Dim lCountLen As Integer

Private Enum ListColumns
    lcLines = 1    ' arr(i, 1) -> Number of code lines
    lcType = 2    ' arr(i, 2) -> Module type
    lcName = 3    ' arr(i, 3) -> Module name (used for access)
    lcFilterIndex = 4    ' arr(i, 4) -> Filter index
    lcSortKey = 5    ' arr(i, 5) -> Sort key
End Enum

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub lbCancel_Click()
    Call btnCancel_Click
End Sub

Private Sub lbClearSerch_Click()
    txtSerchVBAModule.value = vbNullString
End Sub

Private Sub lbCopytModule_Click()
    If MsgBox("Copy selected modules from project [" & cmbMain.value & "] to project [" & cmbCopyToFile.value & "]?", vbYesNo + vbQuestion, "Copying Modules:") = vbNo Then Exit Sub

    Dim i           As Long
    Dim iCount      As Long
    iCount = ListMain.ListCount - 1
    If iCount < 1 Then Exit Sub
    Dim vbProjCopy  As VBIDE.vbProject
    Dim vbProjPaste As VBIDE.vbProject
    Set vbProjCopy = Workbooks(cmbMain.value).vbProject
    Set vbProjPaste = Workbooks(cmbCopyToFile.value).vbProject
    With ListMain
        For i = 0 To iCount
            If .Selected(i) Then
                Call CopyModuleToProject(vbProjPaste, vbProjCopy.VBComponents(.List(i, 2)))
            End If
        Next
    End With
    Call loadDataToList
End Sub

Private Sub lbExportModule_Click()
    If MsgBox("Export selected modules from project [" & cmbMain.value & "]?", vbYesNo + vbQuestion, "Exporting Modules:") = vbNo Then Exit Sub

    Dim i           As Long
    Dim iCount      As Long
    iCount = ListMain.ListCount - 1
    If iCount < 1 Then Exit Sub
    Dim vbProj      As VBIDE.vbProject
    Dim sPath       As String
    With Workbooks(cmbMain.value)
        Set vbProj = .vbProject
        sPath = .Path & Application.PathSeparator & .Name & "_vba" & Application.PathSeparator
        If Not FileHave(sPath, vbDirectory) Then Call MkDir(sPath)
    End With
    Dim j           As Long
    With ListMain
        For i = 0 To iCount
            If .Selected(i) Then
                If exportModuleToFile(vbProj.VBComponents(.List(i, 2)), sPath) Then j = j + 1
            End If
        Next
    End With
    Call loadDataToList
    Call MsgBox("Exported: [" & j & "] modules", vbInformation)
End Sub

Private Sub lbImportModule_Click()
    Dim arrFiles()  As String
    arrFiles = fileDialogFun(Workbooks(cmbMain.value).Path, True, "*.bas;*.cls;*.frm")
    If (Not (Not (arrFiles))) = 0 Then Exit Sub
    Dim vbProj      As VBIDE.vbProject
    Set vbProj = Workbooks(cmbMain.value).vbProject

    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long

    iCount = UBound(arrFiles, 1)
    With vbProj.VBComponents
        On Error Resume Next
        For i = 1 To iCount
            Call .Import(arrFiles(i, 1))
            Select Case Err.Number
                     Case 0
                    Debug.Print ">> Module: [" & arrFiles(i, 1) & "] was imported to: [" & cmbMain.value & "]"
                    j = j + 1
                Case 60061
                    Debug.Print ">> Module: [" & arrFiles(i, 1) & "] already exists in: [" & cmbMain.value & "], not added to project"
                Case Else
                    Debug.Print ">> Error in ImportModule" & vbCrLf & Err.Number & vbCrLf & Err.Description & vbCrLf & "at line" & Erl
            End Select
        Next i
    End With
    Call loadDataToList
    Call MsgBox("Imported: [" & j & "] modules", vbInformation)
End Sub

Private Sub lbRemoveModule_Click()
    If MsgBox("Delete selected modules from project [" & cmbMain.value & "]?", vbYesNo + vbQuestion, "Deleting Modules:") = vbNo Then Exit Sub

    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    iCount = ListMain.ListCount - 1
    If iCount < 1 Then Exit Sub
    Dim vbProj      As VBIDE.vbProject
    Set vbProj = Workbooks(cmbMain.value).vbProject
    With ListMain
        For i = 0 To iCount
            If .Selected(i) Then
                If DeleteModuleToProject(vbProj, .List(i, 2)) Then j = j + 1
            End If
        Next
    End With
    Call loadDataToList
    Call MsgBox("Deleted: [" & j & "] modules", vbInformation)
End Sub

Private Sub txtSerchVBAModule_Change()
    Dim sVal        As String
    sVal = VBA.UCase$(txtSerchVBAModule.value)
    If VBA.Len(sVal) = 0 Then
        ListMain.List = arrList
        Call selectAllFilter
        Exit Sub
    End If

    Dim arrFiltered As Variant
    With ListMain
        If VBA.Len(sVal) > lCountLen Then
            arrFiltered = FilterArrayByText(.List, 2, sVal)
        Else
            arrFiltered = FilterArrayByText(arrList, 2, sVal)
        End If

        If Not IsEmpty(arrFiltered) Then
            .List = arrFiltered
        Else
            ListMain.Clear
        End If
    End With
    lCountLen = VBA.Len(sVal)
    Call selectAllFilter
End Sub



Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With

    Dim objVBProject As VBIDE.vbProject

    On Error Resume Next
    For Each objVBProject In Application.VBE.VBProjects
        cmbMain.AddItem sGetFileName(objVBProject.FileName)
    Next objVBProject

    If FileHave(ActiveWorkbook.FullName, vbNormal) Then cmbMain.value = ActiveWorkbook.Name
    cmbCopyToFile.List = cmbMain.List

    chbAll.value = True
    optAllSelect.value = True
    lbLockedFile.Caption = VBA.ChrW$(60848)
    lbLockedFileCopy.Caption = lbLockedFile.Caption
End Sub

Private Sub cmbMain_Change()
    Call loadDataToList
    Call lockedCopyBtn
End Sub

Private Sub cmbCopyToFile_Change()
    lbLockedFileCopy.Visible = Workbooks(cmbCopyToFile.value).vbProject.Protection = vbext_pp_locked
    Call lockedCopyBtn
End Sub

Private Sub lockedCopyBtn()
    lbCopytModule.Enabled = (Not (lbLockedFile.Visible Or lbLockedFileCopy.Visible)) And cmbCopyToFile.value <> vbNullString
    lbRemoveModule.Enabled = Not lbLockedFile.Visible
    lbExportModule.Enabled = lbRemoveModule.Enabled
    lbImportModule.Enabled = lbRemoveModule.Enabled
End Sub

Private Sub chbAll_Change()
    Call selectAllFilter
End Sub

Private Sub selectAllFilter()
    If ListFilter.ListCount = 0 Then Exit Sub
    Dim i           As Byte
    For i = 0 To 4
        ListFilter.Selected(i) = chbAll.value
    Next i
End Sub

Private Sub loadDataToList()
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.value)
    Dim objVBProject As VBIDE.vbProject
    Set objVBProject = wb.vbProject
    ListMain.Clear
    lbLockedFile.Visible = False
    ReDim arrFilter(0 To 4, 0 To 1)
    arrFilter(4, 0) = "ActiveX Designer"
    arrFilter(4, 1) = 0
    arrFilter(3, 0) = "Class Module"
    arrFilter(3, 1) = 0
    arrFilter(2, 0) = "Document Module"
    arrFilter(2, 1) = 0
    arrFilter(1, 0) = "UserForm"
    arrFilter(1, 1) = 0
    arrFilter(0, 0) = "Code Module"
    arrFilter(0, 1) = 0
    With objVBProject
        Dim iCount  As Long
        If .Protection = vbext_pp_none Then
            Dim i   As Long

            iCount = .VBComponents.Count
            If iCount < 1 Then Exit Sub
            ReDim arr(1 To iCount, 1 To 5) As String
            For i = 1 To iCount
                With .VBComponents(i)
                    arr(i, lcType) = moduleTypeName(.Type)
                    arr(i, lcName) = .Name
                    arr(i, lcSortKey) = arr(i, lcType) & "|" & arr(i, lcName)
                    Select Case .Type
                             Case vbext_ct_ActiveXDesigner
                            arr(i, lcFilterIndex) = 4
                            arrFilter(4, 0) = "ActiveX Designer"
                            arrFilter(4, 1) = arrFilter(4, 1) + 1
                        Case vbext_ct_ClassModule
                            arr(i, lcFilterIndex) = 3
                            arrFilter(3, 0) = "Class Module"
                            arrFilter(3, 1) = arrFilter(3, 1) + 1
                        Case vbext_ct_Document
                            arr(i, lcFilterIndex) = 2
                            arrFilter(2, 0) = "Document Module"
                            arrFilter(2, 1) = arrFilter(2, 1) + 1
                        Case vbext_ct_MSForm
                            arr(i, lcFilterIndex) = 1
                            arrFilter(1, 0) = "UserForm"
                            arrFilter(1, 1) = arrFilter(1, 1) + 1
                        Case vbext_ct_StdModule
                            arr(i, lcFilterIndex) = 0
                            arrFilter(0, 0) = "Code Module"
                            arrFilter(0, 1) = arrFilter(0, 1) + 1
                    End Select
                End With
                arr(i, lcLines) = moduleLineCount(.VBComponents(i))
            Next i
            arr = sortArray2D(arr)
            ListMain.List = arr
        Else
            lbLockedFile.Visible = True
        End If
    End With
    lbTotal.Caption = iCount
    ListFilter.List = arrFilter
    arrList = ListMain.List
    Call selectItemList
    Call selectAllFilter
End Sub

Private Function sortArray2D(ByRef arr As Variant) As Variant
    Dim clsSort     As clsSort2DArray
    Set clsSort = New clsSort2DArray
    With clsSort
        Call .SortQuick(arr, 5, True, sdtString)
        sortArray2D = .ListArray
    End With
    Set clsSort = Nothing
End Function

Private Sub chbEmptySelect_Click()
    Call selectItemList
End Sub

Private Sub optAllSelect_Change()
    Call selectItemList
End Sub

Private Sub ListFilter_Change()
    Call selectItemList
End Sub

Private Sub selectItemList()
    Dim i           As Long
    Dim iCount      As Long
    With ListMain
        iCount = .ListCount - 1
        If iCount < 0 Then Exit Sub
        For i = 0 To iCount
            .Selected(i) = optAllSelect.value And ListFilter.Selected(.List(i, 3))
            If .List(i, 0) = "0" And Not chbEmptySelect.value Then .Selected(i) = False
        Next i
    End With
End Sub
