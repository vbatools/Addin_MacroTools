Attribute VB_Name = "modFilePassWBook"
Option Explicit
Option Private Module

Public Sub delPasswordWBook()
    Dim oForma      As frmDelPaswortSheetBook
    Set oForma = New frmDelPaswortSheetBook
    oForma.Show
    Application.ScreenUpdating = False
    Dim bFlag       As Boolean
    bFlag = VBA.CBool(oForma.lbValue.Caption)
    If Not bFlag Then Exit Sub
    Dim wb          As Workbook
    Set wb = Workbooks(oForma.cmbMain.value)
    Dim sFullName   As String
    With wb
        sFullName = .FullName
        wb.Close True
    End With

    Dim clsZIP      As clsOfficeArchiveManager
    Set clsZIP = New clsOfficeArchiveManager
    With clsZIP
        If .Initialize(sFullName, True) Then
            If .UnZipFile Then
                If oForma.lbMsg.Visible Then Call .delPasswordWBook
                Dim arr As Variant
                arr = .getArraySheetsName()
                If Not IsEmpty(arr) Then
                    Dim i As Long
                    For i = 1 To UBound(arr, 1)
                        Call .delPasswordSheet(arr(i, 2))
                    Next i
                End If
                Call .ZipFilesInFolder
            End If
        End If
    End With
    Set clsZIP = Nothing
    Call Workbooks.Open(sFullName)
    Application.ScreenUpdating = True
    If bFlag Then
        Set oForma = New frmDelPaswortSheetBook
        oForma.Show
    End If
End Sub
