VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCharsMonitor 
   Caption         =   "Character Monitor:"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9015.001
   OleObjectBlob   =   "frmCharsMonitor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCharsMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : CharsMonitor - Character analysis in string
'* Created    : 23-04-2020 14:27
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Dim clsAnc          As clsAnchors

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With

    Set clsAnc = New clsAnchors
    With clsAnc
        Call .Initialize(Me, 600, 1000)
        Call .SetAnchorStyleByName(lbClearForm.Name, anchorRight)
        Call .SetAnchorStyleByName(txtStr.Name, anchorRight Or anchorLeft Or anchorTop Or anchorBottom)
        Call .SetAnchorStyleByName(lbLoadStr.Name, anchorRight Or anchorBottom)
        Call .SetAnchorStyleByName(chStrAfter.Name, anchorRight Or anchorBottom)
        Call .SetAnchorStyleByName(lbMsg.Name, anchorLeft Or anchorBottom)
        Call .SetAnchorStyleByName(Label3.Name, anchorLeft Or anchorBottom)
        Call .SetAnchorStyleByName(ListChars.Name, anchorLeft Or anchorRight Or anchorBottom)
        Call .SetAnchorStyleByName(lbExportStr.Name, anchorLeft Or anchorBottom)
        Call .SetAnchorStyleByName(lbCancel.Name, anchorRight Or anchorBottom)
        Call .SetAnchorStyleByName(lbAddASIITable.Name, anchorBottom)
    End With
    txtStr.value = ActiveCell.value
End Sub

Private Sub UserForm_Terminate()
    Set clsAnc = Nothing
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub lbCancel_Click()
    Call btnCancel_Click
End Sub

Private Sub lbClearForm_Click()
    Me.txtStr = vbNullString
    Me.ListChars.Clear
    Me.lbMsg.Caption = vbNullString
End Sub
Private Sub txtStr_Change()
    Call subStartParser
End Sub
Private Sub subStartParser()
    Dim lRows       As Long
    Dim lWords      As Long

    With Me.ListChars
        .Clear
        Me.lbMsg.Caption = vbNullString
        If Me.txtStr.text <> vbNullString Then
            .List = addArrayFormString(Me.txtStr.text)
            lRows = UBound(VBA.Split(Me.txtStr.text, vbNewLine)) + 1
            lWords = UBound(VBA.Split(WorksheetFunction.Trim((VBA.Replace(Me.txtStr.text, vbNewLine, VBA.Chr(32)))), VBA.Chr(32))) + 1
            If lRows < 0 Then lRows = 0
            Me.lbMsg.Caption = "String length: " & VBA.Len(Me.txtStr.text) & " chars. Lines: " & lRows & " Words: " & lWords
        End If
    End With
End Sub

Private Function addArrayFormString(ByVal sTxt As String) As Variant

    Dim n           As Long
    Dim i           As Long
    Dim sChar       As String * 1

    On Error Resume Next
    n = Len(sTxt): ReDim arr(1 To n, 1 To 5)
    For i = LBound(arr) To UBound(arr)
        arr(i, 1) = i
        sChar = VBA.mid$(sTxt, i, 1)
        arr(i, 2) = sChar
        arr(i, 3) = VBA.Asc(sChar)
        arr(i, 4) = VBA.AscW(sChar)
        arr(i, 5) = VBA.Hex$(VBA.Asc(sChar))
    Next i
    addArrayFormString = arr
End Function
Private Sub lbLoadStr_Click()
    Dim objRng      As Range
    Dim itemRng     As Range
    Dim sStrTemp    As String
    Dim iColCount   As Integer
    Dim i           As Integer
    Dim sChrDel     As String

    i = 1
    sChrDel = VBA.Chr$(32)
    Me.Hide
    Set objRng = GetAddressCell()
    If objRng Is Nothing Then Exit Sub
    iColCount = objRng.Columns.Count
    For Each itemRng In objRng
        If chStrAfter.value Then
            If i > iColCount Then
                i = 1
                sChrDel = vbNewLine
            Else
                sChrDel = VBA.Chr$(32)
            End If
        End If
        i = i + 1
        sStrTemp = sStrTemp & sChrDel & itemRng.value
    Next itemRng
    Me.txtStr = VBA.Right$(sStrTemp, VBA.Len(sStrTemp) - VBA.Len(sChrDel))
    Call subStartParser
    Me.Show
End Sub
Private Sub lbExportStr_Click()
    Dim objRng      As Range
    Me.Hide
    Set objRng = GetAddressCell("Select cell for insertion:")
    If objRng Is Nothing Then Exit Sub
    With Me.ListChars
        If .ListCount > 0 Then
            objRng.Cells(1, 1).value = Me.txtStr
            objRng.Offset(1, 0).Resize(1, 5).Value2 = Array("№", "Char", "Asc", "AscW", "Hex")
            objRng.Offset(2, 0).Resize(.ListCount, 5).Value2 = .List
        End If
    End With
    Me.Show
End Sub
Private Sub lbAddASIITable_Click()
    Dim objRng      As Range
    Dim i           As Integer
    ReDim arr(1 To 256, 1 To 5)
    Me.Hide
    Set objRng = GetAddressCell("Select cell for insertion:")
    If objRng Is Nothing Then Exit Sub
    For i = 1 To 256
        arr(i, 1) = i - 1
        arr(i, 2) = VBA.Hex$(i - 1)
        arr(i, 3) = VBA.Chr$(i - 1)
        arr(i, 4) = VBA.AscW(arr(i, 3))
        arr(i, 5) = GetDiscriptionSpeshelChar(i - 1)
    Next i
    With objRng
        .Resize(1, 5).Value2 = Array("Dec/Asc", "Hex", "Char", "AscW", "Description")
        .Offset(1, 0).Resize(256, 5).Value2 = arr
    End With
End Sub

Private Function GetDiscriptionSpeshelChar(ByVal i As Byte) As String
    Select Case i
             Case 0: GetDiscriptionSpeshelChar = "NOP"
        Case 1: GetDiscriptionSpeshelChar = "SOH"
        Case 2: GetDiscriptionSpeshelChar = "STX"
        Case 3: GetDiscriptionSpeshelChar = "ETX"
        Case 4: GetDiscriptionSpeshelChar = "EOT"
        Case 5: GetDiscriptionSpeshelChar = "ENQ"
        Case 6: GetDiscriptionSpeshelChar = "ACK"
        Case 7: GetDiscriptionSpeshelChar = "BEL"
        Case 8: GetDiscriptionSpeshelChar = "BS"
        Case 9: GetDiscriptionSpeshelChar = "Tab"
        Case 10: GetDiscriptionSpeshelChar = "LF (Line Feed)"
        Case 11: GetDiscriptionSpeshelChar = "VT"
        Case 12: GetDiscriptionSpeshelChar = "FF"
        Case 13: GetDiscriptionSpeshelChar = "CR (Carriage Return)"
        Case 14: GetDiscriptionSpeshelChar = "SO"
        Case 15: GetDiscriptionSpeshelChar = "SI"
        Case 16: GetDiscriptionSpeshelChar = "DLE"
        Case 17: GetDiscriptionSpeshelChar = "DC1"
        Case 18: GetDiscriptionSpeshelChar = "DC2"
        Case 19: GetDiscriptionSpeshelChar = "DC3"
        Case 20: GetDiscriptionSpeshelChar = "DC4"
        Case 21: GetDiscriptionSpeshelChar = "NAK"
        Case 22: GetDiscriptionSpeshelChar = "SYN"
        Case 23: GetDiscriptionSpeshelChar = "ETB"
        Case 24: GetDiscriptionSpeshelChar = "CAN"
        Case 25: GetDiscriptionSpeshelChar = "EM"
        Case 26: GetDiscriptionSpeshelChar = "SUB"
        Case 27: GetDiscriptionSpeshelChar = "ESC"
        Case 28: GetDiscriptionSpeshelChar = "FS"
        Case 29: GetDiscriptionSpeshelChar = "GS"
        Case 30: GetDiscriptionSpeshelChar = "RS"
        Case 31: GetDiscriptionSpeshelChar = "US"
        Case 32: GetDiscriptionSpeshelChar = "SP (Space)"
    End Select
End Function
Private Function GetAddressCell(Optional sMsg As String = "Select data range:") As Range
    Dim sDefault    As String
    On Error GoTo Canceled
    If TypeName(Selection) = "Range" Then
        sDefault = Selection.Address
    Else
        sDefault = vbNullString
    End If
    Set GetAddressCell = Application.InputBox(Prompt:=sMsg, Type:=8, Default:=sDefault)
    Exit Function
Canceled:
    Set GetAddressCell = Nothing
End Function