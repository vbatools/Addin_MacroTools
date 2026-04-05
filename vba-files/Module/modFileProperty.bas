Attribute VB_Name = "modFileProperty"
Option Explicit
Option Private Module

Public Function GetOneProp(ByRef wb As Workbook, ByVal NameProp As String) As String
      GetOneProp = wb.BuiltinDocumentProperties(NameProp).value
End Function

Public Function getFilePropertiesCustomList(ByRef wb As Workbook) As Variant
      Dim i           As Long
    Dim iCount      As Long

    With wb
        iCount = .CustomDocumentProperties.Count
        If iCount = 0 Then Exit Function
        ReDim arr(1 To iCount, 1 To 3)
        For i = 1 To iCount
            With .CustomDocumentProperties(i)
                arr(i, 1) = i
                arr(i, 2) = .Name
                arr(i, 3) = .value
            End With
        Next i
    End With
    getFilePropertiesCustomList = arr
End Function

Public Function getFilePropertiesList(ByRef wb As Workbook) As Variant
    Dim i           As Long
    Dim iCount      As Long

    With wb
        iCount = .BuiltinDocumentProperties.Count
        If iCount = 0 Then Exit Function
        ReDim arr(1 To iCount, 1 To 3)
        For i = 1 To iCount
            With .BuiltinDocumentProperties(i)
                arr(i, 1) = i
                arr(i, 2) = .Name
                On Error Resume Next
                arr(i, 3) = .value
                On Error GoTo 0
            End With
        Next i
    End With
    getFilePropertiesList = arr
End Function

Public Sub addFilePropertyCustom(ByRef wb As Workbook, ByVal nameProperty As String, ByVal valProperty As String)
    Call delFilePropertyCustom(wb, nameProperty)
    Call wb.CustomDocumentProperties.Add(nameProperty, False, msoPropertyTypeString, valProperty)
End Sub

Public Function addFileProperty(ByRef wb As Workbook, ByVal nameProperty As String, ByVal valProperty As String) As Boolean
    Dim docProp     As DocumentProperty
    On Error GoTo endfun
    Set docProp = wb.BuiltinDocumentProperties(nameProperty)
    docProp.value = valProperty
    addFileProperty = True
    Exit Function
endfun:
    On Error GoTo 0
End Function

Public Function delFilePropertiesCustomAll(ByRef wb As Workbook) As Byte
    Dim i           As Long
    Dim iCount      As Long
    Dim k           As Long
    With wb
        iCount = .CustomDocumentProperties.Count
        If iCount > 0 Then
            For i = iCount To 1 Step -1
                .CustomDocumentProperties(i).Delete
                k = k + 1
            Next i
        End If
    End With
    delFilePropertiesCustomAll = k
End Function

Public Function delFilePropertiesAll(ByRef wb As Workbook) As Long
    Dim i           As Long
    Dim iCount      As Long
    Dim k           As Long

    With wb
        iCount = .BuiltinDocumentProperties.Count
        If iCount = 0 Then Exit Function
        For i = 1 To iCount
            With .BuiltinDocumentProperties(i)
                On Error Resume Next
                .value = vbNullString
                If Err.Number = 0 Then k = k + 1
                On Error GoTo 0
            End With
        Next i
    End With
    delFilePropertiesAll = k
End Function

Public Function delFilePropertyCustom(ByRef wb As Workbook, ByVal nameProperty As String) As Boolean
    Dim i           As Long
    Dim iCount      As Long
    With wb
        iCount = .CustomDocumentProperties.Count
        If iCount > 0 Then
            For i = iCount To 1 Step -1
                If .CustomDocumentProperties(i).Name = nameProperty Then
                    .CustomDocumentProperties(i).Delete
                    delFilePropertyCustom = True
                    Exit Function
                End If
            Next i
        End If
    End With
End Function
