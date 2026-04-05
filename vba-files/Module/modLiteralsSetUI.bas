Attribute VB_Name = "modLiteralsSetUI"
Option Explicit
Option Private Module

'==============================================================================
' Enumeration: UIColumns
' Purpose: Defines column indices for the arrUI array.
'==============================================================================
Public Enum UIColumns
    UI_ModuleType = 1  ' Module type (CustomUI / CustomUI14)
    UI_XMLNodeName = 2
    UI_TagName = 3     ' XML tag name (e.g., "button", "menu")
    UI_IdOriginal = 4  ' Identifier for search (current ID)
    UI_IdNew = 5       ' New identifier (also used as an alternative for search)
    UI_AttrName = 6    ' Name of the attribute to be changed
    UI_AttrText = 7
    UI_AttrTextNew = 8    ' New attribute value
    UI_Status = 9      ' Column for recording the operation status (for feedback)
    [_First] = UI_ModuleType
    [_Last] = UI_Status
End Enum

Public Function renameLiteralsToUI(ByRef wb As Workbook, ByRef arrUI As Variant) As Boolean
    On Error GoTo ErrorHandler

    Const KEY_CUSTOM_UI As String = "CustomUI"
    Const KEY_CUSTOM_UI_14 As String = "CustomUI14"

    Dim sFullNameFile As String
    Dim oZipManager As clsOfficeArchiveManager
    Dim lRowCount   As Long
    Dim i           As Long
    Dim sCurrentKey As String
    Dim sPathUI     As String
    Dim oXMLDoc     As MSXML2.DOMDocument
    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList
    Dim bOperationSuccess As Boolean
    Dim sID         As String

    bOperationSuccess = False

    If Not IsArray(arrUI) Then Exit Function
    If wb Is Nothing Then Exit Function

    sFullNameFile = wb.FullName
    wb.Close savechanges:=True

    Set oZipManager = New clsOfficeArchiveManager
    lRowCount = UBound(arrUI, 1)

    With oZipManager
        If .Initialize(sFullNameFile, True) Then
            If .UnZipFile Then
                ' 4. Main data processing loop
                For i = 1 To lRowCount
                    ' Optimization: reload XML only when the interface type changes
                    If sCurrentKey <> CStr(arrUI(i, UIColumns.UI_ModuleType)) Then
                        ' Save the previous document if it was loaded
                        If Not oXMLDoc Is Nothing Then
                            Call oXMLDoc.Save(sPathUI)
                        End If

                        ' Setup for the new type
                        sCurrentKey = CStr(arrUI(i, UIColumns.UI_ModuleType))
                        Select Case sCurrentKey
                            Case KEY_CUSTOM_UI
                                sPathUI = .GetSettings(FileCustomUI)
                            Case KEY_CUSTOM_UI_14
                                sPathUI = .GetSettings(FileCustomUI14)
                            Case Else
                                sPathUI = vbNullString
                        End Select

                        If Len(sPathUI) > 0 Then Set oXMLDoc = .getXMLDOC(sPathUI)
                    End If
                    ' 5. XML modification
                    Call WriteXML(arrUI, i, oXMLDoc)
                Next i

                If Not oXMLDoc Is Nothing Then Call oXMLDoc.Save(sPathUI)
                Call .ZipFilesInFolder
                bOperationSuccess = True
            End If
        End If
    End With

CleanUp:
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(FileName:=sFullNameFile, UpdateLinks:=0)
    Application.DisplayAlerts = True
    On Error GoTo 0

    renameLiteralsToUI = bOperationSuccess
    Exit Function

ErrorHandler:
    bOperationSuccess = False
    Resume CleanUp
End Function

Public Sub WriteXML(ByRef arrUI As Variant, ByRef i As Long, ByRef oXMLDoc As MSXML2.DOMDocument)
    Dim sID         As String
    Dim oXMLNodeList As MSXML2.IXMLDOMNodeList
    
    ' 5. XML modification
    If oXMLDoc Is Nothing Then
        arrUI(i, UIColumns.UI_Status) = "XML not found"
    Else
        sID = arrUI(i, UIColumns.UI_IdOriginal)
        If sID <> vbNullString Then sID = "[@id='" & arrUI(i, UIColumns.UI_IdOriginal) & "']"


        Set oXMLNodeList = oXMLDoc.SelectNodes( _
                arrUI(i, UIColumns.UI_TagName) & sID)

        If oXMLNodeList.Length = 0 Then
            Set oXMLNodeList = oXMLDoc.SelectNodes( _
                    arrUI(i, UIColumns.UI_TagName) & "[@id='" & arrUI(i, UIColumns.UI_IdNew) & "']")
        End If

        If oXMLNodeList Is Nothing Then
            If oXMLNodeList.Length = 0 Then arrUI(i, UIColumns.UI_Status) = "id attribute not found"
        Else
            If ChangeAttribute(oXMLNodeList.Item(0).Attributes, _
                    CStr(arrUI(i, UIColumns.UI_AttrName)), _
                    CStr(arrUI(i, UIColumns.UI_AttrTextNew))) Then
                arrUI(i, UIColumns.UI_Status) = "modified"
            Else
                arrUI(i, UIColumns.UI_Status) = "not modified"
            End If

            If CStr(arrUI(i, UIColumns.UI_IdOriginal)) <> CStr(arrUI(i, UIColumns.UI_IdNew)) _
                    And Len(CStr(arrUI(i, UIColumns.UI_IdNew))) > 0 Then

                Call ChangeAttribute(oXMLNodeList.Item(0).Attributes, "id", CStr(arrUI(i, UIColumns.UI_IdNew)))
                arrUI(i, UIColumns.UI_Status) = arrUI(i, UIColumns.UI_Status) & vbNewLine & "id modified"
            End If
        End If
    End If
End Sub

Private Function ChangeAttribute(ByRef oXMLNodeMap As MSXML2.IXMLDOMNamedNodeMap, _
        ByVal sNameAttr As String, _
        ByVal sValue As String) As Boolean
    If sValue <> vbNullString Then
        If Not oXMLNodeMap Is Nothing Then
            Dim oNode As MSXML2.IXMLDOMNode
            Set oNode = oXMLNodeMap.getNamedItem(sNameAttr)

            If Not oNode Is Nothing Then
                oNode.text = sValue
                ChangeAttribute = True
            End If
        End If
    End If
End Function
