Attribute VB_Name = "modLiteralsGetUI"
Option Explicit
Option Private Module

Public Function parserLiteralsFormUIOnlyProcedures(ByRef wb As Workbook, ByRef bZIPFile As Boolean) As Dictionary
      Dim arrAtr      As Variant
      arrAtr = Array("getContent", "getDescription", "getEnabled", "getHelperText", "getImage", "getImageMso", "getItemCount", _
            "getItemHeight", "getItemID", "getItemImage", "getItemLabel", "getItemScreentip", "getItemSupertip", "getItemWidth", _
            "getKeytip", "getLabel", "getPressed", "getScreentip", "getSelectedItemID", "getSelectedItemIndex", "getShowImage", _
            "getShowLabel", "getSize", "getStyle", "getSupertip", "getText", "getVisible", "loadImage", "onChange", "onHide", "onLoad", _
            "onAction", "onShow")
    Set parserLiteralsFormUIOnlyProcedures = parserLiteralsFormUI(wb, arrAtr, True, bZIPFile)
End Function


Public Function parserLiteralsFormUI(ByRef wb As Workbook, ByRef arrAtr As Variant, ByRef bProc As Boolean, ByRef bZIPFile As Boolean) As Dictionary
    Dim sFullNameFile As String
    Dim clsZIP      As clsOfficeArchiveManager
    Dim oDic        As Dictionary

    Set oDic = New Dictionary
    sFullNameFile = wb.FullName
    wb.Close True
    On Error GoTo ErrorHandler

    Set clsZIP = New clsOfficeArchiveManager
    With clsZIP
        If .Initialize(sFullNameFile, True) Then
            If .UnZipFile Then
                Call ProcessXMLPart(clsZIP, FileCustomUI, "CustomUI", arrAtr, oDic, bProc)
                Call ProcessXMLPart(clsZIP, FileCustomUI14, "CustomUI14", arrAtr, oDic, bProc)
                If bZIPFile Then Call .ZipFilesInFolder
            End If
        End If
    End With
CleanUp:
    If bZIPFile Then
        On Error Resume Next
        Application.DisplayAlerts = False
        Set wb = Workbooks.Open(FileName:=sFullNameFile, UpdateLinks:=0)
        Application.DisplayAlerts = True
        On Error GoTo 0
    End If

    Set parserLiteralsFormUI = oDic
    Exit Function

ErrorHandler:
    Debug.Print ">> Error in parserLiteralsFormUI: " & Err.Description
    Resume CleanUp
End Function

' Helper procedure for processing a specific XML part
Private Sub ProcessXMLPart(ByRef zipMgr As clsOfficeArchiveManager, ByVal fileSetting As Long, _
        ByVal partName As String, ByRef arrAtr As Variant, ByRef oDic As Dictionary, ByRef bProc As Boolean)
    Dim oXMLDoc     As MSXML2.DOMDocument
    Dim oXMLNode    As MSXML2.IXMLDOMNode
    Dim sInitialPath As String

    Set oXMLDoc = zipMgr.getXMLDOC(zipMgr.GetSettings(fileSetting))

    If Not oXMLDoc Is Nothing Then
        If bProc Then
            Set oXMLNode = oXMLDoc.SelectSingleNode("customUI")
            Call getLitersFromXMLNode(partName, oXMLNode, Array("onLoad", "loadImage"), oDic, "customUI")
        End If

        Set oXMLNode = oXMLDoc.SelectSingleNode("customUI/ribbon/tabs")

        ' Build the initial path. Since SelectSingleNode already points to "customUI/ribbon/tabs",
        ' we pass this path as the base.
        sInitialPath = "customUI/ribbon/tabs"

        ' Call the recursive function with the initial path
        Call getLitersFromXML(partName, oXMLNode, arrAtr, oDic, sInitialPath)
    End If
End Sub

' Recursive procedure for traversing XML
Private Sub getLitersFromXML(ByRef sNameUIXML As String, ByRef oXMLNode As MSXML2.IXMLDOMNode, ByRef arrAtr As Variant, ByRef oDic As Dictionary, ByVal sCurrentPath As String)
    Dim i           As Long
    Dim childNode   As MSXML2.IXMLDOMNode
    Dim sChildPath  As String

    If oXMLNode Is Nothing Then Exit Sub

    ' First, process the attributes of the current node using the current path
    If Not oXMLNode.Attributes Is Nothing Then
        Call getLitersFromXMLNode(sNameUIXML, oXMLNode, arrAtr, oDic, sCurrentPath)
    End If

    ' Recursive traversal of child nodes
    For i = 0 To oXMLNode.ChildNodes.Length - 1
        Set childNode = oXMLNode.ChildNodes(i)

        ' If this is an element (tag), not text or a comment
        If childNode.NodeType = NODE_ELEMENT Then
            ' Build the path for the descendant: CurrentPath/DescendantName
            sChildPath = sCurrentPath & "/" & childNode.BaseName

            ' Recursive call
            Call getLitersFromXML(sNameUIXML, childNode, arrAtr, oDic, sChildPath)
        End If
    Next i
End Sub

' Procedure for extracting data from a node
Private Sub getLitersFromXMLNode(ByRef sNameUIXML As String, ByRef oXMLNode As MSXML2.IXMLDOMNode, ByRef arrAtr As Variant, ByRef oDic As Dictionary, ByVal sXPath As String)
    Dim i           As Long
    Dim sKey        As String
    ' Expand the array to 5 elements: 1-PartName, 2-NodeName, 3-AttrName, 4-AttrValue, 5-FullPath
    Dim arrData(1 To 1, 1 To UIColumns.[_Last]) As String
    Dim oAttr       As MSXML2.IXMLDOMAttribute
    Dim id          As String

    If oXMLNode Is Nothing Then Exit Sub

    With oXMLNode.Attributes
        For i = 0 To UBound(arrAtr)
            Set oAttr = .getNamedItem(arrAtr(i))

            If Not oAttr Is Nothing Then
                If Not .getNamedItem("id") Is Nothing Then id = .getNamedItem("id").text
                arrData(1, UIColumns.UI_ModuleType) = sNameUIXML
                arrData(1, UIColumns.UI_XMLNodeName) = oXMLNode.BaseName
                arrData(1, UIColumns.UI_TagName) = sXPath
                arrData(1, UIColumns.UI_IdOriginal) = id
                arrData(1, UIColumns.UI_AttrName) = arrAtr(i)
                arrData(1, UIColumns.UI_AttrText) = oAttr.text
                ' Build the key (keep your logic, can be changed if needed)
                sKey = arrData(1, UIColumns.UI_ModuleType) & "." & arrData(1, UIColumns.UI_TagName) & _
                        "." & arrData(1, UIColumns.UI_IdOriginal) & "." & arrData(1, UIColumns.UI_AttrName)

                If Not oDic.Exists(sKey) Then Call oDic.Add(sKey, arrData)
            End If
        Next i
    End With
End Sub