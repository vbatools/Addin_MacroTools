Attribute VB_Name = "modToolsSnippetsParserJSON"
Option Explicit

'Public Sub test()
'    Dim sPathJSON   As String
'    sPathJSON = ""
'
'    ' Проверка существования файла
'    If Not FileHave(sPathJSON, vbNormal) Then
'        MsgBox "Файл не найден: " & sPathJSON, vbExclamation, "Внимание"
'        Exit Sub
'    End If
'
'    Dim sJSON       As String
'    sJSON = loadTextFromTextFile(sPathJSON)
'    Call parserJSONtoTableSnippets(sJSON)
'End Sub


Sub ReadSnippetsJsonFromGitHub()
    If MsgBox("Download the snippet database from GitHub?", vbYesNo + vbQuestion, "Download Snippet Database:") = vbNo Then Exit Sub

    Dim httpReq     As Object
    Set httpReq = CreateObject("MSXML2.XMLHTTP")

    Dim jsonResponse As String
    Const JSON_URL  As String = "https://raw.githubusercontent.com/vbatools/Addin_MacroToolsVBA_Snippets/refs/heads/main/SNIPPETS.json"

    With httpReq
        .Open "GET", JSON_URL, False
        .send
        If .Status = 200 Then
            Call parserJSONtoTableSnippets(.responseText)
        Else
            MsgBox "Download error: " & .Status, vbCritical
        End If
    End With

    Set httpReq = Nothing
End Sub

Public Sub parserJSONtoTableSnippets(ByVal sJSON As String)

    ' Константы размерности массива
    Const COL_COUNT As Long = 9

    ' Номера столбцов
    Const COL_CODE_GRUP As Long = 1
    Const COL_CODE_SNIPPET As Long = 2
    Const COL_CODE  As Long = 3
    Const COL_DESCRIPTION As Long = 4
    Const COL_CLS_NAME As Long = 5
    Const COL_CLS_REF As Long = 6
    Const COL_FRM_NAME As Long = 7
    Const COL_FRM_REF As Long = 8
    Const COL_FRX_REF As Long = 9

    On Error GoTo ErrorHandler



    Dim objJSON     As Object
    Set objJSON = ParseJson(sJSON)

    If objJSON Is Nothing Then
        MsgBox "JSON parsing error.", vbExclamation
        Exit Sub
    End If

    Set objJSON = objJSON("data")
    If objJSON Is Nothing Then
        MsgBox "Missing 'data' section in JSON.", vbExclamation
        Exit Sub
    End If

    Dim i           As Long
    Dim iCount      As Long
    Dim JsonItem    As Object
    Dim vItem       As Variant
    Dim arrtable    As Variant

    iCount = objJSON.Count
    If iCount = 0 Then
        MsgBox "No data available for loading.", vbExclamation
        Exit Sub
    End If

    ReDim arrtable(1 To iCount, 1 To COL_COUNT)

    Dim sBaseName   As String
    Dim sEXP        As String
    Dim lColName    As Long
    Dim lColRef     As Long
    Dim sPrefix     As String

    For i = 1 To iCount
        Set JsonItem = objJSON(i)
        For Each vItem In JsonItem
            Select Case vItem
                Case "CODE_GRUP"
                    arrtable(i, COL_CODE_GRUP) = JsonItem(vItem)

                Case "CODE_SNIPPET"
                    arrtable(i, COL_CODE_SNIPPET) = JsonItem(vItem)

                Case "CODE"
                    arrtable(i, COL_CODE) = JsonItem(vItem)

                Case "DISCRIPTION", "DESCRIPTION"
                    arrtable(i, COL_DESCRIPTION) = JsonItem(vItem)

                Case Else
                    sBaseName = sGetBaseName(vItem)
                    sEXP = VBA.LCase$(sGetExtensionName(vItem))

                    Select Case sEXP
                        Case "cls"
                            lColName = COL_CLS_NAME
                            lColRef = COL_CLS_REF
                            sPrefix = "TB_CLS_"

                        Case "bas"
                            lColName = COL_CLS_NAME
                            lColRef = COL_CLS_REF
                            sPrefix = "TB_MOD_"

                        Case "frm"
                            lColName = COL_FRM_NAME
                            lColRef = COL_FRM_REF
                            sPrefix = "TB_FRM_"

                        Case "frx"
                            lColName = COL_FRX_REF
                            lColRef = COL_FRX_REF
                            sPrefix = "TB_FRX_"

                        Case Else
                            GoTo NextItem
                    End Select

                    ' Для frx не добавляем имя файла
                    If sEXP <> "frx" Then
                        If arrtable(i, lColName) <> vbNullString Then
                            arrtable(i, lColName) = arrtable(i, lColName) & ";"
                        End If
                        arrtable(i, lColName) = arrtable(i, lColName) & sBaseName
                    End If

                    ' Добавление ссылки
                    If arrtable(i, lColRef) <> vbNullString Then
                        arrtable(i, lColRef) = arrtable(i, lColRef) & ";"
                    End If
                    arrtable(i, lColRef) = arrtable(i, lColRef) & sPrefix & sBaseName

                    Call AddShape(sPrefix & sBaseName, JsonItem(vItem))

            End Select

NextItem:
        Next vItem
    Next i



    ' Выгрузка в таблицу
    With shSettings.ListObjects("TB_SNIPETS")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
        .Range(2, 1).Resize(iCount, COL_COUNT).Value2 = arrtable
        .DataBodyRange.RowHeight = 15
    End With
    
    MsgBox "Snippets imported successfully!", vbInformation

ExitProc:
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Critical Error"
    Resume ExitProc

End Sub

Private Sub DeleteShapes()
    Dim shp         As Shape
    Dim patterns    As Variant
    Dim i           As Long

    patterns = Array("TB_CLS_*", "TB_MOD_*", "TB_FRM_*", "TB_FRX_*")

    For Each shp In shSettings.Shapes
        For i = LBound(patterns) To UBound(patterns)
            If shp.Name Like patterns(i) Then
                shp.Delete
                Exit For
            End If
        Next i
    Next shp
End Sub

Private Sub AddShape(ByVal sNameShape, ByVal sValue As String)
    If haveShape(sNameShape) Then Exit Sub
    Dim shp         As Shape
    Set shp = shSettings.Shapes.AddShape(msoShapeRectangle, 20, 20, 20, 20)
    With shp
        .Name = sNameShape
        .TextFrame2.TextRange.text = sValue
    End With
End Sub

Private Function haveShape(ByVal sNameShape As String) As Boolean
    On Error Resume Next
    haveShape = shSettings.Shapes(sNameShape).Name = sNameShape
    On Error GoTo 0
End Function

Private Function loadTextFromTextFile(ByVal fileName As String, Optional ByVal encoding As String) As String

    On Error Resume Next
    If Len(Trim(encoding)) = 0 Then encoding = "windows-1251"
    With CreateObject("ADODB.Stream")
        .Type = 2:
        If Len(encoding) Then .Charset = encoding
        .Open
        .LoadFromFile fileName
        loadTextFromTextFile = .ReadText
        .Close
    End With
End Function
