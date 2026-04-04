Attribute VB_Name = "modFilePassVBAHideModule"
Option Explicit
Option Private Module
Const DELIMETR      As String = "||"
Const MODULE_NAME   As String = "Module="

Public Sub hideModules(ByVal sFullNameFile As String, ByRef arrNameModule As Variant)
    Dim iCount      As Long
    iCount = UBound(arrNameModule, 1)
    If iCount = 0 Then Exit Sub
    Dim clsZIP      As clsOfficeArchiveManager
    Set clsZIP = New clsOfficeArchiveManager
    With clsZIP
        If .Initialize(sFullNameFile, False) Then
            If .UnZipFile Then
                Dim fileData() As Byte
                fileData = .getBinaryArrayVBAProject(adTypeBinary)
                If Not (Not (fileData)) = 0 Then
                    Dim sFileData As String
                    Dim i As Long
                    Dim k As Long
                    Dim sModule As String
                    Dim sAllName As String
                    Dim sNameModuleByte As String
                    Dim sEmptyString As String
                    sModule = arrayByteJoin(VBA.StrConv(MODULE_NAME, vbFromUnicode))
                    sFileData = arrayByteJoin(fileData)
                    For i = 1 To iCount
                        sNameModuleByte = arrNameModule(i, 1)
                        If sNameModuleByte <> vbNullString Then
                            sNameModuleByte = arrayByteJoin(VBA.StrConv(sNameModuleByte, vbFromUnicode))
                            sAllName = sModule & DELIMETR & sNameModuleByte & DELIMETR & 13 & DELIMETR & 10
                            If sFileData Like "*" & sModule & "*" Then
                                sEmptyString = addEmptyString(VBA.Len(MODULE_NAME & arrNameModule(i, 1))) & "13||10"
                                sFileData = VBA.Replace(sFileData, sAllName, sEmptyString)
                                k = k + 1
                            End If
                        End If
                    Next i
                    Call .putBinaryArrayVBAProject(arrStringToByte(VBA.Split(sFileData, DELIMETR)), adTypeBinary)
                Else
                    Debug.Print ">> No VBA project in file: " & sFullNameFile
                End If
            Else
                Debug.Print ">> No VBA project in file: " & sFullNameFile
            End If
            If .ZipFilesInFolder Then Debug.Print ">> Modules hidden [" & k & "]: " & sFullNameFile
        Else
            Debug.Print ">> Failed to unpack file: " & sFullNameFile
        End If
    End With
End Sub

Private Function addEmptyString(ByRef iLen As Long) As String
    Dim i           As Long
    For i = 1 To iLen
        addEmptyString = addEmptyString & 32 & DELIMETR
    Next
    addEmptyString = addEmptyString
End Function

Private Function arrStringToByte(ByRef arr As Variant) As Byte()
    Dim i           As Long
    ReDim arrRes(0 To UBound(arr, 1)) As Byte
    For i = 0 To UBound(arr, 1)
        arrRes(i) = VBA.CByte(arr(i))
    Next i
    arrStringToByte = arrRes
End Function


Private Function arrayByteJoin(ByRef fileData() As Byte) As String
    Dim i           As Long
    Dim lCountL      As Long
    Dim lCountU      As Long
    Dim vTemp()     As Variant

    ' Check if the array is initialized
    If (Not fileData) = True Then
        arrayByteJoin = ""
        Exit Function
    End If

    lCountL = LBound(fileData)
    lCountU = UBound(fileData)
    ReDim vTemp(lCountL To lCountU)

    For i = lCountL To lCountU
        vTemp(i) = CStr(fileData(i))
    Next i
    arrayByteJoin = Join(vTemp, DELIMETR)
End Function