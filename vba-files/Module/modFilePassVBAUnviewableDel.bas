Attribute VB_Name = "modFilePassVBAUnviewableDel"
Option Explicit
Option Private Module

Public Sub delProtectVBAUnviewable()

    Dim wb          As Workbook
    If Not GetTargetWorkbook(wb, "Removing [Unviewable] Password:", "REMOVE PASSWORD") Then Exit Sub

    Dim sFullNameFile As String
    Application.ScreenUpdating = False
    sFullNameFile = wb.FullName
    wb.Close True
    If unProtectVBAUnviewable(sFullNameFile) Then
        'Call Workbooks.Open(sFullNameFile)
        Application.ScreenUpdating = True
        Call MsgBox("[Unviewable] password has been removed!", vbInformation)
    Else
        Call MsgBox("[Unviewable] password was NOT removed!", vbCritical)
    End If
End Sub

Private Function unProtectVBAUnviewable(ByVal sFullNameFile As String) As Boolean
    Dim clsZIP      As clsOfficeArchiveManager
    Set clsZIP = New clsOfficeArchiveManager
    With clsZIP
        If .Initialize(sFullNameFile, True) Then
            If .UnZipFile Then
                Dim fileData As String
                fileData = .getBinaryArrayVBAProject(adTypeText)
                If fileData <> vbNullString Then
                    Dim sVBA As String
                    sVBA = fileData
                    Dim arr As Variant
                    Dim i As Long
                    Dim sQ As String
                    sQ = VBA.Chr$(34) & "*" & VBA.Chr$(34)
                    arr = VBA.Split(sVBA, vbNewLine)

                    For i = 0 To UBound(arr, 1)
                        If arr(i) Like "CMG=*" Then
                            arr(i) = "CMC="
                        ElseIf arr(i) Like "DPB=*" Then
                            arr(i) = "DPC="
                        ElseIf arr(i) Like "GC=*" Then
                            arr(i) = "CC="
                        End If
                    Next i
                    sVBA = VBA.Join(arr, vbNewLine)

                    unProtectVBAUnviewable = .putBinaryArrayVBAProject(sVBA, adTypeText)
                Else
                    Debug.Print ">> No VBA project in file: " & sFullNameFile
                End If
            Else
                Debug.Print ">> No VBA project in file: " & sFullNameFile
            End If
            If .ZipFilesInFolder Then Debug.Print ">> [Unviewable] password has been removed: " & sFullNameFile
        Else
            Debug.Print ">> Failed to unpack file: " & sFullNameFile
        End If
    End With
End Function