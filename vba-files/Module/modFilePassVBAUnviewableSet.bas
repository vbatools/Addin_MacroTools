Attribute VB_Name = "modFilePassVBAUnviewableSet"
Option Explicit
Option Private Module
Const HOST_INFO     As String = "[Host Extender Info]"

Public Sub setProtectVBAUnviewable()

    Dim wb          As Workbook
    If Not GetTargetWorkbook(wb, "Setting [Unviewable] Password:", "SET") Then Exit Sub

    Application.ScreenUpdating = False
    Dim sFullNameFile As String
    sFullNameFile = wb.FullName
    wb.Close True
    If ProtectVBAUnviewable(sFullNameFile) Then
        'Call Workbooks.Open(sFullNameFile)
        Application.ScreenUpdating = True
        Call MsgBox("[Unviewable] password has been set!", vbInformation)
    Else
        Call MsgBox("[Unviewable] password was NOT set!", vbCritical)
    End If
End Sub

Private Function ProtectVBAUnviewable(ByVal sFullNameFile As String) As Boolean
    Dim clsZIP      As clsOfficeArchiveManager
    Set clsZIP = New clsOfficeArchiveManager
    With clsZIP
        If .Initialize(sFullNameFile, True) Then
            If .UnZipFile Then
                Dim fileData As String
                fileData = .getBinaryArrayVBAProject(adTypeText)
                If fileData <> vbNullString Then
                    Dim sSalt As String
                    Dim sUpPart As String
                    Dim sDownPart As String
                    Dim iLen As Long
                    sSalt = addSaltString(20)

                    iLen = VBA.InStr(fileData, "CMG=") - 1
                    If iLen > 0 Then
                        sUpPart = VBA.Left(fileData, iLen)
                    Else
                        Debug.Print ">> Found key [CMG=]:" & sFullNameFile
                    End If

                    iLen = InStrRev(fileData, HOST_INFO)
                    If iLen > 0 Then
                        sDownPart = VBA.Right(fileData, VBA.Len(fileData) - iLen + 1)
                    Else
                        Debug.Print ">> Found key [" & HOST_INFO & "]: " & sFullNameFile
                    End If

                    If sUpPart <> vbNullString And sDownPart <> vbNullString Then
                        ProtectVBAUnviewable = .putBinaryArrayVBAProject(sUpPart & vbNewLine & sSalt & sDownPart, adTypeText)
                    End If

                    If Not FileHave(.GetSettings(FolderPrinterSettings), vbDirectory) Then
                        Call MoveFile(.getPathVBAProject(), .GetSettings(FilePrinterSettings))
                        'VBAProject
                        sUpPart = sUpPart & sDownPart
                        sUpPart = VBA.Replace(sUpPart, "/", "\")
                        sUpPart = VBA.Replace(sUpPart, " ", VBA.Chr(5))
                        Call .putBinaryArrayVBAProject(VBA.Replace(sUpPart & sDownPart, "/", "\"), adTypeText)
                        ' changed path to project
                        Dim sXML As String
                        sXML = .readXMLFromFile(.GetSettings(ExlFileWorkBookRels))
                        sXML = VBA.Replace(sXML, "vbaProject.bin", "printerSettings.bin")
                        Call .writeXMLToFile(.GetSettings(ExlFileWorkBookRels), sXML)
                    End If
                Else
                    Debug.Print ">> No VBA project in file: " & sFullNameFile
                End If
            Else
                Debug.Print ">> No VBA project in file: " & sFullNameFile
            End If
            If .ZipFilesInFolder Then Debug.Print ">> [Unviewable] password has been set: " & sFullNameFile
        Else
            Debug.Print ">> Failed to unpack file: " & sFullNameFile
        End If
    End With
End Function

Private Function addSaltString(ByVal iLen As Integer) As String
    Dim i           As Integer
    For i = 1 To iLen
        addSaltString = addSaltString & "CMG=" & vbNewLine & "DPB=" & vbNewLine & "GC=" & vbNewLine
    Next i
End Function