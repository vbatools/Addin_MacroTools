VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInfoFileLastAutor 
   Caption         =   "File Properties:"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9135.001
   OleObjectBlob   =   "frmInfoFileLastAutor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInfoFileLastAutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : InfoFile2 - Change Last Author and Last Save Time properties
'* Created    : 20-07-2020 15:34
'* Author     : VBATools
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private Sub cmbMain_Change()
    txtLastAuthor.value = GetOneProp(Workbooks(cmbMain.value), "Last author")
    txtLastAuthorOld.value = txtLastAuthor.value
    txtLastSaveTime.value = GetOneProp(Workbooks(cmbMain.value), "Last save time")
    txtLastSaveTimeOld.value = txtLastSaveTime.value
End Sub

Private Sub lbOK_Click()
    If txtLastSaveTimeOld.value = txtLastSaveTime.value And txtLastAuthorOld.value = txtLastAuthor.value Then Exit Sub
    If IsDate(txtLastSaveTime.value) Then
        Dim wb      As Workbook
        Dim sFullNameFile As String
        Set wb = Workbooks(cmbMain.value)
        sFullNameFile = wb.FullName
        wb.Close savechanges:=True
        Call WriteXML(sFullNameFile, txtLastAuthor.text, CDate(txtLastSaveTime.value))
        Workbooks.Open FileName:=sFullNameFile
        Call MsgBox("Changes saved to file!", vbInformation, "Changes:")
        Unload Me
    Else
        Call MsgBox("[ Last save time ] field does not contain a valid date!", vbCritical, "Error:")
    End If
End Sub

Private Sub WriteXML(ByVal sFullNameFile As String, ByVal LastAuthor As String, ByVal lastTime As Date, Optional bBackUp As Boolean = False)

    Dim cls         As clsOfficeArchiveManager
    Dim sPathXML    As String
    Dim objXMLDOC   As MSXML2.DOMDocument
    Set cls = New clsOfficeArchiveManager

    With cls
        If .Initialize(sFullNameFile, bBackUp) Then
            If .UnZipFile() Then
                sPathXML = .GetSettings(FolderDocProps) & Application.PathSeparator & "core.xml"
                Set objXMLDOC = .getXMLDOC(sPathXML)
                If Not objXMLDOC Is Nothing Then
                    With objXMLDOC.SelectSingleNode("cp:coreProperties")
                        .SelectSingleNode("cp:lastModifiedBy").nodeTypedValue = LastAuthor
                        .SelectSingleNode("dcterms:modified").nodeTypedValue = VBA.Format$(lastTime, "yyyy-mm-ddThh:mm:ssZ")
                    End With
                Else
                    Call MsgBox("core.xml not found", vbCritical)
                End If
                Call objXMLDOC.Save(sPathXML)
                Call .ZipFilesInFolder
            Else
                Call MsgBox("Could not open file archive", vbCritical)
            End If
        End If
    End With

    Set objXMLDOC = Nothing
    Set cls = Nothing
End Sub

Private Sub UserForm_Activate()
    Dim vbProj      As VBIDE.vbProject
    If Workbooks.Count = 0 Then
        Unload Me
        Call MsgBox("No open" & Chr(34) & "Excel files" & Chr(34) & "!", vbOKOnly + vbExclamation, "Error:")
        Exit Sub
    End If
    With cmbMain
        .Clear
        On Error Resume Next
        For Each vbProj In Application.VBE.VBProjects
            .AddItem sGetFileName(vbProj.FileName)
        Next
        On Error GoTo 0
        .value = ActiveWorkbook.Name
    End With
End Sub

Private Sub cmbCancel_Click()
    Unload Me
End Sub
Private Sub lbCancel_Click()
    Call cmbCancel_Click
End Sub
Private Sub UserForm_Initialize()
    StartUpPosition = 0
    Left = Application.Left + (0.5 * Application.Width) - (0.5 * Width)
    Top = Application.Top + (0.5 * Application.Height) - (0.5 * Height)
End Sub