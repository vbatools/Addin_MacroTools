VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsModule 
   Caption         =   "OPTION:"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   OleObjectBlob   =   "frmOptionsModule.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOptionsModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : addOptions - Create OPTIONs in project modules
'* Created    : 17-09-2020 14:06
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub chAll_Change()
    Dim bFlag       As Boolean
    bFlag = chAll.value
    chOptionExplicit.value = bFlag
    chOptionPrivate.value = bFlag
    chOptionCompare.value = bFlag
    chOptionBase.value = bFlag
    chModuleName.value = bFlag
End Sub

Private Sub lbOK_Click()
    Unload Me
End Sub

Private Sub lbBase_Click()
    Dim sTxt        As String
    sTxt = "Used at module level to declare the default lower bound for arrays." & vbNewLine & vbNewLine
    sTxt = sTxt & "Syntax" & vbNewLine & "Option Base { 0 | 1 }" & vbNewLine & vbNewLine
    sTxt = sTxt & "Since Option Base defaults to 0, the Option Base statement is never required. The statement must appear in a module before any procedures." & vbNewLine
    sTxt = sTxt & "The Option Base statement can appear only once in a module and must precede array declarations that include dimensions." & vbNewLine & vbNewLine
    sTxt = sTxt & "Notes" & vbNewLine & vbNewLine
    sTxt = sTxt & "The To clause in Dim, Private, Public, ReDim, and Static statements provides a more flexible way to control the index range of an array." & vbNewLine
    sTxt = sTxt & "However, if the lower bound is not explicitly specified in the To clause, you can use the Option Base statement," & vbNewLine
    sTxt = sTxt & "to set the default lower bound to 1. The lower bound of arrays," & vbNewLine
    sTxt = sTxt & "created with the Array function is always zero, regardless of the Option Base statement."
    sTxt = sTxt & vbNewLine & vbNewLine & "The Option Base statement affects only the lower bound of arrays in the module where it is located."
    Debug.Print sTxt
End Sub
Private Sub lbCompare_Click()
    Dim sTxt        As String
    sTxt = "Used at module level to declare the default comparison method for string data comparison." & vbNewLine & vbNewLine
    sTxt = sTxt & "Syntax" & vbNewLine & "Option Compare { Binary | Text | Database }" & vbNewLine & vbNewLine
    sTxt = sTxt & "Notes" & vbNewLine & vbNewLine
    sTxt = sTxt & "The Option Compare statement, if used, must appear in a module before any procedures." & vbNewLine
    sTxt = sTxt & "The Option Compare statement specifies the string comparison method (Binary, Text, or Database) for a module." & vbNewLine
    sTxt = sTxt & "If a module does not contain an Option Compare statement, the default comparison method is Binary." & vbNewLine
    sTxt = sTxt & "Option Compare Binary results in string comparisons based on a sort order derived from the internal binary representations of the characters." & vbNewLine
    sTxt = sTxt & "In Microsoft Windows, the sort order is determined by the code page." & vbNewLine
    sTxt = sTxt & "A typical binary sort order is shown in the following example:" & vbNewLine & vbNewLine
    sTxt = sTxt & "A < B < E < Z < a < b < e < z < Б < Л < Ш < б < л < ш" & vbNewLine & vbNewLine
    sTxt = sTxt & "Option Compare Text results in case-insensitive string comparisons based on your system's locale settings." & vbNewLine
    sTxt = sTxt & "The same characters sorted with Option Compare Text produce the following order:" & vbNewLine & vbNewLine
    sTxt = sTxt & "(A=a) < (B=b) < (E=e) < (Z=z) < (Б=б) < (Л=л) < (Ш=ш)" & vbNewLine & vbNewLine
    sTxt = sTxt & "Option Compare Database can only be used in Microsoft Access. It results in string comparisons based on the sort order," & vbNewLine
    sTxt = sTxt & "determined by the locale ID of the database where the string comparison occurs."
    Debug.Print sTxt
End Sub
Private Sub lbExplicit_Click()
    Dim sTxt        As String
    sTxt = "Used at module level to force explicit declaration of all variables in that module." & vbNewLine & vbNewLine
    sTxt = sTxt & "Syntax" & vbNewLine & "Option Explicit" & vbNewLine & vbNewLine
    sTxt = sTxt & "Notes" & vbNewLine & vbNewLine
    sTxt = sTxt & "The Option Explicit statement, if used, must appear in a module before any procedures." & vbNewLine
    sTxt = sTxt & "When Option Explicit is used, you must explicitly declare all variables using Dim, Private, Public, ReDim, or Static statements." & vbNewLine
    sTxt = sTxt & "If you attempt to use an undeclared variable name, a compile-time error occurs." & vbNewLine
    sTxt = sTxt & "If Option Explicit is not used, all undeclared variables are of Variant type unless the default type is specified with a DefType statement." & vbNewLine
    sTxt = sTxt & "Use Option Explicit to avoid misspelling existing variable names or to avoid confusion when variable scope is unclear."
    Debug.Print sTxt
End Sub
Private Sub lbPrivate_Click()
    Dim sTxt        As String
    sTxt = "Used at module level to prevent references to module content from outside the project." & vbNewLine & vbNewLine
    sTxt = sTxt & "Syntax" & vbNewLine & "Option Private Module" & vbNewLine & vbNewLine
    sTxt = sTxt & "Notes" & vbNewLine & vbNewLine
    sTxt = sTxt & "When a module contains Option Private Module, public elements such as variables, objects, and user-defined types declared at module level," & vbNewLine
    sTxt = sTxt & "remain available within the project containing the module, but are not available to other applications or projects." & vbNewLine
    sTxt = sTxt & "Microsoft Excel supports loading multiple projects. In this case, Option Private Module restricts mutual visibility between projects."
    Debug.Print sTxt
End Sub

Private Sub cmbCancel_Click()
    Unload Me
End Sub
Private Sub lbCancel_Click()
    Call cmbCancel_Click
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With
End Sub

Private Sub UserForm_Activate()
    On Error GoTo ErrorHandler
    With Application.CommandBars
        lbExplicit.Picture = .GetImageMso("Help", 18, 18)
        lbPrivate.Picture = .GetImageMso("Help", 18, 18)
        lbCompare.Picture = .GetImageMso("Help", 18, 18)
        lbBase.Picture = .GetImageMso("Help", 18, 18)
    End With
    lbModule.Caption = Application.VBE.ActiveCodePane.codeModule.Parent.Name

    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Unload Me
            Debug.Print ">> No active module, please switch to a code module!"
            Exit Sub
        Case 76:
            Exit Sub
        Case Else:
            Call WriteErrorLog(Me.Name & ".UserForm_Activate", True)
    End Select
    Err.Clear
End Sub