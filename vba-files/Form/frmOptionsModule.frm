VERSION 5.0
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsModule 
   Caption         =   "OPTION:"
   ClientHeight    =   5190
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

Option Compare Text
Option Base 1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : addOptions - Adds OPTIONs to module headers
'* Created    : 17-09-2020 14:06
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



Private Sub chAll_Change()
    Dim bFlag       As Boolean
    bFlag = chAll.Value
    chOptionExplicit.Value = bFlag
    chOptionPrivate.Value = bFlag
    chOptionCompare.Value = bFlag
    chOptionBase.Value = bFlag
End Sub

Private Sub lbOK_Click()
    Unload Me
End Sub

Private Sub lbBase_Click()
    Dim sTxt        As String
    sTxt = "Sets the default lower bound for array subscripts. The default lower bound is 0." & vbNewLine & vbNewLine
    sTxt = sTxt & "Syntax" & vbNewLine & "Option Base { 0 | 1 }" & vbNewLine & vbNewLine
    sTxt = sTxt & "When Option Base appears in a file, it must precede any procedure-level declarations. If Option Base is not specified, the default lower bound is 0." & vbNewLine
    sTxt = sTxt & "Option Base statement must appear in the declarations section of a module before any procedures." & vbNewLine & vbNewLine
    sTxt = sTxt & "Remarks" & vbNewLine & vbNewLine
    sTxt = sTxt & "Used with To keyword in Dim, Private, Public, ReDim, and Static statements to declare array subscript ranges." & vbNewLine
    sTxt = sTxt & "If you don't specify a lower bound for an array, the lower bound is determined by the setting of Option Base," & vbNewLine
    sTxt = sTxt & "or by the default lower bound, which is 0. When calling functions that return arrays, such as Array, the lower bound" & vbNewLine
    sTxt = sTxt & "depends on the setting of Option Base unless the call to Array is preceded by LBound or UBound." & vbNewLine
    sTxt = sTxt & "Note that Option Base affects only arrays whose lower bound isn't explicitly stated in the Dim, Private, Public, ReDim, or Static statement that declares them."
    Debug.Print sTxt
End Sub
Private Sub lbCompare_Click()
    Dim sTxt        As String
    sTxt = "Sets the default method for comparing string data. The default method is binary." & vbNewLine & vbNewLine
    sTxt = sTxt & "Syntax" & vbNewLine & "Option Compare { Binary | Text | Database }" & vbNewLine & vbNewLine
    sTxt = sTxt & "Remarks" & vbNewLine & vbNewLine
    sTxt = sTxt & "Used when performing string comparisons. The compare setting affects all string comparison functions." & vbNewLine
    sTxt = sTxt & "Option Compare statement must appear in the declarations section of a module before any procedures." & vbNewLine
    sTxt = sTxt & "If no compare method is specified, the default is Binary. Option Compare Binary performs comparisons based on the internal binary representations of the characters." & vbNewLine
    sTxt = sTxt & "In Microsoft Windows, sorting is generally based on the internal binary representations of the characters." & vbNewLine
    sTxt = sTxt & "In this case, the following characters are sorted in ascending order:" & vbNewLine
    sTxt = sTxt & "A < B < E < Z < a < b < e < z < А < Б < В < Г < Д < Е" & vbNewLine
    sTxt = sTxt & "Option Compare Text performs comparisons based on a case-insensitive alphabetical comparison. " & vbNewLine
    sTxt = sTxt & "Thus, the following characters are sorted in ascending order: " & vbNewLine & vbNewLine
    sTxt = sTxt & "(A=a) < (B=b) < (E=e) < (Z=z) < (А=а) < (Б=б) < (В=в)" & vbNewLine & vbNewLine
    sTxt = sTxt & "Option Compare Database can only be used with Microsoft Access. When used, string comparisons are performed based on the sort order" & vbNewLine
    sTxt = sTxt & "of the database, which may differ from the default sort order. "
    Debug.Print sTxt
End Sub
Private Sub lbExplicit_Click()
    Dim sTxt        As String
    sTxt = "Forces explicit declaration of all variables in a file." & vbNewLine & vbNewLine
    sTxt = sTxt & "Syntax" & vbNewLine & "Option Explicit" & vbNewLine
    sTxt = sTxt & "Remarks" & vbNewLine & vbNewLine
    sTxt = sTxt & "Used to force explicit declaration of all variables in a module. When Option Explicit is used, all variables must be explicitly declared using Dim, Private, Public, ReDim or Static statements." & vbNewLine
    sTxt = sTxt & "If Option Explicit is not specified, variables are automatically declared as Variant type, which can lead to unexpected behavior when variable names are misspelled." & vbNewLine
    sTxt = sTxt & "Using Option Explicit helps catch typos and undeclared variables during compilation, which can prevent runtime errors and improve code reliability."
    Debug.Print sTxt
End Sub
Private Sub lbPrivate_Click()
    Dim sTxt        As String
    sTxt = "Makes all procedures and variables declared in a module private by default." & vbNewLine & vbNewLine
    sTxt = sTxt & "Syntax" & vbNewLine & "Option Private Module" & vbNewLine & vbNewLine
    sTxt = sTxt & "Remarks" & vbNewLine & vbNewLine
    sTxt = sTxt & "When Option Private Module is present, all procedures, variables, constants, properties, and events declared in the module are private by default," & vbNewLine
    sTxt = sTxt & "unless explicitly declared as Public, Friend, or Global. This prevents unintended exposure of internal module members." & vbNewLine
    sTxt = sTxt & "Microsoft Excel does not support this option. When Option Private Module is used, it makes the module behave as if it had private visibility by default."
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
        .top = Application.top + 0.5 * (Application.Height - .Height)
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
    lbModule.Caption = Application.VBE.ActiveCodePane.CodeModule.Parent.Name

    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 91:
            Unload Me
            Debug.Print "An error occurred, please check the form!"
            Exit Sub
        Case 76:
            Exit Sub
        Case Else:
            Debug.Print "Error in addOptions" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
            'Call WriteErrorLog("addOptions")
    End Select
    Err.Clear
End Sub
