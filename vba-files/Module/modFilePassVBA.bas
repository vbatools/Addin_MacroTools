Attribute VB_Name = "modFilePassVBA"
Option Explicit
Option Private Module
#If VBA7 Then
    Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function VirtualProtect Lib "kernel32" (lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flNewProtect As LongPtr, lpflOldProtect As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare PtrSafe Function DialogBoxParam Lib "USER32" Alias "DialogBoxParamA" (ByVal hInstance As LongPtr, ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer
#Else
    Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Any)
    Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Any, ByVal flNewProtect As Any, lpflOldProtect As Any) As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Any, ByVal lpProcName As String) As Long
    Private Declare Function DialogBoxParam Lib "USER32" Alias "DialogBoxParamA" (ByVal hInstance As Any, ByVal pTemplateName As Any, ByVal hWndParent As Any, ByVal lpDialogFunc As Any, ByVal dwInitParam As Any) As Integer
#End If

Private Const PAGE_EXECUTE_READWRITE = &H40
Dim HookBytes(0 To 11) As Byte
Dim OriginBytes(0 To 11) As Byte
Dim pFunc           As LongPtr
Dim Flag            As Boolean

Public Sub unProtectVBA()
    If unProtectVBAProjects Then
        Call MsgBox("VBA project passwords removed!", vbInformation)
        Else
        Call MsgBox("Failed to remove VBA project passwords!", vbCritical)
        End If
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : fHook - remove passwords from VBA projects
'* Created    : 23-03-2023 09:25
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function unProtectVBAProjects() As Boolean
    Dim TmpBytes(0 To 11) As Byte
    Dim p As LongPtr, osi As Byte
    Dim OriginProtect As LongPtr

    unProtectVBAProjects = False
    #If Win64 Then
        osi = 1
    #Else
        osi = 0
    #End If

    pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
    If VirtualProtect(ByVal pFunc, 12, PAGE_EXECUTE_READWRITE, OriginProtect) <> 0 Then

        MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, osi + 1
        If TmpBytes(osi) <> &HB8 Then

            MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 12
            p = GetPtr(AddressOf MyDialogBoxParam)

            If osi Then HookBytes(0) = &H48
            HookBytes(osi) = &HB8
            osi = osi + 1
            MoveMemory ByVal VarPtr(HookBytes(osi)), ByVal VarPtr(p), 4 * osi
            HookBytes(osi + 4 * osi) = &HFF
            HookBytes(osi + 4 * osi + 1) = &HE0

            MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 12
            Flag = True
            unProtectVBAProjects = True
            End If
        End If
End Function

Private Function GetPtr(ByVal lpValue As LongPtr) As LongPtr
    GetPtr = lpValue
End Function

Private Sub RecoverBytes()
    If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 12
End Sub

Private Function MyDialogBoxParam(ByVal hInstance As LongPtr, _
        ByVal pTemplateName As LongPtr, ByVal hWndParent As LongPtr, _
        ByVal lpDialogFunc As LongPtr, ByVal dwInitParam As LongPtr) As Integer

    If pTemplateName = 4070 Then
        MyDialogBoxParam = 1
        Else
        Call RecoverBytes
        MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                hWndParent, lpDialogFunc, dwInitParam)
        Call unProtectVBAProjects
        End If
End Function