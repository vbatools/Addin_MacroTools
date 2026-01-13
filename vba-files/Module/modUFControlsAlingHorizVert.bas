Attribute VB_Name = "modUFControlsAlingHorizVert"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1



'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : vbaCntAlingHoriz - выравнивание контролов по строкам
'* Created    : 04-07-2022 14:39
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub vbaCntAlingHoriz()
    If Application.VBE.ActiveWindow.Type <> vbext_wt_Designer Then Exit Sub

    Dim lCnt        As Long
    Dim dTop        As Double
    Dim dLeft       As Double
    Dim dHeight     As Double
    Dim dWidth      As Double
    Dim dSPACE      As Variant
    Dim lColCnt     As Variant
    Dim dStart      As Double
    Dim dMaxWidth   As Double
    Dim objActiveModule As VBComponent
    Set objActiveModule = getActiveModule()

    lColCnt = Application.InputBox("Введите количество строк", "Выровнять по горизонтальной сетке:", Type:=1)
    If lColCnt <= 0 Or lColCnt = False Then Exit Sub
 
    dSPACE = Application.InputBox("Введите расстояние между фигурами", "Выровнять по горизонтальной сетке:", Type:=1)
    If TypeName(dSPACE) = "Boolean" Then Exit Sub

    lCnt = 1
    Dim cnt         As Control

    For Each cnt In objActiveModule.Designer.Selected
        With cnt
            If lCnt = 1 Then
                dStart = .top
            Else
                If lCnt Mod lColCnt = 1 Or lColCnt = 1 Then
                    'New column, move shape right
                    .top = dStart
                    .Left = dLeft + dMaxWidth + dSPACE
                    dMaxWidth = .Width
                Else
                    'Same column, move shape down
                    .top = dTop + dHeight + dSPACE
                    .Left = dLeft
                End If
            End If
            dTop = .top
            dLeft = .Left
            dHeight = .Height
            dWidth = .Width
            dMaxWidth = WorksheetFunction.Max(dMaxWidth, .Width)
        End With
        lCnt = lCnt + 1
    Next cnt
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : vbaCntAlingVert - выравнивание контролов по столбцам
'* Created    : 04-07-2022 14:39
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub vbaCntAlingVert()
    If Application.VBE.ActiveWindow.Type <> vbext_wt_Designer Then Exit Sub

    Dim lCnt        As Long
    Dim dTop        As Double
    Dim dLeft       As Double
    Dim dHeight     As Double
    Dim dWidth      As Double
    Dim dSPACE      As Variant
    Dim lColCnt     As Variant
    Dim dStart      As Double
    Dim dMaxHeight  As Double
    Dim objActiveModule As VBComponent
    Set objActiveModule = getActiveModule()

    lColCnt = Application.InputBox("Введите количество столбцов", "Выровнять по вертикальной сетке:", Type:=1)
    If lColCnt <= 0 Or lColCnt = False Then Exit Sub

    dSPACE = Application.InputBox("Введите расстояние между фигурами", "Выровнять по вертикальной сетке:", Type:=1)
    If TypeName(dSPACE) = "Boolean" Then Exit Sub

    lCnt = 1
    Dim cnt         As Control
    For Each cnt In objActiveModule.Designer.Selected
        With cnt
            If lCnt = 1 Then
                dStart = .Left
            Else
                If lCnt Mod lColCnt = 1 Or lColCnt = 1 Then
                    .top = dTop + dMaxHeight + dSPACE
                    .Left = dStart
                    dMaxHeight = .Height
                Else
                    .top = dTop
                    .Left = dLeft + dWidth + dSPACE
                End If
            End If
            dTop = .top
            dLeft = .Left
            dHeight = .Height
            dWidth = .Width
            dMaxHeight = WorksheetFunction.Max(dMaxHeight, .Height)
        End With
        lCnt = lCnt + 1
    Next cnt
End Sub
