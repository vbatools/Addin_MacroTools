VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAnchors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'*  Class: clsAnchors
'*  Author: VBATools
'*  Version: 1.1.0
'*  Creation Date: 03.05.2019 09:31
'*  Last Update Date: 12.11.2025 23:57
'*  License: Apache License
'*  Description: Provides anchor-based resizing for MSForms UserForm controls
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'==========================================================================================
' ENUMERATIONS
'==========================================================================================

Public Enum AnchorEdge
    anchorNone = 0
    anchorTop = 1
    anchorBottom = 2
    anchorLeft = 4
    anchorRight = 8
End Enum

Private Enum ControlParamIndex
    idxAnchorStyle = 1
    idxMinimumWidth
    idxMinimumHeight
    idxMinimumTop
    idxMinimumLeft
    idxOriginalLeft
    idxOriginalTop
    idxOriginalWidth
    idxOriginalHeight
    idxControlName
    [_First] = idxAnchorStyle
    [_Last] = idxControlName
End Enum

'==========================================================================================
' CONSTANTS
'==========================================================================================

Private Const CLASS_NAME As String = "clsAnchors"
Private Const CLASS_VERSION As String = "1.1.0"

Private Const RESIZE_HANDLE_SUFFIX As String = "ResizeGrabHandle"
Private Const HANDLE_SUFFIX_BOTTOM As String = "_BOTTOM"
Private Const HANDLE_SUFFIX_RIGHT As String = "_RIGHT"

Private Const MOUSE_BUTTON_LEFT As Integer = 1

Private Const DEFAULT_HANDLE_SIZE As Single = 2
Private Const DEFAULT_HANDLE_FONT_SIZE As Single = 14
Private Const HANDLE_COLOR As Long = &HC0C0C0

'==========================================================================================
' MODULE-LEVEL VARIABLES
'==========================================================================================

' Parent form reference
Private mParent As Object

' Collection of control parameters
Private mItems As Collection

' Resize handles
Private WithEvents mResizeHandle As MSForms.Label
Attribute mResizeHandle.VB_VarHelpID = -1
Private WithEvents mResizeHandleBottom As MSForms.Label
Attribute mResizeHandleBottom.VB_VarHelpID = -1
Private WithEvents mResizeHandleRight As MSForms.Label
Attribute mResizeHandleRight.VB_VarHelpID = -1

' State flags
Private mIsResizing As Boolean
Private mUpdateWhileDragging As Boolean
Private mIsInitialized As Boolean

' Resize tracking
Private mResizeStartX As Single
Private mResizeStartY As Single

' Form dimension constraints
Private mMinWidth As Single
Private mMinHeight As Single
Private mMaxHeight As Single
Private mMaxWidth As Single

' Original form dimensions
Private mOriginalInsideWidth As Single
Private mOriginalInsideHeight As Single

' Control count at initialization
Private mControlCount As Integer

' Handle base name
Private mHandleBaseName As String

'==========================================================================================
' PUBLIC PROPERTIES
'==========================================================================================

Public Property Get Version() As String
    Version = "Class: " & CLASS_NAME & vbNewLine & _
              "Author: VBATools" & vbNewLine & _
              "Version: " & CLASS_VERSION & vbNewLine & _
              "Creation Date: 03.05.2019 09:31" & vbNewLine & _
              "Last Update: 12.11.2025 23:57" & vbNewLine & _
              "License: Apache License" & vbNewLine & _
              "Description: Anchor-based control resizing for UserForms"
End Property

Public Property Get CountAnchors() As Byte
    If Not mItems Is Nothing Then
        CountAnchors = mItems.Count
    End If
End Property

'==========================================================================================
' PUBLIC METHODS
'==========================================================================================

Public Sub Initialize(ByVal objParent As Object, _
                       Optional ByVal MaximumHeight As Single, _
                       Optional ByVal MaximumWidth As Single)

    If mIsInitialized Then
        Err.Raise vbObjectError + 1, CLASS_NAME, "The class has already been initialized"
        Exit Sub
    End If

    If objParent Is Nothing Then
        Err.Raise vbObjectError + 3, CLASS_NAME, "Parent object cannot be Nothing"
        Exit Sub
    End If

    Call CreateResizeHandles(objParent, MaximumHeight, MaximumWidth)
    Call InitializeControlCollection
    Call SetDefaultHandleAnchors

    mIsInitialized = True

End Sub

Public Sub SetAnchorStyleByName(ByVal ControlName As String, _
                                 ByVal anchorStyle As AnchorEdge)

    If mItems Is Nothing Then
        Err.Raise vbObjectError + 2, CLASS_NAME, "Item collection not initialized"
        Exit Sub
    End If

    Dim controlParams As Variant
    Dim i As Integer

    On Error Resume Next
    controlParams = mItems(ControlName)

    If Err.Number <> 0 Then
        Err.Clear
        Err.Raise vbObjectError + 4, CLASS_NAME, "Control '" & ControlName & "' not found"
        Exit Sub
    End If
    On Error GoTo 0

    controlParams(ControlParamIndex.idxAnchorStyle) = anchorStyle

    mItems.Remove ControlName
    mItems.Add controlParams, ControlName

End Sub

'==========================================================================================
' PRIVATE METHODS - INITIALIZATION
'==========================================================================================

Private Sub CreateResizeHandles(ByVal objParent As Object, _
                                 ByVal MaximumHeight As Single, _
                                 ByVal MaximumWidth As Single)

    mHandleBaseName = objParent.Name & RESIZE_HANDLE_SUFFIX

    ' Create corner handle (diagonal resize)
    Call CreateCornerHandle(objParent)

    ' Create bottom edge handle (vertical resize)
    Call CreateBottomHandle(objParent)

    ' Create right edge handle (horizontal resize)
    Call CreateRightHandle(objParent)

    ' Store parent reference and initialize dimensions
    Set mParent = objParent

    With objParent
        mMinWidth = .Width
        mMinHeight = .Height
        mOriginalInsideWidth = .InsideWidth
        mOriginalInsideHeight = .InsideHeight
    End With

    ' Set maximum dimensions (default to minimum if not specified)
    mMaxHeight = IIf(MaximumHeight > 0, MaximumHeight, mMinHeight)
    mMaxWidth = IIf(MaximumWidth > 0, MaximumWidth, mMinWidth)

    mControlCount = mParent.Controls.Count
    Set mItems = New Collection
    mUpdateWhileDragging = True

End Sub

Private Sub CreateCornerHandle(ByVal objParent As Object)

    Set mResizeHandle = objParent.Controls.Add("Forms.Label.1", mHandleBaseName, True)

    With mResizeHandle
        With .Font
            .Name = "Marlett"
            .Charset = 2
            .Size = DEFAULT_HANDLE_FONT_SIZE
            .Bold = True
        End With

        .BackStyle = fmBackStyleTransparent
        .AutoSize = True
        .BorderStyle = fmBorderStyleNone
        .Caption = "o"
        .MousePointer = fmMousePointerSizeNWSE
        .ForeColor = HANDLE_COLOR
        .ZOrder
        .Top = objParent.InsideHeight - .Height
        .Left = objParent.InsideWidth - .Width
    End With

End Sub

Private Sub CreateBottomHandle(ByVal objParent As Object)

    Set mResizeHandleBottom = objParent.Controls.Add( _
        "Forms.Label.1", _
        mHandleBaseName & HANDLE_SUFFIX_BOTTOM, _
        True)

    With mResizeHandleBottom
        .MousePointer = fmMousePointerSizeNS
        .ZOrder
        .Left = 0
        .Height = DEFAULT_HANDLE_SIZE
        .Top = objParent.InsideHeight - .Height
        .Width = objParent.InsideWidth
    End With

End Sub

Private Sub CreateRightHandle(ByVal objParent As Object)

    Set mResizeHandleRight = objParent.Controls.Add( _
        "Forms.Label.1", _
        mHandleBaseName & HANDLE_SUFFIX_RIGHT, _
        True)

    With mResizeHandleRight
        .MousePointer = fmMousePointerSizeWE
        .ZOrder
        .Width = DEFAULT_HANDLE_SIZE
        .Left = objParent.InsideWidth - .Width
        .Height = objParent.InsideHeight
        .Top = 0
    End With

End Sub

Private Sub InitializeControlCollection()

    Dim controlParams() As Variant
    ReDim controlParams(ControlParamIndex.[_First] To ControlParamIndex.[_Last]) As Variant

    Dim i As Integer

    For i = 0 To mControlCount - 1
        With mParent.Controls(i)
            controlParams(ControlParamIndex.idxAnchorStyle) = anchorLeft Or anchorTop
            controlParams(ControlParamIndex.idxMinimumWidth) = .Width
            controlParams(ControlParamIndex.idxMinimumHeight) = .Height
            controlParams(ControlParamIndex.idxMinimumTop) = .Top
            controlParams(ControlParamIndex.idxMinimumLeft) = .Left
            controlParams(ControlParamIndex.idxOriginalLeft) = .Left
            controlParams(ControlParamIndex.idxOriginalTop) = .Top
            controlParams(ControlParamIndex.idxOriginalWidth) = .Width
            controlParams(ControlParamIndex.idxOriginalHeight) = .Height
            controlParams(ControlParamIndex.idxControlName) = .Name

            mItems.Add controlParams, .Name
        End With
    Next i

End Sub

Private Sub SetDefaultHandleAnchors()

    Call SetAnchorStyleByName( _
        mHandleBaseName & HANDLE_SUFFIX_BOTTOM, _
        anchorBottom Or anchorRight Or anchorLeft)

    Call SetAnchorStyleByName( _
        mHandleBaseName & HANDLE_SUFFIX_RIGHT, _
        anchorBottom Or anchorRight Or anchorTop)

End Sub

'==========================================================================================
' PRIVATE METHODS - MOUSE EVENT HANDLERS
'==========================================================================================

' --- MouseDown Events ---

Private Sub mResizeHandle_MouseDown(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)
    Call HandleMouseDown(Button, X, Y)
End Sub

Private Sub mResizeHandleBottom_MouseDown(ByVal Button As Integer, _
                                           ByVal Shift As Integer, _
                                           ByVal X As Single, _
                                           ByVal Y As Single)
    Call HandleMouseDown(Button, X, Y)
End Sub

Private Sub mResizeHandleRight_MouseDown(ByVal Button As Integer, _
                                          ByVal Shift As Integer, _
                                          ByVal X As Single, _
                                          ByVal Y As Single)
    Call HandleMouseDown(Button, X, Y)
End Sub

' --- MouseUp Events ---

Private Sub mResizeHandle_MouseUp(ByVal Button As Integer, _
                                   ByVal Shift As Integer, _
                                   ByVal X As Single, _
                                   ByVal Y As Single)
    Call HandleMouseUp(Button)
End Sub

Private Sub mResizeHandleBottom_MouseUp(ByVal Button As Integer, _
                                         ByVal Shift As Integer, _
                                         ByVal X As Single, _
                                         ByVal Y As Single)
    Call HandleMouseUp(Button)
End Sub

Private Sub mResizeHandleRight_MouseUp(ByVal Button As Integer, _
                                        ByVal Shift As Integer, _
                                        ByVal X As Single, _
                                        ByVal Y As Single)
    Call HandleMouseUp(Button)
End Sub

' --- MouseMove Events ---

Private Sub mResizeHandle_MouseMove(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single)

    If Button <> MOUSE_BUTTON_LEFT Or Not mIsResizing Then Exit Sub

    Dim newWidth As Single
    Dim newHeight As Single

    newWidth = CalculateNewDimension(mParent.Width, X - mResizeStartX, mMinWidth, mMaxWidth)
    newHeight = CalculateNewDimension(mParent.Height, Y - mResizeStartY, mMinHeight, mMaxHeight)

    mParent.Width = newWidth
    mParent.Height = newHeight

    Call UpdateHandlePositions
    Call UpdateControlsIfNeeded

End Sub

Private Sub mResizeHandleBottom_MouseMove(ByVal Button As Integer, _
                                           ByVal Shift As Integer, _
                                           ByVal X As Single, _
                                           ByVal Y As Single)

    If Button <> MOUSE_BUTTON_LEFT Or Not mIsResizing Then Exit Sub

    Dim newHeight As Single

    newHeight = CalculateNewDimension(mParent.Height, Y - mResizeStartY, mMinHeight, mMaxHeight)
    mParent.Height = newHeight

    Call UpdateHandlePositions
    Call UpdateControlsIfNeeded

End Sub

Private Sub mResizeHandleRight_MouseMove(ByVal Button As Integer, _
                                          ByVal Shift As Integer, _
                                          ByVal X As Single, _
                                          ByVal Y As Single)

    If Button <> MOUSE_BUTTON_LEFT Or Not mIsResizing Then Exit Sub

    Dim newWidth As Single

    newWidth = CalculateNewDimension(mParent.Width, X - mResizeStartX, mMinWidth, mMaxWidth)
    mParent.Width = newWidth

    Call UpdateHandlePositions
    Call UpdateControlsIfNeeded

End Sub

'==========================================================================================
' PRIVATE METHODS - MOUSE HANDLER HELPERS
'==========================================================================================

Private Sub HandleMouseDown(ByVal Button As Integer, _
                             ByVal X As Single, _
                             ByVal Y As Single)

    If Button = MOUSE_BUTTON_LEFT Then
        mResizeStartX = X
        mResizeStartY = Y
        mIsResizing = True
    End If

End Sub

Private Sub HandleMouseUp(ByVal Button As Integer)

    If Button = MOUSE_BUTTON_LEFT Then
        If Not mUpdateWhileDragging Then
            Call UpdateControlPositions
        End If
        mIsResizing = False
    End If

End Sub

Private Function CalculateNewDimension(ByVal currentValue As Single, _
                                         ByVal delta As Single, _
                                         ByVal minValue As Single, _
                                         ByVal maxValue As Single) As Single

    Dim result As Single
    result = currentValue + delta

    ' Clamp to min/max bounds
    If result < minValue Then
        result = minValue
    ElseIf result > maxValue Then
        result = maxValue
    End If

    CalculateNewDimension = result

End Function

Private Sub UpdateHandlePositions()

    With mResizeHandle
        .Left = mParent.InsideWidth - .Width
        .Top = mParent.InsideHeight - .Height
    End With

    With mResizeHandleBottom
        .Top = mParent.InsideHeight - .Height
        .Width = mParent.InsideWidth
    End With

    With mResizeHandleRight
        .Left = mParent.InsideWidth - .Width
        .Height = mParent.InsideHeight
    End With

End Sub

Private Sub UpdateControlsIfNeeded()
    If mUpdateWhileDragging Then Call UpdateControlPositions
End Sub

'==========================================================================================
' PRIVATE METHODS - CONTROL POSITIONING
'==========================================================================================

Private Sub UpdateControlPositions()

    Dim i As Integer
    Dim controlParams As Variant
    Dim anchorStyle As AnchorEdge

    For i = 0 To mControlCount - 1
        controlParams = mItems(mParent.Controls(i).Name)
        anchorStyle = controlParams(ControlParamIndex.idxAnchorStyle)

        If anchorStyle <> anchorNone Then
            Call ApplyAnchorStyle(i, controlParams, anchorStyle)
        End If
    Next i

    DoEvents

End Sub

Private Sub ApplyAnchorStyle(ByVal controlIndex As Integer, _
                              ByVal params As Variant, _
                              ByVal anchorStyle As AnchorEdge)

    Dim hasTop As Boolean
    Dim hasBottom As Boolean
    Dim hasLeft As Boolean
    Dim hasRight As Boolean

    hasTop = (anchorStyle And anchorTop) = anchorTop
    hasBottom = (anchorStyle And anchorBottom) = anchorBottom
    hasLeft = (anchorStyle And anchorLeft) = anchorLeft
    hasRight = (anchorStyle And anchorRight) = anchorRight

    ' Apply vertical anchoring
    Call ApplyVerticalAnchor(controlIndex, params, hasTop, hasBottom)

    ' Apply horizontal anchoring
    Call ApplyHorizontalAnchor(controlIndex, params, hasLeft, hasRight)

End Sub

Private Sub ApplyVerticalAnchor(ByVal controlIndex As Integer, _
                                 ByVal params As Variant, _
                                 ByVal hasTop As Boolean, _
                                 ByVal hasBottom As Boolean)

    Dim newHeight As Single
    Dim newTop As Single

    With mParent
        If hasTop And hasBottom Then
            ' Stretch vertically - adjust height
            newHeight = .InsideHeight - (mOriginalInsideHeight - params(ControlParamIndex.idxOriginalTop) - params(ControlParamIndex.idxOriginalHeight)) - params(ControlParamIndex.idxOriginalTop)
            newHeight = ClampValue(newHeight, params(ControlParamIndex.idxMinimumHeight))
            .Controls(controlIndex).Height = newHeight

        ElseIf hasBottom Then
            ' Anchor to bottom - adjust top position
            newTop = .InsideHeight - (mOriginalInsideHeight - params(ControlParamIndex.idxOriginalTop) - params(ControlParamIndex.idxOriginalHeight)) - params(ControlParamIndex.idxOriginalHeight)
            newTop = ClampValue(newTop, params(ControlParamIndex.idxMinimumTop))
            .Controls(controlIndex).Top = newTop

        ' ElseIf hasTop Only - no change needed
        End If
    End With

End Sub

Private Sub ApplyHorizontalAnchor(ByVal controlIndex As Integer, _
                                   ByVal params As Variant, _
                                   ByVal hasLeft As Boolean, _
                                   ByVal hasRight As Boolean)

    Dim newWidth As Single
    Dim newLeft As Single

    With mParent
        If hasLeft And hasRight Then
            ' Stretch horizontally - adjust width
            newWidth = .InsideWidth - (mOriginalInsideWidth - params(ControlParamIndex.idxOriginalLeft) - params(ControlParamIndex.idxOriginalWidth)) - params(ControlParamIndex.idxOriginalLeft)
            newWidth = ClampValue(newWidth, params(ControlParamIndex.idxMinimumWidth))
            .Controls(controlIndex).Width = newWidth

        ElseIf hasRight Then
            ' Anchor to right - adjust left position
            newLeft = .InsideWidth - (mOriginalInsideWidth - params(ControlParamIndex.idxOriginalLeft) - params(ControlParamIndex.idxOriginalWidth)) - params(ControlParamIndex.idxOriginalWidth)
            newLeft = ClampValue(newLeft, params(ControlParamIndex.idxMinimumLeft))
            .Controls(controlIndex).Left = newLeft

        ' ElseIf hasLeft Only - no change needed
        End If
    End With

End Sub

Private Function ClampValue(ByVal value As Single, ByVal minValue As Single) As Single

    If value < minValue Then
        value = minValue
    End If

    If value < 0 Then
        value = 0
    End If

    ClampValue = value

End Function