VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSort2DArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : clsSort2DArray
'* Purpose    : Class for sorting VBA 2D arrays
'* Algorithms : QuickSort (unstable, in-place) and MergeSort (stable)
'*
'* Version    : 3.0 (Refactored)
'* License    : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'==========================================================================================
' ENUMERATIONS
'==========================================================================================

Public Enum SortDataType
    sdtString = 1
    sdtNumber = 2
End Enum

'==========================================================================================
' PRIVATE MEMBERS
'==========================================================================================

Private mCol        As Long
Private mSortUp     As Boolean
Private mArr        As Variant
Private mDataType   As SortDataType

'==========================================================================================
' ERROR CONSTANTS
'==========================================================================================

Private Const ERR_EMPTY_ARRAY      As Long = vbObjectError + 513
Private Const ERR_NOT_2D_ARRAY     As Long = vbObjectError + 514
Private Const ERR_INVALID_COLUMN   As Long = vbObjectError + 515
Private Const ERR_SINGLE_ROW       As Long = vbObjectError + 516

Private Const ERR_MSG_EMPTY        As String = "Array is empty or not initialized"
Private Const ERR_MSG_NOT_2D       As String = "Two-dimensional array required"
Private Const ERR_MSG_INVALID_COL  As String = "Column index out of range"

'==========================================================================================
' PROPERTIES
'==========================================================================================

Public Property Get ColumnSort() As Long
    ColumnSort = mCol
End Property

Public Property Get SortUp() As Boolean
    SortUp = mSortUp
End Property

Public Property Get ListArray() As Variant
    ListArray = mArr
End Property

Public Property Get CountColumns() As Long
    If IsArrayAllocated(mArr) Then
        CountColumns = UBound(mArr, 2) - LBound(mArr, 2) + 1
    End If
End Property

Public Property Get CountRows() As Long
    If IsArrayAllocated(mArr) Then
        CountRows = UBound(mArr, 1) - LBound(mArr, 1) + 1
    End If
End Property

Public Property Get DataType() As SortDataType
    DataType = mDataType
End Property

'==========================================================================================
' PUBLIC METHODS
'==========================================================================================

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Method: SortQuick
'* Purpose: Quick sort for array (QuickSort algorithm)
'*
'* Parameters:
'*   vArr       - 2D array to sort
'*   iCol       - Sort column index (default - first column)
'*   bSortUp    - True = ascending (default), False = descending
'*   eDataType  - Data type: sdtString or sdtNumber
'*
'* Remarks:
'*   - In-place sorting (no additional memory)
'*   - Unstable algorithm (may reorder equal elements)
'*   - O(N log N) average complexity, O(N^2) worst case
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Public Sub SortQuick(ByVal vArr As Variant, _
                     Optional ByVal iCol As Long = -1, _
                     Optional ByVal bSortUp As Boolean = True, _
                     Optional ByVal eDataType As SortDataType = sdtString)

    ' Validation and initialization
    Call ValidateArray(vArr, iCol)

    mArr = vArr
    mCol = iCol
    mSortUp = bSortUp
    mDataType = eDataType

    ' Initialize random number generator for median of three
    Call Randomize

    ' Check for trivial case (1 row)
    If UBound(mArr, 1) - LBound(mArr, 1) < 1 Then Exit Sub

    Call QuickSortInternal(mArr, LBound(mArr, 1), UBound(mArr, 1), mCol, mSortUp, mDataType)

End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Method: SortMerge
'* Purpose: Stable array sorting (MergeSort algorithm)
'*
'* Parameters:
'*   vArr       - 2D array to sort
'*   iCol       - Sort column index (default - first column)
'*   bSortUp    - True = ascending (default), False = descending
'*   eDataType  - Data type: sdtString or sdtNumber
'*
'* Remarks:
'*   - Stable algorithm (preserves order of equal elements)
'*   - Requires O(N) additional memory
'*   - O(N log N) guaranteed complexity
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Public Sub SortMerge(ByVal vArr As Variant, _
                     Optional ByVal iCol As Long = -1, _
                     Optional ByVal bSortUp As Boolean = True, _
                     Optional ByVal eDataType As SortDataType = sdtString)

    ' Validation and initialization
    Call ValidateArray(vArr, iCol)

    mArr = vArr
    mCol = iCol
    mSortUp = bSortUp
    mDataType = eDataType

    ' Check for trivial case (1 row)
    If UBound(mArr, 1) - LBound(mArr, 1) < 1 Then Exit Sub

    Call MergeSortInternal(mArr, LBound(mArr, 1), UBound(mArr, 1), mCol, mSortUp, mDataType)

End Sub

'==========================================================================================
' PRIVATE METHODS - VALIDATION
'==========================================================================================

Private Sub ValidateArray(ByRef vArr As Variant, ByRef iCol As Long)
    ' Check that a variant is an array
    If Not IsArray(vArr) Then
        Err.Raise ERR_NOT_2D_ARRAY, "clsSort2DArray.ValidateArray", ERR_MSG_NOT_2D
    End If

    ' Check for empty/uninitialized array
    If Not IsArrayAllocated(vArr) Then
        Err.Raise ERR_EMPTY_ARRAY, "clsSort2DArray.ValidateArray", ERR_MSG_EMPTY
    End If

    ' Check dimensionality (must be 2D)
    If ArrayDimensions(vArr) <> 2 Then
        Err.Raise ERR_NOT_2D_ARRAY, "clsSort2DArray.ValidateArray", ERR_MSG_NOT_2D
    End If

    ' Determine sort column
    If iCol < 0 Then
        iCol = LBound(vArr, 2)
    End If

    ' Validate column index
    If iCol < LBound(vArr, 2) Or iCol > UBound(vArr, 2) Then
        Err.Raise ERR_INVALID_COLUMN, "clsSort2DArray.ValidateArray", ERR_MSG_INVALID_COL
    End If

End Sub

Private Function IsArrayAllocated(ByRef vArr As Variant) As Boolean
    On Error Resume Next
    Dim i As Long
    i = UBound(vArr, 1)
    IsArrayAllocated = Err.Number = 0
End Function

Private Function ArrayDimensions(ByRef vArr As Variant) As Integer
    On Error Resume Next

    Dim i As Integer
    i = 1
    Do
        i = i + 1
    Loop While UBound(vArr, i)

    ArrayDimensions = i - 1
    Err.Clear
    On Error GoTo 0
End Function

'==========================================================================================
' PRIVATE METHODS - QUICKSORT
'==========================================================================================

Private Sub QuickSortInternal(ByRef pvarArray As Variant, _
                              ByVal lngLeft As Long, _
                              ByVal lngRight As Long, _
                              ByVal iCol As Long, _
                              ByVal bSortAsc As Boolean, _
                              ByVal eDataType As SortDataType)

    Dim lngFirst        As Long
    Dim lngLast         As Long
    Dim varPivot        As Variant
    Dim varSwap         As Variant
    Dim j               As Long
    Dim lngPivotIndex   As Long

    lngFirst = lngLeft
    lngLast = lngRight

    ' Choose pivot element: median of three
    lngPivotIndex = GetMedianOfThree(pvarArray, lngLeft, lngRight, iCol, eDataType)
    varPivot = GetComparableValue(pvarArray(lngPivotIndex, iCol), eDataType)

    Do
        ' Find element from left that should be on right side of pivot
        Do While lngFirst < lngRight
            If ShouldSwap(varPivot, pvarArray(lngFirst, iCol), eDataType, bSortAsc) Then
                Exit Do
            End If
            lngFirst = lngFirst + 1
        Loop

        ' Find element from right that should be on left side of pivot
        Do While lngLast > lngLeft
            If ShouldSwap(pvarArray(lngLast, iCol), varPivot, eDataType, bSortAsc) Then
                Exit Do
            End If
            lngLast = lngLast - 1
        Loop

        ' Swap rows
        If lngFirst <= lngLast Then
            Call SwapRows(pvarArray, lngFirst, lngLast)
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If

    Loop Until lngFirst > lngLast

    ' Recursive call (tail recursion optimization)
    If (lngLast - lngLeft) < (lngRight - lngFirst) Then
        If lngLeft < lngLast Then Call QuickSortInternal(pvarArray, lngLeft, lngLast, iCol, bSortAsc, eDataType)
        If lngFirst < lngRight Then Call QuickSortInternal(pvarArray, lngFirst, lngRight, iCol, bSortAsc, eDataType)
    Else
        If lngFirst < lngRight Then Call QuickSortInternal(pvarArray, lngFirst, lngRight, iCol, bSortAsc, eDataType)
        If lngLeft < lngLast Then Call QuickSortInternal(pvarArray, lngLeft, lngLast, iCol, bSortAsc, eDataType)
    End If

End Sub

Private Function GetMedianOfThree(ByRef pvarArray As Variant, _
                                   ByVal lngLeft As Long, _
                                   ByVal lngRight As Long, _
                                   ByVal iCol As Long, _
                                   ByVal eDataType As SortDataType) As Long

    Dim lngCount    As Long
    Dim a           As Long
    Dim b           As Long
    Dim c           As Long
    Dim vA          As Variant
    Dim vB          As Variant
    Dim vC          As Variant

    lngCount = lngRight - lngLeft + 1

    ' Three random indices
    a = Int(lngCount * Rnd) + lngLeft
    b = Int(lngCount * Rnd) + lngLeft
    c = Int(lngCount * Rnd) + lngLeft

    ' Get comparable values
    vA = GetComparableValue(pvarArray(a, iCol), eDataType)
    vB = GetComparableValue(pvarArray(b, iCol), eDataType)
    vC = GetComparableValue(pvarArray(c, iCol), eDataType)

    ' Determine median
    If (vA <= vB And vB <= vC) Or (vC <= vB And vB <= vA) Then
        GetMedianOfThree = b
    ElseIf (vB <= vA And vA <= vC) Or (vC <= vA And vA <= vB) Then
        GetMedianOfThree = a
    Else
        GetMedianOfThree = c
    End If

End Function

'==========================================================================================
' PRIVATE METHODS - MERGESORT
'==========================================================================================

Private Sub MergeSortInternal(ByRef pvarArray As Variant, _
                              ByVal lngLeft As Long, _
                              ByVal lngRight As Long, _
                              ByVal iCol As Long, _
                              ByVal bSortAsc As Boolean, _
                              ByVal eDataType As SortDataType)

    Dim lngMid As Long

    If lngLeft < lngRight Then
        lngMid = (lngLeft + lngRight) \ 2

        Call MergeSortInternal(pvarArray, lngLeft, lngMid, iCol, bSortAsc, eDataType)
        Call MergeSortInternal(pvarArray, lngMid + 1, lngRight, iCol, bSortAsc, eDataType)

        Call MergeArrays(pvarArray, lngLeft, lngMid, lngRight, iCol, bSortAsc, eDataType)
    End If

End Sub

Private Sub MergeArrays(ByRef pvarArray As Variant, _
                        ByVal lngLeft As Long, _
                        ByVal lngMid As Long, _
                        ByVal lngRight As Long, _
                        ByVal iCol As Long, _
                        ByVal bSortAsc As Boolean, _
                        ByVal eDataType As SortDataType)

    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim n1          As Long
    Dim n2          As Long
    Dim lCol        As Long
    Dim uCol        As Long
    Dim c           As Long
    Dim LeftArr()   As Variant
    Dim RightArr()  As Variant

    n1 = lngMid - lngLeft + 1
    n2 = lngRight - lngMid

    lCol = LBound(pvarArray, 2)
    uCol = UBound(pvarArray, 2)

    ' Create temporary arrays
    ReDim LeftArr(1 To n1, lCol To uCol)
    ReDim RightArr(1 To n2, lCol To uCol)

    ' Copy data to temporary arrays
    For i = 1 To n1
        For c = lCol To uCol
            LeftArr(i, c) = pvarArray(lngLeft + i - 1, c)
        Next c
    Next i

    For j = 1 To n2
        For c = lCol To uCol
            RightArr(j, c) = pvarArray(lngMid + j, c)
        Next c
    Next j

    ' Merge
    i = 1
    j = 1
    k = lngLeft

    Do While i <= n1 And j <= n2
        If ShouldTakeLeft(LeftArr(i, iCol), RightArr(j, iCol), eDataType, bSortAsc) Then
            Call CopyRowToDest(LeftArr, i, pvarArray, k, lCol, uCol)
            i = i + 1
        Else
            Call CopyRowToDest(RightArr, j, pvarArray, k, lCol, uCol)
            j = j + 1
        End If
        k = k + 1
    Loop

    ' Copy remaining elements
    Do While i <= n1
        Call CopyRowToDest(LeftArr, i, pvarArray, k, lCol, uCol)
        i = i + 1
        k = k + 1
    Loop

    Do While j <= n2
        Call CopyRowToDest(RightArr, j, pvarArray, k, lCol, uCol)
        j = j + 1
        k = k + 1
    Loop

End Sub

'==========================================================================================
' PRIVATE METHODS - HELPERS
'==========================================================================================

Private Function GetComparableValue(ByVal vVal As Variant, ByVal eType As SortDataType) As Variant
    On Error GoTo ErrHandler

    If eType = sdtNumber Then
        If IsNumeric(vVal) Then
            GetComparableValue = CDbl(vVal)
        Else
            GetComparableValue = 0
        End If
    Else
        GetComparableValue = CStr(vVal)
    End If
    Exit Function

ErrHandler:
    GetComparableValue = IIf(eType = sdtNumber, 0, vbNullString)
End Function

Private Function ShouldSwap(ByVal vLeft As Variant, _
                            ByVal vRight As Variant, _
                            ByVal eDataType As SortDataType, _
                            ByVal bSortAsc As Boolean) As Boolean
    ' Returns True if elements need to be swapped
    Dim vLeftVal    As Variant
    Dim vRightVal   As Variant

    vLeftVal = GetComparableValue(vLeft, eDataType)
    vRightVal = GetComparableValue(vRight, eDataType)

    If bSortAsc Then
        ShouldSwap = vLeftVal < vRightVal
    Else
        ShouldSwap = vLeftVal > vRightVal
    End If
End Function

Private Function ShouldTakeLeft(ByVal vLeft As Variant, _
                                ByVal vRight As Variant, _
                                ByVal eDataType As SortDataType, _
                                ByVal bSortAsc As Boolean) As Boolean
    ' For MergeSort: returns True if element should be taken from left array
    Dim vLeftVal    As Variant
    Dim vRightVal   As Variant

    vLeftVal = GetComparableValue(vLeft, eDataType)
    vRightVal = GetComparableValue(vRight, eDataType)

    If bSortAsc Then
        ShouldTakeLeft = vLeftVal <= vRightVal
    Else
        ShouldTakeLeft = vLeftVal >= vRightVal
    End If
End Function

Private Sub SwapRows(ByRef pvarArray As Variant, _
                     ByVal lngRow1 As Long, _
                     ByVal lngRow2 As Long)
    Dim c           As Long
    Dim varSwap     As Variant
    Dim lCol        As Long
    Dim uCol        As Long

    lCol = LBound(pvarArray, 2)
    uCol = UBound(pvarArray, 2)

    For c = lCol To uCol
        varSwap = pvarArray(lngRow1, c)
        pvarArray(lngRow1, c) = pvarArray(lngRow2, c)
        pvarArray(lngRow2, c) = varSwap
    Next c
End Sub

Private Sub CopyRowToDest(ByRef srcArr As Variant, _
                          ByVal srcRow As Long, _
                          ByRef destArr As Variant, _
                          ByVal destRow As Long, _
                          ByVal lCol As Long, _
                          ByVal uCol As Long)
    Dim c As Long

    For c = lCol To uCol
        destArr(destRow, c) = srcArr(srcRow, c)
    Next c
End Sub
