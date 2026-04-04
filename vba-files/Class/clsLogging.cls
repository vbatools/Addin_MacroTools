VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Class        :   clsLogging - High-performance CSV logger with Unicode support
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Version      :   2.0
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'---------------------------------------------------------------------------------------------------
' ENUMERATIONS
'---------------------------------------------------------------------------------------------------
Public Enum LOG_LEVEL
    LOG_LEVEL_INFO = 0
    LOG_LEVEL_WARNING = 1
    LOG_LEVEL_ERROR = 2
End Enum

'---------------------------------------------------------------------------------------------------
' CONSTANTS
'---------------------------------------------------------------------------------------------------
Private Const CLASS_NAME As String = "clsLogging"

' Log file configuration
Private Const DEFAULT_LOG_EXTENSION As String = ".csv"
Private Const DEFAULT_LOG_FILENAME As String = NAME_ADDIN & "_logs"

' CSV configuration
Private Const CSV_DELIMITER As String = ";"
Private Const CSV_HEADER As String = "Timestamp" & CSV_DELIMITER & "Level" & CSV_DELIMITER & "Event" & CSV_DELIMITER & "Info"
Private Const QUOTE_CHAR As String = """"

' FSO constants
Private Const FSO_FOR_APPENDING As Integer = 8
Private Const FSO_TRISTATE_TRUE As Long = -1

' Buffer defaults
Private Const DEFAULT_BUFFER_SIZE As Long = 5000
Private Const MAX_MESSAGE_LENGTH As Long = 1000

'---------------------------------------------------------------------------------------------------
' MEMBER VARIABLES
'---------------------------------------------------------------------------------------------------
Private m_FSO As Object
Private m_WorkbookPath As String
Private m_LogFilePath As String
Private m_Buffer As Collection
Private m_MaxBufferSize As Long

'===================================================================================================
' PROPERTIES
'===================================================================================================

' Returns the full path to the log file
Public Property Get FilePathLog() As String
    FilePathLog = m_LogFilePath
End Property

' Sets the log file name (without extension)
Public Property Let FileNameLog(ByVal FileName As String)
    Dim sName As String

    sName = Trim$(FileName)
    If Len(sName) = 0 Then
        sName = DEFAULT_LOG_FILENAME
    End If

    m_LogFilePath = m_WorkbookPath & sName & DEFAULT_LOG_EXTENSION
End Property

' Sets the buffer size before forced write
Public Property Let MaxBufferSize(ByVal Size As Long)

    Select Case Size
             Case 0: m_MaxBufferSize = DEFAULT_BUFFER_SIZE
        Case Is > 0:  m_MaxBufferSize = Size
        Case Else: Err.Raise 5, CLASS_NAME & ".MaxBufferSize", "Buffer size must be greater than zero."
    End Select

End Property

'===================================================================================================
' INITIALIZATION / TERMINATION
'===================================================================================================

Private Sub Class_Initialize()
    On Error GoTo ErrorHandler

    Set m_FSO = CreateObject("Scripting.FileSystemObject")
    Set m_Buffer = New Collection

    m_WorkbookPath = ThisWorkbook.Path & Application.PathSeparator


    Me.FileNameLog = DEFAULT_LOG_FILENAME
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, CLASS_NAME & ".Class_Initialize", _
              "Failed to initialize logger: " & Err.Description
End Sub

Private Sub Class_Terminate()
    ' Force save buffer when object is destroyed
    If Not m_Buffer Is Nothing Then
        If m_Buffer.Count > 0 Then Call FlushBuffer
        Set m_Buffer = Nothing
    End If
    Set m_FSO = Nothing
End Sub

'===================================================================================================
' PUBLIC API
'===================================================================================================

' Logs an informational message
Public Sub LogInfo(ByVal EventX As String, Optional ByVal Info As String)
    Call AddRecord(EventX, Info, LOG_LEVEL_INFO)
End Sub

' Logs a warning message
Public Sub LogWarning(ByVal EventX As String, Optional ByVal Info As String)
    Call AddRecord(EventX, Info, LOG_LEVEL_WARNING)
End Sub

' Logs an error with information from the global Err object
Public Sub LogError(Optional ByVal FunctionName As String)
    Dim sContext As String

    If Len(Trim$(FunctionName)) = 0 Then FunctionName = "Unknown Procedure"

    sContext = "Err #" & Err.Number & _
               " | Line: " & Erl & _
               " | Description: " & Err.Description

    Call AddRecord(FunctionName, sContext, LOG_LEVEL_ERROR, True)
End Sub

'===================================================================================================
' CORE LOGIC
'===================================================================================================
' Main method to add a log entry
Public Sub AddRecord( _
    ByVal EventX As String, _
    Optional ByVal Info As String, _
    Optional ByVal Level As LOG_LEVEL = LOG_LEVEL_INFO, _
    Optional ByVal ForceSave As Boolean = False)

    ' Validate input data
    Call ValidateInput(EventX, "EventX")
    If IsMissing(Info) Then Info = vbNullString

    Dim sLevelTag As String
    Dim csvLine As String

    ' Determine text level tag
    sLevelTag = Choose(Level + 1, "INFO", "WARNING", "ERROR")
    If Len(sLevelTag) = 0 Then sLevelTag = "INFO"

    ' Build CSV line string
    csvLine = EscapeCSV(Format$(Now, "yyyy-mm-dd hh:nn:ss")) & CSV_DELIMITER & _
              EscapeCSV(sLevelTag) & CSV_DELIMITER & _
              EscapeCSV(EventX) & CSV_DELIMITER & _
              EscapeCSV(Info)

    ' Add to buffer (Collection instead of concatenation)
    m_Buffer.Add csvLine

    ' Check conditions for disk write
    If ForceSave Or (m_Buffer.Count >= m_MaxBufferSize) Then
        Call FlushBuffer
    End If
End Sub

' Force write buffer to file
Public Sub FlushBuffer()
    If m_Buffer.Count = 0 Then Exit Sub

    Dim sContent As String
    Dim vItem As Variant

    ' Assemble buffer into a single string
    For Each vItem In m_Buffer
        sContent = sContent & vItem & vbNewLine
    Next vItem

    ' Write and clear
    If AppendToFile(sContent) Then
        Call ClearBuffer
    Else
        Call MsgBox("CRITICAL: Failed to write log buffer to " & m_LogFilePath, vbCritical)
    End If
End Sub

' Open log in external application
Public Sub ShowLog()
    If Not m_FSO.FileExists(m_LogFilePath) Then
        Call MsgBox("Log file does not exist", vbInformation, NAME_ADDIN)
        Exit Sub
    End If

    On Error Resume Next
    Call CreateObject("WScript.Shell").Run("""" & m_LogFilePath & """")
    On Error GoTo 0
End Sub

' Delete the log file
Public Sub ResetLogs()
    If Not m_FSO.FileExists(m_LogFilePath) Then
        Call MsgBox("Log file not found.", vbInformation, NAME_ADDIN)
        Exit Sub
    End If

    On Error Resume Next
    Call m_FSO.DeleteFile(m_LogFilePath, True)

    If Err.Number = 0 Then
        Call ClearBuffer
        Call MsgBox("Log successfully cleared.", vbInformation, NAME_ADDIN)
    Else
        Call MsgBox("Error deleting log:" & Err.Description, vbExclamation, NAME_ADDIN)
    End If
    On Error GoTo 0
End Sub

'===================================================================================================
' PRIVATE HELPERS
'===================================================================================================

' Clear the buffer
Private Sub ClearBuffer()
    Do While m_Buffer.Count > 0
        m_Buffer.Remove 1
    Loop
End Sub

' Escape special characters for CSV
Private Function EscapeCSV(ByVal txt As String) As String
    Dim needsQuotes As Boolean

    ' Check if quotes are needed
    needsQuotes = InStr(txt, CSV_DELIMITER) > 0 Or _
                  InStr(txt, vbCr) > 0 Or _
                  InStr(txt, vbLf) > 0 Or _
                  InStr(txt, QUOTE_CHAR) > 0

    ' Double the quotes (CSV standard RFC 4180)
    txt = Replace$(txt, QUOTE_CHAR, QUOTE_CHAR & QUOTE_CHAR)

    If needsQuotes Then
        EscapeCSV = QUOTE_CHAR & txt & QUOTE_CHAR
    Else
        EscapeCSV = txt
    End If
End Function

' Append text to file
Private Function AppendToFile(ByVal text As String, Optional ByVal FilePath As String) As Boolean
    Dim ts As Object
    Dim targetPath As String
    Dim bIsNewFile As Boolean

    On Error GoTo ErrorHandler

    targetPath = IIf(Len(FilePath) = 0, m_LogFilePath, FilePath)
    bIsNewFile = Not m_FSO.FileExists(targetPath)

    Set ts = m_FSO.OpenTextFile(targetPath, FSO_FOR_APPENDING, True, FSO_TRISTATE_TRUE)

    With ts
        If bIsNewFile Then Call .Write(CSV_HEADER & vbNewLine)
        Call .Write(text)
        .Close
    End With

    AppendToFile = True
    Set ts = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "Error writing to log file [" & targetPath & "]: " & Err.Description
    AppendToFile = False
    Set ts = Nothing
End Function

' Validate input parameters
Private Sub ValidateInput(ByVal value As String, ByVal ParamName As String)
    If Len(Trim$(value)) = 0 Then
        Err.Raise 5, CLASS_NAME & "." & ParamName, "Parameter cannot be empty or whitespace."
    End If

    If Len(value) > MAX_MESSAGE_LENGTH Then
        Err.Raise 5, CLASS_NAME & "." & ParamName, _
              "Parameter exceeds maximum length of " & MAX_MESSAGE_LENGTH & " characters."
    End If
End Sub