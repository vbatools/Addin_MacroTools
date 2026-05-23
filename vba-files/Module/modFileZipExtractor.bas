Attribute VB_Name = "modFileZipExtractor"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module       :   modFileZipExtractor - Extracting attached files to excel
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   14-04-2026 13:23:37
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module

' ============================================
' КОНСТАНТЫ СИГНАТУР ФАЙЛОВ
' ============================================
Private Const SIG_PK_ZIP As String = "504B0304"
Private Const SIG_RAR As String = "52617221"
Private Const SIG_7Z As String = "377ABCAF"
Private Const SIG_PDF As String = "25504446"
Private Const SIG_JPEG As String = "FFD8FF"
Private Const SIG_PNG As String = "89504E47"
Private Const SIG_GIF As String = "47494638"
Private Const SIG_OLE2 As String = "D0CF11E0"

' ============================================
' КОНСТАНТЫ РАСШИРЕНИЙ ФАЙЛОВ
' ============================================
Private Const EXT_ZIP As String = ".zip"
Private Const EXT_RAR As String = ".rar"
Private Const EXT_7Z As String = ".7z"
Private Const EXT_PDF As String = ".pdf"
Private Const EXT_JPEG As String = ".jpg"
Private Const EXT_PNG As String = ".png"
Private Const EXT_GIF As String = ".gif"
Private Const EXT_XLSX As String = ".xlsx"
Private Const EXT_DOCX As String = ".docx"
Private Const EXT_PPTX As String = ".pptx"
Private Const EXT_XLS As String = ".xls"
Private Const EXT_DOC As String = ".doc"
Private Const EXT_PPT As String = ".ppt"
Private Const EXT_OLE As String = ".ole"
Private Const EXT_BIN As String = ".bin"

' ============================================
' КОНСТАНТЫ ПОИСКА
' ============================================
Private Const MAX_OLE_HEADER_SIZE As Long = 5000
Private Const MAX_ZIP_TYPE_SEARCH As Long = 2000
Private Const MAX_OLE_CONTENT_SEARCH As Long = 100000

' ============================================
' КОНСТАНТЫ ДЛЯ АРХИВАТОРОВ
' ============================================
Private Const PATH_7ZIP As String = "C:\Program Files\7-Zip\7z.exe"
Private Const PATH_7ZIP_X86 As String = "C:\Program Files (x86)\7-Zip\7z.exe"
Private Const PATH_WINRAR As String = "C:\Program Files\WinRAR\WinRAR.exe"
Private Const PATH_WINRAR_X86 As String = "C:\Program Files (x86)\WinRAR\WinRAR.exe"

' ============================================
' ГЛАВНАЯ ПРОЦЕДУРА ИЗВЛЕЧЕНИЯ ФАЙЛОВ
' ============================================
Public Sub fileExtractorFromExcelFile()
    Dim arrFiles()  As String
    Dim clsZIP      As clsOfficeArchiveManager
    Dim arrFilesTable As Variant
    Dim sPathExtractFiles As String
    Dim sFullNameFile As String
    Dim lI          As Long
    Dim oFSO        As Object
    Dim bSuccess    As Boolean

    On Error GoTo ErrorHandler

    ' Получаем список файлов через диалог
    arrFiles = fileDialogFun(ActiveWorkbook.Path, False, TYPE_FILES)
    If (Not (Not (arrFiles))) = 0 Then Exit Sub

    ' Инициализируем архивный менеджер
    Set clsZIP = New clsOfficeArchiveManager

    With clsZIP
        If Not .Initialize(arrFiles(1, 1), True) Then GoTo CleanUp
        If Not .UnZipFile Then GoTo CleanUp

        arrFilesTable = GetFilesTable(.GetSettings(ExlFolderEmbeddings))

        If IsEmpty(arrFilesTable) Then GoTo CleanUp

        ' Создаём папку для извлечённых файлов
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        sPathExtractFiles = BuildExtractPath(sGetParentFolderName(arrFiles(1, 1)), sGetBaseName(arrFiles(1, 1)))

        If Not FileHave(sPathExtractFiles, vbDirectory) Then
            MkDir sPathExtractFiles
        End If

        sPathExtractFiles = sPathExtractFiles & Application.PathSeparator

        ' Извлекаем файлы
        For lI = 1 To UBound(arrFilesTable, 1)
            sFullNameFile = arrFilesTable(lI, 2)
            If VBA.LCase$(sGetExtensionName(sFullNameFile)) = "bin" Then
                ConvertBinToFile sFullNameFile, sPathExtractFiles, oFSO
            Else
                Call MoveFile(sFullNameFile, sPathExtractFiles & Application.PathSeparator & sGetFileName(sFullNameFile))
            End If
        Next lI
        .ZipFilesInFolder
    End With

    Call UnpackAllArchives(sPathExtractFiles, oFSO)

CleanUp:
    Set oFSO = Nothing
    Set clsZIP = Nothing
    Call MsgBox("Files are extracted!", vbInformation)
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при извлечении файлов:" & vbCrLf & _
            "Номер: " & Err.Number & vbCrLf & _
            "Описание: " & Err.Description, vbCritical, "fileExtractorFromExcelFile"
    Resume CleanUp
End Sub

' ============================================
' ФОРМИРОВАНИЕ ПУТИ ДЛЯ ИЗВЛЕЧЕНИЯ
' ============================================
Private Function BuildExtractPath(ByVal sPath As String, ByVal sFileName As String) As String
    BuildExtractPath = sPath & Application.PathSeparator & _
            sFileName & "_Extract_" & VBA.Format$(VBA.Now(), "dd_mm_yyyy_hhmmss")
End Function

' ============================================
' КОНВЕРТАЦИЯ BIN В ИСХОДНЫЙ ФОРМАТ
' ============================================
Public Sub ConvertBinToFile(ByVal sBinPath As String, _
        ByVal sDestPath As String, _
        ByRef oFSO As Object)
    Dim oStream     As Object
    Dim aFileBytes() As Byte
    Dim aOutBytes() As Byte
    Dim lStartPos   As Long
    Dim sExt        As String
    Dim sOutPath    As String
    Dim sBaseName   As String
    Dim bIsOleContainer As Boolean

    On Error GoTo ErrorHandler

    ' Проверка существования исходного файла
    If Not oFSO.FileExists(sBinPath) Then Exit Sub

    ' Читаем .bin файл в байтовый массив
    Set oStream = CreateObject("ADODB.Stream")

    With oStream
        .Type = 1    ' adTypeBinary
        .Open
        .LoadFromFile sBinPath
        aFileBytes = .Read
        .Close
    End With

    ' Проверяем, является ли файл OLE-контейнером
    bIsOleContainer = CheckSignature(aFileBytes, 0, SIG_OLE2)
    
    If bIsOleContainer Then
        ' Извлекаем содержимое из OLE-контейнера
        aOutBytes = ExtractFromOleContainer(aFileBytes, sExt)
        
        ' Если не удалось извлечь - сохраняем как OLE
        If StrPtr(sExt) = 0 Then
            sExt = EXT_OLE
            aOutBytes = aFileBytes
        End If
    Else
        ' Ищем начало встроенных данных
        lStartPos = FindEmbeddedData(aFileBytes)

        If lStartPos >= 0 Then
            ' Определяем расширение по сигнатуре
            sExt = DetectFileExtension(aFileBytes, lStartPos)

            ' Создаём массив байтов без OLE-заголовка
            aOutBytes = ExtractBytes(aFileBytes, lStartPos)
        Else
            ' Неизвестный формат
            sExt = EXT_BIN
            aOutBytes = aFileBytes
        End If
    End If

    ' Формируем имя выходного файла
    sBaseName = oFSO.GetBaseName(sBinPath)
    sOutPath = BuildOutputPath(sDestPath, sBaseName, sExt, oFSO)

    ' Сохраняем файл
    With oStream
        .Type = 1
        .Open
        .Write aOutBytes
        .SaveToFile sOutPath, 2    ' adSaveCreateOverWrite
        .Close
    End With

CleanUp:
    If Not oStream Is Nothing Then
        If oStream.State = 1 Then oStream.Close
    End If
    Set oStream = Nothing
    Exit Sub

ErrorHandler:
    Resume CleanUp
End Sub

' ============================================
' ИЗВЛЕЧЕНИЕ ДАННЫХ ИЗ OLE-КОНТЕЙНЕРА
' ============================================
Private Function ExtractFromOleContainer(ByRef aBytes() As Byte, ByRef sOutExt As String) As Byte()
    Dim lI          As Long
    Dim lBoundArr   As Long
    Dim lMaxSearch  As Long
    Dim lFoundPos   As Long
    Dim aResult()   As Byte
    
    sOutExt = vbNullString
    lBoundArr = UBound(aBytes)
    lMaxSearch = IIf(lBoundArr > MAX_OLE_CONTENT_SEARCH, MAX_OLE_CONTENT_SEARCH, lBoundArr)
    
    ' Ищем сигнатуры файлов внутри OLE-контейнера
    For lI = 512 To lMaxSearch - 4  ' Пропускаем заголовок OLE (минимум 512 байт)
        ' PDF
        If CheckSignature(aBytes, lI, SIG_PDF) Then
            lFoundPos = lI
            sOutExt = EXT_PDF
            Exit For
        End If
        
        ' ZIP / DOCX / XLSX / PPTX
        If CheckSignature(aBytes, lI, SIG_PK_ZIP) Then
            lFoundPos = lI
            sOutExt = DetectZipType(aBytes, lFoundPos)
            Exit For
        End If
        
        ' JPEG
        If CheckSignature(aBytes, lI, SIG_JPEG, 3) Then
            lFoundPos = lI
            sOutExt = EXT_JPEG
            Exit For
        End If
        
        ' PNG
        If CheckSignature(aBytes, lI, SIG_PNG) Then
            lFoundPos = lI
            sOutExt = EXT_PNG
            Exit For
        End If
        
        ' GIF
        If CheckSignature(aBytes, lI, SIG_GIF) Then
            lFoundPos = lI
            sOutExt = EXT_GIF
            Exit For
        End If
        
        ' RAR
        If CheckSignature(aBytes, lI, SIG_RAR) Then
            lFoundPos = lI
            sOutExt = EXT_RAR
            Exit For
        End If
        
        ' 7Z
        If CheckSignature(aBytes, lI, SIG_7Z) Then
            lFoundPos = lI
            sOutExt = EXT_7Z
            Exit For
        End If
    Next lI
    
    ' Если нашли сигнатуру - извлекаем данные
    If StrPtr(sOutExt) <> 0 And lFoundPos >= 0 Then
        aResult = ExtractBytesFromOle(aBytes, lFoundPos)
        ExtractFromOleContainer = aResult
    Else
        ' Возвращаем исходные данные
        ExtractFromOleContainer = aBytes
        sOutExt = EXT_OLE
    End If
End Function

' ============================================
' ИЗВЛЕЧЕНИЕ БАЙТОВ ИЗ OLE С УЧЁТОМ РАЗМЕРА
' ============================================
Private Function ExtractBytesFromOle(ByRef aSource() As Byte, ByVal lStartPos As Long) As Byte()
    Dim lOutSize    As Long
    Dim aResult()   As Byte
    Dim lI          As Long
    Dim lJ          As Long
    Dim lEndPos     As Long
    
    ' Пытаемся найти конец данных по сигнатуре или по размеру
    lEndPos = FindDataEnd(aSource, lStartPos)
    
    If lEndPos > lStartPos Then
        lOutSize = lEndPos - lStartPos
    Else
        lOutSize = UBound(aSource) - lStartPos + 1
    End If
    
    ReDim aResult(lOutSize - 1)
    
    lJ = 0
    For lI = lStartPos To lStartPos + lOutSize - 1
        If lI <= UBound(aSource) Then
            aResult(lJ) = aSource(lI)
            lJ = lJ + 1
        End If
    Next lI
    
    ExtractBytesFromOle = aResult
End Function

' ============================================
' ПОИСК КОНЦА ДАННЫХ В OLE
' ============================================
Private Function FindDataEnd(ByRef aBytes() As Byte, ByVal lStartPos As Long) As Long
    Dim lI          As Long
    Dim lBoundArr   As Long
    Dim lZeroCount  As Long
    
    lBoundArr = UBound(aBytes)
    lZeroCount = 0
    
    ' Ищем серию нулей (признак конца данных)
    For lI = lStartPos + 100 To lBoundArr  ' Минимум 100 байт данных
        If aBytes(lI) = 0 Then
            lZeroCount = lZeroCount + 1
            ' 16 нулей подряд - вероятный конец данных
            If lZeroCount >= 16 Then
                FindDataEnd = lI - lZeroCount
                Exit Function
            End If
        Else
            lZeroCount = 0
        End If
    Next lI
    
    FindDataEnd = -1
End Function

' ============================================
' ИЗВЛЕЧЕНИЕ БАЙТОВ БЕЗ OLE-ЗАГОЛОВКА
' ============================================
Private Function ExtractBytes(ByRef aSource() As Byte, ByVal lStartPos As Long) As Byte()
    Dim lOutSize    As Long
    Dim aResult()   As Byte
    Dim lI          As Long
    Dim lJ          As Long

    lOutSize = UBound(aSource) - lStartPos
    ReDim aResult(lOutSize)

    lJ = 0
    For lI = lStartPos To UBound(aSource)
        aResult(lJ) = aSource(lI)
        lJ = lJ + 1
    Next lI

    ExtractBytes = aResult
End Function

' ============================================
' ФОРМИРОВАНИЕ УНИКАЛЬНОГО ПУТИ ВЫХОДНОГО ФАЙЛА
' ============================================
Private Function BuildOutputPath(ByVal sDestPath As String, _
        ByVal sBaseName As String, _
        ByVal sExt As String, _
        ByRef oFSO As Object) As String
    Dim sPath       As String

    sPath = sDestPath & sBaseName & sExt

    If oFSO.FileExists(sPath) Then
        sPath = sDestPath & sBaseName & "_" & Format(Now, "hhmmss") & sExt
    End If

    BuildOutputPath = sPath
End Function

' ============================================
' ПОИСК НАЧАЛА ВСТРОЕННЫХ ДАННЫХ
' ============================================
Private Function FindEmbeddedData(ByRef aBytes() As Byte) As Long
    Dim lI          As Long
    Dim lMaxSearch  As Long
    Dim lBoundArr   As Long

    lBoundArr = UBound(aBytes)
    lMaxSearch = IIf(lBoundArr > MAX_OLE_HEADER_SIZE, MAX_OLE_HEADER_SIZE, lBoundArr)

    ' Минимальный размер для проверки сигнатуры
    If lMaxSearch < 4 Then
        FindEmbeddedData = -1
        Exit Function
    End If

    For lI = 0 To lMaxSearch - 4
        ' ZIP / DOCX / XLSX / PPTX
        If CheckSignature(aBytes, lI, SIG_PK_ZIP) Then
            FindEmbeddedData = lI
            Exit Function
        End If

        ' RAR
        If CheckSignature(aBytes, lI, SIG_RAR) Then
            FindEmbeddedData = lI
            Exit Function
        End If

        ' 7Z
        If CheckSignature(aBytes, lI, SIG_7Z) Then
            FindEmbeddedData = lI
            Exit Function
        End If

        ' PDF
        If CheckSignature(aBytes, lI, SIG_PDF) Then
            FindEmbeddedData = lI
            Exit Function
        End If

        ' JPEG
        If CheckSignature(aBytes, lI, SIG_JPEG, 3) Then
            FindEmbeddedData = lI
            Exit Function
        End If

        ' PNG
        If CheckSignature(aBytes, lI, SIG_PNG) Then
            FindEmbeddedData = lI
            Exit Function
        End If

        ' GIF
        If CheckSignature(aBytes, lI, SIG_GIF) Then
            FindEmbeddedData = lI
            Exit Function
        End If

        ' OLE2 - пропускаем, обрабатываем отдельно
    Next lI

    FindEmbeddedData = -1
End Function

' ============================================
' ПРОВЕРКА СИГНАТУРЫ ФАЙЛА
' ============================================
Private Function CheckSignature(ByRef aBytes() As Byte, _
        ByVal lPos As Long, _
        ByVal sSignature As String, _
        Optional ByVal lLength As Long = 4) As Boolean
    Dim sHex        As String
    Dim lI          As Long

    CheckSignature = False

    If lPos + lLength > UBound(aBytes) + 1 Then Exit Function

    sHex = vbNullString
    For lI = 0 To lLength - 1
        sHex = sHex & Right$("0" & Hex$(aBytes(lPos + lI)), 2)
    Next lI

    CheckSignature = (sHex = sSignature)
End Function

' ============================================
' ОПРЕДЕЛЕНИЕ РАСШИРЕНИЯ ФАЙЛА ПО СИГНАТУРЕ
' ============================================
Private Function DetectFileExtension(ByRef aBytes() As Byte, ByVal lStartPos As Long) As String
    ' ZIP (также DOCX, XLSX, PPTX)
    If CheckSignature(aBytes, lStartPos, SIG_PK_ZIP) Then
        DetectFileExtension = DetectZipType(aBytes, lStartPos)
        Exit Function
    End If

    ' RAR
    If CheckSignature(aBytes, lStartPos, SIG_RAR) Then
        DetectFileExtension = EXT_RAR
        Exit Function
    End If

    ' 7Z
    If CheckSignature(aBytes, lStartPos, SIG_7Z) Then
        DetectFileExtension = EXT_7Z
        Exit Function
    End If

    ' PDF
    If CheckSignature(aBytes, lStartPos, SIG_PDF) Then
        DetectFileExtension = EXT_PDF
        Exit Function
    End If

    ' JPEG
    If CheckSignature(aBytes, lStartPos, SIG_JPEG, 3) Then
        DetectFileExtension = EXT_JPEG
        Exit Function
    End If

    ' PNG
    If CheckSignature(aBytes, lStartPos, SIG_PNG) Then
        DetectFileExtension = EXT_PNG
        Exit Function
    End If

    ' GIF
    If CheckSignature(aBytes, lStartPos, SIG_GIF) Then
        DetectFileExtension = EXT_GIF
        Exit Function
    End If

    ' OLE2 / CFB (xls, doc, ppt)
    If CheckSignature(aBytes, lStartPos, SIG_OLE2) Then
        DetectFileExtension = DetectOle2Type(aBytes, lStartPos)
        Exit Function
    End If

    ' По умолчанию
    DetectFileExtension = EXT_BIN
End Function

' ============================================
' ОПРЕДЕЛЕНИЕ ТИПА OFFICE ВНУТРИ ZIP
' ============================================
Private Function DetectZipType(ByRef aBytes() As Byte, ByVal lStartPos As Long) As String
    Dim lI          As Long
    Dim lMaxSearch  As Long
    Dim lBoundArr   As Long

    lBoundArr = UBound(aBytes)
    lMaxSearch = IIf(lBoundArr > lStartPos + MAX_ZIP_TYPE_SEARCH, _
            lStartPos + MAX_ZIP_TYPE_SEARCH, lBoundArr)

    For lI = lStartPos To lMaxSearch - 10
        ' xl\ = Excel
        If aBytes(lI) = &H78 And aBytes(lI + 1) = &H6C And aBytes(lI + 2) = &H5C Then
            DetectZipType = EXT_XLSX
            Exit Function
        End If

        ' word\ = Word
        If aBytes(lI) = &H77 And aBytes(lI + 1) = &H6F And _
                aBytes(lI + 2) = &H72 And aBytes(lI + 3) = &H64 Then
            DetectZipType = EXT_DOCX
            Exit Function
        End If

        ' ppt\ = PowerPoint
        If aBytes(lI) = &H70 And aBytes(lI + 1) = &H70 And aBytes(lI + 2) = &H74 Then
            DetectZipType = EXT_PPTX
            Exit Function
        End If
    Next lI

    DetectZipType = EXT_ZIP
End Function

' ============================================
' ОПРЕДЕЛЕНИЕ ТИПА OLE2 (xls, doc, ppt)
' ============================================
Private Function DetectOle2Type(ByRef aBytes() As Byte, ByVal lStartPos As Long) As String
    Dim lI          As Long
    Dim lMaxSearch  As Long
    Dim lBoundArr   As Long

    lBoundArr = UBound(aBytes)
    lMaxSearch = IIf(lBoundArr > lStartPos + MAX_ZIP_TYPE_SEARCH, _
            lStartPos + MAX_ZIP_TYPE_SEARCH, lBoundArr)

    For lI = lStartPos To lMaxSearch - 12
        ' Workbook = Excel (xls)
        If aBytes(lI) = &H57 And aBytes(lI + 1) = &H6F And _
                aBytes(lI + 2) = &H72 And aBytes(lI + 3) = &H6B And _
                aBytes(lI + 4) = &H62 And aBytes(lI + 5) = &H6F And _
                aBytes(lI + 6) = &H6F And aBytes(lI + 7) = &H6B Then
            DetectOle2Type = EXT_XLS
            Exit Function
        End If

        ' WordDocument = Word (doc)
        If aBytes(lI) = &H57 And aBytes(lI + 1) = &H6F And _
                aBytes(lI + 2) = &H72 And aBytes(lI + 3) = &H64 And _
                aBytes(lI + 4) = &H44 And aBytes(lI + 5) = &H6F And _
                aBytes(lI + 6) = &H63 And aBytes(lI + 7) = &H75 And _
                aBytes(lI + 8) = &H6D And aBytes(lI + 9) = &H65 And _
                aBytes(lI + 10) = &H6E And aBytes(lI + 11) = &H74 Then
            DetectOle2Type = EXT_DOC
            Exit Function
        End If

        ' PowerPoint Document = PowerPoint (ppt)
        If aBytes(lI) = &H50 And aBytes(lI + 1) = &H6F And _
                aBytes(lI + 2) = &H77 And aBytes(lI + 3) = &H65 And _
                aBytes(lI + 4) = &H72 And aBytes(lI + 5) = &H50 And _
                aBytes(lI + 6) = &H6F And aBytes(lI + 7) = &H69 And _
                aBytes(lI + 8) = &H6E And aBytes(lI + 9) = &H74 Then
            DetectOle2Type = EXT_PPT
            Exit Function
        End If
    Next lI

    ' Если тип не определён — используем универсальное расширение
    DetectOle2Type = EXT_OLE
End Function

' ============================================
' РАСПАКОВКА ВСЕХ АРХИВОВ В ПАПКЕ
' ============================================
Private Sub UnpackAllArchives(ByVal sFolderPath As String, ByRef oFSO As Object)
    Dim oFolder     As Object
    Dim oFile       As Object
    Dim oFiles      As Object
    Dim sExt        As String
    Dim sArchivePath As String
    Dim sDestFolder As String
    Dim bUnpacked   As Boolean

    On Error Resume Next

    Set oFolder = oFSO.GetFolder(sFolderPath)
    If oFolder Is Nothing Then Exit Sub

    Set oFiles = oFolder.Files

    ' Проходим по всем файлам в папке
    For Each oFile In oFiles
        sExt = LCase$(oFSO.GetExtensionName(oFile.Path))
        sArchivePath = oFile.Path

        bUnpacked = False

        Select Case sExt
            Case "zip"
                bUnpacked = UnpackZip(sArchivePath, sFolderPath, oFSO)

            Case "rar"
                bUnpacked = UnpackRar(sArchivePath, sFolderPath, oFSO)

            Case "7z"
                bUnpacked = Unpack7z(sArchivePath, sFolderPath, oFSO)
        End Select

        ' Удаляем архив после успешной распаковки
        If bUnpacked Then
            On Error Resume Next
            oFSO.DeleteFile sArchivePath, True
            On Error GoTo 0
        End If
    Next oFile

    ' Рекурсивно проверяем подпапки на наличие архивов
    Dim oSubFolder  As Object
    For Each oSubFolder In oFolder.SubFolders
        UnpackAllArchives oSubFolder.Path & Application.PathSeparator, oFSO
    Next oSubFolder
End Sub

' ============================================
' РАСПАКОВКА ZIP (через Shell.Application)
' ============================================
Private Function UnpackZip(ByVal sZipPath As String, _
        ByVal sDestPath As String, _
        ByRef oFSO As Object) As Boolean
    Dim oShell      As Object
    Dim oZipFile    As Object
    Dim oDestFolder As Object
    Dim oItems      As Object
    Dim sDestFolder As String
    Dim lRetry      As Long

    On Error GoTo ErrorHandler

    UnpackZip = False

    ' Проверяем существование архива
    If Not oFSO.FileExists(sZipPath) Then Exit Function

    ' Создаём папку для распаковки
    sDestFolder = sDestPath & oFSO.GetBaseName(sZipPath)

    If Not oFSO.FolderExists(sDestFolder) Then
        oFSO.CreateFolder sDestFolder
    End If

    ' Используем Shell.Application для распаковки
    Set oShell = CreateObject("Shell.Application")

    ' Ожидание освобождения файла
    For lRetry = 1 To 10
        On Error Resume Next
        Set oZipFile = oShell.Namespace(sZipPath)
        On Error GoTo ErrorHandler

        If Not oZipFile Is Nothing Then Exit For
        Call VBA.DoEvents
        Call Sleep(100)
    Next lRetry

    If oZipFile Is Nothing Then Exit Function

    Set oDestFolder = oShell.Namespace(sDestFolder)
    If oDestFolder Is Nothing Then Exit Function

    Set oItems = oZipFile.Items
    If oItems.Count = 0 Then Exit Function

    ' Копируем все файлы из архива
    oDestFolder.CopyHere oItems, 4 + 16    ' Скрытое + Да для всех

    ' Ожидание завершения распаковки
    Call WaitForExtraction(sDestFolder, oItems.Count, oFSO)

    UnpackZip = True

CleanUp:
    Set oItems = Nothing
    Set oDestFolder = Nothing
    Set oZipFile = Nothing
    Set oShell = Nothing
    Exit Function

ErrorHandler:
    Resume CleanUp
End Function

' ============================================
' РАСПАКОВКА RAR (через WinRAR или 7-Zip)
' ============================================
Private Function UnpackRar(ByVal sRarPath As String, _
        ByVal sDestPath As String, _
        ByRef oFSO As Object) As Boolean
    Dim sExePath    As String
    Dim sDestFolder As String
    Dim sCmd        As String
    Dim lResult     As Long

    On Error GoTo ErrorHandler

    UnpackRar = False

    ' Проверяем существование архива
    If Not oFSO.FileExists(sRarPath) Then Exit Function

    ' Создаём папку для распаковки
    sDestFolder = sDestPath

    If Not oFSO.FolderExists(sDestFolder) Then
        oFSO.CreateFolder sDestFolder
    End If

    ' Определяем путь к архиватору
    sExePath = GetArchiverPath("rar")
    If sExePath = vbNullString Then
        ' Пробуем через 7-Zip
        sExePath = GetArchiverPath("7z")
    End If

    If sExePath = vbNullString Then
        Debug.Print "Архиватор для RAR не найден"
        Exit Function
    End If

    ' Формируем командную строку
    sCmd = """" & sExePath & """ x -y -o+ """ & sRarPath & """ """ & sDestFolder & """"

    ' Выполняем распаковку
    lResult = Shell(sCmd, vbHide)

    ' Ожидание завершения (простейший вариант)
    Call Sleep(500)
    Call WaitForProcess(lResult)

    ' Проверяем результат
    UnpackRar = True

    Exit Function

ErrorHandler:
    UnpackRar = False
End Function

' ============================================
' РАСПАКОВКА 7Z (через 7-Zip)
' ============================================
Private Function Unpack7z(ByVal s7zPath As String, _
        ByVal sDestPath As String, _
        ByRef oFSO As Object) As Boolean
    Dim sExePath    As String
    Dim sDestFolder As String
    Dim sCmd        As String
    Dim lResult     As Long

    On Error GoTo ErrorHandler

    Unpack7z = False

    ' Проверяем существование архива
    If Not oFSO.FileExists(s7zPath) Then Exit Function

    ' Создаём папку для распаковки
    sDestFolder = sDestPath & oFSO.GetBaseName(s7zPath)

    If Not oFSO.FolderExists(sDestFolder) Then
        oFSO.CreateFolder sDestFolder
    End If

    ' Определяем путь к 7-Zip
    sExePath = GetArchiverPath("7z")

    If sExePath = vbNullString Then
        Debug.Print "7-Zip не найден"
        Exit Function
    End If

    ' Формируем командную строку
    sCmd = """" & sExePath & """ x -y -o""" & sDestFolder & """ """ & s7zPath & """"

    ' Выполняем распаковку
    lResult = Shell(sCmd, vbHide)

    ' Ожидание завершения
    Call Sleep(500)
    Call WaitForProcess(lResult)

    ' Проверяем результат
    Unpack7z = True

    Exit Function

ErrorHandler:
    Unpack7z = False
End Function

' ============================================
' ПОЛУЧЕНИЕ ПУТИ К АРХИВАТОРУ
' ============================================
Private Function GetArchiverPath(ByVal sType As String) As String
    Dim oFSO        As Object
    Dim sPath       As String

    Set oFSO = CreateObject("Scripting.FileSystemObject")

    GetArchiverPath = vbNullString

    Select Case LCase$(sType)
        Case "7z"
            ' Проверяем разные варианты расположения 7-Zip
            If oFSO.FileExists(PATH_7ZIP) Then
                GetArchiverPath = PATH_7ZIP
            ElseIf oFSO.FileExists(PATH_7ZIP_X86) Then
                GetArchiverPath = PATH_7ZIP_X86
            Else
                ' Пробуем найти через PATH
                sPath = Environ$("ProgramFiles") & "\7-Zip\7z.exe"
                If oFSO.FileExists(sPath) Then
                    GetArchiverPath = sPath
                End If
            End If

        Case "rar"
            ' Проверяем WinRAR
            If oFSO.FileExists(PATH_WINRAR) Then
                GetArchiverPath = PATH_WINRAR
            ElseIf oFSO.FileExists(PATH_WINRAR_X86) Then
                GetArchiverPath = PATH_WINRAR_X86
            Else
                sPath = Environ$("ProgramFiles") & "\WinRAR\WinRAR.exe"
                If oFSO.FileExists(sPath) Then
                    GetArchiverPath = sPath
                End If
            End If
    End Select

    Set oFSO = Nothing
End Function

' ============================================
' ОЖИДАНИЕ ЗАВЕРШЕНИЯ РАСПАКОВКИ (Shell.Application)
' ============================================
Private Sub WaitForExtraction(ByVal sFolderPath As String, _
        ByVal lExpectedCount As Long, _
        ByRef oFSO As Object)
    Dim lTimeout    As Long
    Dim lFileCount  As Long
    Dim oFolder     As Object

    lTimeout = 0

    Do While lTimeout < 30000    ' Максимум 30 секунд
        On Error Resume Next
        Set oFolder = oFSO.GetFolder(sFolderPath)
        lFileCount = oFolder.Files.Count
        On Error GoTo 0

        If lFileCount >= lExpectedCount Then Exit Do

        Call Sleep(100)
        lTimeout = lTimeout + 100
        Call VBA.DoEvents
    Loop
End Sub

' ============================================
' ОЖИДАНИЕ ЗАВЕРШЕНИЯ ВНЕШНЕГО ПРОЦЕССА
' ============================================
Private Sub WaitForProcess(ByVal lProcessId As Long)
    Dim oWMI        As Object
    Dim oProcess    As Object
    Dim oProcesses  As Object
    Dim lTimeout    As Long

    On Error Resume Next

    Set oWMI = GetObject("winmgmts:\\.\root\cimv2")

    lTimeout = 0

    Do While lTimeout < 60000    ' Максимум 60 секунд
        Set oProcesses = oWMI.ExecQuery( _
                "SELECT * FROM Win32_Process WHERE ProcessId = " & lProcessId)

        If oProcesses.Count = 0 Then Exit Do

        Call Sleep(200)
        lTimeout = lTimeout + 200
        Call VBA.DoEvents
    Loop

    Set oProcesses = Nothing
    Set oWMI = Nothing
End Sub

