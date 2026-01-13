VERSION 5.00
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
'* Module     : addOptions - создание OPTIONs в модулях проекта
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
    sTxt = "Используется на уровне модуля для объявления нижней границы массивов, по умолчанию." & vbNewLine & vbNewLine
    sTxt = sTxt & "Синтаксис" & vbNewLine & "Option Base { 0 | 1 }" & vbNewLine & vbNewLine
    sTxt = sTxt & "Поскольку Option Base по умолчанию равна 0, оператор Option Base никогда не используется. Оператор должен находиться в модуле до всех процедур." & vbNewLine
    sTxt = sTxt & "Оператор Option Base может указываться в модуле только один раз и должен предшествовать объявлениям массивов, включающим размерности." & vbNewLine & vbNewLine
    sTxt = sTxt & "Примечание" & vbNewLine & vbNewLine
    sTxt = sTxt & "Предложение To в инструкциях Dim, Private, Public, ReDim и Static предоставляет более гибкий способ управления диапазоном индексов массива." & vbNewLine
    sTxt = sTxt & "Однако если нижняя граница индексов не задается явно в предложении To, можно воспользоваться инструкцией Option Base," & vbNewLine
    sTxt = sTxt & "чтобы установить используемую по умолчанию нижнюю границу индексов, равную 1. Нижняя граница значений индексов массивов," & vbNewLine
    sTxt = sTxt & "создаваемых с помощью функции Array, всегда равняется нулю; вне зависимости от инструкции Option Base."
    sTxt = sTxt & vbNewLine & vbNewLine & "Инструкция Option Base действует на нижнюю границу индексов массивов только того модуля, в котором расположена сама эта инструкция."
    Debug.Print sTxt
End Sub
Private Sub lbCompare_Click()
    Dim sTxt        As String
    sTxt = "Используется на уровне модуля для объявления метода сравнения по умолчанию, который будет использоваться при сравнении строковых данных." & vbNewLine & vbNewLine
    sTxt = sTxt & "Синтаксис" & vbNewLine & "Option Compare { Binary | Text | Database }" & vbNewLine & vbNewLine
    sTxt = sTxt & "Примечание" & vbNewLine & vbNewLine
    sTxt = sTxt & "Инструкция Option Compare при ее использовании должна находиться в модуле перед любой процедурой." & vbNewLine
    sTxt = sTxt & "Инструкция Option Compare указывает способ сравнения строк (Binary, Text или Database) для модуля." & vbNewLine
    sTxt = sTxt & "Если модуль не содержит инструкцию Option Compare, по умолчанию используется способ сравнения Binary." & vbNewLine
    sTxt = sTxt & "Инструкция Option Compare Binary задает сравнение строк на основе порядка сортировки, определяемого внутренним двоичным представлением символов." & vbNewLine
    sTxt = sTxt & "В Microsoft Windows порядок сортировки определяется кодовой страницей символов." & vbNewLine
    sTxt = sTxt & "В следующем примере представлен типичный результат двоичного порядка сортировки:" & vbNewLine & vbNewLine
    sTxt = sTxt & "A < B < E < Z < a < b < e < z < Б < Л < Ш < б < л < ш" & vbNewLine & vbNewLine
    sTxt = sTxt & "Инструкция Option Compare Text задает сравнение строк без учета регистра символов на основе системной национальной настройки." & vbNewLine
    sTxt = sTxt & "Тем же символам, что и выше, при сортировке с инструкцией Option Compare Text соответствует следующий порядок: " & vbNewLine & vbNewLine
    sTxt = sTxt & "(A=a) < (B=b) < (E=e) < (Z=z) < (Б=б) < (Л=л) < (Ш=ш)" & vbNewLine & vbNewLine
    sTxt = sTxt & "Инструкция Option Compare Database может использоваться только в Microsoft Access. При этом задает сравнение строк на основе порядка сортировки," & vbNewLine
    sTxt = sTxt & "определяемого национальной настройкой базы данных, в которой производится сравнение строк. "
    Debug.Print sTxt
End Sub
Private Sub lbExplicit_Click()
    Dim sTxt        As String
    sTxt = "Используется на уровне модуля для принудительного явного объявления всех переменных в этом модуле." & vbNewLine & vbNewLine
    sTxt = sTxt & "Синтаксис" & vbNewLine & "Option Explicit" & vbNewLine & vbNewLine
    sTxt = sTxt & "Примечание" & vbNewLine & vbNewLine
    sTxt = sTxt & "Инструкция Option Explicit при ее использовании должна находиться в модуле до любой процедуры." & vbNewLine
    sTxt = sTxt & "При использовании инструкции Option Explicit необходимо явно описать все переменные с помощью инструкций Dim, Private, Public, ReDim или Static." & vbNewLine
    sTxt = sTxt & "При попытке использовать неописанное имя переменной возникает ошибка во время компиляции." & vbNewLine
    sTxt = sTxt & "Когда инструкция Option Explicit не используется, все неописанные переменные имеют тип Variant, если используемый по умолчанию тип данных не задается с помощью инструкции DefТип." & vbNewLine
    sTxt = sTxt & "Используйте инструкцию Option Explicit, чтобы избежать неверного ввода имени имеющейся переменной или риска конфликтов в программе, когда область определения переменной не совсем ясна."
    Debug.Print sTxt
End Sub
Private Sub lbPrivate_Click()
    Dim sTxt        As String
    sTxt = "Используется на уровне модуля для запрета ссылок на контент модуля извне проекта." & vbNewLine & vbNewLine
    sTxt = sTxt & "Синтаксис" & vbNewLine & "Option Private Module" & vbNewLine & vbNewLine
    sTxt = sTxt & "Примечание" & vbNewLine & vbNewLine
    sTxt = sTxt & "Когда модуль содержит инструкцию Option Private Module, общие элементы, например, переменные, объекты и определяемые пользователем типы, описанные на уровне модуля," & vbNewLine
    sTxt = sTxt & "остаются доступными внутри проекта, содержащего этот модуль, но недоступными для других приложений или проектов." & vbNewLine
    sTxt = sTxt & "Microsoft Excel поддерживает загрузку нескольких проектов. В этом случае инструкция Option Private Module позволяет ограничить взаимную видимость проектов."
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
            Debug.Print "Нет активного модуля, перейдите в модуль кода!"
            Exit Sub
        Case 76:
            Exit Sub
        Case Else:
            Debug.Print "Ошибка! в addOptions" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "в строке " & Erl
            'Call WriteErrorLog("addOptions")
    End Select
    Err.Clear
End Sub
