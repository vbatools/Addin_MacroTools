Attribute VB_Name = "modAddinCreateMenu"
Option Explicit
Option Private Module
Private ToolContextEventHandlers As Collection
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RefreshMenu - Reload add-in tools menu in the VBE code editor
'* Created    : 22-03-2023 14:36
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub RefreshMenu()
    Call DeleteContextMenus
    Call AddContextMenus
    Call MsgBox("Add-in reload " & modAddinConst.NAME_ADDIN & "completed!", vbInformation, "Add-in reload " & modAddinConst.NAME_ADDIN & ":")
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : Auto_Open - Create menu in the VBE code editor on startup
'* Created    : 22-03-2023 14:27
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Auto_Open()
    If VBAIsTrusted And ThisWorkbook.Name = modAddinConst.NAME_ADDIN & ".xlam" Then Call AddContextMenus
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : Auto_Close - Remove menu in the VBE code editor when add-in is closed
'* Created    : 22-03-2023 14:31
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Auto_Close()
    Call DeleteContextMenus
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddContextMenus - Create menu in the VBE code editor
'* Created    : 22-03-2023 14:27
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddContextMenus()

    Call AddNewCommandBarMenu(modAddinConst.MENU_MOVE_CONTROLS)
    Call AddButtom(984, "Tool Help", "HelpMoveControl", modAddinConst.MENU_MOVE_CONTROLS, vbNullString, False, True)
    Call AddButtom(38, "Move Up", "MoveControlUp", modAddinConst.MENU_MOVE_CONTROLS, vbNullString)
    Call AddButtom(40, "Move Down", "MoveControlDown", modAddinConst.MENU_MOVE_CONTROLS, vbNullString, False, True)
    Call AddButtom(39, "Move Left", "MoveControlLeft", modAddinConst.MENU_MOVE_CONTROLS, vbNullString)
    Call AddButtom(41, "Move Right", "MoveControlRight", modAddinConst.MENU_MOVE_CONTROLS, vbNullString)
    Call AddComboBox(modAddinConst.MENU_MOVE_CONTROLS, modAddinConst.MENU_MOVE_CONTROLS, "Control", Array("Control", "Top Left", "Bottom Right"))

    Call AddNewCommandBarMenu(modAddinConst.MENU_TOOLS)

    Call AddButtom(277, "Hotkeys", "AddLegendHotKeys", modAddinConst.MENU_TOOLS, vbNullString, False, True)
    Call AddButtom(0, "FormatBuilder", "showBilderFormat", modAddinConst.MENU_TOOLS, vbNullString, True, True)
    Call AddButtom(0, "MsgBoxBuilder", "showMsgBoxGenerator", modAddinConst.MENU_TOOLS, vbNullString, True, True)
    Call AddButtom(0, "ProcedureBuilder", "showBilderProcedure", modAddinConst.MENU_TOOLS, vbNullString, True, True)

    Call AddButtom(107, "Option's Explicit and Private Module", "insertOptionsExplicitAndPrivateModule", modAddinConst.MENU_TOOLS, vbNullString, False, False)
    Call AddButtom(0, "Option's", "subOptionsForm", modAddinConst.MENU_TOOLS, vbNullString, True, True)

    Call AddButtom(3838, "Close All VBE Windows", "CloseAllWindowsVBE", modAddinConst.MENU_TOOLS, vbNullString, False, True)

    Call AddButtom(8, "TODO List", "ShowTODOList", modAddinConst.MENU_TOOLS, vbNullString, False, False)
    Call AddButtom(1972, "Create TODO", "sysAddTODOTop", modAddinConst.MENU_TOOLS, vbNullString, False, False)
    Call AddButtom(456, "Create Update Comment Line", "sysAddModifiedTop", modAddinConst.MENU_TOOLS, vbNullString, False, False)
    Call AddButtom(1546, "Create Comment", "sysAddHeaderTop", modAddinConst.MENU_TOOLS, vbNullString, False, True)

    Call AddButtom(2474, "Remove Code Line Breaks", "delBreaksLinesInCodeVBA", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(6939, "Remove All Comments from Code", "delCommentsInCodeVBA", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(699, "Remove Double Blank Lines", "delTwoEmptyStrings", modAddinConst.MENU_TOOLS, vbNullString, False, True)

    Call AddButtom(7460, "Multi-line Dim's", "dimMultiLine", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(7772, "Single-line Dim's", "dimOneLine", modAddinConst.MENU_TOOLS, vbNullString, False, True)

    Call AddButtom(7770, "Comment Out 'Debug.print", "debugOff", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(7771, "Uncomment Debug.print", "debugOn", modAddinConst.MENU_TOOLS, vbNullString, False, True)


    Call AddButtom(3917, "Remove Code Formatting", "CutTab", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(3919, "Format Code", "ReBild", modAddinConst.MENU_TOOLS, vbNullString, False, True)
    Call AddButtom(12, "Remove Line Numbers", "RemoveLineNumbersVBProject", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(11, "Add Line Numbers", "AddLineNumbersVBProject", modAddinConst.MENU_TOOLS, vbNullString)

    Call AddComboBox(modAddinConst.MENU_TOOLS, modAddinConst.MENU_TOOLS, modAddinConst.TYPE_SELECTED_MODULE, Array(modAddinConst.TYPE_ALL_VBAPROJECT, modAddinConst.TYPE_SELECTED_MODULE), True)

    Call AddButtom(21, "Delete Module", "DeleteSnippetEnumModule", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(1753, "Insert Module", "AddSnippetEnumModule", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(22, "Insert Code", "InsertCodeFromSnippet", modAddinConst.MENU_TOOLS, vbNullString, False, False)


    Call AddButtom(9634, "Swap Relative/Absolute [=]", "SwapEgual", modAddinConst.MENU_CODE_WINDOW, "SwapEgual", True, False)
    Call AddButtom(0, "UPPER Case", "toUpperCase", modAddinConst.MENU_CODE_WINDOW, "toUpperCase", True, False)
    Call AddButtom(0, "lower Case", "toLowerCase", modAddinConst.MENU_CODE_WINDOW, "toLowerCase", True, False)
    Call AddButtom(22, "Insert Code", "InsertCodeFromSnippet", modAddinConst.MENU_CODE_WINDOW, "InsertCodeFromSnippet", True, False)

    Call AddButtom(1650, "Aling Horiz", "vbaCntAlingHoriz", modAddinConst.MENU_FORMS, "Aling Horiz", True)
    Call AddButtom(1653, "Aling Vert", "vbaCntAlingVert", modAddinConst.MENU_FORMS, "Aling Vert", True)

    Call AddButtom(162, "ReName Control", "RenameControl", modAddinConst.MENU_FORMS, "ReName Control", True)

    Call AddButtom(22, "Paste Style", "PasteStyleControl", modAddinConst.MENU_FORMS, "Paste Style", True)
    Call AddButtom(1076, "Copy Style", "CopyStyleControl", modAddinConst.MENU_FORMS, "Copy Style", True)

    Call AddButtom(0, "UPPER CASE", "UperTextInControl", modAddinConst.MENU_FORMS, "UPPER CASE", True, False)
    Call AddButtom(0, "lower case", "LowerTextInControl", modAddinConst.MENU_FORMS, "lower case", True, False)

    Call AddButtom(2045, "Copy Module", "CopyModyleVBE", modAddinConst.MENU_PROJECT_WINDOW, "Copy Module", True, False)

    Call AddButtom(22, "Paste Style", "PasteStyleForms", modAddinConst.MENU_MS_FORMS, "Paste Style", True)
    Call AddButtom(1076, "Copy Style", "CopyStyleForms", modAddinConst.MENU_MS_FORMS, "Copy Style", True)

    Call AddButtom(0, "UPPER CASE", "UperTextInForm", modAddinConst.MENU_MS_FORMS, "UPPER CASE", True, False)
    Call AddButtom(0, "lower case", "LowerTextInForm", modAddinConst.MENU_MS_FORMS, "lower case", True, False)
End Sub


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddNewCommandBarMenu - Create main menu in the VBE editor
'* Created    : 22-03-2023 14:28
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                     Description
'*
'* ByVal sNameCommandBar As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function AddNewCommandBarMenu(ByVal sNameCommandBar As String) As CommandBar
    Dim myCommandBar As CommandBar
    On Error GoTo AddNewCommandBar
    Set myCommandBar = Application.VBE.CommandBars(sNameCommandBar)
    If myCommandBar Is Nothing Then
AddNewCommandBar:
        Set myCommandBar = Application.VBE.CommandBars.Add(Name:=sNameCommandBar, Position:=msoBarTop)
        With myCommandBar
            .Visible = True
            .rowIndex = 3
        End With
    End If
    Set AddNewCommandBarMenu = myCommandBar
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : DeleteContextMenus - Remove all child controls in the VBE code editor menu
'* Created    : 22-03-2023 14:32
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub DeleteContextMenus()
    Dim myCommandBar As CommandBar
    On Error Resume Next

    Set myCommandBar = Application.VBE.CommandBars(modAddinConst.MENU_MOVE_CONTROLS)
    If Not myCommandBar Is Nothing Then myCommandBar.Delete

    Set myCommandBar = Application.VBE.CommandBars(modAddinConst.MENU_TOOLS)
    If Not myCommandBar Is Nothing Then myCommandBar.Delete

    On Error GoTo 0

    Call DeleteButton("SwapEgual", modAddinConst.MENU_CODE_WINDOW)
    Call DeleteButton("toUpperCase", modAddinConst.MENU_CODE_WINDOW)
    Call DeleteButton("toLowerCase", modAddinConst.MENU_CODE_WINDOW)
    Call DeleteButton("InsertCodeFromSnippet", modAddinConst.MENU_CODE_WINDOW)

    Call DeleteButton("Aling Horiz", modAddinConst.MENU_FORMS)
    Call DeleteButton("Aling Vert", modAddinConst.MENU_FORMS)
    Call DeleteButton("ReName Control", modAddinConst.MENU_FORMS)
    Call DeleteButton("Paste Style", modAddinConst.MENU_FORMS)
    Call DeleteButton("Copy Style", modAddinConst.MENU_FORMS)
    Call DeleteButton("UPPER CASE", modAddinConst.MENU_FORMS)
    Call DeleteButton("lower case", modAddinConst.MENU_FORMS)

    Call DeleteButton("Copy Module", modAddinConst.MENU_PROJECT_WINDOW)

    Call DeleteButton("Paste Style", modAddinConst.MENU_MS_FORMS)
    Call DeleteButton("Copy Style", modAddinConst.MENU_MS_FORMS)
    Call DeleteButton("UPPER CASE", modAddinConst.MENU_MS_FORMS)
    Call DeleteButton("lower case", modAddinConst.MENU_MS_FORMS)

    Set ToolContextEventHandlers = Nothing
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddButtom - Create a button in the VBE code editor
'* Created    : 22-03-2023 14:29
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                             Description
'*
'* ByVal Face As Long                                   :
'* ByVal Capitan As String                              :
'* ByVal sOnAction As String                            :
'* ByVal sMenu As String                                :
'* Optional ByRef VisibleCapiton As Boolean = False     :
'* Optional ByVal BeginGroup As Boolean = False         :
'* Optional ByVal ShortcutText As String = vbNullString :
'* Optional ByVal Before As Byte = 1                    :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddButtom(ByVal Face As Long, _
        ByVal Capitan As String, _
        ByVal sOnAction As String, _
        ByVal sNameCommandBar As String, _
        ByVal sTag As String, _
        Optional ByRef VisibleCapiton As Boolean = False, _
        Optional ByVal bBeginGroup As Boolean = False, _
        Optional ByVal sShortcutText As String = vbNullString, _
        Optional ByVal byBefore As Byte = 1)
    Dim btn         As CommandBarButton
    Dim evtContextMenu As clsVBECommandHandler
    On Error GoTo ErrorHandler
    Set btn = Application.VBE.CommandBars(sNameCommandBar).Controls.Add(Type:=msoControlButton, Before:=byBefore)
    With btn
        .FaceId = Face
        If VisibleCapiton Then .Caption = Capitan
        .TooltipText = Capitan
        .Tag = sTag
        .OnAction = "'" & ThisWorkbook.Name & "'!" & sOnAction
        .Style = msoButtonIconAndCaption
        .BeginGroup = bBeginGroup
        .ShortcutText = sShortcutText
    End With
    Set evtContextMenu = New clsVBECommandHandler
    Set evtContextMenu.cmdButton = btn
    If ToolContextEventHandlers Is Nothing Then Set ToolContextEventHandlers = New Collection
    Call ToolContextEventHandlers.Add(evtContextMenu)
    Exit Sub
ErrorHandler:
    Call WriteErrorLog("AddButtom", False)
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : DeleteButton - Remove a button in the VBE code editor
'* Created    : 22-03-2023 14:33
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByRef sTag As String  :
'* ByVal sMenu As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub DeleteButton(ByVal sTag As String, ByVal sMenu As String)
    Dim Cbar        As CommandBar
    Dim ctrl        As CommandBarControl
    On Error Resume Next
    Set Cbar = Application.VBE.CommandBars(sMenu)
    If Cbar Is Nothing Then Exit Sub
    On Error GoTo ErrorHandler
    For Each ctrl In Cbar.Controls
        If ctrl.Tag = sTag Then
            ctrl.Delete
            Exit Sub
        End If
    Next ctrl
    Exit Sub
ErrorHandler:
    Call WriteErrorLog("DeleteButton", False)
    Err.Clear
    Resume Next
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddComboBox - Create a ComboBox in the VBE for working with modules
'* Created    : 22-03-2023 14:29
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByVal sMenu As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddComboBox(ByVal sNameCommandBar As String, _
        ByVal sTag As String, _
        ByVal sText As String, _
        ByVal arrItem As Variant, _
        Optional ByVal bBeginGroup As Boolean = False)
    Dim combox      As CommandBarComboBox
    On Error GoTo ErrorHandler
    Set combox = Application.VBE.CommandBars(sNameCommandBar).Controls.Add(Type:=msoControlComboBox, Before:=1)
    Dim sVar        As Variant
    With combox
        .Tag = sTag
        For Each sVar In arrItem
            Call .AddItem(sVar)
        Next sVar
        .BeginGroup = bBeginGroup
        .text = sText
    End With
    Exit Sub
ErrorHandler:
    Call WriteErrorLog("AddComboBox", False)
    Err.Clear
    Resume Next
End Sub