Attribute VB_Name = "modAddinCreateMenus"
Option Explicit
Option Private Module
Option Compare Text
Option Base 1
Private ToolContextEventHandlers As Collection
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : RefreshMenu - Refreshes the add-in menu in the VBE main menu
'* Created    : 22-03-2023 14:36
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub RefreshMenu()
    Call DeleteContextMenus
    Call AddContextMenus
    Call MsgBox("Refresh of " & modAddinConst.NAME_ADDIN & " add-in completed!", vbInformation, "Refresh " & modAddinConst.NAME_ADDIN & ":")
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : Auto_Open - Runs when the add-in is loaded in the VBE environment
'* Created    : 22-03-2023 14:27
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Auto_Open()
    'If VBAIsTrusted And ThisWorkbook.Name = modAddinConst.NAME_ADDIN & ".xlam" Then
    Call AddContextMenus
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : Auto_Close - Runs when the add-in is unloaded from the VBE environment, removing menus
'* Created    : 22-03-2023 14:31
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Auto_Close()
    Call DeleteContextMenus
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddContextMenus - Adds menus to the VBE environment
'* Created    : 22-03-2023 14:27
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddContextMenus()

    Call AddNewCommandBarMenu(modAddinConst.MENU_MOVE_CONTROLS)
    Call AddButtom(984, "Help Move Controls", "HelpMoveControl", modAddinConst.MENU_MOVE_CONTROLS, vbNullString, False, True)
    Call AddButtom(38, "Move Up", "MoveControlUp", modAddinConst.MENU_MOVE_CONTROLS, vbNullString)
    Call AddButtom(40, "Move Down", "MoveControlDown", modAddinConst.MENU_MOVE_CONTROLS, vbNullString, False, True)
    Call AddButtom(39, "Move Left", "MoveControlLeft", modAddinConst.MENU_MOVE_CONTROLS, vbNullString)
    Call AddButtom(41, "Move Right", "MoveControlRight", modAddinConst.MENU_MOVE_CONTROLS, vbNullString)
    Call AddComboBox(modAddinConst.MENU_MOVE_CONTROLS, modAddinConst.MENU_MOVE_CONTROLS, "Control", Array("Control", "Top Left", "Bottom Right"))

    Call AddNewCommandBarMenu(modAddinConst.MENU_TOOLS)
    Call AddButtom(107, "Option's Explicit and Private Module", "insertOptionsExplicitAndPrivateModule", modAddinConst.MENU_TOOLS, vbNullString, False, False)
    Call AddButtom(0, "Option's", "subOptionsForm", modAddinConst.MENU_TOOLS, vbNullString, True, True)
    Call AddButtom(984, "Help Main Add-in", "HelpMainAddin", modAddinConst.MENU_TOOLS, vbNullString, False, True)
    Call AddButtom(7460, "Multiline Dim's", "dimMultiLine", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(7772, "Single Line Dim's", "dimOneLine", modAddinConst.MENU_TOOLS, vbNullString, False, True)
    Call AddButtom(7770, "Disable 'Debug.print'", "debugOff", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(7771, "Enable Debug.print", "debugOn", modAddinConst.MENU_TOOLS, vbNullString, False, True)
    Call AddButtom(699, "Remove Double Empty Lines", "delTwoEmptyStrings", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(3917, "Cut Tab", "CutTab", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(3919, "Rebuild", "ReBild", modAddinConst.MENU_TOOLS, vbNullString, False, True)
    Call AddButtom(12, "Remove Line Numbers", "RemoveLineNumbersPublic", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddButtom(11, "Add Line Numbers", "AddLineNumbers_", modAddinConst.MENU_TOOLS, vbNullString)
    Call AddComboBox(modAddinConst.MENU_TOOLS, modAddinConst.MENU_TOOLS, modAddinConst.TYPE_SELECTED_MODULE, Array(modAddinConst.TYPE_ALL_VBAPROJECT, modAddinConst.TYPE_SELECTED_MODULE))


    Call AddButtom(9634, "Swap Assignment [=]", "SwapEgual", modAddinConst.MENU_CODE_WINDOW, "SwapEgual", True, False)
    Call AddButtom(0, "UPPER Case", "toUpperCase", modAddinConst.MENU_CODE_WINDOW, "toUpperCase", True, False)
    Call AddButtom(0, "lower Case", "toLowerCase", modAddinConst.MENU_CODE_WINDOW, "toLowerCase", True, False)
    Call AddButtom(22, "Insert Code", "InsertCode", modAddinConst.MENU_CODE_WINDOW, "InsertCode", True, False)

    Call AddButtom(1650, "Align Horiz", "vbaCntAlingHoriz", modAddinConst.MENU_FORMS, "Align Horiz", True)
    Call AddButtom(1653, "Align Vert", "vbaCntAlingVert", modAddinConst.MENU_FORMS, "Align Vert", True)
    Call AddButtom(162, "Rename Control", "RenameControl", modAddinConst.MENU_FORMS, "Rename Control", True)
    Call AddButtom(22, "Paste Style", "PasteStyleControl", modAddinConst.MENU_FORMS, "Paste Style", True)
    Call AddButtom(1076, "Copy Style", "CopyStyleControl", modAddinConst.MENU_FORMS, "Copy Style", True)
    Call AddButtom(0, "UPPER CASE", "UperTextInControl", modAddinConst.MENU_FORMS, "UPPER CASE", True, False)
    Call AddButtom(0, "lower case", "LowerTextInControl", modAddinConst.MENU_FORMS, "lower case", True, False)

    Call AddButtom(2045, "Copy Module", "CopyModyleVBE", modAddinConst.MENU_PROJECT_WINDOW, "Copy Module", True, False)

    Call AddButtom(22, "Paste Style", "PasteStyleForms", modAddinConst.MENU_MS_FORMS, "Paste Style", True)
    Call AddButtom(1076, "Copy Style", "CopyStyleControl", modAddinConst.MENU_MS_FORMS, "Copy Style", True)
    Call AddButtom(0, "UPPER CASE", "UperTextInForm", modAddinConst.MENU_MS_FORMS, "UPPER CASE", True, False)
    Call AddButtom(0, "lower case", "LowerTextInForm", modAddinConst.MENU_MS_FORMS, "lower case", True, False)
End Sub


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddNewCommandBarMenu - Creates a new menu in the VBE environment
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
            .RowIndex = 3
        End With
    End If
    Set AddNewCommandBarMenu = myCommandBar
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : DeleteContextMenus - Removes all added menus from the VBE environment
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
    Call DeleteButton("InsertCode", modAddinConst.MENU_CODE_WINDOW)

    Call DeleteButton("Align Horiz", modAddinConst.MENU_FORMS)
    Call DeleteButton("Align Vert", modAddinConst.MENU_FORMS)
    Call DeleteButton("Rename Control", modAddinConst.MENU_FORMS)
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
'* Sub        : AddButtom - Adds a button to the VBE environment
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
'* Optional ByVal BeginGroup As Boolean = False        :
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
        Optional ByVal BeginGroup As Boolean = False, _
        Optional ByVal ShortcutText As String = vbNullString, _
        Optional ByVal Before As Byte = 1)
    Dim btn         As CommandBarButton
    Dim evtContextMenu As clsVBECommandHandler
    Set btn = Application.VBE.CommandBars(sNameCommandBar).Controls.Add(Type:=msoControlButton, Before:=Before)
    With btn
        .FaceId = Face
        If VisibleCapiton Then .Caption = Capitan
        .TooltipText = Capitan
        .Tag = sTag
        .OnAction = "'" & ThisWorkbook.Name & "'!" & sOnAction
        .Style = msoButtonIconAndCaption
        .BeginGroup = BeginGroup
        .ShortcutText = ShortcutText
    End With
    Set evtContextMenu = New clsVBECommandHandler
    Set evtContextMenu.EvtHandler = btn
    If ToolContextEventHandlers Is Nothing Then Set ToolContextEventHandlers = New Collection
    Call ToolContextEventHandlers.Add(evtContextMenu)
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : DeleteButton - Removes a button from the VBE environment
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
    Dim Ctrl        As CommandBarControl
    On Error Resume Next
    Set Cbar = Application.VBE.CommandBars(sMenu)
    If Cbar Is Nothing Then Exit Sub
    'On Error GoTo ErrorHandler
    For Each Ctrl In Cbar.Controls
        If Ctrl.Tag = sTag Then
            Ctrl.Delete
            Exit Sub
        End If
    Next Ctrl
    Exit Sub
ErrorHandler:
    Debug.Print "Error in DeleteButton" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line " & Erl
    'Call WriteErrorLog("DeleteButton")
    Err.Clear
    Resume Next
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : AddComboBox - Adds a ComboBox to the VBE environment for selection purposes
'* Created    : 22-03-2023 14:29
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):             Description
'*
'* ByVal sMenu As String :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddComboBox(ByVal sNameCommandBar As String, ByVal sTag As String, ByVal sText As String, ByVal arrItem As Variant)
    Dim combox      As CommandBarComboBox
    Set combox = Application.VBE.CommandBars(sNameCommandBar).Controls.Add(Type:=msoControlComboBox, Before:=1)
    Dim sVar        As Variant
    With combox
        .Tag = sTag
        For Each sVar In arrItem
            Call .AddItem(sVar)
        Next sVar
        .Text = sText
    End With
End Sub
