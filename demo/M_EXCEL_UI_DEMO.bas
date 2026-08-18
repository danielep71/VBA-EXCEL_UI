Attribute VB_Name = "M_EXCEL_UI_DEMO"
'==============================================================================
'                           MODULE: EXCEL_UI_DEMO
'------------------------------------------------------------------------------
' PURPOSE
'   Provide a worksheet-based showcase for the EXCEL_UI component, including:
'     - selective SHOW / HIDE actions driven by worksheet check boxes
'     - explicit window-target scope selection for window-level UI elements
'     - current-state synchronization back into the check boxes
'     - selection helpers and preset profiles
'     - explicit capture / reset actions for the UI snapshot feature
'     - a reproducible demo-sheet builder
'
' PUBLIC SURFACE
'   - Demo_CreateDemoSheet
'   - Demo_ShowSelectedUI
'   - Demo_HideSelectedUI
'   - Demo_SyncCheckBoxesToUI
'   - Demo_SelectAllUI
'   - Demo_ClearAllUI
'   - Demo_PresetKiosk
'   - Demo_PresetAnalyst
'   - Demo_PresetMinimal
'   - Demo_CaptureUIState
'   - Demo_ResetUIToCapturedState
'
' DEMO SEMANTICS
'   - Checked   => selected for the next SHOW / HIDE action
'   - Unchecked => leave unchanged
'   - Target scope applies only to Headings, Workbook Tabs, and Gridlines
'   - Ribbon, Status Bar, Scroll Bars, Formula Bar, and Title Bar retain their
'     established application / main-window scope
'   - Synchronization reads window-level state from ActiveWindow because one
'     checkbox cannot represent potentially heterogeneous multi-window state
'   - A blank TargetScope on a pre-v1.1.0 demo sheet defaults to All Excel
'     windows so existing SHOW/HIDE buttons remain usable before sheet rebuild
'
' COMPATIBILITY
'   - Windows Excel only for Title Bar readback
'   - Supports Forms and ActiveX check boxes for demo state reads/writes
'   - Relies on the public EXCEL_UI API:
'       * UIVisibility
'       * UIWindowTargetScope
'       * UI_SetExcelUI
'       * UI_CaptureExcelUIState
'       * UI_ResetExcelUIToSnapshot
'       * UI_HasExcelUIStateSnapshot
'
' UPDATED
'   2026-08-18
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.0
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit

'------------------------------------------------------------------------------
' DEMO CONFIGURATION
'------------------------------------------------------------------------------
    Private Const DEMO_SHEET_NAME As String = "DEMO_UI"

    Private Const CB_RIBBON       As String = "chkRibbon"
    Private Const CB_STATUSBAR    As String = "chkStatusBar"
    Private Const CB_SCROLLBARS   As String = "chkScrollBars"
    Private Const CB_FORMULABAR   As String = "chkFormulaBar"
    Private Const CB_HEADINGS     As String = "chkHeadings"
    Private Const CB_WORKBOOKTABS As String = "chkWorkbookTabs"
    Private Const CB_GRIDLINES    As String = "chkGridlines"
    Private Const CB_TITLEBAR     As String = "chkTitleBar"

    Private Const TARGET_SCOPE_CELL As String = "H17"
    Private Const TARGET_SCOPE_HELPER_RANGE As String = "AA2:AA4"

    Private Const TARGET_SCOPE_ALL_TEXT As String = "All Excel windows"
    Private Const TARGET_SCOPE_ACTIVE_TEXT As String = "Active window"
    Private Const TARGET_SCOPE_WORKBOOK_TEXT As String = "Active workbook windows"

    Private Const NOTE_SCOPE_TEXT As String = "Scope / semantics note:" & vbLf & _
        "- Checked means SELECTED for the next SHOW or HIDE action." & vbLf & _
        "- WINDOW TARGET SCOPE affects Headings, Workbook Tabs, and Gridlines only." & vbLf & _
        "- Title Bar is a MAIN WINDOW FRAME setting and is never limited by TargetScope." & vbLf & _
        "- Ribbon, Status Bar, Scroll Bars, and Formula Bar remain application-level." & vbLf & _
        "- SYNC CHECKBOXES reads window-level state from ActiveWindow." & vbLf & _
        "- Presets change selections only; they do not change target scope or apply UI state."

    Private Const NOTE_RESTORE_TEXT As String = "Restore note:" & vbLf & _
        "UI_ShowExcelUI shows all managed UI elements; it does NOT restore a captured custom baseline." & vbLf & _
        "Use CAPTURE STATE and RESET STATE for explicit snapshot / restore."

'------------------------------------------------------------------------------
' WIN32 / WIN64 API FOR TITLE-BAR STATE READ
'------------------------------------------------------------------------------
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function Demo_GetWindowLongPtr Lib "user32" Alias _
            "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    #Else
        Private Declare PtrSafe Function Demo_GetWindowLong Lib "user32" Alias _
            "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    #End If

    Private Declare PtrSafe Function Demo_GetLastError Lib "kernel32" Alias _
        "GetLastError" () As Long

    Private Declare PtrSafe Sub Demo_SetLastError Lib "kernel32" Alias _
        "SetLastError" (ByVal dwErrCode As Long)
#Else
    Private Declare Function Demo_GetWindowLong Lib "user32" Alias _
        "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

    Private Declare Function Demo_GetLastError Lib "kernel32" Alias _
        "GetLastError" () As Long

    Private Declare Sub Demo_SetLastError Lib "kernel32" Alias _
        "SetLastError" (ByVal dwErrCode As Long)
#End If

'------------------------------------------------------------------------------
' API CONSTANTS FOR TITLE-BAR STATE READ
'------------------------------------------------------------------------------
    Private Const DEMO_GWL_STYLE  As Long = -16
    Private Const DEMO_WS_CAPTION As Long = &HC00000


'------------------------------------------------------------------------------
' PUBLIC DEMO BUILDER
'------------------------------------------------------------------------------

Public Sub Demo_CreateDemoSheet()
'
'==============================================================================
'                            Demo_CreateDemoSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Build or rebuild the Excel UI demo sheet and its interactive control panel.
'
' BEHAVIOR
'   - Rebuilds the demo template
'   - Creates application-level and window/main-frame selection check boxes
'   - Creates a validated target-scope selector defaulting to All Excel windows
'   - Adds SHOW/HIDE, preset, capture/reset, sync, and reset-sheet controls
'   - Writes v1.1.0 scope and restore guidance
'   - Synchronizes check boxes to the current Excel UI state
'
' ERROR POLICY
'   Raises after best-effort cleanup so builder defects remain visible.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Dim WB                  As Workbook
    Dim WS                  As Worksheet
    Dim ButtonSpecs         As Variant
    Dim FastModeState       As tDEMOFastModeState
    Dim FastModeOn          As Boolean
    Dim SavedErrNumber      As Long
    Dim SavedErrSource      As String
    Dim SavedErrDescription As String

    Const PROC As String = "Demo_CreateDemoSheet"

        On Error GoTo Clean_Fail

        Set WB = ThisWorkbook

        DEMO_Btn_Click
        DEMO_FastMode_Begin FastModeState
        FastModeOn = True
        Application.Cursor = xlWait

        DEMO_Sheet_BuildTemplate DEMO_SHEET_NAME, "EXCEL UI", "Demo Sheet", , , , _
            , , "C:H", , , , , , , , , , 31

        Set WS = WB.Worksheets(DEMO_SHEET_NAME)

'------------------------------------------------------------------------------
' APPLICATION-LEVEL UI STATE
'------------------------------------------------------------------------------
        DEMO_Prepare_LabeledInputSection WS, WS.Range("G4:H4"), _
            "APPLICATION LEVEL UI STATE", WS.Range("G5:G8"), WS.Range("H5:H8")

        DEMO_Write_NamedInputRow WB, WS, WS.Range("G5"), WS.Range("H5"), _
            "Ribbon", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G6"), WS.Range("H6"), _
            "Status bar", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G7"), WS.Range("H7"), _
            "Scroll bars", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G8"), WS.Range("H8"), _
            "Formula bar", ""

        DEMO_CB_AddForms WS, WS.Range("H5:H8"), Array(CB_RIBBON, _
            CB_STATUSBAR, CB_SCROLLBARS, CB_FORMULABAR)

'
'------------------------------------------------------------------------------
' WINDOW-LEVEL UI STATE
'------------------------------------------------------------------------------
        DEMO_Prepare_LabeledInputSection WS, WS.Range("G11:H11"), _
            "WINDOW LEVEL UI STATE", WS.Range("G12:G14"), WS.Range("H12:H14")

        DEMO_Write_NamedInputRow WB, WS, WS.Range("G12"), WS.Range("H12"), _
            "Headings", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G13"), WS.Range("H13"), _
            "Workbook tabs", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G14"), WS.Range("H14"), _
            "Gridlines", ""

        DEMO_CB_AddForms WS, WS.Range("H12:H14"), Array(CB_HEADINGS, _
            CB_WORKBOOKTABS, CB_GRIDLINES)

'------------------------------------------------------------------------------
' WINDOW TARGET SCOPE
'------------------------------------------------------------------------------
        Demo_BuildTargetScopeSelector WB, WS

'------------------------------------------------------------------------------
' MAIN WINDOW FRAME
'------------------------------------------------------------------------------
        DEMO_Prepare_LabeledInputSection WS, WS.Range("G19:H19"), _
            "MAIN WINDOW FRAME", WS.Range("G20"), WS.Range("H20")

        DEMO_Write_NamedInputRow WB, WS, WS.Range("G20"), WS.Range("H20"), _
            "Title bar", ""

        DEMO_CB_AddForms WS, WS.Range("H20"), Array(CB_TITLEBAR)

'------------------------------------------------------------------------------
' ACTION BUTTONS
'------------------------------------------------------------------------------
        DEMO_Write_BandHeader WS.Range("C4:E4"), "ACTIONS"

        ButtonSpecs = Array( _
            Array("btn_UI_Show", "SHOW SELECTED UI", "Demo_ShowSelectedUI"), _
            Array("btn_UI_Hide", "HIDE SELECTED UI", "Demo_HideSelectedUI"))

        DEMO_Btn_AddGrid WS, WS.Range("C5"), ButtonSpecs, 2, 150, 25
        DEMO_Set_RangeBorder WS.Range("C4:E6")

'------------------------------------------------------------------------------
' SELECT / CLEAR BUTTONS
'------------------------------------------------------------------------------
        DEMO_Write_BandHeader WS.Range("C8:E8"), "SELECT / CLEAR"

        ButtonSpecs = Array( _
            Array("btn_UI_SelectAll", "SELECT ALL", "Demo_SelectAllUI"), _
            Array("btn_UI_ClearAll", "CLEAR ALL", "Demo_ClearAllUI"))

        DEMO_Btn_AddGrid WS, WS.Range("C9"), ButtonSpecs, 2, 150, 25
        DEMO_Set_RangeBorder WS.Range("C8:E10")

'------------------------------------------------------------------------------
' PRESET BUTTONS
'------------------------------------------------------------------------------
        DEMO_Write_BandHeader WS.Range("C12:E12"), "PRESET SELECTION"

        ButtonSpecs = Array( _
            Array("btn_UI_Kiosk", "KIOSK", "Demo_PresetKiosk"), _
            Array("btn_UI_Analyst", "ANALYST", "Demo_PresetAnalyst"), _
            Array("btn_UI_Minimal", "MINIMAL", "Demo_PresetMinimal"))

        DEMO_Btn_AddGrid WS, WS.Range("C13"), ButtonSpecs, 2, 150, 25, , 13, , 8
        DEMO_Set_RangeBorder WS.Range("C12:E16")

'------------------------------------------------------------------------------
' CAPTURE / RESET STATE BUTTONS
'------------------------------------------------------------------------------
        DEMO_Write_BandHeader WS.Range("C18:E18"), "CAPTURE / RESET STATE"

        ButtonSpecs = Array( _
            Array("btn_UI_CaptureState", "CAPTURE STATE", "Demo_CaptureUIState"), _
            Array("btn_UI_ResetState", "RESET STATE", "Demo_ResetUIToCapturedState"))

        DEMO_Btn_AddGrid WS, WS.Range("C19"), ButtonSpecs, 2, 150, 25, , 13, , 8
        DEMO_Set_RangeBorder WS.Range("C18:E20")

'------------------------------------------------------------------------------
' SYNC BUTTON
'------------------------------------------------------------------------------
        DEMO_Btn_Add WS, "btn_UI_Sync", "SYNC CHECKBOXES", _
            WS.Range("G22").Left, WS.Range("G22").Top, WS.Range("G22:H22").Width, 25, _
            "Demo_SyncCheckBoxesToUI"

'------------------------------------------------------------------------------
' NOTES
'------------------------------------------------------------------------------
        With WS.Range("B24:H28")
            .Merge
            .Interior.Color = RGB(255, 242, 204)
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        DEMO_Set_RangeBorder WS.Range("B24:H28")

        With WS.Range("B29:H30")
            .Merge
            .Interior.Color = RGB(217, 225, 242)
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        DEMO_Set_RangeBorder WS.Range("B29:H30")

        WS.Range("B24").Value = NOTE_SCOPE_TEXT
        WS.Range("B29").Value = NOTE_RESTORE_TEXT

'------------------------------------------------------------------------------
' INITIAL SYNCHRONIZATION
'------------------------------------------------------------------------------
        Demo_SyncCheckBoxesToUI False

'------------------------------------------------------------------------------
' RESET SHEET BUTTON
'------------------------------------------------------------------------------
        DEMO_Btn_Add WS, "btn_UI_ResetSheet", "RESET SHEET", _
            WS.Range("H2").Left + 1, WS.Range("H2").Top + 1, 105, 25, _
            "Demo_CreateDemoSheet"

Clean_Exit:
        On Error Resume Next
        Application.Cursor = xlDefault

        If FastModeOn Then
            DEMO_FastMode_End FastModeState
        End If

        On Error GoTo 0

        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

        Exit Sub

Clean_Fail:
        SavedErrNumber = Err.Number
        SavedErrSource = IIf(Len(Err.Source) > 0, Err.Source, PROC)
        SavedErrDescription = Err.Description
        Resume Clean_Exit

End Sub


'------------------------------------------------------------------------------
' PUBLIC DEMO ACTIONS
'------------------------------------------------------------------------------

Public Sub Demo_ShowSelectedUI()
'
'==============================================================================
'                        Demo_ShowSelectedUI
'------------------------------------------------------------------------------
' PURPOSE
'   Show the selected UI elements using the target scope selected on the sheet.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        DEMO_Btn_PlayFeedback
        Demo_ApplySelectedUI UI_Show, "Demo_ShowSelectedUI"

End Sub


Public Sub Demo_HideSelectedUI()
'
'==============================================================================
'                        Demo_HideSelectedUI
'------------------------------------------------------------------------------
' PURPOSE
'   Hide the selected UI elements using the target scope selected on the sheet.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        DEMO_Btn_PlayFeedback
        Demo_ApplySelectedUI UI_Hide, "Demo_HideSelectedUI"

End Sub


Public Sub Demo_SyncCheckBoxesToUI(Optional ByVal PlayFeedback As Boolean = True)
'
'==============================================================================
'                     Demo_SyncCheckBoxesToUI
'------------------------------------------------------------------------------
' PURPOSE
'   Synchronize demo check boxes to the current Excel UI state.
'
' NOTES
'   Window-level values are read from ActiveWindow regardless of the selected
'   target scope. A single checkbox cannot represent heterogeneous states across
'   multiple windows.
'
' ERROR POLICY
'   Fail-soft; partial failures are written to the Immediate Window.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Dim WS        As Worksheet
    Dim ActiveWin As Window
    Dim IsVisible As Boolean
    Dim FailMsg   As String

    Const PROC As String = "Demo_SyncCheckBoxesToUI"

        On Error GoTo Fail

        If PlayFeedback Then
            DEMO_Btn_PlayFeedback
        End If

        Set WS = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)
        Set ActiveWin = Application.ActiveWindow

        If Demo_TryGetRibbonVisibility(IsVisible, FailMsg) Then
            If Not Demo_TrySetCheckBoxState(WS, CB_RIBBON, IsVisible, FailMsg) Then
                Demo_LogFailure PROC, CB_RIBBON, FailMsg
            End If
        Else
            Demo_LogFailure PROC, "RibbonState", FailMsg
        End If

        If Not Demo_TrySetCheckBoxState(WS, CB_STATUSBAR, _
            Application.DisplayStatusBar, FailMsg) Then
            Demo_LogFailure PROC, CB_STATUSBAR, FailMsg
        End If

        If Not Demo_TrySetCheckBoxState(WS, CB_SCROLLBARS, _
            Application.DisplayScrollBars, FailMsg) Then
            Demo_LogFailure PROC, CB_SCROLLBARS, FailMsg
        End If

        If Not Demo_TrySetCheckBoxState(WS, CB_FORMULABAR, _
            Application.DisplayFormulaBar, FailMsg) Then
            Demo_LogFailure PROC, CB_FORMULABAR, FailMsg
        End If

        If ActiveWin Is Nothing Then
            Demo_LogFailure PROC, "ActiveWindow", _
                "no active window available for window-level synchronization"
        Else
            If Not Demo_TrySetCheckBoxState(WS, CB_HEADINGS, _
                ActiveWin.DisplayHeadings, FailMsg) Then
                Demo_LogFailure PROC, CB_HEADINGS, FailMsg
            End If

            If Not Demo_TrySetCheckBoxState(WS, CB_WORKBOOKTABS, _
                ActiveWin.DisplayWorkbookTabs, FailMsg) Then
                Demo_LogFailure PROC, CB_WORKBOOKTABS, FailMsg
            End If

            If Not Demo_TrySetCheckBoxState(WS, CB_GRIDLINES, _
                ActiveWin.DisplayGridlines, FailMsg) Then
                Demo_LogFailure PROC, CB_GRIDLINES, FailMsg
            End If
        End If

        If Demo_TryGetTitleBarVisibility(IsVisible, FailMsg) Then
            If Not Demo_TrySetCheckBoxState(WS, CB_TITLEBAR, IsVisible, FailMsg) Then
                Demo_LogFailure PROC, CB_TITLEBAR, FailMsg
            End If
        Else
            Demo_LogFailure PROC, "TitleBarState", FailMsg
        End If

SafeExit:
        Exit Sub

Fail:
        Demo_LogFailure PROC, "Unexpected", Demo_GetRuntimeErrorText
        Resume SafeExit

End Sub


Public Sub Demo_SelectAllUI()
'
'==============================================================================
'                             Demo_SelectAllUI
'------------------------------------------------------------------------------
' PURPOSE
'   Select every managed UI element for the next SHOW or HIDE action.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        DEMO_Btn_PlayFeedback

        Demo_SetSelectionProfile CallerProc:="Demo_SelectAllUI", _
            RibbonSelected:=True, StatusBarSelected:=True, ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, HeadingsSelected:=True, _
            WorkbookTabsSelected:=True, GridlinesSelected:=True, _
            TitleBarSelected:=True

End Sub


Public Sub Demo_ClearAllUI()
'
'==============================================================================
'                             Demo_ClearAllUI
'------------------------------------------------------------------------------
' PURPOSE
'   Clear every selection for the next SHOW or HIDE action.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        DEMO_Btn_PlayFeedback

        Demo_SetSelectionProfile CallerProc:="Demo_ClearAllUI", _
            RibbonSelected:=False, StatusBarSelected:=False, _
            ScrollBarsSelected:=False, FormulaBarSelected:=False, _
            HeadingsSelected:=False, WorkbookTabsSelected:=False, _
            GridlinesSelected:=False, TitleBarSelected:=False

End Sub


Public Sub Demo_PresetKiosk()
'
'==============================================================================
'                            Demo_PresetKiosk
'------------------------------------------------------------------------------
' PURPOSE
'   Select a broad kiosk-style UI bundle. Target scope remains unchanged.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        DEMO_Btn_PlayFeedback

        Demo_SetSelectionProfile CallerProc:="Demo_PresetKiosk", _
            RibbonSelected:=True, StatusBarSelected:=True, ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, HeadingsSelected:=True, _
            WorkbookTabsSelected:=True, GridlinesSelected:=True, _
            TitleBarSelected:=True

End Sub


Public Sub Demo_PresetAnalyst()
'
'==============================================================================
'                           Demo_PresetAnalyst
'------------------------------------------------------------------------------
' PURPOSE
'   Select a worksheet-analysis-oriented UI bundle. Target scope is unchanged.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        DEMO_Btn_PlayFeedback

        Demo_SetSelectionProfile CallerProc:="Demo_PresetAnalyst", _
            RibbonSelected:=False, StatusBarSelected:=True, ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, HeadingsSelected:=True, _
            WorkbookTabsSelected:=True, GridlinesSelected:=True, _
            TitleBarSelected:=False

End Sub


Public Sub Demo_PresetMinimal()
'
'==============================================================================
'                           Demo_PresetMinimal
'------------------------------------------------------------------------------
' PURPOSE
'   Select a major-application-chrome bundle. Target scope remains unchanged.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        DEMO_Btn_PlayFeedback

        Demo_SetSelectionProfile CallerProc:="Demo_PresetMinimal", _
            RibbonSelected:=True, StatusBarSelected:=True, ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, HeadingsSelected:=False, _
            WorkbookTabsSelected:=False, GridlinesSelected:=False, _
            TitleBarSelected:=True

End Sub


Public Sub Demo_CaptureUIState(Optional ByVal ShowConfirmation As Boolean = True)
'
'==============================================================================
'                    Demo_CaptureUIState
'------------------------------------------------------------------------------
' PURPOSE
'   Capture the current managed Excel UI state through the public snapshot API.
'
' NOTES
'   Snapshot capture remains independent of the selective TargetScope setting.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Const PROC As String = "Demo_CaptureUIState"

        On Error GoTo Fail

        DEMO_Btn_PlayFeedback
        UI_CaptureExcelUIState

        If ShowConfirmation Then
            MsgBox "Current Excel UI state captured.", vbInformation, "Excel UI Demo"
        End If

SafeExit:
        Exit Sub

Fail:
        Demo_LogFailure PROC, "Unexpected", Demo_GetRuntimeErrorText
        Resume SafeExit

End Sub


Public Sub Demo_ResetUIToCapturedState()
'
'==============================================================================
'                    Demo_ResetUIToCapturedState
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the most recently captured managed Excel UI snapshot.
'
' NOTES
'   Snapshot restore remains independent of the selective TargetScope setting.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Const PROC As String = "Demo_ResetUIToCapturedState"

        On Error GoTo Fail

        DEMO_Btn_PlayFeedback

        If Not UI_HasExcelUIStateSnapshot Then
            MsgBox "No captured Excel UI state is available.", vbExclamation, _
                "Excel UI Demo"
            GoTo SafeExit
        End If

        UI_ResetExcelUIToSnapshot
        Demo_SyncCheckBoxesToUI False

SafeExit:
        Exit Sub

Fail:
        Demo_LogFailure PROC, "Unexpected", Demo_GetRuntimeErrorText
        Resume SafeExit

End Sub


'------------------------------------------------------------------------------
' TARGET-SCOPE UI
'------------------------------------------------------------------------------

Private Sub Demo_BuildTargetScopeSelector(ByVal WB As Workbook, ByVal WS As Worksheet)
'
'==============================================================================
'                    Demo_BuildTargetScopeSelector
'------------------------------------------------------------------------------
' PURPOSE
'   Build the v1.1.0 target-scope selector and its locale-independent helper list.
'
' ERROR POLICY
'   Raises normally to Demo_CreateDemoSheet.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        DEMO_Write_BandHeader WS.Range("G16:H16"), "WINDOW TARGET SCOPE"

        WS.Range("G17").Value = "Target scope"
        WS.Range(TARGET_SCOPE_CELL).Value = TARGET_SCOPE_ALL_TEXT

        DEMO_Format_Labels WS.Range("G17")
        DEMO_Format_InputCell WS.Range(TARGET_SCOPE_CELL)

        WS.Range(TARGET_SCOPE_HELPER_RANGE).NumberFormat = "@"
        WS.Range("AA2").Value = TARGET_SCOPE_ALL_TEXT
        WS.Range("AA3").Value = TARGET_SCOPE_ACTIVE_TEXT
        WS.Range("AA4").Value = TARGET_SCOPE_WORKBOOK_TEXT

        DEMO_Apply_ValidationList WS.Range(TARGET_SCOPE_CELL), "=$AA$2:$AA$4"
        DEMO_Hide_HelperColumns WS, "AA:AZ"

        DEMO_Set_RangeBorder WS.Range("G16:H17")

End Sub


Private Function Demo_TryResolveTargetScope(ByVal WS As Worksheet, _
    ByRef TargetScope As UIWindowTargetScope, ByRef FailMsg As String) As Boolean
'
'==============================================================================
'                    Demo_TryResolveTargetScope
'------------------------------------------------------------------------------
' PURPOSE
'   Map the demo dropdown text to the public UIWindowTargetScope enum.
'
' RETURNS
'   TRUE on a recognized value; FALSE with FailMsg otherwise.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Dim ScopeText As String

        On Error GoTo Fail

        Demo_TryResolveTargetScope = False
        TargetScope = UI_TargetAllExcelWindows
        FailMsg = vbNullString

        If WS Is Nothing Then
            FailMsg = "worksheet reference is Nothing"
            GoTo SafeExit
        End If

        ScopeText = Trim$(CStr(WS.Range(TARGET_SCOPE_CELL).Value2))

        'Backward-compatible upgrade behavior:
        'an old demo sheet has no TargetScope selector. Treat blank as the
        'established all-windows default so SHOW/HIDE continues to work before
        'the sheet is rebuilt with Demo_CreateDemoSheet.
        If Len(ScopeText) = 0 Then
            TargetScope = UI_TargetAllExcelWindows
            Demo_TryResolveTargetScope = True
            GoTo SafeExit
        End If

        Select Case StrComp(ScopeText, TARGET_SCOPE_ALL_TEXT, vbTextCompare)
            Case 0
                TargetScope = UI_TargetAllExcelWindows
                Demo_TryResolveTargetScope = True
                GoTo SafeExit
        End Select

        Select Case StrComp(ScopeText, TARGET_SCOPE_ACTIVE_TEXT, vbTextCompare)
            Case 0
                TargetScope = UI_TargetActiveWindow
                Demo_TryResolveTargetScope = True
                GoTo SafeExit
        End Select

        Select Case StrComp(ScopeText, TARGET_SCOPE_WORKBOOK_TEXT, vbTextCompare)
            Case 0
                TargetScope = UI_TargetActiveWorkbookWindows
                Demo_TryResolveTargetScope = True
                GoTo SafeExit
        End Select

        FailMsg = "unsupported target-scope value: '" & ScopeText & "'"

SafeExit:
        Exit Function

Fail:
        FailMsg = Demo_GetRuntimeErrorText

End Function


'------------------------------------------------------------------------------
' SHARED PRIVATE ORCHESTRATION
'------------------------------------------------------------------------------

Private Sub Demo_ApplySelectedUI(ByVal SelectedVisibility As UIVisibility, _
    ByVal CallerProc As String)
'
'==============================================================================
'                        Demo_ApplySelectedUI
'------------------------------------------------------------------------------
' PURPOSE
'   Apply SHOW or HIDE to the selected UI elements under the selected target scope.
'
' BEHAVIOR
'   - Checked controls map to SelectedVisibility
'   - Unchecked controls map to UI_LeaveUnchanged
'   - TargetScope applies only to window-level production operations
'   - Empty selection is rejected with an informational message
'   - Invalid demo scope text is rejected instead of silently widening scope
'
' ERROR POLICY
'   Fail-soft; unexpected failures are logged to the Immediate Window.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Dim WS              As Worksheet
    Dim RibbonVis       As UIVisibility
    Dim StatusBarVis    As UIVisibility
    Dim ScrollBarsVis   As UIVisibility
    Dim FormulaBarVis   As UIVisibility
    Dim HeadingsVis     As UIVisibility
    Dim WorkbookTabsVis As UIVisibility
    Dim GridlinesVis    As UIVisibility
    Dim TitleBarVis     As UIVisibility
    Dim TargetScope     As UIWindowTargetScope
    Dim FailMsg         As String

        On Error GoTo Fail

        Set WS = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)

        RibbonVis = Demo_CheckBoxToUIVisibility(WS, CB_RIBBON, _
            SelectedVisibility, CallerProc)
        StatusBarVis = Demo_CheckBoxToUIVisibility(WS, CB_STATUSBAR, _
            SelectedVisibility, CallerProc)
        ScrollBarsVis = Demo_CheckBoxToUIVisibility(WS, CB_SCROLLBARS, _
            SelectedVisibility, CallerProc)
        FormulaBarVis = Demo_CheckBoxToUIVisibility(WS, CB_FORMULABAR, _
            SelectedVisibility, CallerProc)
        HeadingsVis = Demo_CheckBoxToUIVisibility(WS, CB_HEADINGS, _
            SelectedVisibility, CallerProc)
        WorkbookTabsVis = Demo_CheckBoxToUIVisibility(WS, CB_WORKBOOKTABS, _
            SelectedVisibility, CallerProc)
        GridlinesVis = Demo_CheckBoxToUIVisibility(WS, CB_GRIDLINES, _
            SelectedVisibility, CallerProc)
        TitleBarVis = Demo_CheckBoxToUIVisibility(WS, CB_TITLEBAR, _
            SelectedVisibility, CallerProc)

        If Not Demo_HasAnySelectedChange(RibbonVis, StatusBarVis, ScrollBarsVis, _
            FormulaBarVis, HeadingsVis, WorkbookTabsVis, GridlinesVis, TitleBarVis) _
            Then
            MsgBox "No UI elements are selected.", vbInformation, "Excel UI Demo"
            GoTo SafeExit
        End If

        If Not Demo_TryResolveTargetScope(WS, TargetScope, FailMsg) Then
            Demo_LogFailure CallerProc, "TargetScope", FailMsg
            MsgBox "Select a valid window target scope before applying UI state.", _
                vbExclamation, "Excel UI Demo"
            GoTo SafeExit
        End If

        UI_SetExcelUI Ribbon:=RibbonVis, StatusBar:=StatusBarVis, _
            ScrollBars:=ScrollBarsVis, FormulaBar:=FormulaBarVis, _
            Headings:=HeadingsVis, WorkbookTabs:=WorkbookTabsVis, _
            Gridlines:=GridlinesVis, TitleBar:=TitleBarVis, _
            TargetScope:=TargetScope

SafeExit:
        Exit Sub

Fail:
        Demo_LogFailure CallerProc, "Unexpected", Demo_GetRuntimeErrorText
        Resume SafeExit

End Sub


Private Sub Demo_SetSelectionProfile(ByVal CallerProc As String, _
    ByVal RibbonSelected As Boolean, ByVal StatusBarSelected As Boolean, _
    ByVal ScrollBarsSelected As Boolean, ByVal FormulaBarSelected As Boolean, _
    ByVal HeadingsSelected As Boolean, ByVal WorkbookTabsSelected As Boolean, _
    ByVal GridlinesSelected As Boolean, ByVal TitleBarSelected As Boolean)
'
'==============================================================================
'                         Demo_SetSelectionProfile
'------------------------------------------------------------------------------
' PURPOSE
'   Set all demo check boxes according to one Boolean selection profile.
'
' NOTES
'   Target scope is intentionally not changed by selection profiles.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Dim WS      As Worksheet
    Dim FailMsg As String

        On Error GoTo Fail

        Set WS = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)

        If Not Demo_TrySetCheckBoxState(WS, CB_RIBBON, RibbonSelected, FailMsg) Then
            Demo_LogFailure CallerProc, CB_RIBBON, FailMsg
        End If
        If Not Demo_TrySetCheckBoxState(WS, CB_STATUSBAR, StatusBarSelected, FailMsg) Then
            Demo_LogFailure CallerProc, CB_STATUSBAR, FailMsg
        End If
        If Not Demo_TrySetCheckBoxState(WS, CB_SCROLLBARS, ScrollBarsSelected, FailMsg) Then
            Demo_LogFailure CallerProc, CB_SCROLLBARS, FailMsg
        End If
        If Not Demo_TrySetCheckBoxState(WS, CB_FORMULABAR, FormulaBarSelected, FailMsg) Then
            Demo_LogFailure CallerProc, CB_FORMULABAR, FailMsg
        End If
        If Not Demo_TrySetCheckBoxState(WS, CB_HEADINGS, HeadingsSelected, FailMsg) Then
            Demo_LogFailure CallerProc, CB_HEADINGS, FailMsg
        End If
        If Not Demo_TrySetCheckBoxState(WS, CB_WORKBOOKTABS, _
            WorkbookTabsSelected, FailMsg) Then
            Demo_LogFailure CallerProc, CB_WORKBOOKTABS, FailMsg
        End If
        If Not Demo_TrySetCheckBoxState(WS, CB_GRIDLINES, GridlinesSelected, FailMsg) Then
            Demo_LogFailure CallerProc, CB_GRIDLINES, FailMsg
        End If
        If Not Demo_TrySetCheckBoxState(WS, CB_TITLEBAR, TitleBarSelected, FailMsg) Then
            Demo_LogFailure CallerProc, CB_TITLEBAR, FailMsg
        End If

SafeExit:
        Exit Sub

Fail:
        Demo_LogFailure CallerProc, "Unexpected", Demo_GetRuntimeErrorText
        Resume SafeExit

End Sub


Private Function Demo_HasAnySelectedChange(ByVal RibbonVis As UIVisibility, _
    ByVal StatusBarVis As UIVisibility, ByVal ScrollBarsVis As UIVisibility, _
    ByVal FormulaBarVis As UIVisibility, ByVal HeadingsVis As UIVisibility, _
    ByVal WorkbookTabsVis As UIVisibility, ByVal GridlinesVis As UIVisibility, _
    ByVal TitleBarVis As UIVisibility) As Boolean
'
'==============================================================================
'                        Demo_HasAnySelectedChange
'------------------------------------------------------------------------------
' PURPOSE
'   Return TRUE when at least one requested UI element is not LeaveUnchanged.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        Demo_HasAnySelectedChange = (RibbonVis <> UI_LeaveUnchanged Or _
            StatusBarVis <> UI_LeaveUnchanged Or ScrollBarsVis <> UI_LeaveUnchanged Or _
            FormulaBarVis <> UI_LeaveUnchanged Or HeadingsVis <> UI_LeaveUnchanged Or _
            WorkbookTabsVis <> UI_LeaveUnchanged Or GridlinesVis <> UI_LeaveUnchanged Or _
            TitleBarVis <> UI_LeaveUnchanged)

End Function


'------------------------------------------------------------------------------
' CHECKBOX / STATE TRANSLATION HELPERS
'------------------------------------------------------------------------------

Private Function Demo_CheckBoxToUIVisibility(ByVal WS As Worksheet, _
    ByVal CheckBoxName As String, ByVal SelectedVisibility As UIVisibility, _
    ByVal CallerProc As String) As UIVisibility
'
'==============================================================================
'                        Demo_CheckBoxToUIVisibility
'------------------------------------------------------------------------------
' PURPOSE
'   Map checked => selected visibility and unchecked/unavailable => unchanged.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Dim IsChecked As Boolean
    Dim FailMsg   As String

        Demo_CheckBoxToUIVisibility = UI_LeaveUnchanged

        If Not Demo_TryGetCheckBoxState(WS, CheckBoxName, IsChecked, FailMsg) Then
            Demo_LogFailure CallerProc, CheckBoxName, FailMsg
            Exit Function
        End If

        If IsChecked Then
            Demo_CheckBoxToUIVisibility = SelectedVisibility
        End If

End Function


Private Function Demo_TryGetCheckBoxState(ByVal WS As Worksheet, _
    ByVal ControlName As String, ByRef IsChecked As Boolean, _
    ByRef FailMsg As String) As Boolean
'
'==============================================================================
'                        Demo_TryGetCheckBoxState
'------------------------------------------------------------------------------
' PURPOSE
'   Read a Forms or ActiveX check-box state without raising to the caller.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Dim Shp         As Shape
    Dim CheckBoxOle As OLEObject
    Dim ValueOut    As Variant

        On Error GoTo Fail

        Demo_TryGetCheckBoxState = False
        IsChecked = False
        FailMsg = vbNullString

        If WS Is Nothing Then
            FailMsg = "worksheet reference is Nothing"
            GoTo SafeExit
        End If

        If Len(Trim$(ControlName)) = 0 Then
            FailMsg = "control name is blank"
            GoTo SafeExit
        End If

        On Error Resume Next
            Set Shp = WS.Shapes(ControlName)
        On Error GoTo Fail

        If Not Shp Is Nothing Then
            If Shp.Type <> msoFormControl Then
                FailMsg = "shape exists but is not a Forms control"
                GoTo SafeExit
            End If

            If Shp.FormControlType <> xlCheckBox Then
                FailMsg = "Forms control exists but is not a CheckBox"
                GoTo SafeExit
            End If

            IsChecked = (Shp.ControlFormat.Value = xlOn)
            Demo_TryGetCheckBoxState = True
            GoTo SafeExit
        End If

        On Error Resume Next
            Set CheckBoxOle = WS.OLEObjects(ControlName)
        On Error GoTo Fail

        If Not CheckBoxOle Is Nothing Then
            If InStr(1, CheckBoxOle.progID, "CheckBox", vbTextCompare) = 0 Then
                FailMsg = "ActiveX control exists but is not a CheckBox"
                GoTo SafeExit
            End If

            ValueOut = CallByName(CheckBoxOle.Object, "Value", VbGet)
            IsChecked = CBool(ValueOut)
            Demo_TryGetCheckBoxState = True
            GoTo SafeExit
        End If

        FailMsg = "check box not found"

SafeExit:
        Exit Function

Fail:
        FailMsg = Demo_GetRuntimeErrorText

End Function


Private Function Demo_TrySetCheckBoxState(ByVal WS As Worksheet, _
    ByVal ControlName As String, ByVal IsChecked As Boolean, _
    ByRef FailMsg As String) As Boolean
'
'==============================================================================
'                        Demo_TrySetCheckBoxState
'------------------------------------------------------------------------------
' PURPOSE
'   Write a Forms or ActiveX check-box state without raising to the caller.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Dim Shp         As Shape
    Dim CheckBoxOle As OLEObject

        On Error GoTo Fail

        Demo_TrySetCheckBoxState = False
        FailMsg = vbNullString

        If WS Is Nothing Then
            FailMsg = "worksheet reference is Nothing"
            GoTo SafeExit
        End If

        If Len(Trim$(ControlName)) = 0 Then
            FailMsg = "control name is blank"
            GoTo SafeExit
        End If

        On Error Resume Next
            Set Shp = WS.Shapes(ControlName)
        On Error GoTo Fail

        If Not Shp Is Nothing Then
            If Shp.Type <> msoFormControl Then
                FailMsg = "shape exists but is not a Forms control"
                GoTo SafeExit
            End If

            If Shp.FormControlType <> xlCheckBox Then
                FailMsg = "Forms control exists but is not a CheckBox"
                GoTo SafeExit
            End If

            If IsChecked Then
                Shp.ControlFormat.Value = xlOn
            Else
                Shp.ControlFormat.Value = xlOff
            End If

            Demo_TrySetCheckBoxState = True
            GoTo SafeExit
        End If

        On Error Resume Next
            Set CheckBoxOle = WS.OLEObjects(ControlName)
        On Error GoTo Fail

        If Not CheckBoxOle Is Nothing Then
            If InStr(1, CheckBoxOle.progID, "CheckBox", vbTextCompare) = 0 Then
                FailMsg = "ActiveX control exists but is not a CheckBox"
                GoTo SafeExit
            End If

            CallByName CheckBoxOle.Object, "Value", VbLet, IsChecked
            Demo_TrySetCheckBoxState = True
            GoTo SafeExit
        End If

        FailMsg = "check box not found"

SafeExit:
        Exit Function

Fail:
        FailMsg = Demo_GetRuntimeErrorText

End Function


'------------------------------------------------------------------------------
' UI-STATE READ HELPERS
'------------------------------------------------------------------------------

Private Function Demo_TryGetRibbonVisibility(ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean
'
'==============================================================================
'                         Demo_TryGetRibbonVisibility
'------------------------------------------------------------------------------
' PURPOSE
'   Read current Ribbon visibility using CommandBars with an Excel4 fallback.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
    Dim ValueOut As Variant

        On Error GoTo Fail

        Demo_TryGetRibbonVisibility = False
        IsVisible = False
        FailMsg = vbNullString

        On Error Resume Next
            IsVisible = Application.CommandBars("Ribbon").Visible
        If Err.Number = 0 Then
            On Error GoTo Fail
            Demo_TryGetRibbonVisibility = True
            GoTo SafeExit
        End If

        Err.Clear
        On Error GoTo Fail

        On Error Resume Next
            ValueOut = Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")
        If Err.Number = 0 Then
            On Error GoTo Fail
            IsVisible = CBool(ValueOut)
            Demo_TryGetRibbonVisibility = True
            GoTo SafeExit
        End If

        FailMsg = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        On Error GoTo Fail

SafeExit:
        Exit Function

Fail:
        FailMsg = Demo_GetRuntimeErrorText

End Function


Private Function Demo_TryGetTitleBarVisibility(ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean
'
'==============================================================================
'                        Demo_TryGetTitleBarVisibility
'------------------------------------------------------------------------------
' PURPOSE
'   Read title-bar visibility from Application.Hwnd using the caption style bit.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
#If VBA7 Then
    Dim xlHnd      As LongPtr
    Dim StyleValue As LongPtr
#Else
    Dim xlHnd      As Long
    Dim StyleValue As Long
#End If
    Dim LastErr As Long

        On Error GoTo Fail

        Demo_TryGetTitleBarVisibility = False
        IsVisible = False
        FailMsg = vbNullString
        xlHnd = Application.hWnd

        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo SafeExit
        End If

        Demo_SetLastError 0

#If VBA7 Then
    #If Win64 Then
        StyleValue = Demo_GetWindowLongPtr(xlHnd, DEMO_GWL_STYLE)
    #Else
        StyleValue = Demo_GetWindowLong(xlHnd, DEMO_GWL_STYLE)
    #End If
#Else
        StyleValue = Demo_GetWindowLong(xlHnd, DEMO_GWL_STYLE)
#End If

        LastErr = Demo_GetLastError

        If StyleValue = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)
            GoTo SafeExit
        End If

        IsVisible = ((StyleValue And DEMO_WS_CAPTION) <> 0)
        Demo_TryGetTitleBarVisibility = True

SafeExit:
        Exit Function

Fail:
        FailMsg = Demo_GetRuntimeErrorText

End Function


'------------------------------------------------------------------------------
' DIAGNOSTICS
'------------------------------------------------------------------------------

Private Function Demo_GetRuntimeErrorText() As String
'
'==============================================================================
'                      Demo_GetRuntimeErrorText
'------------------------------------------------------------------------------
' PURPOSE
'   Format the active Err object without raising.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        On Error Resume Next

        Demo_GetRuntimeErrorText = CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

End Function


Private Sub Demo_LogFailure(ByVal ProcName As String, ByVal Stage As String, _
    ByVal Detail As String)
'
'==============================================================================
'                            Demo_LogFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Write one fail-soft diagnostic line to the Immediate Window.
'
' UPDATED
'   2026-08-18
'==============================================================================
'
        On Error Resume Next
        Debug.Print ProcName & " failed @ " & Stage & " | " & Detail

End Sub
