Attribute VB_Name = "M_EXCEL_UI_DEMO"

Option Explicit

'
'==============================================================================
' MODULE: M_EXCEL_UI_DEMO
'==============================================================================
' PURPOSE
'   Provide a worksheet-based showcase for the EXCEL_UI module, including:
'     - a reproducible demo-sheet builder
'     - selective SHOW / HIDE actions driven by worksheet check boxes
'     - current-state synchronization back into the check boxes
'     - selection helpers and preset profiles
'     - explicit capture / reset actions for the UI snapshot feature
'     - explanatory notes rendered on the demo sheet
'
' WHY THIS EXISTS
'   A demo workbook is easier to understand, present, and test when users can
'   interact with the EXCEL_UI module through worksheet controls rather than by
'   editing VBA calls directly
'
'   This module bridges the demo sheet controls to the public EXCEL_UI API and
'   builds the demo surface in a repeatable, presentation-ready way
'
' NOTES
'   - Window-level sync reads the current ActiveWindow state
'   - TitleBar SHOW / HIDE acts on the active window's frame: the handle is
'     read from ActiveWindow.hWnd and used only while Application.Hwnd names
'     the same frame
'   - TitleBar sync reads the Excel window represented by Application.Hwnd
'   - Missing or misnamed controls are logged to the Immediate Window and
'     treated conservatively
'   - Demo_CreateDemoSheet is the canonical builder
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
' EXPECTED DEMO CHECK BOX NAMES
'   Application-level
'     - chkStatusBar
'     - chkScrollBars
'     - chkFormulaBar
'
'   Active-window
'     - chkRibbon
'     - chkTitleBar
'
'   Window-level
'     - chkHeadings
'     - chkWorkbookTabs
'     - chkGridlines
'
' EXPECTED DEMO SHEET
'   - Worksheet name: DEMO_UI
'
' DEMO SEMANTICS
'   - Checked   => selected for the next SHOW / HIDE action
'   - Unchecked => leave unchanged
'
'   The sync action reads the current Excel UI state and marks currently
'   visible elements as checked for convenience and reference
'
' COMPATIBILITY
'   - Supports both Forms check boxes and ActiveX check boxes when reading or
'     writing control state, through DEMO_CB_TryGetState and
'     DEMO_CB_TrySetState of M_DEMO_BUILDER; a mixed check box is treated as
'     unavailable and leaves its UI element unchanged
'   - Relies on the public API exposed by the EXCEL_UI module:
'       * UIVisibility
'       * UI_SetExcelUI
'       * UI_CaptureExcelUIState
'       * UI_ResetExcelUIToSnapshot
'       * UI_HasExcelUIStateSnapshot
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.2
'
' HISTORY
'   2026-11-18 - The title-bar read tests the complete WS_CAPTION mask, so a
'                style with only WS_BORDER or only WS_DLGFRAME is not reported
'                as a visible title bar (#6).
'   2026-11-15 - The title-bar read captures Err.LastDllError right after
'                the WinAPI call; GetLastError is no longer declared (#144).
'   2026-11-14 - Check-box reads and writes moved to M_DEMO_BUILDER, which
'                reaches ActiveX check boxes (#73).
'   2026-11-14 - Demo_CreateDemoSheet keeps fast-mode restoration failures and
'                the entry cursor (#70).
'   2026-11-06 - Header and comments aligned with the code (#114).
'   2026-08-21 - Captured the Err fields before suppressing errors, so the
'                runtime-error text describes the error it was given.
'   2026-08-18 - Reformatted to the project house style. No behavior change.
'
' UPDATED
'   2026-11-18
'==============================================================================

'------------------------------------------------------------------------------
' DEMO CONFIGURATION
'------------------------------------------------------------------------------
    Private Const DEMO_SHEET_NAME   As String = "DEMO_UI"   'Demo worksheet name

    Private Const CB_RIBBON         As String = "chkRibbon"         'Ribbon check box name
    Private Const CB_STATUSBAR      As String = "chkStatusBar"      'StatusBar check box name
    Private Const CB_SCROLLBARS     As String = "chkScrollBars"     'ScrollBars check box name
    Private Const CB_FORMULABAR     As String = "chkFormulaBar"     'FormulaBar check box name
    Private Const CB_HEADINGS       As String = "chkHeadings"       'Headings check box name
    Private Const CB_WORKBOOKTABS   As String = "chkWorkbookTabs"   'WorkbookTabs check box name
    Private Const CB_GRIDLINES      As String = "chkGridlines"      'Gridlines check box name
    Private Const CB_TITLEBAR       As String = "chkTitleBar"       'TitleBar check box name

    Private Const NOTE_SCOPE_TEXT As String = "Scope / semantics note:" & vbLf & _
        "- Checked means SELECTED for the next SHOW or HIDE action." & vbLf & _
        "- Application-level items affect the current Excel instance." & vbLf & _
        "- Window-level sync reads ActiveWindow; apply actions target each open Excel window." _
        & vbLf & _
        "- TitleBar is Windows-only and uses WinAPI against Application.Hwnd." & _
        vbLf & _
        "- Preset buttons only set selections; they do not apply SHOW or HIDE by themselves."

    Private Const NOTE_RESTORE_TEXT As String = "Restore note:" & vbLf & _
        "UI_ShowExcelUI shows all managed UI elements. It does NOT restore a previously captured user-specific UI state." _
        & vbLf & _
        "Use CAPTURE STATE and RESET STATE for explicit snapshot / restore."

'------------------------------------------------------------------------------
' WIN32 / WIN64 API FOR TITLE-BAR STATE READ
'------------------------------------------------------------------------------
    #If VBA7 Then
        #If Win64 Then
            Private Declare PtrSafe Function Demo_GetWindowLongPtr Lib "user32" _
                Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As _
                Long) As LongPtr
        #Else
            Private Declare PtrSafe Function Demo_GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) _
                As Long
        #End If
        Private Declare PtrSafe Sub Demo_SetLastError Lib "kernel32" Alias _
            "SetLastError" (ByVal dwErrCode As Long)
    #Else
        Private Declare Function Demo_GetWindowLong Lib "user32" Alias _
            "GetWindowLongA" ( ByVal hWnd As Long, ByVal nIndex As Long) As Long
        Private Declare Sub Demo_SetLastError Lib "kernel32" Alias _
            "SetLastError" ( ByVal dwErrCode As Long)
    #End If

'------------------------------------------------------------------------------
' API CONSTANTS FOR TITLE-BAR STATE READ
'------------------------------------------------------------------------------
    Private Const DEMO_GWL_STYLE    As Long = -16       'Window style index
    Private Const DEMO_WS_CAPTION   As Long = &HC00000  'Caption mask (WS_BORDER Or WS_DLGFRAME)

    '                               '
    '------------------------------------------------------------------------------
    '
    '                           PUBLIC DEMO BUILDER
    '
    '------------------------------------------------------------------------------
    '


Public Sub Demo_CreateDemoSheet()

'
'------------------------------------------------------------------------------
'                            DEMO CREATE DEMO SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Build or rebuild the Excel UI demo sheet and its interactive control panel
'
' WHY THIS EXISTS
'   The demo workbook needs a single repeatable builder that:
'     - prepares the base demo template
'     - writes the application-level and window-level control sections
'     - creates the action, preset, and utility buttons
'     - writes the explanatory notes
'     - restores Excel Application state cleanly even on failure
'
'   Centralizing that logic keeps the demo easier to rebuild, maintain, and
'   present
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Enters a fast-mode scope for smoother rebuild
'   - Rebuilds the demo template sheet
'   - Writes two checkbox-driven control sections:
'       * application-level UI state
'       * window-level UI state
'   - Adds the action, selection, preset, sync, capture, and reset controls
'   - Writes the scope and restore notes
'   - Synchronizes the check boxes to the current Excel UI state
'   - Restores the five fast-mode settings, the entry cursor included, through
'     a centralized cleanup path; a partial fast-mode entry is rolled back by
'     DEMO_FastMode_Begin itself
'   - Leaves the demo view the template applied, DisplayFormulaBar included,
'     in place: it is the requested view, not a temporary setting
'
' ERROR POLICY
'   - Raises errors normally
'   - An original error keeps its number and source; a failed restoration is
'     appended to its description as a "Cleanup failure: " line
'   - A restoration failure without an original error is raised itself
'
' DEPENDENCIES
'   - DEMO_Btn_Click
'   - DEMO_FastMode_Begin
'   - DEMO_FastMode_EndPreservingError
'   - DEMO_Sheet_BuildTemplate
'   - DEMO_Prepare_LabeledInputSection
'   - DEMO_Write_NamedInputRow
'   - DEMO_Write_BandHeader
'   - DEMO_Btn_AddGrid
'   - DEMO_Btn_Add
'   - DEMO_Set_RangeBorder
'   - DEMO_CB_AddForms
'   - Demo_SyncCheckBoxesToUI
'
' HISTORY
'   2026-11-14 - Fast-mode cleanup restores the entry cursor and reports a
'                failed restoration (#70).
'   2026-07-25
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "Demo_CreateDemoSheet"  'Procedure name for diagnostics

    Dim WB                  As Workbook             'Target workbook
    Dim WS                  As Worksheet            'Demo worksheet
    Dim ButtonSpecs         As Variant              'Button name / caption / macro specification
    Dim FastModeState       As tDEMOFastModeState   'Saved Application state and owned settings
    Dim SavedErrNumber      As Long                 'Captured error number
    Dim SavedErrSource      As String               'Captured error source
    Dim SavedErrDescription As String               'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Target the workbook that contains this module
        Set WB = ThisWorkbook
    'Play the standard demo-button click feedback
        DEMO_Btn_Click
    'Capture and apply fast-mode Application settings, the wait cursor
    'included; a failed entry has already rolled itself back
        DEMO_FastMode_Begin FastModeState
    'Build or rebuild the generic template for the demo sheet
        DEMO_Sheet_BuildTemplate DEMO_SHEET_NAME, "EXCEL UI", "Demo Sheet", , , , _
            , , "C:H", , , , , , , , , , 29
    'Resolve the main demo sheet after template preparation
        Set WS = WB.Worksheets(DEMO_SHEET_NAME)

'------------------------------------------------------------------------------
' BUILD CONTROL PANEL APPLICATION-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Apply the standard section, label, and input formatting
        DEMO_Prepare_LabeledInputSection WS, WS.Range("G4:H4"), _
            "APPLICATION LEVEL UI STATE", WS.Range("G5:G8"), WS.Range("H5:H8")

'------------------------------------------------------------------------------
' WRITE INPUT ROWS APPLICATION-LEVEL UI STATE
'------------------------------------------------------------------------------
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G5"), WS.Range("H5"), _
            "Ribbon", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G6"), WS.Range("H6"), _
            "Status bar", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G7"), WS.Range("H7"), _
            "Scroll bars", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G8"), WS.Range("H8"), _
            "Formula bar", ""
    'Convert the application-level input cells to checkboxes
        DEMO_CB_AddForms WS, WS.Range("H5:H8"), Array(CB_RIBBON, _
            CB_STATUSBAR, CB_SCROLLBARS, CB_FORMULABAR)

'------------------------------------------------------------------------------
' BUILD CONTROL PANEL WINDOW-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Apply the standard section, label, and input formatting
        DEMO_Prepare_LabeledInputSection WS, WS.Range("G11:H11"), _
            "WINDOW LEVEL UI STATE", WS.Range("G12:G15"), WS.Range("H12:H15")

'------------------------------------------------------------------------------
' WRITE INPUT ROWS WINDOW-LEVEL UI STATE
'------------------------------------------------------------------------------
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G12"), WS.Range("H12"), _
            "Headings", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G13"), WS.Range("H13"), _
            "Workbook tabs", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G14"), WS.Range("H14"), _
            "Gridlines", ""
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G15"), WS.Range("H15"), _
            "Title bar", ""
    'Convert the window-level input cells to checkboxes
        DEMO_CB_AddForms WS, WS.Range("H12:H15"), Array(CB_HEADINGS, _
            CB_WORKBOOKTABS, CB_GRIDLINES, CB_TITLEBAR)

'------------------------------------------------------------------------------
' BUILD ACTION BUTTONS
'------------------------------------------------------------------------------
        DEMO_Write_BandHeader WS.Range("C4:E4"), "ACTIONS"
    'Define the action button grid
        ButtonSpecs = Array(Array("btn_UI_Show", "SHOW SELECTED UI", _
            "Demo_ShowSelectedUI"), Array("btn_UI_Hide", "HIDE SELECTED UI", _
            "Demo_HideSelectedUI"))
        DEMO_Btn_AddGrid WS, WS.Range("C5"), ButtonSpecs, 2, 150, 25
        DEMO_Set_RangeBorder WS.Range("C4:E6")

'------------------------------------------------------------------------------
' BUILD SELECT / CLEAR BUTTONS
'------------------------------------------------------------------------------
        DEMO_Write_BandHeader WS.Range("C8:E8"), "SELECT / CLEAR"
    'Define the select / clear button grid
        ButtonSpecs = Array(Array("btn_UI_SelectAll", "SELECT ALL", _
            "Demo_SelectAllUI"), Array("btn_UI_ClearAll", "CLEAR ALL", _
            "Demo_ClearAllUI"))
        DEMO_Btn_AddGrid WS, WS.Range("C9"), ButtonSpecs, 2, 150, 25
        DEMO_Set_RangeBorder WS.Range("C8:E10")

'------------------------------------------------------------------------------
' BUILD PRESET / UTILITY BUTTONS
'------------------------------------------------------------------------------
        DEMO_Write_BandHeader WS.Range("C12:E12"), "PRESET SELECTION"
    'Define the preset / utility button grid
        ButtonSpecs = Array(Array("btn_UI_Kiosk", "KIOSK", "Demo_PresetKiosk"), _
            Array("btn_UI_Analyst", "ANALYST", "Demo_PresetAnalyst"), _
            Array("btn_UI_Minimal", "MINIMAL", "Demo_PresetMinimal"))
        DEMO_Btn_AddGrid WS, WS.Range("C13"), ButtonSpecs, 2, 150, 25, , 13, , _
            8
        DEMO_Set_RangeBorder WS.Range("C12:E16")

'------------------------------------------------------------------------------
' BUILD CAPTURE / RESET STATE BUTTONS
'------------------------------------------------------------------------------
        DEMO_Write_BandHeader WS.Range("C18:E18"), "CAPTURE / RESET STATE"
    'Define the capture / reset button grid
        ButtonSpecs = Array(Array("btn_UI_CaptureState", "CAPTURE STATE", _
            "Demo_CaptureUIState"), Array("btn_UI_ResetState", "RESET STATE", _
            "Demo_ResetUIToCapturedState"))
        DEMO_Btn_AddGrid WS, WS.Range("C19"), ButtonSpecs, 2, 150, 25, , 13, , _
            8
        DEMO_Set_RangeBorder WS.Range("C18:E20")

'------------------------------------------------------------------------------
' BUILD SYNC CHECK BOXES BUTTON
'------------------------------------------------------------------------------
        DEMO_Btn_Add WS, "btn_UI_Sync", "SYNC CHECKBOXES", _
            WS.Range("G17").Left, WS.Range("G17").Top, WS.Range("G17:H18").Width, 25, _
            "Demo_SyncCheckBoxesToUI"

'------------------------------------------------------------------------------
' FORMAT NOTE AREAS
'------------------------------------------------------------------------------
    'Format the scope / semantics note area
        With WS.Range("B22:H25")
            .Merge
            .Interior.Color = RGB(255, 242, 204)
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        DEMO_Set_RangeBorder WS.Range("B22:H25")
    'Format the restore note area
        With WS.Range("B26:H27")
            .Merge
            .Interior.Color = RGB(217, 225, 242)
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        DEMO_Set_RangeBorder WS.Range("B26:H27")

'------------------------------------------------------------------------------
' WRITE NOTES
'------------------------------------------------------------------------------
        WS.Range("B22").Value = NOTE_SCOPE_TEXT
        WS.Range("B26").Value = NOTE_RESTORE_TEXT

'------------------------------------------------------------------------------
' SYNCHRONIZE CHECK BOXES
'------------------------------------------------------------------------------
        Demo_SyncCheckBoxesToUI False

'------------------------------------------------------------------------------
' BUILD RESET SHEET BUTTON
'------------------------------------------------------------------------------
        DEMO_Btn_Add WS, "btn_UI_ResetSheet", "RESET SHEET", _
            WS.Range("H2").Left + 1, WS.Range("H2").Top + 1, 105, 25, _
            "Demo_CreateDemoSheet"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanExit:
    'Disable the handler first, so cleanup can never re-enter it
        On Error GoTo 0
    'Restore every fast-mode setting this call applied; a failed restoration
    'joins the original error, or becomes the error, instead of being dropped
        DEMO_FastMode_EndPreservingError FastModeState, SavedErrNumber, _
            SavedErrSource, SavedErrDescription
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CleanFail:
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = IIf(Len(Err.Source) > 0, Err.Source, PROC_NAME)
        SavedErrDescription = Err.Description
        Resume CleanExit

End Sub

'
'------------------------------------------------------------------------------
'
'                             PUBLIC DEMO ACTIONS
'
'------------------------------------------------------------------------------
'

Public Sub Demo_ShowSelectedUI()

'
'------------------------------------------------------------------------------
'                            DEMO SHOW SELECTED UI
'------------------------------------------------------------------------------
' PURPOSE
'   Show only the UI elements currently selected by the user on the demo sheet
'
' WHY THIS EXISTS
'   The demo sheet uses check boxes as a user-friendly selector for which UI
'   elements should be affected by the EXCEL_UI module
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Interprets checked boxes as selected targets
'   - Applies UI_Show to selected elements
'   - Leaves unchecked elements unchanged
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_ApplySelectedUI
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLY ACTION
'------------------------------------------------------------------------------
        DEMO_Btn_PlayFeedback
        Demo_ApplySelectedUI UI_Show, "Demo_ShowSelectedUI"

End Sub

Public Sub Demo_HideSelectedUI()

'
'------------------------------------------------------------------------------
'                            DEMO HIDE SELECTED UI
'------------------------------------------------------------------------------
' PURPOSE
'   Hide only the UI elements currently selected by the user on the demo sheet
'
' WHY THIS EXISTS
'   The demo sheet uses check boxes as a user-friendly selector for which UI
'   elements should be affected by the EXCEL_UI module
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Interprets checked boxes as selected targets
'   - Applies UI_Hide to selected elements
'   - Leaves unchecked elements unchanged
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_ApplySelectedUI
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLY ACTION
'------------------------------------------------------------------------------
        DEMO_Btn_PlayFeedback
        Demo_ApplySelectedUI UI_Hide, "Demo_HideSelectedUI"

End Sub

Public Sub Demo_SyncCheckBoxesToUI(Optional ByVal PlayFeedback As Boolean = _
    True)
'
'------------------------------------------------------------------------------
'                         DEMO SYNC CHECK BOXES TO UI
'------------------------------------------------------------------------------
' PURPOSE
'   Read the current Excel UI state and synchronize the demo check boxes so
'   currently visible elements are checked
'
' WHY THIS EXISTS
'   A demo is more useful when users can see the current state before deciding
'   what to show or hide next
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Reads current application-level state directly from Excel
'   - Reads current window-level state from ActiveWindow
'   - Reads title-bar visibility from the Excel window represented by
'     Application.Hwnd
'   - Updates the check boxes to reflect the visible state
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Partial failures are written to the Immediate Window
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_TryGetRibbonVisibility
'   - Demo_TryGetTitleBarVisibility
'   - DEMO_CB_TrySetState
'   - Demo_LogFailure
'
' NOTES
'   Window-level sync uses ActiveWindow as the reference window
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "Demo_SyncCheckBoxesToUI"   'Procedure name for diagnostics

    Dim WS          As Worksheet    'Demo worksheet
    Dim ActiveWin   As Window       'Active Excel window for window-level reads
    Dim IsVisible   As Boolean      'Resolved current visibility state
    Dim FailMsg     As String       'Diagnostic message from reader / writer helpers

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
    'Play the optional button-click visual feedback only when requested
        If PlayFeedback Then
            DEMO_Btn_PlayFeedback
        End If
        Set WS = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)
        Set ActiveWin = Application.ActiveWindow

'------------------------------------------------------------------------------
' SYNC APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Read current Ribbon visibility and update the related check box
        If Demo_TryGetRibbonVisibility(IsVisible, FailMsg) Then
            If Not DEMO_CB_TrySetState(WS, CB_RIBBON, IsVisible, FailMsg) _
                Then
                Demo_LogFailure PROC_NAME, CB_RIBBON, FailMsg
            End If
        Else
            Demo_LogFailure PROC_NAME, "RibbonState", FailMsg
        End If
    'Read current StatusBar visibility and update the related check box
        If Not DEMO_CB_TrySetState(WS, CB_STATUSBAR, _
            Application.DisplayStatusBar, FailMsg) Then
            Demo_LogFailure PROC_NAME, CB_STATUSBAR, FailMsg
        End If
    'Read current ScrollBars visibility and update the related check box
        If Not DEMO_CB_TrySetState(WS, CB_SCROLLBARS, _
            Application.DisplayScrollBars, FailMsg) Then
            Demo_LogFailure PROC_NAME, CB_SCROLLBARS, FailMsg
        End If
    'Read current FormulaBar visibility and update the related check box
        If Not DEMO_CB_TrySetState(WS, CB_FORMULABAR, _
            Application.DisplayFormulaBar, FailMsg) Then
            Demo_LogFailure PROC_NAME, CB_FORMULABAR, FailMsg
        End If

'------------------------------------------------------------------------------
' SYNC WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Reject missing ActiveWindow deterministically for window-level sync
        If ActiveWin Is Nothing Then
            Demo_LogFailure PROC_NAME, "ActiveWindow", _
                "no active window available for window-level sync"
        Else
        'Update the Headings check box from ActiveWindow
            If Not DEMO_CB_TrySetState(WS, CB_HEADINGS, _
                ActiveWin.DisplayHeadings, FailMsg) Then
                Demo_LogFailure PROC_NAME, CB_HEADINGS, FailMsg
            End If
        'Update the WorkbookTabs check box from ActiveWindow
            If Not DEMO_CB_TrySetState(WS, CB_WORKBOOKTABS, _
                ActiveWin.DisplayWorkbookTabs, FailMsg) Then
                Demo_LogFailure PROC_NAME, CB_WORKBOOKTABS, FailMsg
            End If
        'Update the Gridlines check box from ActiveWindow
            If Not DEMO_CB_TrySetState(WS, CB_GRIDLINES, _
                ActiveWin.DisplayGridlines, FailMsg) Then
                Demo_LogFailure PROC_NAME, CB_GRIDLINES, FailMsg
            End If
        End If

'------------------------------------------------------------------------------
' SYNC TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Read current title-bar visibility and update the related check box
        If Demo_TryGetTitleBarVisibility(IsVisible, FailMsg) Then
            If Not DEMO_CB_TrySetState(WS, CB_TITLEBAR, IsVisible, FailMsg) _
                Then
                Demo_LogFailure PROC_NAME, CB_TITLEBAR, FailMsg
            End If
        Else
            Demo_LogFailure PROC_NAME, "TitleBarState", FailMsg
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        Demo_LogFailure PROC_NAME, "Unexpected", Demo_GetRuntimeErrorText
        Resume SafeExit

End Sub

Public Sub Demo_SelectAllUI()

'
'------------------------------------------------------------------------------
'                              DEMO SELECT ALL UI
'------------------------------------------------------------------------------
' PURPOSE
'   Check all demo check boxes so all managed UI elements are selected for the
'   next SHOW or HIDE action
'
' WHY THIS EXISTS
'   Select All is a useful convenience action during demos and testing
'
' RETURNS
'   Nothing
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
        DEMO_Btn_PlayFeedback

    'Select all managed UI elements
        Demo_SetSelectionProfile CallerProc:="Demo_SelectAllUI", _
            RibbonSelected:=True, StatusBarSelected:=True, ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, HeadingsSelected:=True, _
            WorkbookTabsSelected:=True, GridlinesSelected:=True, _
            TitleBarSelected:=True

End Sub

Public Sub Demo_ClearAllUI()

'
'------------------------------------------------------------------------------
'                              DEMO CLEAR ALL UI
'------------------------------------------------------------------------------
' PURPOSE
'   Clear all demo check boxes so no UI elements are selected for the next SHOW
'   or HIDE action
'
' WHY THIS EXISTS
'   Clear All is a useful convenience action during demos and testing
'
' RETURNS
'   Nothing
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
        DEMO_Btn_PlayFeedback
    'Clear all managed UI selections
        Demo_SetSelectionProfile CallerProc:="Demo_ClearAllUI", _
            RibbonSelected:=False, StatusBarSelected:=False, _
            ScrollBarsSelected:=False, FormulaBarSelected:=False, _
            HeadingsSelected:=False, WorkbookTabsSelected:=False, _
            GridlinesSelected:=False, TitleBarSelected:=False

End Sub

Public Sub Demo_PresetKiosk()

'
'------------------------------------------------------------------------------
'                              DEMO PRESET KIOSK
'------------------------------------------------------------------------------
' PURPOSE
'   Pre-select a broad kiosk-style profile covering all managed UI elements
'
' WHY THIS EXISTS
'   A kiosk-like presentation typically considers all major Excel chrome and
'   worksheet aids as candidate targets
'
' RETURNS
'   Nothing
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_SetSelectionProfile
'
' NOTES
'   - This preset only sets the check boxes
'   - It does NOT apply SHOW or HIDE by itself
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
        DEMO_Btn_PlayFeedback
    'Select all managed UI elements for a kiosk-style bundle
        Demo_SetSelectionProfile CallerProc:="Demo_PresetKiosk", _
            RibbonSelected:=True, StatusBarSelected:=True, ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, HeadingsSelected:=True, _
            WorkbookTabsSelected:=True, GridlinesSelected:=True, _
            TitleBarSelected:=True

End Sub

Public Sub Demo_PresetAnalyst()

'
'------------------------------------------------------------------------------
'                             DEMO PRESET ANALYST
'------------------------------------------------------------------------------
' PURPOSE
'   Pre-select a profile focused on worksheet navigation and analysis aids
'
' WHY THIS EXISTS
'   Analytical use cases often care most about sheet aids such as headings,
'   tabs, gridlines, formula bar, and related navigation cues
'
' RETURNS
'   Nothing
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_SetSelectionProfile
'
' NOTES
'   - This preset only sets the check boxes
'   - It does NOT apply SHOW or HIDE by itself
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
        DEMO_Btn_PlayFeedback
    'Select a worksheet-analysis-oriented bundle
        Demo_SetSelectionProfile CallerProc:="Demo_PresetAnalyst", _
            RibbonSelected:=False, StatusBarSelected:=True, ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, HeadingsSelected:=True, _
            WorkbookTabsSelected:=True, GridlinesSelected:=True, _
            TitleBarSelected:=False

End Sub

Public Sub Demo_PresetMinimal()

'
'------------------------------------------------------------------------------
'                             DEMO PRESET MINIMAL
'------------------------------------------------------------------------------
' PURPOSE
'   Pre-select a profile focused on major application chrome rather than
'   worksheet aids
'
' WHY THIS EXISTS
'   A minimal-shell scenario often focuses on the application frame, bars, and
'   navigation chrome
'
' RETURNS
'   Nothing
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_SetSelectionProfile
'
' NOTES
'   - This preset only sets the check boxes
'   - It does NOT apply SHOW or HIDE by itself
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
        DEMO_Btn_PlayFeedback
    'Select a minimal-shell-oriented bundle
        Demo_SetSelectionProfile CallerProc:="Demo_PresetMinimal", _
            RibbonSelected:=True, StatusBarSelected:=True, ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, HeadingsSelected:=False, _
            WorkbookTabsSelected:=False, GridlinesSelected:=False, _
            TitleBarSelected:=True

End Sub

Public Sub Demo_CaptureUIState(Optional ByVal ShowConfirmation As Boolean = _
    True)
'
'------------------------------------------------------------------------------
'                            DEMO CAPTURE UI STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Capture the current managed Excel UI state through the core module snapshot
'   API
'
' WHY THIS EXISTS
'   The demo becomes more useful when users can:
'     - capture the current baseline
'     - experiment with show / hide actions
'     - restore the captured baseline later
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Delegates to UI_CaptureExcelUIState
'   - Optionally shows a confirmation message
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Unexpected failures are written to the Immediate Window
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - UI_CaptureExcelUIState
'   - Demo_LogFailure
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "Demo_CaptureUIState"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
        DEMO_Btn_PlayFeedback

'------------------------------------------------------------------------------
' CAPTURE SNAPSHOT
'------------------------------------------------------------------------------
        UI_CaptureExcelUIState

'------------------------------------------------------------------------------
' INFORM USER
'------------------------------------------------------------------------------
    'Show the demo capture confirmation when requested
        If ShowConfirmation Then
            MsgBox "Current Excel UI state captured.", vbInformation, _
                "Excel UI Demo"
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        Demo_LogFailure PROC_NAME, "Unexpected", Demo_GetRuntimeErrorText
        Resume SafeExit

End Sub

Public Sub Demo_ResetUIToCapturedState()

'
'------------------------------------------------------------------------------
'                       DEMO RESET UI TO CAPTURED STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the managed Excel UI to the most recently captured explicit
'   snapshot through the core module reset API
'
' WHY THIS EXISTS
'   The demo becomes more useful when users can experiment safely and then
'   return to a captured baseline
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Rejects reset when no snapshot is available
'   - Delegates to UI_ResetExcelUIToSnapshot
'   - Synchronizes the check boxes back to the current visible state
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Unexpected failures are written to the Immediate Window
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - UI_HasExcelUIStateSnapshot
'   - UI_ResetExcelUIToSnapshot
'   - Demo_SyncCheckBoxesToUI
'   - Demo_LogFailure
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "Demo_ResetUIToCapturedState"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
        DEMO_Btn_PlayFeedback

'------------------------------------------------------------------------------
' VALIDATE SNAPSHOT AVAILABILITY
'------------------------------------------------------------------------------
    'Reject reset when no explicit snapshot is currently available
        If Not UI_HasExcelUIStateSnapshot Then
            MsgBox "No captured Excel UI state is available.", vbExclamation, _
                "Excel UI Demo"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RESET TO SNAPSHOT
'------------------------------------------------------------------------------
        UI_ResetExcelUIToSnapshot

'------------------------------------------------------------------------------
' RESYNCHRONIZE DEMO CHECK BOXES
'------------------------------------------------------------------------------
    'Synchronize the demo check boxes back to the current visible state after reset
        Demo_SyncCheckBoxesToUI False

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        Demo_LogFailure PROC_NAME, "Unexpected", Demo_GetRuntimeErrorText
        Resume SafeExit

End Sub

'
'------------------------------------------------------------------------------
'
'                         SHARED PRIVATE ORCHESTRATION
'
'------------------------------------------------------------------------------
'

Private Sub Demo_ApplySelectedUI(ByVal SelectedVisibility As UIVisibility, ByVal _
    CallerProc As String)
'
'------------------------------------------------------------------------------
'                            DEMO APPLY SELECTED UI
'------------------------------------------------------------------------------
' PURPOSE
'   Shared worker for applying SHOW or HIDE to the UI elements selected on the
'   demo worksheet
'
' WHY THIS EXISTS
'   The public SHOW and HIDE entry points are structurally identical except for
'   the requested tri-state action, so shared logic is centralized here
'
' INPUTS
'   SelectedVisibility
'     Requested action for checked elements:
'       - UI_Show
'       - UI_Hide
'
'   CallerProc
'     Public caller procedure name used for diagnostics
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Reads each demo check box
'   - Maps checked => SelectedVisibility
'   - Maps unchecked => UI_LeaveUnchanged
'   - Applies UI_SetExcelUI when at least one control is selected
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Unexpected failures are written to the Immediate Window
'
' DEPENDENCIES
'   - Demo_CheckBoxToUIVisibility
'   - Demo_HasAnySelectedChange
'   - UI_SetExcelUI
'   - Demo_LogFailure
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_DEMO.Demo_ApplySelectedUI"

    Dim WS              As Worksheet    'Demo worksheet
    Dim RibbonVis       As UIVisibility 'Resolved Ribbon visibility
    Dim StatusBarVis    As UIVisibility 'Resolved StatusBar visibility
    Dim ScrollBarsVis   As UIVisibility 'Resolved ScrollBars visibility
    Dim FormulaBarVis   As UIVisibility 'Resolved FormulaBar visibility
    Dim HeadingsVis     As UIVisibility 'Resolved Headings visibility
    Dim WorkbookTabsVis As UIVisibility 'Resolved WorkbookTabs visibility
    Dim GridlinesVis    As UIVisibility 'Resolved Gridlines visibility
    Dim TitleBarVis     As UIVisibility 'Resolved TitleBar visibility

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
        Set WS = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)

'------------------------------------------------------------------------------
' RESOLVE REQUESTED UI STATE
'------------------------------------------------------------------------------
    'Resolve Ribbon request from the related check box
        RibbonVis = Demo_CheckBoxToUIVisibility(WS, CB_RIBBON, _
            SelectedVisibility, CallerProc)
    'Resolve StatusBar request from the related check box
        StatusBarVis = Demo_CheckBoxToUIVisibility(WS, CB_STATUSBAR, _
            SelectedVisibility, CallerProc)
    'Resolve ScrollBars request from the related check box
        ScrollBarsVis = Demo_CheckBoxToUIVisibility(WS, CB_SCROLLBARS, _
            SelectedVisibility, CallerProc)
    'Resolve FormulaBar request from the related check box
        FormulaBarVis = Demo_CheckBoxToUIVisibility(WS, CB_FORMULABAR, _
            SelectedVisibility, CallerProc)
    'Resolve Headings request from the related check box
        HeadingsVis = Demo_CheckBoxToUIVisibility(WS, CB_HEADINGS, _
            SelectedVisibility, CallerProc)
    'Resolve WorkbookTabs request from the related check box
        WorkbookTabsVis = Demo_CheckBoxToUIVisibility(WS, CB_WORKBOOKTABS, _
            SelectedVisibility, CallerProc)
    'Resolve Gridlines request from the related check box
        GridlinesVis = Demo_CheckBoxToUIVisibility(WS, CB_GRIDLINES, _
            SelectedVisibility, CallerProc)
    'Resolve TitleBar request from the related check box
        TitleBarVis = Demo_CheckBoxToUIVisibility(WS, CB_TITLEBAR, _
            SelectedVisibility, CallerProc)

'------------------------------------------------------------------------------
' VALIDATE SELECTION
'------------------------------------------------------------------------------
    'Reject empty selection so the user understands why nothing happened
        If Not Demo_HasAnySelectedChange(RibbonVis, StatusBarVis, ScrollBarsVis, _
            FormulaBarVis, HeadingsVis, WorkbookTabsVis, GridlinesVis, TitleBarVis) _
            Then
        'Inform the user that no options were selected
            MsgBox "No UI elements are selected.", vbInformation, _
                "Excel UI Demo"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' APPLY REQUESTED STATE
'------------------------------------------------------------------------------
    'Apply the requested visibility only to the selected UI elements
        UI_SetExcelUI Ribbon:=RibbonVis, StatusBar:=StatusBarVis, _
            ScrollBars:=ScrollBarsVis, FormulaBar:=FormulaBarVis, _
            Headings:=HeadingsVis, WorkbookTabs:=WorkbookTabsVis, _
            Gridlines:=GridlinesVis, TitleBar:=TitleBarVis

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        Demo_LogFailure CallerProc, "Unexpected", Demo_GetRuntimeErrorText
        Resume SafeExit

End Sub

Private Sub Demo_SetSelectionProfile(ByVal CallerProc As String, ByVal _
    RibbonSelected As Boolean, ByVal StatusBarSelected As Boolean, ByVal _
    ScrollBarsSelected As Boolean, ByVal FormulaBarSelected As Boolean, ByVal _
    HeadingsSelected As Boolean, ByVal WorkbookTabsSelected As Boolean, ByVal _
    GridlinesSelected As Boolean, ByVal TitleBarSelected As Boolean)
'
'------------------------------------------------------------------------------
'                          DEMO SET SELECTION PROFILE
'------------------------------------------------------------------------------
' PURPOSE
'   Set all demo check boxes in one call according to the supplied Boolean
'   selection profile
'
' WHY THIS EXISTS
'   The demo exposes convenience actions such as:
'     - Select All
'     - Clear All
'     - preset bundles
'
'   A shared writer keeps those actions concise and consistent
'
' INPUTS
'   CallerProc
'     Public caller procedure name used for diagnostics
'
'   [Boolean selection flags]
'     Requested checked state for each demo check box
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Per-control failures are written to the Immediate Window
'
' DEPENDENCIES
'   - DEMO_CB_TrySetState
'   - Demo_LogFailure
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_DEMO.Demo_SetSelectionProfile"

    Dim WS      As Worksheet    'Demo worksheet
    Dim FailMsg As String       'Diagnostic message from the writer helper

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
        Set WS = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)

'------------------------------------------------------------------------------
' WRITE SELECTION PROFILE
'------------------------------------------------------------------------------
        If Not DEMO_CB_TrySetState(WS, CB_RIBBON, RibbonSelected, FailMsg) _
            Then
            Demo_LogFailure CallerProc, CB_RIBBON, FailMsg
        End If
        If Not DEMO_CB_TrySetState(WS, CB_STATUSBAR, StatusBarSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_STATUSBAR, FailMsg
        End If
        If Not DEMO_CB_TrySetState(WS, CB_SCROLLBARS, ScrollBarsSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_SCROLLBARS, FailMsg
        End If
        If Not DEMO_CB_TrySetState(WS, CB_FORMULABAR, FormulaBarSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_FORMULABAR, FailMsg
        End If
        If Not DEMO_CB_TrySetState(WS, CB_HEADINGS, HeadingsSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_HEADINGS, FailMsg
        End If
        If Not DEMO_CB_TrySetState(WS, CB_WORKBOOKTABS, _
            WorkbookTabsSelected, FailMsg) Then
            Demo_LogFailure CallerProc, CB_WORKBOOKTABS, FailMsg
        End If
        If Not DEMO_CB_TrySetState(WS, CB_GRIDLINES, GridlinesSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_GRIDLINES, FailMsg
        End If
        If Not DEMO_CB_TrySetState(WS, CB_TITLEBAR, TitleBarSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_TITLEBAR, FailMsg
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        Demo_LogFailure CallerProc, "Unexpected", Demo_GetRuntimeErrorText
        Resume SafeExit

End Sub

Private Function Demo_HasAnySelectedChange(ByVal RibbonVis As UIVisibility, _
    ByVal StatusBarVis As UIVisibility, ByVal ScrollBarsVis As UIVisibility, ByVal _
    FormulaBarVis As UIVisibility, ByVal HeadingsVis As UIVisibility, ByVal _
    WorkbookTabsVis As UIVisibility, ByVal GridlinesVis As UIVisibility, ByVal _
    TitleBarVis As UIVisibility) As Boolean
'
'------------------------------------------------------------------------------
'                         DEMO HAS ANY SELECTED CHANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Determine whether at least one UI element has been selected for change
'
' WHY THIS EXISTS
'   The demo macros should inform the user when no check boxes are selected,
'   rather than silently doing nothing
'
' RETURNS
'   TRUE  => at least one argument differs from UI_LeaveUnchanged
'   FALSE => all arguments are UI_LeaveUnchanged
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
        Demo_HasAnySelectedChange = (RibbonVis <> UI_LeaveUnchanged Or _
            StatusBarVis <> UI_LeaveUnchanged Or ScrollBarsVis <> UI_LeaveUnchanged _
            Or FormulaBarVis <> UI_LeaveUnchanged Or HeadingsVis <> _
            UI_LeaveUnchanged Or WorkbookTabsVis <> UI_LeaveUnchanged Or _
            GridlinesVis <> UI_LeaveUnchanged Or TitleBarVis <> UI_LeaveUnchanged)

End Function

'
'------------------------------------------------------------------------------
'
'                  PRIVATE CHECKBOX/STATE TRANSLATION HELPERS
'
'------------------------------------------------------------------------------
'

Private Function Demo_CheckBoxToUIVisibility(ByVal WS As Worksheet, ByVal _
    CheckBoxName As String, ByVal SelectedVisibility As UIVisibility, ByVal _
    CallerProc As String) As UIVisibility
'
'------------------------------------------------------------------------------
'                       DEMO CHECK BOX TO UI VISIBILITY
'------------------------------------------------------------------------------
' PURPOSE
'   Convert the checked state of a demo worksheet check box into a tri-state
'   UIVisibility value suitable for UI_SetExcelUI
'
' WHY THIS EXISTS
'   The demo uses check boxes to express selection semantics:
'     - checked   => affect this UI element
'     - unchecked => leave this UI element unchanged
'
' INPUTS
'   WS
'     Demo worksheet containing the check box control
'
'   CheckBoxName
'     Name of the Forms or ActiveX check box
'
'   SelectedVisibility
'     The visibility to apply when the check box is checked:
'       - UI_Show
'       - UI_Hide
'
'   CallerProc
'     Calling procedure name used for diagnostics
'
' RETURNS
'   UI_Show / UI_Hide when the check box is checked
'   UI_LeaveUnchanged when the check box is unchecked or unavailable
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Missing or invalid controls are written to the Immediate Window and
'     treated as UI_LeaveUnchanged
'
' DEPENDENCIES
'   - DEMO_CB_TryGetState
'   - Demo_LogFailure
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_DEMO.Demo_CheckBoxToUIVisibility"

    Dim IsChecked   As Boolean  'Resolved check-box state
    Dim FailMsg     As String   'Diagnostic message from the reader helper

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to UI_LeaveUnchanged unless a checked state is confirmed
        Demo_CheckBoxToUIVisibility = UI_LeaveUnchanged

'------------------------------------------------------------------------------
' READ CHECK-BOX STATE
'------------------------------------------------------------------------------
        If Not DEMO_CB_TryGetState(WS, CheckBoxName, IsChecked, FailMsg) _
            Then
        'Log the control-resolution failure and keep UI_LeaveUnchanged
            Demo_LogFailure CallerProc, CheckBoxName, FailMsg
            Exit Function
        End If

'------------------------------------------------------------------------------
' MAP CHECK-BOX STATE TO TRI-STATE VISIBILITY
'------------------------------------------------------------------------------
    'Apply the requested visibility only when the check box is checked
        If IsChecked Then
            Demo_CheckBoxToUIVisibility = SelectedVisibility
        Else
            Demo_CheckBoxToUIVisibility = UI_LeaveUnchanged
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                        PRIVATE UI-STATE READ HELPERS
'
'------------------------------------------------------------------------------
'

Private Function Demo_TryGetRibbonVisibility(ByRef IsVisible As Boolean, ByRef _
    FailMsg As String) As Boolean
'
'------------------------------------------------------------------------------
'                        DEMO TRY GET RIBBON VISIBILITY
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read current Ribbon visibility
'
' WHY THIS EXISTS
'   The demo needs to synchronize worksheet check boxes with the current Excel
'   UI state, but Ribbon visibility is not exposed through a simple dedicated
'   Application Boolean property
'
' INPUTS
'   IsVisible
'     Receives current Ribbon visibility on success
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => Ribbon visibility was read successfully
'   FALSE => read failed
'
' BEHAVIOR
'   - First attempts CommandBars("Ribbon").Visible
'   - Falls back to an Excel4 macro read when needed
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_DEMO.Demo_TryGetRibbonVisibility"

    Dim ValueOut    As Variant  'Fallback Excel4 macro result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
    'Initialize outputs and default result
        Demo_TryGetRibbonVisibility = False
        IsVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' TRY COMMANDBARS
'------------------------------------------------------------------------------
    'Attempt to read Ribbon visibility from the CommandBars collection
        On Error Resume Next
        IsVisible = Application.CommandBars("Ribbon").Visible
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler
            Demo_TryGetRibbonVisibility = True
            GoTo SafeExit
        End If
        Err.Clear
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' TRY EXCEL4 MACRO FALLBACK
'------------------------------------------------------------------------------
    'Attempt a fallback read using an Excel4 macro
        On Error Resume Next
        ValueOut = _
            Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")
        If Err.Number = 0 Then
            On Error GoTo ErrorHandler
            IsVisible = CBool(ValueOut)
            Demo_TryGetRibbonVisibility = True
            GoTo SafeExit
        End If
        FailMsg = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        FailMsg = Demo_GetRuntimeErrorText

End Function

Private Function Demo_TryGetTitleBarVisibility(ByRef IsVisible As Boolean, ByRef _
    FailMsg As String) As Boolean
'
'------------------------------------------------------------------------------
'                      DEMO TRY GET TITLE BAR VISIBILITY
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read current title-bar visibility for the Excel window
'   represented by Application.Hwnd
'
' WHY THIS EXISTS
'   Title-bar state is managed through WinAPI in EXCEL_UI, so the demo needs a
'   corresponding read-side helper for synchronization
'
' INPUTS
'   IsVisible
'     Receives current title-bar visibility on success: TRUE only when the
'     complete WS_CAPTION mask (WS_BORDER and WS_DLGFRAME) is set
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => title-bar visibility was read successfully
'   FALSE => read failed
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' WINAPI ERROR CAPTURE
'   SetLastError 0, the style read into a local and the Err.LastDllError
'   capture are three adjacent statements in every compilation branch, as in
'   src/. GetWindowLong[Ptr] can return zero legitimately, so only a zero
'   style with a non-zero captured error is a failure.
'
' HISTORY
'   2026-11-18 - Tests the complete WS_CAPTION mask instead of any caption
'                bit (#6).
'   2026-11-15 - Err.LastDllError captured right after the call (#144).
'   2026-07-25
'
' UPDATED
'   2026-11-18
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_DEMO.Demo_TryGetTitleBarVisibility"

#If VBA7 Then
    Dim xlHnd       As LongPtr  'Excel window handle from Application.Hwnd
    Dim StyleValue  As LongPtr  'Current window style value
#Else
    Dim xlHnd       As Long 'Excel window handle from Application.Hwnd
    Dim StyleValue  As Long 'Current window style value
#End If
    Dim LastErr As Long 'Last Win32 error after API call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
    'Initialize outputs and default result
        Demo_TryGetTitleBarVisibility = False
        IsVisible = False
        FailMsg = vbNullString
        xlHnd = Application.hWnd
    'Reject invalid window handle deterministically
        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ WINDOW STYLE
'------------------------------------------------------------------------------
    'Clear last-error state before the API call
        Demo_SetLastError 0

    #If VBA7 Then
        #If Win64 Then
    'Read the current window style using the 64-bit API
        StyleValue = Demo_GetWindowLongPtr(xlHnd, DEMO_GWL_STYLE)
        #Else
    'Read the current window style using the 32-bit API under VBA7
        StyleValue = Demo_GetWindowLong(xlHnd, DEMO_GWL_STYLE)
        #End If
    #Else
    'Read the current window style using the legacy 32-bit API
        StyleValue = Demo_GetWindowLong(xlHnd, DEMO_GWL_STYLE)
    #End If

    'Capture the native error in the very next statement, before any other
    'call or On Error statement can replace it
        LastErr = Err.LastDllError

    'A zero style with a captured error of zero is a valid zero; a zero style
    'with a non-zero captured error is a failure
        If StyleValue = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed reading the window style; " & _
                "Err.LastDllError=" & CStr(LastErr) & "; hWnd=" & CStr(xlHnd)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' MAP STYLE TO TITLE-BAR VISIBILITY
'------------------------------------------------------------------------------
    'Treat the complete caption mask as the title-bar visibility signal.
    'WS_CAPTION is WS_BORDER Or WS_DLGFRAME; one of the two is not a caption.
        IsVisible = ((StyleValue And DEMO_WS_CAPTION) = DEMO_WS_CAPTION)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
        Demo_TryGetTitleBarVisibility = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        FailMsg = Demo_GetRuntimeErrorText

End Function

'
'------------------------------------------------------------------------------
'
'                             PRIVATE DIAGNOSTICS
'
'------------------------------------------------------------------------------
'

Private Function Demo_GetRuntimeErrorText() As String

'
'------------------------------------------------------------------------------
'                         DEMO GET RUNTIME ERROR TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Build a consistent runtime diagnostic string from the active Err object
'
' WHY THIS EXISTS
'   Multiple routines in this module need the same failure-text formatting
'
' RETURNS
'   String
'     Best-effort diagnostic text including Err.Number, Err.Description,
'     Err.Source when available, and Erl when available
'
' ERROR POLICY
'   Does NOT raise
'
' NOTES
'   The Err object is read before the On Error statement, not after. Any form of
'   On Error resets Err, so the obvious ordering silently produces "0:" with an
'   empty description for every diagnostic this builder is asked to format.
'
' HISTORY
'   2026-08-21 - Capture Err before suppressing errors, so diagnostics carry the
'                failure they were invoked to describe.
'   2026-07-25
'
' UPDATED
'   2026-08-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_DEMO.Demo_GetRuntimeErrorText"

    Dim ErrNumber       As Long     'Err.Number captured on entry
    Dim ErrDescription  As String   'Err.Description captured on entry
    Dim ErrSource       As String   'Err.Source captured on entry
    Dim ErrLine         As Long     'Erl captured on entry

'------------------------------------------------------------------------------
' CAPTURE ERROR STATE
'------------------------------------------------------------------------------
    'Read the Err object BEFORE any On Error statement. Any form of On Error
    'resets Err, so a read taken afterwards returns zero and an empty
    'description, and every diagnostic this builder produces reports "0:" with
    'nothing after it.
        ErrNumber = Err.Number
        ErrDescription = Err.Description
        ErrSource = Err.Source
        ErrLine = Erl

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected issue while formatting the diagnostic
        On Error Resume Next

'------------------------------------------------------------------------------
' BUILD ERROR TEXT
'------------------------------------------------------------------------------
        Demo_GetRuntimeErrorText = CStr(ErrNumber) & ": " & ErrDescription & _
            IIf(Len(ErrSource) > 0, " | Source: " & ErrSource, vbNullString) & _
            IIf(ErrLine <> 0, " | Line: " & CStr(ErrLine), vbNullString)

End Function

Private Sub Demo_LogFailure(ByVal ProcName As String, ByVal Stage As String, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                               DEMO LOG FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Write a consistent diagnostic line to the Immediate Window for the demo
'   module
'
' WHY THIS EXISTS
'   The demo uses fail-soft behavior and needs a single place to format
'   diagnostics consistently
'
' INPUTS
'   ProcName
'     Procedure name associated with the failure
'
'   Stage
'     Logical stage associated with the failure
'
'   Detail
'     Failure detail to append
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Suppresses any unexpected logging failure locally
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected logging failure
        On Error Resume Next

'------------------------------------------------------------------------------
' WRITE DIAGNOSTIC LINE
'------------------------------------------------------------------------------
        Debug.Print ProcName & " failed @ " & Stage & " | " & Detail

End Sub
