Attribute VB_Name = "M_EXCEL_UI_DEMO"
Option Explicit

'
'==============================================================================
'                           MODULE: EXCEL_UI_DEMO
'------------------------------------------------------------------------------
' PURPOSE
'   Provide a worksheet-based showcase for the EXCEL_UI module, including:
'     - selective SHOW / HIDE actions driven by worksheet check boxes
'     - a demo-sheet builder
'     - current-state synchronization back into the check boxes
'     - selection helpers (Select All / Clear All)
'     - preset selection profiles
'     - explanatory notes rendered on the demo sheet
'
' WHY THIS EXISTS
'   A demo workbook is easier to understand and present when non-technical
'   users can interact with the EXCEL_UI module through worksheet controls
'   rather than by editing VBA calls directly.
'
'   This module bridges the demo sheet controls to the public EXCEL_UI API and
'   also builds the demo surface itself so the showcase is reproducible.
'
' PUBLIC SURFACE
'   - Demo_ShowSelectedExcelUI
'   - Demo_HideSelectedExcelUI
'   - Demo_SyncCheckBoxesToCurrentUI
'   - Demo_SelectAllUI
'   - Demo_ClearAllUI
'   - Demo_PresetKiosk
'   - Demo_PresetAnalyst
'   - Demo_PresetMinimal
'   - Demo_CreateExcelUISheet
'
' EXPECTED DEMO CHECK BOX NAMES
'   Application-level
'     - chkRibbon
'     - chkStatusBar
'     - chkScrollBars
'     - chkFormulaBar
'
'   Window-level
'     - chkHeadings
'     - chkWorkbookTabs
'     - chkGridlines
'     - chkTitleBar
'
' EXPECTED DEMO SHEET
'   - Worksheet name: Demo
'
' DEMO SEMANTICS
'   - Checked     => selected for the next SHOW / HIDE action
'   - Unchecked   => leave unchanged
'
'   The "Sync Checkboxes" feature reads the current UI state and marks currently
'   visible elements as checked for convenience / reference.
'
' COMPATIBILITY
'   - Supports both Forms check boxes and ActiveX check boxes when reading or
'     writing control state.
'   - Relies on the public API exposed by the EXCEL_UI module:
'       * K_UIVisibility
'       * K_SetExcelUI
'       * K_HideExcelUI
'       * K_ShowExcelUI
'
' NOTES
'   - Window-level sync reads the current ActiveWindow state.
'   - TitleBar sync reads the Excel window represented by Application.Hwnd.
'   - Missing or misnamed controls are logged to the Immediate Window and
'     treated conservatively.
'   - The demo builder performs a destructive rebuild of the Demo sheet.
'
' UPDATED
'   2026-04-04
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.0.0
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE: DEMO CONFIGURATION
'------------------------------------------------------------------------------
Private Const DEMO_SHEET_NAME         As String = "Demo"                    'Demo worksheet name

Private Const CB_RIBBON               As String = "chkRibbon"               'Ribbon check box name
Private Const CB_STATUSBAR            As String = "chkStatusBar"            'StatusBar check box name
Private Const CB_SCROLLBARS           As String = "chkScrollBars"           'ScrollBars check box name
Private Const CB_FORMULABAR           As String = "chkFormulaBar"           'FormulaBar check box name

Private Const CB_HEADINGS             As String = "chkHeadings"             'Headings check box name
Private Const CB_WORKBOOKTABS         As String = "chkWorkbookTabs"         'WorkbookTabs check box name
Private Const CB_GRIDLINES            As String = "chkGridlines"            'Gridlines check box name
Private Const CB_TITLEBAR             As String = "chkTitleBar"             'TitleBar check box name

Private Const BTN_SHOW_NAME           As String = "btnShowExcelUI"          'Show button shape name
Private Const BTN_HIDE_NAME           As String = "btnHideExcelUI"          'Hide button shape name
Private Const BTN_SYNC_NAME           As String = "btnSyncExcelUI"          'Sync button shape name
Private Const BTN_SELECTALL_NAME      As String = "btnSelectAllUI"          'Select-all button shape name
Private Const BTN_CLEARALL_NAME       As String = "btnClearAllUI"           'Clear-all button shape name
Private Const BTN_PRESET_KIOSK_NAME   As String = "btnPresetKioskUI"        'Preset-Kiosk button shape name
Private Const BTN_PRESET_ANALYST_NAME As String = "btnPresetAnalystUI"      'Preset-Analyst button shape name
Private Const BTN_PRESET_MINIMAL_NAME As String = "btnPresetMinimalUI"      'Preset-Minimal button shape name

Private Const BTN_SHOW_MACRO          As String = "Demo_ShowSelectedExcelUI"     'Show button macro
Private Const BTN_HIDE_MACRO          As String = "Demo_HideSelectedExcelUI"     'Hide button macro
Private Const BTN_SYNC_MACRO          As String = "Demo_SyncCheckBoxesToCurrentUI" 'Sync button macro
Private Const BTN_SELECTALL_MACRO     As String = "Demo_SelectAllUI"             'Select-all button macro
Private Const BTN_CLEARALL_MACRO      As String = "Demo_ClearAllUI"              'Clear-all button macro
Private Const BTN_PRESET_KIOSK_MACRO  As String = "Demo_PresetKiosk"             'Preset-Kiosk button macro
Private Const BTN_PRESET_ANALYST_MACRO As String = "Demo_PresetAnalyst"          'Preset-Analyst button macro
Private Const BTN_PRESET_MINIMAL_MACRO As String = "Demo_PresetMinimal"          'Preset-Minimal button macro

Private Const NOTE_SCOPE_TEXT As String = _
    "Scope / semantics note:" & vbLf & _
    "- Checked means SELECTED for the next SHOW or HIDE action." & vbLf & _
    "- Application-level items affect the current Excel instance." & vbLf & _
    "- Window-level sync reads ActiveWindow; apply actions target each open Excel window." & vbLf & _
    "- TitleBar is Windows-only and uses WinAPI against Application.Hwnd." & vbLf & _
    "- Preset buttons only set selections; they do not apply SHOW or HIDE by themselves."

Private Const NOTE_RESTORE_TEXT As String = _
    "Restore note:" & vbLf & _
    "K_ShowExcelUI shows all managed UI elements. It does NOT restore a previously captured user-specific UI state."

'------------------------------------------------------------------------------
' DECLARE: WIN32 / WIN64 API (TITLE-BAR STATE READ FOR SYNC)
'------------------------------------------------------------------------------
#If VBA7 Then

    #If Win64 Then

        Private Declare PtrSafe Function Demo_GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
            ByVal hWnd As LongPtr, _
            ByVal nIndex As Long) _
            As LongPtr

    #Else

        Private Declare PtrSafe Function Demo_GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
            ByVal hWnd As LongPtr, _
            ByVal nIndex As Long) _
            As Long

    #End If

    Private Declare PtrSafe Function Demo_GetLastError Lib "kernel32" Alias "GetLastError" () As Long

    Private Declare PtrSafe Sub Demo_SetLastError Lib "kernel32" Alias "SetLastError" ( _
        ByVal dwErrCode As Long)

#Else

    Private Declare Function Demo_GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long) _
        As Long

    Private Declare Function Demo_GetLastError Lib "kernel32" Alias "GetLastError" () As Long

    Private Declare Sub Demo_SetLastError Lib "kernel32" Alias "SetLastError" ( _
        ByVal dwErrCode As Long)

#End If

'------------------------------------------------------------------------------
' DECLARE: WINAPI CONSTANTS FOR TITLE-BAR STATE READ
'------------------------------------------------------------------------------
Private Const DEMO_GWL_STYLE          As Long = -16       'Window style index
Private Const DEMO_WS_CAPTION         As Long = &HC00000  'Caption / title-bar style bit

Public Sub Demo_ShowSelectedExcelUI()

'
'==============================================================================
'                        Demo_ShowSelectedExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Show only the UI elements currently selected by the user on the demo sheet.
'
' WHY THIS EXISTS
'   The demo sheet uses check boxes as a user-friendly selector for which
'   UI elements should be affected by the EXCEL_UI module.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Interprets checked boxes as selected targets.
'   - Applies K_UI_Show to selected elements.
'   - Leaves unchecked elements unchanged.
'
' DEPENDENCIES
'   - Demo_ApplySelectedExcelUI
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Delegate the action to the shared worker
        Demo_ApplySelectedExcelUI K_UI_Show, "Demo_ShowSelectedExcelUI"

End Sub

Public Sub Demo_HideSelectedExcelUI()

'
'==============================================================================
'                        Demo_HideSelectedExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Hide only the UI elements currently selected by the user on the demo sheet.
'
' WHY THIS EXISTS
'   The demo sheet uses check boxes as a user-friendly selector for which
'   UI elements should be affected by the EXCEL_UI module.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Interprets checked boxes as selected targets.
'   - Applies K_UI_Hide to selected elements.
'   - Leaves unchecked elements unchanged.
'
' DEPENDENCIES
'   - Demo_ApplySelectedExcelUI
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Delegate the action to the shared worker
        Demo_ApplySelectedExcelUI K_UI_Hide, "Demo_HideSelectedExcelUI"

End Sub

Public Sub Demo_SyncCheckBoxesToCurrentUI()

'
'==============================================================================
'                     Demo_SyncCheckBoxesToCurrentUI
'------------------------------------------------------------------------------
' PURPOSE
'   Read the current Excel UI state and synchronize the demo check boxes so
'   currently visible elements are checked.
'
' WHY THIS EXISTS
'   A demo is more useful when users can see the current state before deciding
'   what to show or hide next.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads current application-level state directly from Excel.
'   - Reads current window-level state from ActiveWindow.
'   - Reads title-bar visibility from the Excel window represented by
'     Application.Hwnd.
'   - Updates the check boxes to reflect the visible state.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Partial failures are written to the Immediate Window.
'
' DEPENDENCIES
'   - Demo_TryGetRibbonVisible
'   - Demo_TryGetTitleBarVisible
'   - Demo_TrySetCheckBoxState
'   - Demo_LogFailure
'
' NOTES
'   - Window-level sync uses ActiveWindow as the reference window.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ws                  As Worksheet     'Demo worksheet
    Dim Win                 As Window        'Active Excel window for window-level reads
    Dim IsVisible           As Boolean       'Resolved current visibility state
    Dim Msg                 As String        'Diagnostic message from reader / writer helpers

    Const PROC As String = "Demo_SyncCheckBoxesToCurrentUI"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Resolve the demo worksheet
        Set Ws = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)

    'Resolve the active Excel window used for window-level sync
        Set Win = Application.ActiveWindow

'------------------------------------------------------------------------------
' SYNC APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Read current Ribbon visibility and update the related check box
        If Demo_TryGetRibbonVisible(IsVisible, Msg) Then
            If Not Demo_TrySetCheckBoxState(Ws, CB_RIBBON, IsVisible, Msg) Then
                Demo_LogFailure PROC, CB_RIBBON, Msg
            End If
        Else
            Demo_LogFailure PROC, "RibbonState", Msg
        End If

    'Read current StatusBar visibility and update the related check box
        If Not Demo_TrySetCheckBoxState(Ws, CB_STATUSBAR, Application.DisplayStatusBar, Msg) Then
            Demo_LogFailure PROC, CB_STATUSBAR, Msg
        End If

    'Read current ScrollBars visibility and update the related check box
        If Not Demo_TrySetCheckBoxState(Ws, CB_SCROLLBARS, Application.DisplayScrollBars, Msg) Then
            Demo_LogFailure PROC, CB_SCROLLBARS, Msg
        End If

    'Read current FormulaBar visibility and update the related check box
        If Not Demo_TrySetCheckBoxState(Ws, CB_FORMULABAR, Application.DisplayFormulaBar, Msg) Then
            Demo_LogFailure PROC, CB_FORMULABAR, Msg
        End If

'------------------------------------------------------------------------------
' SYNC WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Reject missing ActiveWindow deterministically for window-level sync
        If Win Is Nothing Then

            'Log the missing ActiveWindow state
                Demo_LogFailure PROC, "ActiveWindow", "no active window available for window-level sync"

        Else

            'Update the Headings check box from ActiveWindow
                If Not Demo_TrySetCheckBoxState(Ws, CB_HEADINGS, Win.DisplayHeadings, Msg) Then
                    Demo_LogFailure PROC, CB_HEADINGS, Msg
                End If

            'Update the WorkbookTabs check box from ActiveWindow
                If Not Demo_TrySetCheckBoxState(Ws, CB_WORKBOOKTABS, Win.DisplayWorkbookTabs, Msg) Then
                    Demo_LogFailure PROC, CB_WORKBOOKTABS, Msg
                End If

            'Update the Gridlines check box from ActiveWindow
                If Not Demo_TrySetCheckBoxState(Ws, CB_GRIDLINES, Win.DisplayGridlines, Msg) Then
                    Demo_LogFailure PROC, CB_GRIDLINES, Msg
                End If

        End If

'------------------------------------------------------------------------------
' SYNC TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Read current title-bar visibility and update the related check box
        If Demo_TryGetTitleBarVisible(IsVisible, Msg) Then
            If Not Demo_TrySetCheckBoxState(Ws, CB_TITLEBAR, IsVisible, Msg) Then
                Demo_LogFailure PROC, CB_TITLEBAR, Msg
            End If
        Else
            Demo_LogFailure PROC, "TitleBarState", Msg
        End If

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write a diagnostic line without interrupting callers
        Demo_LogFailure PROC, "Unexpected", _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

    'Exit quietly after logging
        Resume SafeExit

End Sub

Public Sub Demo_SelectAllUI()

'
'==============================================================================
'                             Demo_SelectAllUI
'------------------------------------------------------------------------------
' PURPOSE
'   Check all demo check boxes so all managed UI elements are selected for the
'   next SHOW or HIDE action.
'
' WHY THIS EXISTS
'   Select All is a useful convenience action during demos and testing.
'
' RETURNS
'   None
'
' DEPENDENCIES
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
    'Select all managed UI elements
        Demo_SetSelectionProfile _
            CallerProc:="Demo_SelectAllUI", _
            RibbonSelected:=True, _
            StatusBarSelected:=True, _
            ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, _
            HeadingsSelected:=True, _
            WorkbookTabsSelected:=True, _
            GridlinesSelected:=True, _
            TitleBarSelected:=True

End Sub

Public Sub Demo_ClearAllUI()

'
'==============================================================================
'                             Demo_ClearAllUI
'------------------------------------------------------------------------------
' PURPOSE
'   Clear all demo check boxes so no UI elements are selected for the next
'   SHOW or HIDE action.
'
' WHY THIS EXISTS
'   Clear All is a useful convenience action during demos and testing.
'
' RETURNS
'   None
'
' DEPENDENCIES
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
    'Clear all managed UI selections
        Demo_SetSelectionProfile _
            CallerProc:="Demo_ClearAllUI", _
            RibbonSelected:=False, _
            StatusBarSelected:=False, _
            ScrollBarsSelected:=False, _
            FormulaBarSelected:=False, _
            HeadingsSelected:=False, _
            WorkbookTabsSelected:=False, _
            GridlinesSelected:=False, _
            TitleBarSelected:=False

End Sub

Public Sub Demo_PresetKiosk()

'
'==============================================================================
'                            Demo_PresetKiosk
'------------------------------------------------------------------------------
' PURPOSE
'   Pre-select a broad "kiosk" profile covering all managed UI elements.
'
' WHY THIS EXISTS
'   A kiosk-like presentation typically considers all major Excel chrome and
'   worksheet aids as candidate targets.
'
' RETURNS
'   None
'
' NOTES
'   - This preset only sets the check boxes.
'   - It does NOT apply SHOW or HIDE by itself.
'
' DEPENDENCIES
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
    'Select all managed UI elements for a kiosk-style bundle
        Demo_SetSelectionProfile _
            CallerProc:="Demo_PresetKiosk", _
            RibbonSelected:=True, _
            StatusBarSelected:=True, _
            ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, _
            HeadingsSelected:=True, _
            WorkbookTabsSelected:=True, _
            GridlinesSelected:=True, _
            TitleBarSelected:=True

End Sub

Public Sub Demo_PresetAnalyst()

'
'==============================================================================
'                           Demo_PresetAnalyst
'------------------------------------------------------------------------------
' PURPOSE
'   Pre-select a profile focused on worksheet navigation and analysis aids.
'
' WHY THIS EXISTS
'   Analytical use cases often care most about sheet aids such as headings,
'   tabs, gridlines, formula bar, and related navigation cues.
'
' RETURNS
'   None
'
' NOTES
'   - This preset only sets the check boxes.
'   - It does NOT apply SHOW or HIDE by itself.
'
' DEPENDENCIES
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
    'Select a worksheet-analysis-oriented bundle
        Demo_SetSelectionProfile _
            CallerProc:="Demo_PresetAnalyst", _
            RibbonSelected:=False, _
            StatusBarSelected:=True, _
            ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, _
            HeadingsSelected:=True, _
            WorkbookTabsSelected:=True, _
            GridlinesSelected:=True, _
            TitleBarSelected:=False

End Sub

Public Sub Demo_PresetMinimal()

'
'==============================================================================
'                           Demo_PresetMinimal
'------------------------------------------------------------------------------
' PURPOSE
'   Pre-select a profile focused on major application chrome rather than
'   worksheet aids.
'
' WHY THIS EXISTS
'   A minimal-shell scenario often focuses on the application frame, bars,
'   and navigation chrome.
'
' RETURNS
'   None
'
' NOTES
'   - This preset only sets the check boxes.
'   - It does NOT apply SHOW or HIDE by itself.
'
' DEPENDENCIES
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
    'Select a minimal-shell-oriented bundle
        Demo_SetSelectionProfile _
            CallerProc:="Demo_PresetMinimal", _
            RibbonSelected:=True, _
            StatusBarSelected:=True, _
            ScrollBarsSelected:=True, _
            FormulaBarSelected:=True, _
            HeadingsSelected:=False, _
            WorkbookTabsSelected:=False, _
            GridlinesSelected:=False, _
            TitleBarSelected:=True

End Sub

Private Sub Demo_ApplySelectedExcelUI( _
    ByVal SelectedVisibility As K_UIVisibility, _
    ByVal CallerProc As String)

'
'==============================================================================
'                        Demo_ApplySelectedExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Shared worker for applying SHOW or HIDE to the UI elements selected on the
'   demo worksheet.
'
' WHY THIS EXISTS
'   The public SHOW and HIDE entry points are structurally identical except for
'   the requested tri-state action, so shared logic is centralized here.
'
' INPUTS
'   SelectedVisibility
'     Requested action for checked elements:
'       - K_UI_Show
'       - K_UI_Hide
'
'   CallerProc
'     Public caller procedure name used for diagnostics.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads each demo check box.
'   - Maps checked => SelectedVisibility.
'   - Maps unchecked => K_UI_LeaveUnchanged.
'   - Applies K_SetExcelUI when at least one control is selected.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Unexpected failures are written to the Immediate Window.
'
' DEPENDENCIES
'   - Demo_CheckBoxToVisibility
'   - Demo_HasAnyRequestedChange
'   - K_SetExcelUI
'   - Demo_LogFailure
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ws                  As Worksheet         'Demo worksheet
    Dim RibbonVis           As K_UIVisibility    'Resolved Ribbon visibility
    Dim StatusBarVis        As K_UIVisibility    'Resolved StatusBar visibility
    Dim ScrollBarsVis       As K_UIVisibility    'Resolved ScrollBars visibility
    Dim FormulaBarVis       As K_UIVisibility    'Resolved FormulaBar visibility
    Dim HeadingsVis         As K_UIVisibility    'Resolved Headings visibility
    Dim WorkbookTabsVis     As K_UIVisibility    'Resolved WorkbookTabs visibility
    Dim GridlinesVis        As K_UIVisibility    'Resolved Gridlines visibility
    Dim TitleBarVis         As K_UIVisibility    'Resolved TitleBar visibility

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Resolve the demo worksheet
        Set Ws = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)

'------------------------------------------------------------------------------
' RESOLVE REQUESTED UI STATE
'------------------------------------------------------------------------------
    'Resolve Ribbon request from the related check box
        RibbonVis = Demo_CheckBoxToVisibility(Ws, CB_RIBBON, SelectedVisibility, CallerProc)

    'Resolve StatusBar request from the related check box
        StatusBarVis = Demo_CheckBoxToVisibility(Ws, CB_STATUSBAR, SelectedVisibility, CallerProc)

    'Resolve ScrollBars request from the related check box
        ScrollBarsVis = Demo_CheckBoxToVisibility(Ws, CB_SCROLLBARS, SelectedVisibility, CallerProc)

    'Resolve FormulaBar request from the related check box
        FormulaBarVis = Demo_CheckBoxToVisibility(Ws, CB_FORMULABAR, SelectedVisibility, CallerProc)

    'Resolve Headings request from the related check box
        HeadingsVis = Demo_CheckBoxToVisibility(Ws, CB_HEADINGS, SelectedVisibility, CallerProc)

    'Resolve WorkbookTabs request from the related check box
        WorkbookTabsVis = Demo_CheckBoxToVisibility(Ws, CB_WORKBOOKTABS, SelectedVisibility, CallerProc)

    'Resolve Gridlines request from the related check box
        GridlinesVis = Demo_CheckBoxToVisibility(Ws, CB_GRIDLINES, SelectedVisibility, CallerProc)

    'Resolve TitleBar request from the related check box
        TitleBarVis = Demo_CheckBoxToVisibility(Ws, CB_TITLEBAR, SelectedVisibility, CallerProc)

'------------------------------------------------------------------------------
' VALIDATE SELECTION
'------------------------------------------------------------------------------
    'Reject empty selection so the user understands why nothing happened
        If Not Demo_HasAnyRequestedChange( _
                    RibbonVis, _
                    StatusBarVis, _
                    ScrollBarsVis, _
                    FormulaBarVis, _
                    HeadingsVis, _
                    WorkbookTabsVis, _
                    GridlinesVis, _
                    TitleBarVis) Then

            'Inform the user that no options were selected
                MsgBox "No UI elements are selected.", vbInformation, "Excel UI Demo"

            'Exit quietly
                GoTo SafeExit

        End If

'------------------------------------------------------------------------------
' APPLY REQUESTED STATE
'------------------------------------------------------------------------------
    'Apply the requested visibility only to the selected UI elements
        K_SetExcelUI _
            Ribbon:=RibbonVis, _
            StatusBar:=StatusBarVis, _
            ScrollBars:=ScrollBarsVis, _
            FormulaBar:=FormulaBarVis, _
            Headings:=HeadingsVis, _
            WorkbookTabs:=WorkbookTabsVis, _
            Gridlines:=GridlinesVis, _
            TitleBar:=TitleBarVis

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write a diagnostic line without interrupting callers
        Demo_LogFailure CallerProc, "Unexpected", _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

    'Exit quietly after logging
        Resume SafeExit

End Sub

Private Function Demo_CheckBoxToVisibility( _
    ByVal Ws As Worksheet, _
    ByVal CheckBoxName As String, _
    ByVal SelectedVisibility As K_UIVisibility, _
    ByVal CallerProc As String) As K_UIVisibility

'
'==============================================================================
'                        Demo_CheckBoxToVisibility
'------------------------------------------------------------------------------
' PURPOSE
'   Convert the checked state of a demo worksheet check box into a tri-state
'   K_UIVisibility value suitable for K_SetExcelUI.
'
' WHY THIS EXISTS
'   The demo uses check boxes to express selection semantics:
'     - checked   => affect this UI element
'     - unchecked => leave this UI element unchanged
'
' INPUTS
'   Ws
'     Demo worksheet containing the check box control.
'
'   CheckBoxName
'     Name of the Forms or ActiveX check box.
'
'   SelectedVisibility
'     The visibility to apply when the check box is checked:
'       - K_UI_Show
'       - K_UI_Hide
'
'   CallerProc
'     Calling procedure name used for diagnostics.
'
' RETURNS
'   K_UI_Show / K_UI_Hide when the check box is checked.
'   K_UI_LeaveUnchanged when the check box is unchecked or unavailable.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Missing / invalid controls are written to the Immediate Window and
'     treated as LeaveUnchanged.
'
' DEPENDENCIES
'   - Demo_TryGetCheckBoxState
'   - Demo_LogFailure
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IsChecked           As Boolean   'Resolved check-box state
    Dim Msg                 As String    'Diagnostic message from the reader helper

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to LeaveUnchanged unless a checked state is confirmed
        Demo_CheckBoxToVisibility = K_UI_LeaveUnchanged

'------------------------------------------------------------------------------
' READ CHECK-BOX STATE
'------------------------------------------------------------------------------
    'Attempt to read the requested check box
        If Not Demo_TryGetCheckBoxState(Ws, CheckBoxName, IsChecked, Msg) Then

            'Log the control-resolution failure and keep LeaveUnchanged
                Demo_LogFailure CallerProc, CheckBoxName, Msg

            'Exit with default value
                Exit Function

        End If

'------------------------------------------------------------------------------
' MAP CHECK-BOX STATE TO TRI-STATE VISIBILITY
'------------------------------------------------------------------------------
    'Apply the requested visibility only when the check box is checked
        If IsChecked Then
            Demo_CheckBoxToVisibility = SelectedVisibility
        Else
            Demo_CheckBoxToVisibility = K_UI_LeaveUnchanged
        End If

End Function

Private Function Demo_TryGetCheckBoxState( _
    ByVal Ws As Worksheet, _
    ByVal ControlName As String, _
    ByRef IsChecked As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                        Demo_TryGetCheckBoxState
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read the checked state of a demo worksheet check box.
'
' WHY THIS EXISTS
'   The demo workbook may use either:
'     - Forms check boxes
'     - ActiveX check boxes
'
'   This helper supports both models behind a single reader.
'
' INPUTS
'   Ws
'     Worksheet containing the control.
'
'   ControlName
'     Name of the control to inspect.
'
'   IsChecked
'     Receives TRUE when the check box is checked, FALSE otherwise.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => control found and state read successfully
'   FALSE => control missing or invalid
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Returns FALSE and populates FailMsg on failure.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Shp                 As Shape         'Candidate Forms control
    Dim Ole                 As OLEObject     'Candidate ActiveX control
    Dim V                   As Variant       'Late-bound Value property result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize outputs and default result
        Demo_TryGetCheckBoxState = False
        IsChecked = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' TRY FORMS CHECK BOX
'------------------------------------------------------------------------------
    'Attempt to resolve the control as a worksheet shape
        On Error Resume Next
            Set Shp = Ws.Shapes(ControlName)
        On Error GoTo Fail

    'Process the shape when it exists
        If Not Shp Is Nothing Then

            'Reject shapes that are not Forms controls
                If Shp.Type <> msoFormControl Then
                    FailMsg = "shape exists but is not a Forms control"
                    GoTo SafeExit
                End If

            'Reject Forms controls that are not check boxes
                If Shp.FormControlType <> xlCheckBox Then
                    FailMsg = "Forms control exists but is not a CheckBox"
                    GoTo SafeExit
                End If

            'Read the checked state from the Forms check box
                IsChecked = (Shp.ControlFormat.Value = xlOn)

            'Mark success
                Demo_TryGetCheckBoxState = True

            'Exit after successful Forms-control read
                GoTo SafeExit

        End If

'------------------------------------------------------------------------------
' TRY ACTIVEX CHECK BOX
'------------------------------------------------------------------------------
    'Attempt to resolve the control as an ActiveX OLEObject
        On Error Resume Next
            Set Ole = Ws.OLEObjects(ControlName)
        On Error GoTo Fail

    'Process the OLEObject when it exists
        If Not Ole Is Nothing Then

            'Reject ActiveX controls that are not check boxes
                If InStr(1, Ole.progID, "CheckBox", vbTextCompare) = 0 Then
                    FailMsg = "ActiveX control exists but is not a CheckBox"
                    GoTo SafeExit
                End If

            'Read the checked state through late-bound Value access
                V = CallByName(Ole.Object, "Value", VbGet)
                IsChecked = CBool(V)

            'Mark success
                Demo_TryGetCheckBoxState = True

            'Exit after successful ActiveX-control read
                GoTo SafeExit

        End If

'------------------------------------------------------------------------------
' NOT FOUND
'------------------------------------------------------------------------------
    'Report that neither a Forms nor an ActiveX check box was found
        FailMsg = "check box not found"

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = CStr(Err.Number) & ": " & Err.Description & _
                  IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
                  IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

End Function

Private Function Demo_TrySetCheckBoxState( _
    ByVal Ws As Worksheet, _
    ByVal ControlName As String, _
    ByVal IsChecked As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                        Demo_TrySetCheckBoxState
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to write the checked state of a demo worksheet check box.
'
' WHY THIS EXISTS
'   The demo supports selection profiles and current-state synchronization,
'   both of which need to programmatically set the worksheet controls.
'
' INPUTS
'   Ws
'     Worksheet containing the control.
'
'   ControlName
'     Name of the control to update.
'
'   IsChecked
'     Requested target state.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => control found and updated successfully
'   FALSE => control missing or invalid
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Returns FALSE and populates FailMsg on failure.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Shp                 As Shape         'Candidate Forms control
    Dim Ole                 As OLEObject     'Candidate ActiveX control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize default result
        Demo_TrySetCheckBoxState = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' TRY FORMS CHECK BOX
'------------------------------------------------------------------------------
    'Attempt to resolve the control as a worksheet shape
        On Error Resume Next
            Set Shp = Ws.Shapes(ControlName)
        On Error GoTo Fail

    'Process the shape when it exists
        If Not Shp Is Nothing Then

            'Reject shapes that are not Forms controls
                If Shp.Type <> msoFormControl Then
                    FailMsg = "shape exists but is not a Forms control"
                    GoTo SafeExit
                End If

            'Reject Forms controls that are not check boxes
                If Shp.FormControlType <> xlCheckBox Then
                    FailMsg = "Forms control exists but is not a CheckBox"
                    GoTo SafeExit
                End If

            'Write the checked state to the Forms check box
                Shp.ControlFormat.Value = IIf(IsChecked, xlOn, xlOff)

            'Mark success
                Demo_TrySetCheckBoxState = True

            'Exit after successful Forms-control write
                GoTo SafeExit

        End If

'------------------------------------------------------------------------------
' TRY ACTIVEX CHECK BOX
'------------------------------------------------------------------------------
    'Attempt to resolve the control as an ActiveX OLEObject
        On Error Resume Next
            Set Ole = Ws.OLEObjects(ControlName)
        On Error GoTo Fail

    'Process the OLEObject when it exists
        If Not Ole Is Nothing Then

            'Reject ActiveX controls that are not check boxes
                If InStr(1, Ole.progID, "CheckBox", vbTextCompare) = 0 Then
                    FailMsg = "ActiveX control exists but is not a CheckBox"
                    GoTo SafeExit
                End If

            'Write the checked state through late-bound Value access
                CallByName Ole.Object, "Value", VbLet, IsChecked

            'Mark success
                Demo_TrySetCheckBoxState = True

            'Exit after successful ActiveX-control write
                GoTo SafeExit

        End If

'------------------------------------------------------------------------------
' NOT FOUND
'------------------------------------------------------------------------------
    'Report that neither a Forms nor an ActiveX check box was found
        FailMsg = "check box not found"

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = CStr(Err.Number) & ": " & Err.Description & _
                  IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
                  IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

End Function

Private Sub Demo_SetSelectionProfile( _
    ByVal CallerProc As String, _
    ByVal RibbonSelected As Boolean, _
    ByVal StatusBarSelected As Boolean, _
    ByVal ScrollBarsSelected As Boolean, _
    ByVal FormulaBarSelected As Boolean, _
    ByVal HeadingsSelected As Boolean, _
    ByVal WorkbookTabsSelected As Boolean, _
    ByVal GridlinesSelected As Boolean, _
    ByVal TitleBarSelected As Boolean)

'
'==============================================================================
'                         Demo_SetSelectionProfile
'------------------------------------------------------------------------------
' PURPOSE
'   Set all demo check boxes in one call according to the supplied Boolean
'   selection profile.
'
' WHY THIS EXISTS
'   The demo exposes convenience actions such as:
'     - Select All
'     - Clear All
'     - preset bundles
'
'   A shared writer keeps those actions concise and consistent.
'
' INPUTS
'   CallerProc
'     Public caller procedure name used for diagnostics.
'
'   [Boolean selection flags]
'     Requested checked state for each demo check box.
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Per-control failures are written to the Immediate Window.
'
' DEPENDENCIES
'   - Demo_TrySetCheckBoxState
'   - Demo_LogFailure
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ws                  As Worksheet     'Demo worksheet
    Dim Msg                 As String        'Diagnostic message from the writer helper

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Resolve the demo worksheet
        Set Ws = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)

'------------------------------------------------------------------------------
' WRITE SELECTION PROFILE
'------------------------------------------------------------------------------
    'Write the Ribbon selection state
        If Not Demo_TrySetCheckBoxState(Ws, CB_RIBBON, RibbonSelected, Msg) Then
            Demo_LogFailure CallerProc, CB_RIBBON, Msg
        End If

    'Write the StatusBar selection state
        If Not Demo_TrySetCheckBoxState(Ws, CB_STATUSBAR, StatusBarSelected, Msg) Then
            Demo_LogFailure CallerProc, CB_STATUSBAR, Msg
        End If

    'Write the ScrollBars selection state
        If Not Demo_TrySetCheckBoxState(Ws, CB_SCROLLBARS, ScrollBarsSelected, Msg) Then
            Demo_LogFailure CallerProc, CB_SCROLLBARS, Msg
        End If

    'Write the FormulaBar selection state
        If Not Demo_TrySetCheckBoxState(Ws, CB_FORMULABAR, FormulaBarSelected, Msg) Then
            Demo_LogFailure CallerProc, CB_FORMULABAR, Msg
        End If

    'Write the Headings selection state
        If Not Demo_TrySetCheckBoxState(Ws, CB_HEADINGS, HeadingsSelected, Msg) Then
            Demo_LogFailure CallerProc, CB_HEADINGS, Msg
        End If

    'Write the WorkbookTabs selection state
        If Not Demo_TrySetCheckBoxState(Ws, CB_WORKBOOKTABS, WorkbookTabsSelected, Msg) Then
            Demo_LogFailure CallerProc, CB_WORKBOOKTABS, Msg
        End If

    'Write the Gridlines selection state
        If Not Demo_TrySetCheckBoxState(Ws, CB_GRIDLINES, GridlinesSelected, Msg) Then
            Demo_LogFailure CallerProc, CB_GRIDLINES, Msg
        End If

    'Write the TitleBar selection state
        If Not Demo_TrySetCheckBoxState(Ws, CB_TITLEBAR, TitleBarSelected, Msg) Then
            Demo_LogFailure CallerProc, CB_TITLEBAR, Msg
        End If

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write a diagnostic line without interrupting callers
        Demo_LogFailure CallerProc, "Unexpected", _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

    'Exit quietly after logging
        Resume SafeExit

End Sub

Private Function Demo_TryGetRibbonVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                         Demo_TryGetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read current Ribbon visibility.
'
' WHY THIS EXISTS
'   The demo needs to synchronize worksheet check boxes with the current Excel
'   UI state, but Ribbon visibility is not exposed through a simple dedicated
'   Application Boolean property.
'
' INPUTS
'   IsVisible
'     Receives current Ribbon visibility on success.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => Ribbon visibility was read successfully
'   FALSE => read failed
'
' BEHAVIOR
'   - First attempts CommandBars("Ribbon").Visible.
'   - Falls back to an Excel4 macro read when needed.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Returns FALSE and populates FailMsg on failure.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant    'Fallback Excel4 macro result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize outputs and default result
        Demo_TryGetRibbonVisible = False
        IsVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' TRY COMMANDBARS
'------------------------------------------------------------------------------
    'Attempt to read Ribbon visibility from the CommandBars collection
        On Error Resume Next
            IsVisible = Application.CommandBars("Ribbon").Visible
        If Err.Number = 0 Then
            On Error GoTo Fail
            Demo_TryGetRibbonVisible = True
            GoTo SafeExit
        End If
        Err.Clear
        On Error GoTo Fail

'------------------------------------------------------------------------------
' TRY EXCEL4 MACRO FALLBACK
'------------------------------------------------------------------------------
    'Attempt a fallback read using an Excel4 macro
        On Error Resume Next
            V = Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")
        If Err.Number = 0 Then
            On Error GoTo Fail
            IsVisible = CBool(V)
            Demo_TryGetRibbonVisible = True
            GoTo SafeExit
        End If
        FailMsg = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        On Error GoTo Fail

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = CStr(Err.Number) & ": " & Err.Description & _
                  IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
                  IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

End Function

Private Function Demo_TryGetTitleBarVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                        Demo_TryGetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read current title-bar visibility for the Excel window
'   represented by Application.Hwnd.
'
' WHY THIS EXISTS
'   Title-bar state is managed through WinAPI in EXCEL_UI, so the demo needs a
'   corresponding read-side helper for synchronization.
'
' INPUTS
'   IsVisible
'     Receives current title-bar visibility on success.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => title-bar visibility was read successfully
'   FALSE => read failed
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Returns FALSE and populates FailMsg on failure.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim xlHnd               As LongPtr   'Excel window handle from Application.Hwnd
    Dim StyleValue          As LongPtr   'Current window style value
#Else
    Dim xlHnd               As Long      'Excel window handle from Application.Hwnd
    Dim StyleValue          As Long      'Current window style value
#End If
    Dim LastErr             As Long      'Last Win32 error after API call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize outputs and default result
        Demo_TryGetTitleBarVisible = False
        IsVisible = False
        FailMsg = vbNullString

    'Read the Excel window handle
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

    'Read the Win32 last-error value immediately after the API call
        LastErr = Demo_GetLastError

    'Treat zero + nonzero last error as failure
        If StyleValue = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' MAP STYLE TO TITLE-BAR VISIBILITY
'------------------------------------------------------------------------------
    'Treat the caption style bit as the title-bar visibility signal
        IsVisible = ((StyleValue And DEMO_WS_CAPTION) <> 0)

    'Mark success after a valid style read
        Demo_TryGetTitleBarVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = CStr(Err.Number) & ": " & Err.Description & _
                  IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
                  IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

End Function

Private Function Demo_HasAnyRequestedChange( _
    ByVal RibbonVis As K_UIVisibility, _
    ByVal StatusBarVis As K_UIVisibility, _
    ByVal ScrollBarsVis As K_UIVisibility, _
    ByVal FormulaBarVis As K_UIVisibility, _
    ByVal HeadingsVis As K_UIVisibility, _
    ByVal WorkbookTabsVis As K_UIVisibility, _
    ByVal GridlinesVis As K_UIVisibility, _
    ByVal TitleBarVis As K_UIVisibility) As Boolean

'
'==============================================================================
'                        Demo_HasAnyRequestedChange
'------------------------------------------------------------------------------
' PURPOSE
'   Determine whether at least one UI element has been selected for change.
'
' WHY THIS EXISTS
'   The demo macros should inform the user when no check boxes are selected,
'   rather than silently doing nothing.
'
' RETURNS
'   TRUE  => at least one argument differs from K_UI_LeaveUnchanged
'   FALSE => all arguments are K_UI_LeaveUnchanged
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return TRUE when at least one requested visibility is actionable
        Demo_HasAnyRequestedChange = ( _
            RibbonVis <> K_UI_LeaveUnchanged _
            Or StatusBarVis <> K_UI_LeaveUnchanged _
            Or ScrollBarsVis <> K_UI_LeaveUnchanged _
            Or FormulaBarVis <> K_UI_LeaveUnchanged _
            Or HeadingsVis <> K_UI_LeaveUnchanged _
            Or WorkbookTabsVis <> K_UI_LeaveUnchanged _
            Or GridlinesVis <> K_UI_LeaveUnchanged _
            Or TitleBarVis <> K_UI_LeaveUnchanged)

End Function

Public Sub Demo_CreateExcelUISheet()

'
'==============================================================================
'                         Demo_CreateExcelUISheet
'------------------------------------------------------------------------------
' PURPOSE
'   Create or rebuild the EXCEL_UI demo worksheet and place all required
'   labels, formatting, notes, check boxes, and action shapes.
'
' WHY THIS EXISTS
'   Manually creating the demo sheet is repetitive and error-prone. This
'   routine produces a consistent, presentation-ready layout in one step.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Resolves or creates worksheet "Demo".
'   - Clears prior content and controls.
'   - Applies layout and formatting.
'   - Creates the section blocks and their labels.
'   - Adds eight Forms check boxes with the expected names.
'   - Adds action buttons for:
'       * SHOW selected
'       * HIDE selected
'       * Sync Checkboxes
'       * Select All
'       * Clear All
'       * Kiosk preset
'       * Analyst preset
'       * Minimal preset
'   - Adds explanatory notes to the demo sheet.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Unexpected failures are written to the Immediate Window.
'
' DEPENDENCIES
'   - Demo_GetOrCreateSheet
'   - Demo_ResetSheet
'   - Demo_FormatSheetLayout
'   - Demo_WriteStaticLabels
'   - Demo_WriteNotes
'   - Demo_AddFormsCheckBox
'   - Demo_AddActionButton
'   - Demo_LogFailure
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Wb                  As Workbook     'Owning workbook
    Dim Ws                  As Worksheet    'Demo worksheet
    Dim OldScreenUpdating   As Boolean      'Previous ScreenUpdating state

    Const PROC As String = "Demo_CreateExcelUISheet"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Cache and suppress screen updates during rebuild
        OldScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False

    'Resolve the owning workbook
        Set Wb = ThisWorkbook

    'Resolve the demo worksheet or create it when missing
        Set Ws = Demo_GetOrCreateSheet(Wb, DEMO_SHEET_NAME)

'------------------------------------------------------------------------------
' RESET / FORMAT SHEET
'------------------------------------------------------------------------------
    'Reset the worksheet to a clean state before rebuilding the demo
        Demo_ResetSheet Ws
    'Apply the base layout and formatting for the demo surface
        Demo_FormatSheetLayout Ws
    'Write all static labels and section headers
        Demo_WriteStaticLabels Ws
    'Write explanatory notes to the lower part of the demo sheet
        Demo_WriteNotes Ws

'------------------------------------------------------------------------------
' ADD CHECK BOXES
'------------------------------------------------------------------------------
    'Add the Ribbon check box
        Demo_AddFormsCheckBox Ws, CB_RIBBON, Ws.Range("K6")
    'Add the StatusBar check box
        Demo_AddFormsCheckBox Ws, CB_STATUSBAR, Ws.Range("K7")
    'Add the ScrollBars check box
        Demo_AddFormsCheckBox Ws, CB_SCROLLBARS, Ws.Range("K8")
    'Add the FormulaBar check box
        Demo_AddFormsCheckBox Ws, CB_FORMULABAR, Ws.Range("K9")
    'Add the Headings check box
        Demo_AddFormsCheckBox Ws, CB_HEADINGS, Ws.Range("K12")
    'Add the WorkbookTabs check box
        Demo_AddFormsCheckBox Ws, CB_WORKBOOKTABS, Ws.Range("K13")
    'Add the Gridlines check box
        Demo_AddFormsCheckBox Ws, CB_GRIDLINES, Ws.Range("K14")
    'Add the TitleBar check box
        Demo_AddFormsCheckBox Ws, CB_TITLEBAR, Ws.Range("K15")

'------------------------------------------------------------------------------
' ADD ACTION BUTTONS
'------------------------------------------------------------------------------
    'Add the SHOW action shape and wire it to the show macro
        Demo_AddActionButton _
            Ws:=Ws, _
            ShapeName:=BTN_SHOW_NAME, _
            TargetRange:=Ws.Range("D6:F7"), _
            CaptionText:="SHOW SELECTED UI", _
            MacroName:=BTN_SHOW_MACRO, _
            FillColor:=RGB(0, 102, 153), _
            FontColor:=RGB(255, 255, 255)

    'Add the HIDE action shape and wire it to the hide macro
        Demo_AddActionButton _
            Ws:=Ws, _
            ShapeName:=BTN_HIDE_NAME, _
            TargetRange:=Ws.Range("D9:F10"), _
            CaptionText:="HIDE SELECTED UI", _
            MacroName:=BTN_HIDE_MACRO, _
            FillColor:=RGB(192, 80, 0), _
            FontColor:=RGB(255, 255, 255)

    'Add the Sync Checkboxes shape and wire it to the sync macro
        Demo_AddActionButton _
            Ws:=Ws, _
            ShapeName:=BTN_SYNC_NAME, _
            TargetRange:=Ws.Range("D12:F13"), _
            CaptionText:="SYNC CHECKBOXES", _
            MacroName:=BTN_SYNC_MACRO, _
            FillColor:=RGB(79, 129, 189), _
            FontColor:=RGB(255, 255, 255)

    'Add the Select All shape and wire it to the select-all macro
        Demo_AddActionButton _
            Ws:=Ws, _
            ShapeName:=BTN_SELECTALL_NAME, _
            TargetRange:=Ws.Range("D15:E16"), _
            CaptionText:="SELECT ALL", _
            MacroName:=BTN_SELECTALL_MACRO, _
            FillColor:=RGB(84, 130, 53), _
            FontColor:=RGB(255, 255, 255)

    'Add the Clear All shape and wire it to the clear-all macro
        Demo_AddActionButton _
            Ws:=Ws, _
            ShapeName:=BTN_CLEARALL_NAME, _
            TargetRange:=Ws.Range("F15:G16"), _
            CaptionText:="CLEAR ALL", _
            MacroName:=BTN_CLEARALL_MACRO, _
            FillColor:=RGB(127, 127, 127), _
            FontColor:=RGB(255, 255, 255)

    'Add the Kiosk preset shape and wire it to the preset macro
        Demo_AddActionButton _
            Ws:=Ws, _
            ShapeName:=BTN_PRESET_KIOSK_NAME, _
            TargetRange:=Ws.Range("D19:E20"), _
            CaptionText:="KIOSK", _
            MacroName:=BTN_PRESET_KIOSK_MACRO, _
            FillColor:=RGB(31, 73, 125), _
            FontColor:=RGB(255, 255, 255)

    'Add the Analyst preset shape and wire it to the preset macro
        Demo_AddActionButton _
            Ws:=Ws, _
            ShapeName:=BTN_PRESET_ANALYST_NAME, _
            TargetRange:=Ws.Range("F19:G20"), _
            CaptionText:="ANALYST", _
            MacroName:=BTN_PRESET_ANALYST_MACRO, _
            FillColor:=RGB(49, 133, 156), _
            FontColor:=RGB(255, 255, 255)

    'Add the Minimal preset shape and wire it to the preset macro
        Demo_AddActionButton _
            Ws:=Ws, _
            ShapeName:=BTN_PRESET_MINIMAL_NAME, _
            TargetRange:=Ws.Range("D21:G22"), _
            CaptionText:="MINIMAL", _
            MacroName:=BTN_PRESET_MINIMAL_MACRO, _
            FillColor:=RGB(148, 138, 84), _
            FontColor:=RGB(255, 255, 255)

'------------------------------------------------------------------------------
' FINALIZE
'------------------------------------------------------------------------------
    'Select the demo sheet so the result is immediately visible
        Ws.Activate
    'Position the selection at the top-left useful cell
        Ws.Range("B3").Select

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Restore ScreenUpdating before exiting
        Application.ScreenUpdating = OldScreenUpdating
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write the failure to the Immediate Window without interrupting callers
        Demo_LogFailure PROC, "Unexpected", _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

    'Exit quietly after logging
        Resume SafeExit

End Sub

Private Function Demo_GetOrCreateSheet( _
    ByVal Wb As Workbook, _
    ByVal SheetName As String) As Worksheet

'
'==============================================================================
'                          Demo_GetOrCreateSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Return the requested worksheet when it exists, otherwise create it.
'
' WHY THIS EXISTS
'   The demo builder should work both the first time and on subsequent rebuilds
'   without requiring manual worksheet preparation.
'
' INPUTS
'   Wb
'     Workbook that will contain the sheet.
'
'   SheetName
'     Requested worksheet name.
'
' RETURNS
'   The resolved worksheet.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ws                  As Worksheet    'Resolved worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Attempt to resolve the worksheet by name
        On Error Resume Next
            Set Ws = Wb.Worksheets(SheetName)
        On Error GoTo 0

'------------------------------------------------------------------------------
' CREATE WHEN MISSING
'------------------------------------------------------------------------------
    'Create the worksheet when it does not already exist
        If Ws Is Nothing Then

            'Add a new worksheet at the end of the workbook
                Set Ws = Wb.Worksheets.Add(After:=Wb.Worksheets(Wb.Worksheets.Count))

            'Assign the requested worksheet name
                Ws.Name = SheetName

        End If

'------------------------------------------------------------------------------
' RETURN WORKSHEET
'------------------------------------------------------------------------------
    'Return the resolved worksheet reference
        Set Demo_GetOrCreateSheet = Ws

End Function

Private Sub Demo_ResetSheet(ByVal Ws As Worksheet)

'
'==============================================================================
'                               Demo_ResetSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Clear prior content, formatting, shapes, and OLEObjects from the demo
'   worksheet before rebuilding it.
'
' WHY THIS EXISTS
'   Rebuilding on top of previous controls creates duplicate shapes and check
'   boxes. A clean reset guarantees deterministic output.
'
' INPUTS
'   Ws
'     Demo worksheet to reset.
'
' RETURNS
'   None
'
' NOTES
'   - This is a destructive reset of the Demo sheet.
'   - Do not store unrelated content on that sheet.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long         'Reverse loop index for shapes / OLEObjects

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' DELETE SHAPES
'------------------------------------------------------------------------------
    'Delete all existing shapes in reverse order
        For i = Ws.Shapes.Count To 1 Step -1
            Ws.Shapes(i).Delete
        Next i

'------------------------------------------------------------------------------
' DELETE OLEOBJECTS
'------------------------------------------------------------------------------
    'Delete all existing OLEObjects in reverse order
        For i = Ws.OLEObjects.Count To 1 Step -1
            Ws.OLEObjects(i).Delete
        Next i

'------------------------------------------------------------------------------
' CLEAR CELLS
'------------------------------------------------------------------------------
    'Clear all cell content, formatting, comments, and validation
        Ws.Cells.Clear

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Exit quietly after partial cleanup
        Resume SafeExit

End Sub

Private Sub Demo_FormatSheetLayout(ByVal Ws As Worksheet)

'
'==============================================================================
'                           Demo_FormatSheetLayout
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the worksheet layout and visual formatting for the demo surface.
'
' WHY THIS EXISTS
'   Separating layout logic from control-creation logic makes the demo builder
'   easier to maintain and adjust.
'
' INPUTS
'   Ws
'     Demo worksheet to format.
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY GLOBAL SHEET FORMATTING
'------------------------------------------------------------------------------
    'Set the worksheet tab color
        Ws.Tab.Color = RGB(0, 102, 153)
    'Set the default worksheet font name
        Ws.Cells.Font.Name = "Calibri"
    'Set the default worksheet font size
        Ws.Cells.Font.Size = 11
    'Set horizontal alignment baseline
        Ws.Cells.HorizontalAlignment = xlLeft
    'Set vertical alignment baseline
        Ws.Cells.VerticalAlignment = xlCenter

'------------------------------------------------------------------------------
' APPLY COLUMN / ROW LAYOUT
'------------------------------------------------------------------------------
    'Set a narrow left margin column
        Ws.Columns("A").ColumnWidth = 2

    'Set the title / subtitle / main content block widths
        Ws.Columns("B").ColumnWidth = 2
        Ws.Columns("C").ColumnWidth = 12
        Ws.Columns("D").ColumnWidth = 11
        Ws.Columns("E").ColumnWidth = 11
        Ws.Columns("F").ColumnWidth = 11
        Ws.Columns("G").ColumnWidth = 11
        Ws.Columns("H").ColumnWidth = 7
        Ws.Columns("I").ColumnWidth = 14
        Ws.Columns("J").ColumnWidth = 12
        Ws.Columns("K").ColumnWidth = 8
        Ws.Columns("L").ColumnWidth = 8
        Ws.Columns("M").ColumnWidth = 8

    'Set the main demo row heights
        Ws.Rows("1:32").RowHeight = 22

    'Increase the note rows for wrapped text
        Ws.Rows("25:31").RowHeight = 28

'------------------------------------------------------------------------------
' FORMAT TITLE / SUBTITLE BANDS
'------------------------------------------------------------------------------
    'Merge and format the title band
        With Ws.Range("B1:M1")
            .Merge
            .Value = "EXCEL UI"
            .Interior.Color = RGB(0, 0, 0)
            .Font.Color = RGB(255, 192, 0)
            .Font.Bold = True
            .Font.Size = 16
        End With

    'Merge and format the subtitle band
        With Ws.Range("B2:M2")
            .Merge
            .Value = "Demo"
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 14
        End With

'------------------------------------------------------------------------------
' FORMAT APPLICATION-LEVEL BLOCK
'------------------------------------------------------------------------------
    'Format the application-level section header
        With Ws.Range("I5:K5")
            .Merge
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Value = "APPLICATION LEVEL UI STATE"
        End With

    'Format the application-level label cells
        With Ws.Range("I6:J9")
            .Interior.Color = RGB(220, 230, 241)
            .Font.Bold = True
        End With

    'Format the application-level check-box cells
        With Ws.Range("K6:K9")
            .Interior.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

    'Apply borders to the application-level block
        With Ws.Range("I5:K9").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With

'------------------------------------------------------------------------------
' FORMAT WINDOW-LEVEL BLOCK
'------------------------------------------------------------------------------
    'Format the window-level section header
        With Ws.Range("I11:K11")
            .Merge
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Value = "WINDOW-LEVEL UI STATE"
        End With

    'Format the window-level label cells
        With Ws.Range("I12:J15")
            .Interior.Color = RGB(220, 230, 241)
            .Font.Bold = True
        End With

    'Format the window-level check-box cells
        With Ws.Range("K12:K15")
            .Interior.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With

    'Apply borders to the window-level block
        With Ws.Range("I11:K15").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With

'------------------------------------------------------------------------------
' FORMAT PRESET LABEL AREA
'------------------------------------------------------------------------------
    'Format the preset section label
        With Ws.Range("D18:G18")
            .Merge
            .Interior.Color = RGB(230, 230, 230)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Value = "PRESET SELECTIONS"
        End With

'------------------------------------------------------------------------------
' FORMAT NOTE AREAS
'------------------------------------------------------------------------------
    'Format the scope / semantics note area
        With Ws.Range("B25:M28")
            .Merge
            .Interior.Color = RGB(255, 242, 204)
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
            .WrapText = True
            .VerticalAlignment = xlTop
        End With

    'Format the restore note area
        With Ws.Range("B29:M31")
            .Merge
            .Interior.Color = RGB(217, 225, 242)
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
            .WrapText = True
            .VerticalAlignment = xlTop
        End With

    'Apply borders to the note areas
        With Ws.Range("B25:M31").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Exit quietly after partial formatting
        Resume SafeExit

End Sub

Private Sub Demo_WriteStaticLabels(ByVal Ws As Worksheet)

'
'==============================================================================
'                           Demo_WriteStaticLabels
'------------------------------------------------------------------------------
' PURPOSE
'   Write the fixed labels used by the demo sheet.
'
' WHY THIS EXISTS
'   Keeping text assignment separate from formatting and control creation makes
'   the builder easier to maintain.
'
' INPUTS
'   Ws
'     Demo worksheet receiving the fixed labels.
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' WRITE APPLICATION-LEVEL LABELS
'------------------------------------------------------------------------------
    'Write the Ribbon label
        Ws.Range("I6").Value = "Ribbon"

    'Write the StatusBar label
        Ws.Range("I7").Value = "StatusBar"

    'Write the ScrollBars label
        Ws.Range("I8").Value = "ScrollBars"

    'Write the FormulaBar label
        Ws.Range("I9").Value = "FormulaBar"

'------------------------------------------------------------------------------
' WRITE WINDOW-LEVEL LABELS
'------------------------------------------------------------------------------
    'Write the Headings label
        Ws.Range("I12").Value = "Headings"

    'Write the WorkbookTabs label
        Ws.Range("I13").Value = "WorkbookTabs"

    'Write the Gridlines label
        Ws.Range("I14").Value = "Gridlines"

    'Write the TitleBar label
        Ws.Range("I15").Value = "TitleBar"

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Exit quietly after partial label assignment
        Resume SafeExit

End Sub

Private Sub Demo_WriteNotes(ByVal Ws As Worksheet)

'
'==============================================================================
'                             Demo_WriteNotes
'------------------------------------------------------------------------------
' PURPOSE
'   Write the explanatory notes displayed at the bottom of the demo sheet.
'
' WHY THIS EXISTS
'   The demo becomes easier to understand when users can read the key semantics
'   and limitations directly on the worksheet.
'
' INPUTS
'   Ws
'     Demo worksheet receiving the note text.
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' WRITE NOTES
'------------------------------------------------------------------------------
    'Write the scope / semantics note text
        Ws.Range("B25").Value = NOTE_SCOPE_TEXT
    'Write the restore note text
        Ws.Range("B29").Value = NOTE_RESTORE_TEXT

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Exit quietly after partial note assignment
        Resume SafeExit

End Sub

Private Sub Demo_AddFormsCheckBox( _
    ByVal Ws As Worksheet, _
    ByVal CheckBoxName As String, _
    ByVal AnchorCell As Range)

'
'==============================================================================
'                           Demo_AddFormsCheckBox
'------------------------------------------------------------------------------
' PURPOSE
'   Add a Forms check box centered inside the supplied anchor cell.
'
' WHY THIS EXISTS
'   The demo sheet uses simple worksheet check boxes as selectors for which
'   UI elements the user wants to affect.
'
' INPUTS
'   Ws
'     Worksheet receiving the check box.
'
'   CheckBoxName
'     Name assigned to the Forms check box.
'
'   AnchorCell
'     Cell used as the placement anchor for the control.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a Forms check box.
'   - Centers it inside the anchor cell.
'   - Removes the caption text.
'   - Initializes the control to unchecked.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cb                  As CheckBox    'Created Forms check box
    Dim BoxWidth            As Double      'Check-box width
    Dim BoxHeight           As Double      'Check-box height
    Dim BoxLeft             As Double      'Resolved left position
    Dim BoxTop              As Double      'Resolved top position

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Define the visual check-box size
        BoxWidth = 18
        BoxHeight = 16

    'Center the control horizontally inside the anchor cell
        BoxLeft = AnchorCell.Left + (AnchorCell.Width - BoxWidth) / 2

    'Center the control vertically inside the anchor cell
        BoxTop = AnchorCell.Top + (AnchorCell.Height - BoxHeight) / 2

'------------------------------------------------------------------------------
' CREATE CHECK BOX
'------------------------------------------------------------------------------
    'Create the Forms check box
        Set Cb = Ws.CheckBoxes.Add(BoxLeft, BoxTop, BoxWidth, BoxHeight)

    'Assign the requested control name
        Cb.Name = CheckBoxName

    'Remove the check-box caption
        Cb.Caption = vbNullString

    'Initialize the control to unchecked
        Cb.Value = xlOff

    'Clear any linked cell
        Cb.LinkedCell = vbNullString

    'Tie the control position and size to its cells
        Cb.Placement = xlMoveAndSize

    'Ensure the control is printed with the sheet
        Cb.PrintObject = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Exit quietly after partial control creation
        Resume SafeExit

End Sub

Private Sub Demo_AddActionButton( _
    ByVal Ws As Worksheet, _
    ByVal ShapeName As String, _
    ByVal TargetRange As Range, _
    ByVal CaptionText As String, _
    ByVal MacroName As String, _
    ByVal FillColor As Long, _
    ByVal FontColor As Long)

'
'==============================================================================
'                            Demo_AddActionButton
'------------------------------------------------------------------------------
' PURPOSE
'   Add a rounded-rectangle action shape over a target range and assign its
'   click behavior to the requested macro.
'
' WHY THIS EXISTS
'   The demo sheet needs visually prominent controls that users can click to
'   trigger the various demo routines.
'
' INPUTS
'   Ws
'     Worksheet receiving the action shape.
'
'   ShapeName
'     Name assigned to the created shape.
'
'   TargetRange
'     Range used for shape placement and sizing.
'
'   CaptionText
'     Button caption displayed inside the shape.
'
'   MacroName
'     Macro name assigned to the shape's OnAction property.
'
'   FillColor
'     Button fill color.
'
'   FontColor
'     Button text color.
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Shp                 As Shape    'Created action shape

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' CREATE SHAPE
'------------------------------------------------------------------------------
    'Create the rounded-rectangle shape over the target range
        Set Shp = Ws.Shapes.AddShape( _
                        Type:=msoShapeRoundedRectangle, _
                        Left:=TargetRange.Left, _
                        Top:=TargetRange.Top, _
                        Width:=TargetRange.Width, _
                        Height:=TargetRange.Height)

    'Assign the requested shape name
        Shp.Name = ShapeName

    'Tie the shape position and size to its cells
        Shp.Placement = xlMoveAndSize

    'Hide the outline
        Shp.Line.Visible = msoFalse

    'Apply the requested fill color
        Shp.Fill.ForeColor.RGB = FillColor

    'Assign the workbook-qualified macro to the shape click action
        Shp.OnAction = "'" & ThisWorkbook.Name & "'!" & MacroName

'------------------------------------------------------------------------------
' APPLY TEXT FORMATTING
'------------------------------------------------------------------------------
    'Write the button caption
        Shp.TextFrame2.TextRange.Characters.Text = CaptionText

    'Center the text horizontally
        Shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

    'Center the text vertically
        Shp.TextFrame2.VerticalAnchor = msoAnchorMiddle

    'Apply the button font size
        Shp.TextFrame2.TextRange.Font.Size = 11

    'Apply bold text
        Shp.TextFrame2.TextRange.Font.Bold = msoTrue

    'Apply the requested font color
        Shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = FontColor

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Exit quietly after partial shape creation
        Resume SafeExit

End Sub

Private Sub Demo_LogFailure( _
    ByVal ProcName As String, _
    ByVal Stage As String, _
    ByVal Detail As String)

'
'==============================================================================
'                            Demo_LogFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Write a consistent diagnostic line to the Immediate Window for the demo
'   module.
'
' WHY THIS EXISTS
'   The demo uses fail-soft behavior and needs a single place to format
'   diagnostics consistently.
'
' INPUTS
'   ProcName
'     Procedure name associated with the failure.
'
'   Stage
'     Logical stage associated with the failure.
'
'   Detail
'     Failure detail to append.
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Suppresses any unexpected logging failure locally.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected logging failure
        On Error Resume Next

'------------------------------------------------------------------------------
' WRITE DIAGNOSTIC LINE
'------------------------------------------------------------------------------
    'Write a consistent diagnostic line to the Immediate Window
        Debug.Print ProcName & " failed @ " & Stage & " | " & Detail

End Sub

