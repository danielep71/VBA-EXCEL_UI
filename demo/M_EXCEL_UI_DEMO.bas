Attribute VB_Name = "M_EXCEL_UI_DEMO"
'
'==============================================================================
'                           MODULE: EXCEL_UI_DEMO
'------------------------------------------------------------------------------
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
'     writing control state
'   - Relies on the public API exposed by the EXCEL_UI module:
'       * UIVisibility
'       * UI_SetExcelUI
'       * UI_HideExcelUI
'       * UI_ShowExcelUI
'       * UI_CaptureExcelUIState
'       * UI_ResetExcelUIToSnapshot
'       * UI_HasExcelUIStateSnapshot
'
' NOTES
'   - Window-level sync reads the current ActiveWindow state
'   - TitleBar sync reads the Excel window represented by Application.Hwnd
'   - Missing or misnamed controls are logged to the Immediate Window and
'     treated conservatively
'   - Demo_CreateDemoSheet is the canonical builder
'
' UPDATED
'   2026-04-19
'
' AUTHOR
'   Daniele Penza
'
'==============================================================================
'

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit declaration of all variables
    
'------------------------------------------------------------------------------
' DEMO CONFIGURATION
'------------------------------------------------------------------------------
    Private Const DEMO_SHEET_NAME         As String = "DEMO_UI"                 'Demo worksheet name
    
    Private Const CB_RIBBON               As String = "chkRibbon"               'Ribbon check box name
    Private Const CB_STATUSBAR            As String = "chkStatusBar"            'StatusBar check box name
    Private Const CB_SCROLLBARS           As String = "chkScrollBars"           'ScrollBars check box name
    Private Const CB_FORMULABAR           As String = "chkFormulaBar"           'FormulaBar check box name
    Private Const CB_HEADINGS             As String = "chkHeadings"             'Headings check box name
    Private Const CB_WORKBOOKTABS         As String = "chkWorkbookTabs"         'WorkbookTabs check box name
    Private Const CB_GRIDLINES            As String = "chkGridlines"            'Gridlines check box name
    Private Const CB_TITLEBAR             As String = "chkTitleBar"             'TitleBar check box name
    
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
        Private Declare PtrSafe Function Demo_GetLastError Lib "kernel32" Alias _
            "GetLastError" () As Long
        Private Declare PtrSafe Sub Demo_SetLastError Lib "kernel32" Alias _
            "SetLastError" (ByVal dwErrCode As Long)
    #Else
        Private Declare Function Demo_GetWindowLong Lib "user32" Alias _
            "GetWindowLongA" ( ByVal hWnd As Long, ByVal nIndex As Long) As Long
        Private Declare Function Demo_GetLastError Lib "kernel32" Alias _
            "GetLastError" () As Long
        Private Declare Sub Demo_SetLastError Lib "kernel32" Alias _
            "SetLastError" ( ByVal dwErrCode As Long)
    #End If

'------------------------------------------------------------------------------
' API CONSTANTS FOR TITLE-BAR STATE READ
'------------------------------------------------------------------------------
    Private Const DEMO_GWL_STYLE          As Long = -16       'Window style index
    Private Const DEMO_WS_CAPTION         As Long = &HC00000  'Caption / title-bar style bit



'                               '
'------------------------------------------------------------------------------
'
'                           PUBLIC DEMO BUILDER
'
'------------------------------------------------------------------------------
'

Public Sub Demo_CreateDemoSheet()

'
'==============================================================================
'                            Demo_CreateDemoSheet
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
'   None
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
'   - Restores cursor and fast-mode state through a centralized cleanup path
'
' ERROR POLICY
'   - Raises errors normally
'   - Cleanup is best-effort and should not overwrite the original error
'
' DEPENDENCIES
'   - DEMO_FastMode_Begin
'   - DEMO_FastMode_End
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
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WB                  As Workbook             'Target workbook
    Dim WS                  As Worksheet            'Demo worksheet
    Dim ButtonSpecs         As Variant              'Button name / caption / macro specification
    Dim FastModeState       As tDEMOFastModeState   'Saved Application-state snapshot
    Dim FastModeOn          As Boolean              'TRUE when fast mode was entered
    Dim SavedErrNumber      As Long                 'Captured error number
    Dim SavedErrSource      As String               'Captured error source
    Dim SavedErrDescription As String               'Captured error description

    Const PROC As String = "Demo_CreateDemoSheet"     'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo Clean_Fail
    'Target the workbook that contains this module
        Set WB = ThisWorkbook
    'Simulate Button click
        DEMO_Btn_Click
    'Capture and apply fast-mode Application settings
        DEMO_FastMode_Begin FastModeState
        FastModeOn = True
    'Show the wait cursor while rebuilding the demo workbook
        Application.Cursor = xlWait
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
    'Write the Ribbon input row
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G5"), WS.Range("H5"), _
            "Ribbon", ""
    'Write the Status Bar input row
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G6"), WS.Range("H6"), _
            "Status bar", ""
    'Write the Scroll Bars input row
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G7"), WS.Range("H7"), _
            "Scroll bars", ""
    'Write the Formula Bar input row
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
    'Write the Headings input row
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G12"), WS.Range("H12"), _
            "Headings", ""
    'Write the Workbook Tabs input row
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G13"), WS.Range("H13"), _
            "Workbook tabs", ""
    'Write the Gridlines input row
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G14"), WS.Range("H14"), _
            "Gridlines", ""
    'Write the Title Bar input row
        DEMO_Write_NamedInputRow WB, WS, WS.Range("G15"), WS.Range("H15"), _
            "Title bar", ""
    'Convert the window-level input cells to checkboxes
        DEMO_CB_AddForms WS, WS.Range("H12:H15"), Array(CB_HEADINGS, _
            CB_WORKBOOKTABS, CB_GRIDLINES, CB_TITLEBAR)

'------------------------------------------------------------------------------
' BUILD ACTION BUTTONS
'------------------------------------------------------------------------------
    'Write the action section header
        DEMO_Write_BandHeader WS.Range("C4:E4"), "ACTIONS"
    'Define the action button grid
        ButtonSpecs = Array(Array("btn_UI_Show", "SHOW SELECTED UI", _
            "Demo_ShowSelectedUI"), Array("btn_UI_Hide", "HIDE SELECTED UI", _
            "Demo_HideSelectedUI"))
    'Create the standard two-column action-button grid
        DEMO_Btn_AddGrid WS, WS.Range("C5"), ButtonSpecs, 2, 150, 25
    'Apply a border around the full action-button area
        DEMO_Set_RangeBorder WS.Range("C4:E6")

'------------------------------------------------------------------------------
' BUILD SELECT / CLEAR BUTTONS
'------------------------------------------------------------------------------
    'Write the select / clear section header
        DEMO_Write_BandHeader WS.Range("C8:E8"), "SELECT / CLEAR"
    'Define the select / clear button grid
        ButtonSpecs = Array(Array("btn_UI_SelectAll", "SELECT ALL", _
            "Demo_SelectAllUI"), Array("btn_UI_ClearAll", "CLEAR ALL", _
            "Demo_ClearAllUI"))
    'Create the standard two-column select / clear button grid
        DEMO_Btn_AddGrid WS, WS.Range("C9"), ButtonSpecs, 2, 150, 25
    'Apply a border around the full select / clear area
        DEMO_Set_RangeBorder WS.Range("C8:E10")

'------------------------------------------------------------------------------
' BUILD PRESET / UTILITY BUTTONS
'------------------------------------------------------------------------------
    'Write the preset selection section header
        DEMO_Write_BandHeader WS.Range("C12:E12"), "PRESET SELECTION"
    'Define the preset / utility button grid
        ButtonSpecs = Array(Array("btn_UI_Kiosk", "KIOSK", "Demo_PresetKiosk"), _
            Array("btn_UI_Analyst", "ANALYST", "Demo_PresetAnalyst"), _
            Array("btn_UI_Minimal", "MINIMAL", "Demo_PresetMinimal"))
    'Create the standard two-column preset / utility button grid
        DEMO_Btn_AddGrid WS, WS.Range("C13"), ButtonSpecs, 2, 150, 25, , 13, , _
            8
    'Apply a border around the full preset / utility area
        DEMO_Set_RangeBorder WS.Range("C12:E16")

'------------------------------------------------------------------------------
' BUILD CAPTURE / RESET STATE BUTTONS
'------------------------------------------------------------------------------
    'Write the capture / reset state section header
        DEMO_Write_BandHeader WS.Range("C18:E18"), "CAPTURE / RESET STATE"
    'Define the capture / reset button grid
        ButtonSpecs = Array(Array("btn_UI_CaptureState", "CAPTURE STATE", _
            "Demo_CaptureUIState"), Array("btn_UI_ResetState", "RESET STATE", _
            "Demo_ResetUIToCapturedState"))
    'Create the standard two-column capture / reset button grid
        DEMO_Btn_AddGrid WS, WS.Range("C19"), ButtonSpecs, 2, 150, 25, , 13, , _
            8
    'Apply a border around the full capture / reset area
        DEMO_Set_RangeBorder WS.Range("C18:E20")
        
'------------------------------------------------------------------------------
' BUILD SYNC CHECK BOXES BUTTON
'------------------------------------------------------------------------------
    'Create the sync button anchored to a dedicated layout block
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
    'Apply borders to the scope / semantics note area
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
    'Apply borders to the restore note area
        DEMO_Set_RangeBorder WS.Range("B26:H27")

'------------------------------------------------------------------------------
' WRITE NOTES
'------------------------------------------------------------------------------
    'Write the scope / semantics note text
        WS.Range("B22").Value = NOTE_SCOPE_TEXT
    'Write the restore note text
        WS.Range("B26").Value = NOTE_RESTORE_TEXT

'------------------------------------------------------------------------------
' SYNCHRONIZE CHECK BOXES
'------------------------------------------------------------------------------
    'Synchronize the check boxes back to the current visible UI state
        Demo_SyncCheckBoxesToUI False

'------------------------------------------------------------------------------
' BUILD RESET SHEET BUTTON
'------------------------------------------------------------------------------
    'Create the reset sheet button
        DEMO_Btn_Add WS, "btn_UI_ResetSheet", "RESET SHEET", _
            WS.Range("H2").Left + 1, WS.Range("H2").Top + 1, 105, 25, _
            "Demo_CreateDemoSheet"

Clean_Exit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Protect cleanup so it cannot overwrite the original error
        On Error Resume Next
    'Restore the normal cursor
        Application.Cursor = xlDefault
    'Restore the original Excel Application state only when fast mode was entered
        If FastModeOn Then
            DEMO_FastMode_End FastModeState
        End If
    'Restore normal error handling before any re-raise
        On Error GoTo 0
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If
    'Normal termination point
        Exit Sub

Clean_Fail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = IIf(Len(Err.Source) > 0, Err.Source, PROC)
        SavedErrDescription = Err.Description
    'Continue through the centralized cleanup path
        Resume Clean_Exit

End Sub


'
'------------------------------------------------------------------------------
'
'                           PUBLIC DEMO ACTIONS
'
'------------------------------------------------------------------------------
'

Public Sub Demo_ShowSelectedUI()

'
'==============================================================================
'                        Demo_ShowSelectedUI
'------------------------------------------------------------------------------
' PURPOSE
'   Show only the UI elements currently selected by the user on the demo sheet
'
' WHY THIS EXISTS
'   The demo sheet uses check boxes as a user-friendly selector for which UI
'   elements should be affected by the EXCEL_UI module
'
' RETURNS
'   None
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
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY ACTION
'------------------------------------------------------------------------------
    'Play the optional button-click visual feedback
        DEMO_Btn_PlayFeedback
    'Delegate the action to the shared worker
        Demo_ApplySelectedUI UI_Show, "Demo_ShowSelectedUI"

End Sub

Public Sub Demo_HideSelectedUI()

'
'==============================================================================
'                        Demo_HideSelectedUI
'------------------------------------------------------------------------------
' PURPOSE
'   Hide only the UI elements currently selected by the user on the demo sheet
'
' WHY THIS EXISTS
'   The demo sheet uses check boxes as a user-friendly selector for which UI
'   elements should be affected by the EXCEL_UI module
'
' RETURNS
'   None
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
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY ACTION
'------------------------------------------------------------------------------
    'Play the optional button-click visual feedback
        DEMO_Btn_PlayFeedback
    'Delegate the action to the shared worker
        Demo_ApplySelectedUI UI_Hide, "Demo_HideSelectedUI"

End Sub

Public Sub Demo_SyncCheckBoxesToUI(Optional ByVal PlayFeedback As Boolean = _
    True)
'
'==============================================================================
'                     Demo_SyncCheckBoxesToUI
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
'   None
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
'   - Demo_TrySetCheckBoxState
'   - Demo_LogFailure
'
' NOTES
'   Window-level sync uses ActiveWindow as the reference window
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet     'Demo worksheet
    Dim ActiveWin           As Window        'Active Excel window for window-level reads
    Dim IsVisible           As Boolean       'Resolved current visibility state
    Dim FailMsg             As String        'Diagnostic message from reader / writer helpers

    Const PROC As String = "Demo_SyncCheckBoxesToUI"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail
    'Play the optional button-click visual feedback only when requested
        If PlayFeedback Then
            DEMO_Btn_PlayFeedback
        End If
    'Resolve the demo worksheet
        Set WS = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)
    'Resolve the active Excel window used for window-level sync
        Set ActiveWin = Application.ActiveWindow

'------------------------------------------------------------------------------
' SYNC APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Read current Ribbon visibility and update the related check box
        If Demo_TryGetRibbonVisibility(IsVisible, FailMsg) Then
            If Not Demo_TrySetCheckBoxState(WS, CB_RIBBON, IsVisible, FailMsg) _
                Then
                Demo_LogFailure PROC, CB_RIBBON, FailMsg
            End If
        Else
            Demo_LogFailure PROC, "RibbonState", FailMsg
        End If
    'Read current StatusBar visibility and update the related check box
        If Not Demo_TrySetCheckBoxState(WS, CB_STATUSBAR, _
            Application.DisplayStatusBar, FailMsg) Then
            Demo_LogFailure PROC, CB_STATUSBAR, FailMsg
        End If
    'Read current ScrollBars visibility and update the related check box
        If Not Demo_TrySetCheckBoxState(WS, CB_SCROLLBARS, _
            Application.DisplayScrollBars, FailMsg) Then
            Demo_LogFailure PROC, CB_SCROLLBARS, FailMsg
        End If
    'Read current FormulaBar visibility and update the related check box
        If Not Demo_TrySetCheckBoxState(WS, CB_FORMULABAR, _
            Application.DisplayFormulaBar, FailMsg) Then
            Demo_LogFailure PROC, CB_FORMULABAR, FailMsg
        End If

'------------------------------------------------------------------------------
' SYNC WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Reject missing ActiveWindow deterministically for window-level sync
        If ActiveWin Is Nothing Then
            'Log the missing ActiveWindow state
                Demo_LogFailure PROC, "ActiveWindow", _
                    "no active window available for window-level sync"
        Else
            'Update the Headings check box from ActiveWindow
                If Not Demo_TrySetCheckBoxState(WS, CB_HEADINGS, _
                    ActiveWin.DisplayHeadings, FailMsg) Then
                    Demo_LogFailure PROC, CB_HEADINGS, FailMsg
                End If
            'Update the WorkbookTabs check box from ActiveWindow
                If Not Demo_TrySetCheckBoxState(WS, CB_WORKBOOKTABS, _
                    ActiveWin.DisplayWorkbookTabs, FailMsg) Then
                    Demo_LogFailure PROC, CB_WORKBOOKTABS, FailMsg
                End If
            'Update the Gridlines check box from ActiveWindow
                If Not Demo_TrySetCheckBoxState(WS, CB_GRIDLINES, _
                    ActiveWin.DisplayGridlines, FailMsg) Then
                    Demo_LogFailure PROC, CB_GRIDLINES, FailMsg
                End If
        End If

'------------------------------------------------------------------------------
' SYNC TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Read current title-bar visibility and update the related check box
        If Demo_TryGetTitleBarVisibility(IsVisible, FailMsg) Then
            If Not Demo_TrySetCheckBoxState(WS, CB_TITLEBAR, IsVisible, FailMsg) _
                Then
                Demo_LogFailure PROC, CB_TITLEBAR, FailMsg
            End If
        Else
            Demo_LogFailure PROC, "TitleBarState", FailMsg
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
        Demo_LogFailure PROC, "Unexpected", Demo_GetRuntimeErrorText
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
'   next SHOW or HIDE action
'
' WHY THIS EXISTS
'   Select All is a useful convenience action during demos and testing
'
' RETURNS
'   None
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
    'Play the optional button-click visual feedback
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
'==============================================================================
'                             Demo_ClearAllUI
'------------------------------------------------------------------------------
' PURPOSE
'   Clear all demo check boxes so no UI elements are selected for the next SHOW
'   or HIDE action
'
' WHY THIS EXISTS
'   Clear All is a useful convenience action during demos and testing
'
' RETURNS
'   None
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
    'Play the optional button-click visual feedback
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
'==============================================================================
'                            Demo_PresetKiosk
'------------------------------------------------------------------------------
' PURPOSE
'   Pre-select a broad kiosk-style profile covering all managed UI elements
'
' WHY THIS EXISTS
'   A kiosk-like presentation typically considers all major Excel chrome and
'   worksheet aids as candidate targets
'
' RETURNS
'   None
'
' NOTES
'   - This preset only sets the check boxes
'   - It does NOT apply SHOW or HIDE by itself
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
    'Play the optional button-click visual feedback
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
'==============================================================================
'                           Demo_PresetAnalyst
'------------------------------------------------------------------------------
' PURPOSE
'   Pre-select a profile focused on worksheet navigation and analysis aids
'
' WHY THIS EXISTS
'   Analytical use cases often care most about sheet aids such as headings,
'   tabs, gridlines, formula bar, and related navigation cues
'
' RETURNS
'   None
'
' NOTES
'   - This preset only sets the check boxes
'   - It does NOT apply SHOW or HIDE by itself
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
    'Play the optional button-click visual feedback
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
'==============================================================================
'                           Demo_PresetMinimal
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
'   None
'
' NOTES
'   - This preset only sets the check boxes
'   - It does NOT apply SHOW or HIDE by itself
'
' DEPENDENCIES
'   - DEMO_Btn_PlayFeedback
'   - Demo_SetSelectionProfile
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY PROFILE
'------------------------------------------------------------------------------
    'Play the optional button-click visual feedback
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
'==============================================================================
'                    Demo_CaptureUIState
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
'   None
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
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC As String = "Demo_CaptureUIState"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail
    'Play the optional button-click visual feedback
        DEMO_Btn_PlayFeedback

'------------------------------------------------------------------------------
' CAPTURE SNAPSHOT
'------------------------------------------------------------------------------
    'Capture the current managed Excel UI state through the core module
        UI_CaptureExcelUIState

'------------------------------------------------------------------------------
' INFORM USER
'------------------------------------------------------------------------------
    'Confirm that the state snapshot was captured when requested
        If ShowConfirmation Then
            MsgBox "Current Excel UI state captured.", vbInformation, _
                "Excel UI Demo"
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
        Demo_LogFailure PROC, "Unexpected", Demo_GetRuntimeErrorText
    'Exit quietly after logging
        Resume SafeExit

End Sub

Public Sub Demo_ResetUIToCapturedState()

'
'==============================================================================
'                    Demo_ResetUIToCapturedState
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
'   None
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
'   - UI_HasExcelUIStateSnapshot
'   - UI_ResetExcelUIToSnapshot
'   - Demo_SyncCheckBoxesToUI
'   - Demo_LogFailure
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC As String = "Demo_ResetUIToCapturedState"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail
    'Play the optional button-click visual feedback
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
    'Restore the captured managed Excel UI state through the core module
        UI_ResetExcelUIToSnapshot

'------------------------------------------------------------------------------
' RESYNCHRONIZE DEMO CHECK BOXES
'------------------------------------------------------------------------------
    'Synchronize the demo check boxes back to the current visible state after reset
        Demo_SyncCheckBoxesToUI False

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
        Demo_LogFailure PROC, "Unexpected", Demo_GetRuntimeErrorText
    'Exit quietly after logging
        Resume SafeExit

End Sub

'
'------------------------------------------------------------------------------
'
'                           SHARED PRIVATE ORCHESTRATION
'
'------------------------------------------------------------------------------
'


Private Sub Demo_ApplySelectedUI(ByVal SelectedVisibility As UIVisibility, ByVal _
    CallerProc As String)

'
'==============================================================================
'                        Demo_ApplySelectedUI
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
'   None
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
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet     'Demo worksheet
    Dim RibbonVis           As UIVisibility  'Resolved Ribbon visibility
    Dim StatusBarVis        As UIVisibility  'Resolved StatusBar visibility
    Dim ScrollBarsVis       As UIVisibility  'Resolved ScrollBars visibility
    Dim FormulaBarVis       As UIVisibility  'Resolved FormulaBar visibility
    Dim HeadingsVis         As UIVisibility  'Resolved Headings visibility
    Dim WorkbookTabsVis     As UIVisibility  'Resolved WorkbookTabs visibility
    Dim GridlinesVis        As UIVisibility  'Resolved Gridlines visibility
    Dim TitleBarVis         As UIVisibility  'Resolved TitleBar visibility

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail
    'Resolve the demo worksheet
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
            'Exit quietly
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
        Demo_LogFailure CallerProc, "Unexpected", Demo_GetRuntimeErrorText
    'Exit quietly after logging
        Resume SafeExit

End Sub

Private Sub Demo_SetSelectionProfile(ByVal CallerProc As String, ByVal _
    RibbonSelected As Boolean, ByVal StatusBarSelected As Boolean, ByVal _
    ScrollBarsSelected As Boolean, ByVal FormulaBarSelected As Boolean, ByVal _
    HeadingsSelected As Boolean, ByVal WorkbookTabsSelected As Boolean, ByVal _
    GridlinesSelected As Boolean, ByVal TitleBarSelected As Boolean)

'
'==============================================================================
'                         Demo_SetSelectionProfile
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
'   None
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Per-control failures are written to the Immediate Window
'
' DEPENDENCIES
'   - Demo_TrySetCheckBoxState
'   - Demo_LogFailure
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet     'Demo worksheet
    Dim FailMsg             As String        'Diagnostic message from the writer helper

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail
    'Resolve the demo worksheet
        Set WS = ThisWorkbook.Worksheets(DEMO_SHEET_NAME)

'------------------------------------------------------------------------------
' WRITE SELECTION PROFILE
'------------------------------------------------------------------------------
    'Write the Ribbon selection state
        If Not Demo_TrySetCheckBoxState(WS, CB_RIBBON, RibbonSelected, FailMsg) _
            Then
            Demo_LogFailure CallerProc, CB_RIBBON, FailMsg
        End If
    'Write the StatusBar selection state
        If Not Demo_TrySetCheckBoxState(WS, CB_STATUSBAR, StatusBarSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_STATUSBAR, FailMsg
        End If
    'Write the ScrollBars selection state
        If Not Demo_TrySetCheckBoxState(WS, CB_SCROLLBARS, ScrollBarsSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_SCROLLBARS, FailMsg
        End If
    'Write the FormulaBar selection state
        If Not Demo_TrySetCheckBoxState(WS, CB_FORMULABAR, FormulaBarSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_FORMULABAR, FailMsg
        End If
    'Write the Headings selection state
        If Not Demo_TrySetCheckBoxState(WS, CB_HEADINGS, HeadingsSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_HEADINGS, FailMsg
        End If
    'Write the WorkbookTabs selection state
        If Not Demo_TrySetCheckBoxState(WS, CB_WORKBOOKTABS, _
            WorkbookTabsSelected, FailMsg) Then
            Demo_LogFailure CallerProc, CB_WORKBOOKTABS, FailMsg
        End If
    'Write the Gridlines selection state
        If Not Demo_TrySetCheckBoxState(WS, CB_GRIDLINES, GridlinesSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_GRIDLINES, FailMsg
        End If
    'Write the TitleBar selection state
        If Not Demo_TrySetCheckBoxState(WS, CB_TITLEBAR, TitleBarSelected, _
            FailMsg) Then
            Demo_LogFailure CallerProc, CB_TITLEBAR, FailMsg
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
        Demo_LogFailure CallerProc, "Unexpected", Demo_GetRuntimeErrorText
    'Exit quietly after logging
        Resume SafeExit

End Sub

Private Function Demo_HasAnySelectedChange(ByVal RibbonVis As UIVisibility, _
    ByVal StatusBarVis As UIVisibility, ByVal ScrollBarsVis As UIVisibility, ByVal _
    FormulaBarVis As UIVisibility, ByVal HeadingsVis As UIVisibility, ByVal _
    WorkbookTabsVis As UIVisibility, ByVal GridlinesVis As UIVisibility, ByVal _
    TitleBarVis As UIVisibility) As Boolean

'
'==============================================================================
'                        Demo_HasAnySelectedChange
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
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return TRUE when at least one requested visibility is actionable
        Demo_HasAnySelectedChange = (RibbonVis <> UI_LeaveUnchanged Or _
            StatusBarVis <> UI_LeaveUnchanged Or ScrollBarsVis <> UI_LeaveUnchanged _
            Or FormulaBarVis <> UI_LeaveUnchanged Or HeadingsVis <> _
            UI_LeaveUnchanged Or WorkbookTabsVis <> UI_LeaveUnchanged Or _
            GridlinesVis <> UI_LeaveUnchanged Or TitleBarVis <> UI_LeaveUnchanged)

End Function



'
'------------------------------------------------------------------------------
'
'                   PRIVATE CHECKBOX/STATE TRANSLATION HELPERS
'
'------------------------------------------------------------------------------
'

Private Function Demo_CheckBoxToUIVisibility(ByVal WS As Worksheet, ByVal _
    CheckBoxName As String, ByVal SelectedVisibility As UIVisibility, ByVal _
    CallerProc As String) As UIVisibility

'
'==============================================================================
'                        Demo_CheckBoxToUIVisibility
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
'   - Demo_TryGetCheckBoxState
'   - Demo_LogFailure
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IsChecked           As Boolean   'Resolved check-box state
    Dim FailMsg             As String    'Diagnostic message from the reader helper

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to UI_LeaveUnchanged unless a checked state is confirmed
        Demo_CheckBoxToUIVisibility = UI_LeaveUnchanged

'------------------------------------------------------------------------------
' READ CHECK-BOX STATE
'------------------------------------------------------------------------------
    'Attempt to read the requested check box
        If Not Demo_TryGetCheckBoxState(WS, CheckBoxName, IsChecked, FailMsg) _
            Then
            'Log the control-resolution failure and keep UI_LeaveUnchanged
                Demo_LogFailure CallerProc, CheckBoxName, FailMsg
            'Exit with default value
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

Private Function Demo_TryGetCheckBoxState(ByVal WS As Worksheet, ByVal _
    ControlName As String, ByRef IsChecked As Boolean, ByRef FailMsg As String) As _
    Boolean

'
'==============================================================================
'                        Demo_TryGetCheckBoxState
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read the checked state of a demo worksheet check box
'
' WHY THIS EXISTS
'   The demo workbook may use either:
'     - Forms check boxes
'     - ActiveX check boxes
'
'   This helper supports both models behind a single reader
'
' INPUTS
'   WS
'     Worksheet containing the control
'
'   ControlName
'     Name of the control to inspect
'
'   IsChecked
'     Receives TRUE when the check box is checked, FALSE otherwise
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => control found and state read successfully
'   FALSE => control missing or invalid
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Shp                 As Shape         'Candidate Forms control
    Dim CheckBoxOle         As OLEObject     'Candidate ActiveX control
    Dim ValueOut            As Variant       'Late-bound Value property result

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
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing worksheet reference deterministically
        If WS Is Nothing Then
            FailMsg = "worksheet reference is Nothing"
            GoTo SafeExit
        End If
    'Reject a blank control name deterministically
        If Len(Trim$(ControlName)) = 0 Then
            FailMsg = "control name is blank"
            GoTo SafeExit
        End If
        
'------------------------------------------------------------------------------
' TRY FORMS CHECK BOX
'------------------------------------------------------------------------------
    'Attempt to resolve the control as a worksheet shape
        On Error Resume Next
            Set Shp = WS.Shapes(ControlName)
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
            Set CheckBoxOle = WS.OLEObjects(ControlName)
        On Error GoTo Fail
    'Process the OLEObject when it exists
        If Not CheckBoxOle Is Nothing Then
            'Reject ActiveX controls that are not check boxes
                If InStr(1, CheckBoxOle.progID, "CheckBox", vbTextCompare) = 0 _
                    Then
                    FailMsg = "ActiveX control exists but is not a CheckBox"
                    GoTo SafeExit
                End If
            'Read the checked state through late-bound Value access
                ValueOut = CallByName(CheckBoxOle.Object, "Value", VbGet)
                IsChecked = CBool(ValueOut)
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
        FailMsg = Demo_GetRuntimeErrorText

End Function

Private Function Demo_TrySetCheckBoxState(ByVal WS As Worksheet, ByVal _
    ControlName As String, ByVal IsChecked As Boolean, ByRef FailMsg As String) As _
    Boolean

'
'==============================================================================
'                        Demo_TrySetCheckBoxState
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to write the checked state of a demo worksheet check box
'
' WHY THIS EXISTS
'   The demo supports selection profiles and current-state synchronization,
'   both of which need to programmatically set the worksheet controls
'
' INPUTS
'   WS
'     Worksheet containing the control
'
'   ControlName
'     Name of the control to update
'
'   IsChecked
'     Requested target state
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => control found and updated successfully
'   FALSE => control missing or invalid
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Shp                 As Shape         'Candidate Forms control
    Dim CheckBoxOle         As OLEObject     'Candidate ActiveX control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail
    'Initialize default result
        Demo_TrySetCheckBoxState = False
        FailMsg = vbNullString
'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing worksheet reference deterministically
        If WS Is Nothing Then
            FailMsg = "worksheet reference is Nothing"
            GoTo SafeExit
        End If
    
    'Reject a blank control name deterministically
        If Len(Trim$(ControlName)) = 0 Then
            FailMsg = "control name is blank"
            GoTo SafeExit
        End If
    
'------------------------------------------------------------------------------
' TRY FORMS CHECK BOX
'------------------------------------------------------------------------------
    'Attempt to resolve the control as a worksheet shape
        On Error Resume Next
            Set Shp = WS.Shapes(ControlName)
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
                If IsChecked Then
                    Shp.ControlFormat.Value = xlOn
                Else
                    Shp.ControlFormat.Value = xlOff
                End If
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
            Set CheckBoxOle = WS.OLEObjects(ControlName)
        On Error GoTo Fail
    'Process the OLEObject when it exists
        If Not CheckBoxOle Is Nothing Then
            'Reject ActiveX controls that are not check boxes
                If InStr(1, CheckBoxOle.progID, "CheckBox", vbTextCompare) = 0 _
                    Then
                    FailMsg = "ActiveX control exists but is not a CheckBox"
                    GoTo SafeExit
                End If
            'Write the checked state through late-bound Value access
                CallByName CheckBoxOle.Object, "Value", VbLet, IsChecked
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
        FailMsg = Demo_GetRuntimeErrorText

End Function

'
'------------------------------------------------------------------------------
'
'                       PRIVATE UI-STATE READ HELPERS
'
'------------------------------------------------------------------------------
'

Private Function Demo_TryGetRibbonVisibility(ByRef IsVisible As Boolean, ByRef _
    FailMsg As String) As Boolean

'
'==============================================================================
'                         Demo_TryGetRibbonVisibility
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
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ValueOut            As Variant      'Fallback Excel4 macro result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail
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
            On Error GoTo Fail
            Demo_TryGetRibbonVisibility = True
            GoTo SafeExit
        End If
        Err.Clear
        On Error GoTo Fail

'------------------------------------------------------------------------------
' TRY EXCEL4 MACRO FALLBACK
'------------------------------------------------------------------------------
    'Attempt a fallback read using an Excel4 macro
        On Error Resume Next
            ValueOut = _
                Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")
        If Err.Number = 0 Then
            On Error GoTo Fail
            IsVisible = CBool(ValueOut)
            Demo_TryGetRibbonVisibility = True
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
        FailMsg = Demo_GetRuntimeErrorText

End Function

Private Function Demo_TryGetTitleBarVisibility(ByRef IsVisible As Boolean, ByRef _
    FailMsg As String) As Boolean

'
'==============================================================================
'                        Demo_TryGetTitleBarVisibility
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
'     Receives current title-bar visibility on success
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
' UPDATED
'   2026-04-19
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
        Demo_TryGetTitleBarVisibility = False
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
    'Treat zero plus nonzero last error as failure
        If StyleValue = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' MAP STYLE TO TITLE-BAR VISIBILITY
'------------------------------------------------------------------------------
    'Treat the caption style bit as the title-bar visibility signal
        IsVisible = ((StyleValue And DEMO_WS_CAPTION) <> 0)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after a valid style read
        Demo_TryGetTitleBarVisibility = True

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
        FailMsg = Demo_GetRuntimeErrorText

End Function

'
'------------------------------------------------------------------------------
'
'                           PRIVATE DIAGNOSTICS
'
'------------------------------------------------------------------------------
'

Private Function Demo_GetRuntimeErrorText() As String

'
'==============================================================================
'                      Demo_GetRuntimeErrorText
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
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected issue while formatting the diagnostic
        On Error Resume Next

'------------------------------------------------------------------------------
' BUILD ERROR TEXT
'------------------------------------------------------------------------------
    'Build a consistent diagnostic string from the current Err state
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
'   None
'
' ERROR POLICY
'   Suppresses any unexpected logging failure locally
'
' UPDATED
'   2026-04-19
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

