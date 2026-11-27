Attribute VB_Name = "M_EXCEL_UI_TEST_RUNTIME"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: M_EXCEL_UI_TEST_RUNTIME
'==============================================================================
' PURPOSE
'   Hold the runtime regression cases of EXCEL_UI: visibility, convenience
'   wrappers, structured results, window-target scope, the failure list and
'   the quiet-update scope
'
' WHY THIS EXISTS
'   UI-control code is easy to break accidentally when refining tri-state
'   behavior, selective application, leave-unchanged semantics, the
'   convenience wrappers, the structured-result path or window-target scope.
'   These cases detect such a regression
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   - Each case applies a known UI state through the public API, waits for it
'     to settle, and asserts the host state or the structured result the call
'     returned.
'   - The module keeps no state; its only declarations are the constants that
'     name the quiet-update fault kinds.
'
' ERROR POLICY
'   - A failed assertion raises; no case returns a verdict.
'   - The cases that open windows or workbooks, or arm a test seam, clean up
'     first and then raise the failure they kept. The others have no handler.
'   - Recording a pass or a failure is left to the pack that runs the case.
'
' DEPENDENCIES
'   - M_EXCEL_UI_TEST_HARNESS
'
' NOTES
'   TEST_QUIET_FAULT_NONE, _ENTRY, _WRITE and _READBACK carry the values of the
'   fault kinds of the quiet-update seam in M_EXCEL_UI_RUNTIME.
'
' PUBLIC SURFACE
'   None of its own. The cases are Public so that the regression pack in
'   M_EXCEL_UI_TEST_CERTIFICATION can run them
'
' TEST SCOPE
'   Core tests
'     - show-all baseline
'     - selective hide
'     - selective show
'     - no-op / leave-unchanged
'
'   Wrapper tests
'     - convenience wrappers
'     - executed only in the full pack because they also affect TitleBar
'
'   Structured-result tests
'     - clean success path
'     - no-op / leave-unchanged success path
'     - success path without FailureList capture
'     - invalid UIVisibility structured failure path
'
'   Target-scope tests
'     - active-window-only application
'     - active-workbook-window application
'     - invalid-scope failure with application-level continuation
'
'   Environment-preservation tests
'     - ScreenUpdating restored to prior state
'     - the quiet-update scope restores only the suppression it confirmed
'
'   Failure-list tests
'     - a list that cannot grow degrades visibly without raising
'     - a first entry that cannot be allocated keeps the result contract
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.2
'
' HISTORY
'   2026-11-15 - Split out of the single regression module; the cases moved
'                here unchanged apart from visibility (#90).
'
' UPDATED
'   2026-11-15
'==============================================================================

'------------------------------------------------------------------------------
' TEST CONFIGURATION
'------------------------------------------------------------------------------
    Private Const TEST_QUIET_FAULT_NONE     As Long = 0 'Quiet scope: no fault
    Private Const TEST_QUIET_FAULT_ENTRY    As Long = 1 'Quiet scope: entry read fails
    Private Const TEST_QUIET_FAULT_WRITE    As Long = 2 'Quiet scope: write ineffective
    Private Const TEST_QUIET_FAULT_READBACK As Long = 3 'Quiet scope: readback fails


'
'------------------------------------------------------------------------------
'
'                               REGRESSION CASES
'
'------------------------------------------------------------------------------
'

Public Sub TST_Case_ShowAllBaseline(ByVal IncludeTitleBarTests As Boolean)

'
'------------------------------------------------------------------------------
'                           VERIFY SHOW-ALL BASELINE
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that the module can drive all managed UI elements to visible state
'
' WHY THIS EXISTS
'   This case establishes a known visible baseline and validates that the
'   public API can set every managed element to shown
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => include TitleBar in the show-all assertion
'     FALSE => leave title-bar assertions out of this case
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Applies UI_Show to the Ribbon, StatusBar, ScrollBars, FormulaBar,
'     Headings, WorkbookTabs and Gridlines, and to TitleBar when title-bar
'     tests are included, then waits for the UI to settle.
'   - Asserts the Ribbon, StatusBar, ScrollBars and FormulaBar visible.
'   - Asserts Headings, WorkbookTabs and Gridlines visible on every open
'     window, and TitleBar visible when title-bar tests are included.
'
' ERROR POLICY
'   - Raises on assertion failure
'   - Has no handler and no cleanup: the failure reaches the caller unchanged
'
' DEPENDENCIES
'   - UI_SetExcelUI
'   - TST_TitleBarMode
'   - TST_WaitUI
'   - TST_AssertRibbonVisible
'   - TST_AssertApplicationProperty
'   - TST_AssertAllWindowsProperty
'   - TST_AssertTitleBarVisible
'   - TST_Log
'
' NOTES
'   Changes the managed UI state of the host and does not restore it; the
'   pack that runs the case restores its entry state afterwards.
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log "TST_Case_ShowAllBaseline", "START", "Setting all managed UI visible"

'------------------------------------------------------------------------------
' APPLY SHOW-ALL BASELINE
'------------------------------------------------------------------------------
    'Drive all application- and window-level UI elements to visible state
        UI_SetExcelUI _
            Ribbon:=UI_Show, _
            StatusBar:=UI_Show, _
            ScrollBars:=UI_Show, _
            FormulaBar:=UI_Show, _
            Headings:=UI_Show, _
            WorkbookTabs:=UI_Show, _
            Gridlines:=UI_Show, _
            TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Show)

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Assert Ribbon visible
        TST_AssertRibbonVisible True, "ShowAllBaseline.Ribbon"

    'Assert StatusBar visible
        TST_AssertApplicationProperty True, "DisplayStatusBar", "ShowAllBaseline.StatusBar"

    'Assert ScrollBars visible
        TST_AssertApplicationProperty True, "DisplayScrollBars", "ShowAllBaseline.ScrollBars"

    'Assert FormulaBar visible
        TST_AssertApplicationProperty True, "DisplayFormulaBar", "ShowAllBaseline.FormulaBar"

'------------------------------------------------------------------------------
' ASSERT WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Assert Headings visible across all open Excel windows
        TST_AssertAllWindowsProperty True, "DisplayHeadings", "ShowAllBaseline.Headings"

    'Assert WorkbookTabs visible across all open Excel windows
        TST_AssertAllWindowsProperty True, "DisplayWorkbookTabs", "ShowAllBaseline.WorkbookTabs"

    'Assert Gridlines visible across all open Excel windows
        TST_AssertAllWindowsProperty True, "DisplayGridlines", "ShowAllBaseline.Gridlines"

'------------------------------------------------------------------------------
' ASSERT TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Assert TitleBar visible when title-bar testing is enabled
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "ShowAllBaseline.TitleBar"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_ShowAllBaseline", "PASS", "All requested elements are visible"

End Sub

Public Sub TST_Case_SelectiveHide(ByVal IncludeTitleBarTests As Boolean)

'
'------------------------------------------------------------------------------
'                            VERIFY SELECTIVE HIDE
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that selective hide requests affect only the requested UI elements
'   while leaving the others unchanged
'
' WHY THIS EXISTS
'   Selective application is one of the most important contracts of the
'   tri-state API
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => assert that TitleBar remains visible and unchanged
'     FALSE => skip TitleBar assertions in this case
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Shows every managed element as a baseline, and waits.
'   - Hides only StatusBar and Gridlines, leaving the other arguments at
'     their default, and waits.
'   - Asserts StatusBar and Gridlines hidden, and the Ribbon, ScrollBars,
'     FormulaBar, Headings and WorkbookTabs still visible.
'   - Asserts TitleBar still visible when title-bar tests are included.
'
' ERROR POLICY
'   - Raises on assertion failure
'   - Has no handler and no cleanup: the failure reaches the caller unchanged
'
' DEPENDENCIES
'   - UI_SetExcelUI
'   - TST_TitleBarMode
'   - TST_WaitUI
'   - TST_AssertRibbonVisible
'   - TST_AssertApplicationProperty
'   - TST_AssertAllWindowsProperty
'   - TST_AssertTitleBarVisible
'   - TST_Log
'
' NOTES
'   Changes the managed UI state of the host and does not restore it; the
'   pack that runs the case restores its entry state afterwards.
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log "TST_Case_SelectiveHide", "START", "Hiding only selected elements"

'------------------------------------------------------------------------------
' ESTABLISH VISIBLE BASELINE
'------------------------------------------------------------------------------
    'Start from a known visible baseline
        UI_SetExcelUI _
            Ribbon:=UI_Show, _
            StatusBar:=UI_Show, _
            ScrollBars:=UI_Show, _
            FormulaBar:=UI_Show, _
            Headings:=UI_Show, _
            WorkbookTabs:=UI_Show, _
            Gridlines:=UI_Show, _
            TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Show)

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' APPLY SELECTIVE HIDE
'------------------------------------------------------------------------------
    'Hide only StatusBar and Gridlines while leaving the rest unchanged
        UI_SetExcelUI _
            StatusBar:=UI_Hide, _
            Gridlines:=UI_Hide

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT SELECTIVE RESULT
'------------------------------------------------------------------------------
    'Assert Ribbon remained visible
        TST_AssertRibbonVisible True, "SelectiveHide.Ribbon"

    'Assert StatusBar is hidden
        TST_AssertApplicationProperty False, "DisplayStatusBar", "SelectiveHide.StatusBar"

    'Assert ScrollBars remained visible
        TST_AssertApplicationProperty True, "DisplayScrollBars", "SelectiveHide.ScrollBars"

    'Assert FormulaBar remained visible
        TST_AssertApplicationProperty True, "DisplayFormulaBar", "SelectiveHide.FormulaBar"

    'Assert Headings remained visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayHeadings", "SelectiveHide.Headings"

    'Assert WorkbookTabs remained visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayWorkbookTabs", "SelectiveHide.WorkbookTabs"

    'Assert Gridlines are hidden across all windows
        TST_AssertAllWindowsProperty False, "DisplayGridlines", "SelectiveHide.Gridlines"

    'Assert TitleBar remained visible and unchanged when requested
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "SelectiveHide.TitleBar"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_SelectiveHide", "PASS", "Selective hide behaved as expected"

End Sub

Public Sub TST_Case_SelectiveShow(ByVal IncludeTitleBarTests As Boolean)

'
'------------------------------------------------------------------------------
'                            VERIFY SELECTIVE SHOW
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that selective show requests affect only the requested UI elements
'   while leaving the others unchanged
'
' WHY THIS EXISTS
'   Selective application is one of the most important contracts of the
'   tri-state API
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => keep TitleBar visible and assert it remains unchanged
'     FALSE => skip TitleBar assertions in this case
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Hides every managed element as a baseline, keeping TitleBar visible
'     when title-bar tests are included, and waits.
'   - Shows only StatusBar and WorkbookTabs, leaving the other arguments at
'     their default, and waits.
'   - Asserts StatusBar and WorkbookTabs visible, and the Ribbon,
'     ScrollBars, FormulaBar, Headings and Gridlines still hidden.
'   - Asserts TitleBar still visible when title-bar tests are included.
'
' ERROR POLICY
'   - Raises on assertion failure
'   - Has no handler and no cleanup: the failure reaches the caller unchanged
'
' DEPENDENCIES
'   - UI_SetExcelUI
'   - TST_TitleBarMode
'   - TST_WaitUI
'   - TST_AssertRibbonVisible
'   - TST_AssertApplicationProperty
'   - TST_AssertAllWindowsProperty
'   - TST_AssertTitleBarVisible
'   - TST_Log
'
' NOTES
'   Changes the managed UI state of the host and does not restore it; the
'   pack that runs the case restores its entry state afterwards.
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log "TST_Case_SelectiveShow", "START", "Showing only selected elements"

'------------------------------------------------------------------------------
' ESTABLISH HIDDEN BASELINE
'------------------------------------------------------------------------------
    'Drive application- and window-level elements hidden while keeping TitleBar
    'unchanged or visible according to test scope
        UI_SetExcelUI _
            Ribbon:=UI_Hide, _
            StatusBar:=UI_Hide, _
            ScrollBars:=UI_Hide, _
            FormulaBar:=UI_Hide, _
            Headings:=UI_Hide, _
            WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Hide, _
            TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Show)

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' APPLY SELECTIVE SHOW
'------------------------------------------------------------------------------
    'Show only StatusBar and WorkbookTabs while leaving the rest unchanged
        UI_SetExcelUI _
            StatusBar:=UI_Show, _
            WorkbookTabs:=UI_Show

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT SELECTIVE RESULT
'------------------------------------------------------------------------------
    'Assert Ribbon remained hidden
        TST_AssertRibbonVisible False, "SelectiveShow.Ribbon"

    'Assert StatusBar is visible
        TST_AssertApplicationProperty True, "DisplayStatusBar", "SelectiveShow.StatusBar"

    'Assert ScrollBars remained hidden
        TST_AssertApplicationProperty False, "DisplayScrollBars", "SelectiveShow.ScrollBars"

    'Assert FormulaBar remained hidden
        TST_AssertApplicationProperty False, "DisplayFormulaBar", "SelectiveShow.FormulaBar"

    'Assert Headings remained hidden across all windows
        TST_AssertAllWindowsProperty False, "DisplayHeadings", "SelectiveShow.Headings"

    'Assert WorkbookTabs are visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayWorkbookTabs", "SelectiveShow.WorkbookTabs"

    'Assert Gridlines remained hidden across all windows
        TST_AssertAllWindowsProperty False, "DisplayGridlines", "SelectiveShow.Gridlines"

    'Assert TitleBar remained visible and unchanged when requested
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "SelectiveShow.TitleBar"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_SelectiveShow", "PASS", "Selective show behaved as expected"

End Sub

Public Sub TST_Case_NoOpLeaveUnchanged(ByVal IncludeTitleBarTests As Boolean)

'
'------------------------------------------------------------------------------
'                       VERIFY NO-OP LEAVES UI UNCHANGED
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a no-argument UI_SetExcelUI call behaves as a no-op
'
' WHY THIS EXISTS
'   The tri-state API promises that omitted arguments do not accidentally drive
'   visibility changes
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => include TitleBar in the baseline and assertion
'     FALSE => skip TitleBar assertions in this case
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Establishes a mixed baseline: Ribbon, ScrollBars, Headings and
'     Gridlines shown; StatusBar, FormulaBar and WorkbookTabs hidden;
'     TitleBar shown when title-bar tests are included.
'   - Calls UI_SetExcelUI with no arguments, and waits.
'   - Asserts every element still has its baseline state.
'
' ERROR POLICY
'   - Raises on assertion failure
'   - Has no handler and no cleanup: the failure reaches the caller unchanged
'
' DEPENDENCIES
'   - UI_SetExcelUI
'   - TST_TitleBarMode
'   - TST_WaitUI
'   - TST_AssertRibbonVisible
'   - TST_AssertApplicationProperty
'   - TST_AssertAllWindowsProperty
'   - TST_AssertTitleBarVisible
'   - TST_Log
'
' NOTES
'   Changes the managed UI state of the host and does not restore it; the
'   pack that runs the case restores its entry state afterwards.
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log "TST_Case_NoOpLeaveUnchanged", "START", "Validating no-op and leave-unchanged behavior"

'------------------------------------------------------------------------------
' ESTABLISH MIXED BASELINE
'------------------------------------------------------------------------------
    'Establish a mixed baseline that should remain unchanged
        UI_SetExcelUI _
            Ribbon:=UI_Show, _
            StatusBar:=UI_Hide, _
            ScrollBars:=UI_Show, _
            FormulaBar:=UI_Hide, _
            Headings:=UI_Show, _
            WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Show, _
            TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Show)

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' APPLY NO-OP
'------------------------------------------------------------------------------
    'Invoke the API with no arguments so every element is LeaveUnchanged
        UI_SetExcelUI

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT NO-OP RESULT
'------------------------------------------------------------------------------
    'Assert Ribbon remained visible
        TST_AssertRibbonVisible True, "NoOp.Ribbon"

    'Assert StatusBar remained hidden
        TST_AssertApplicationProperty False, "DisplayStatusBar", "NoOp.StatusBar"

    'Assert ScrollBars remained visible
        TST_AssertApplicationProperty True, "DisplayScrollBars", "NoOp.ScrollBars"

    'Assert FormulaBar remained hidden
        TST_AssertApplicationProperty False, "DisplayFormulaBar", "NoOp.FormulaBar"

    'Assert Headings remained visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayHeadings", "NoOp.Headings"

    'Assert WorkbookTabs remained hidden across all windows
        TST_AssertAllWindowsProperty False, "DisplayWorkbookTabs", "NoOp.WorkbookTabs"

    'Assert Gridlines remained visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayGridlines", "NoOp.Gridlines"

    'Assert TitleBar remained visible and unchanged when requested
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "NoOp.TitleBar"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_NoOpLeaveUnchanged", "PASS", "No-op behavior behaved as expected"

End Sub

Public Sub TST_Case_ConvenienceWrappers(ByVal IncludeTitleBarTests As Boolean)

'
'------------------------------------------------------------------------------
'                        VERIFY HIDE AND SHOW WRAPPERS
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that UI_HideExcelUI and UI_ShowExcelUI drive all managed UI elements
'   to hidden and visible state respectively
'
' WHY THIS EXISTS
'   The convenience wrappers are part of the public surface and should be
'   regression-tested explicitly
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => include TitleBar assertions
'     FALSE => skip TitleBar assertions in this case
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Calls UI_HideExcelUI, waits, and asserts every managed element hidden,
'     TitleBar included when title-bar tests are included.
'   - Calls UI_ShowExcelUI, waits, and asserts every managed element
'     visible, TitleBar included when title-bar tests are included.
'
' ERROR POLICY
'   - Raises on assertion failure
'   - Has no handler and no cleanup: the failure reaches the caller unchanged
'
' DEPENDENCIES
'   - UI_HideExcelUI
'   - UI_ShowExcelUI
'   - TST_WaitUI
'   - TST_AssertRibbonVisible
'   - TST_AssertApplicationProperty
'   - TST_AssertAllWindowsProperty
'   - TST_AssertTitleBarVisible
'   - TST_Log
'
' NOTES
'   The certification pack runs this case only when title-bar tests are
'   included, because the wrappers also act on TitleBar. It leaves the UI
'   visible; the pack restores its entry state afterwards.
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log "TST_Case_ConvenienceWrappers", "START", "Validating UI_HideExcelUI and UI_ShowExcelUI"

'------------------------------------------------------------------------------
' APPLY HIDE-ALL WRAPPER
'------------------------------------------------------------------------------
    'Hide all managed UI elements through the convenience wrapper
        UI_HideExcelUI

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT HIDE-ALL RESULT
'------------------------------------------------------------------------------
    'Assert Ribbon hidden
        TST_AssertRibbonVisible False, "Wrappers.HideAll.Ribbon"

    'Assert StatusBar hidden
        TST_AssertApplicationProperty False, "DisplayStatusBar", "Wrappers.HideAll.StatusBar"

    'Assert ScrollBars hidden
        TST_AssertApplicationProperty False, "DisplayScrollBars", "Wrappers.HideAll.ScrollBars"

    'Assert FormulaBar hidden
        TST_AssertApplicationProperty False, "DisplayFormulaBar", "Wrappers.HideAll.FormulaBar"

    'Assert Headings hidden across all windows
        TST_AssertAllWindowsProperty False, "DisplayHeadings", "Wrappers.HideAll.Headings"

    'Assert WorkbookTabs hidden across all windows
        TST_AssertAllWindowsProperty False, "DisplayWorkbookTabs", "Wrappers.HideAll.WorkbookTabs"

    'Assert Gridlines hidden across all windows
        TST_AssertAllWindowsProperty False, "DisplayGridlines", "Wrappers.HideAll.Gridlines"

    'Assert TitleBar hidden when requested
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible False, "Wrappers.HideAll.TitleBar"
        End If

'------------------------------------------------------------------------------
' APPLY SHOW-ALL WRAPPER
'------------------------------------------------------------------------------
    'Show all managed UI elements through the convenience wrapper
        UI_ShowExcelUI

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT SHOW-ALL RESULT
'------------------------------------------------------------------------------
    'Assert Ribbon visible
        TST_AssertRibbonVisible True, "Wrappers.ShowAll.Ribbon"

    'Assert StatusBar visible
        TST_AssertApplicationProperty True, "DisplayStatusBar", "Wrappers.ShowAll.StatusBar"

    'Assert ScrollBars visible
        TST_AssertApplicationProperty True, "DisplayScrollBars", "Wrappers.ShowAll.ScrollBars"

    'Assert FormulaBar visible
        TST_AssertApplicationProperty True, "DisplayFormulaBar", "Wrappers.ShowAll.FormulaBar"

    'Assert Headings visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayHeadings", "Wrappers.ShowAll.Headings"

    'Assert WorkbookTabs visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayWorkbookTabs", "Wrappers.ShowAll.WorkbookTabs"

    'Assert Gridlines visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayGridlines", "Wrappers.ShowAll.Gridlines"

    'Assert TitleBar visible when requested
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "Wrappers.ShowAll.TitleBar"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_ConvenienceWrappers", "PASS", "Convenience wrappers behaved as expected"

End Sub

Public Sub TST_Case_WithResult_AllSuccess(ByVal IncludeTitleBarTests As Boolean)

'
'------------------------------------------------------------------------------
'                       VERIFY STRUCTURED RESULT SUCCESS
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that UI_SetExcelUI_WithResult returns a clean-success outcome when
'   all requested UI updates succeed
'
' WHY THIS EXISTS
'   The structured-result public path should be regression-tested explicitly for
'   its clean success contract
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => include TitleBar in the success assertion
'     FALSE => leave TitleBar unchanged in this case
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Calls UI_SetExcelUI_WithResult with UI_Show for every managed
'     element, TitleBar through TST_TitleBarMode, and waits.
'   - Asserts the Boolean result, FailureCount and FailureList describe a
'     clean success.
'   - Asserts every element visible, TitleBar included when title-bar tests
'     are included.
'
' ERROR POLICY
'   - Raises on assertion failure
'
' DEPENDENCIES
'   - UI_SetExcelUI_WithResult
'   - TST_TitleBarMode
'   - TST_WaitUI
'   - TST_AssertResultSuccess
'   - TST_AssertRibbonVisible
'   - TST_AssertApplicationProperty
'   - TST_AssertAllWindowsProperty
'   - TST_AssertTitleBarVisible
'   - TST_Log
'
' NOTES
'   Changes the managed UI state of the host and does not restore it; the
'   pack that runs the case restores its entry state afterwards.
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_RUNTIME.TST_Case_WithResult_AllSuccess"

    Dim OK              As Boolean  'Boolean success flag returned by the API
    Dim FailureCount    As Long     'Number of recorded failures
    Dim FailureList     As Variant  'Optional array of recorded failures

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log "TST_Case_WithResult_AllSuccess", "START", "Validating structured-result success path"

'------------------------------------------------------------------------------
' APPLY REQUESTED UI STATE
'------------------------------------------------------------------------------
    'Apply a deterministic visible state through the structured-result API
        OK = UI_SetExcelUI_WithResult( _
                Ribbon:=UI_Show, _
                StatusBar:=UI_Show, _
                ScrollBars:=UI_Show, _
                FormulaBar:=UI_Show, _
                Headings:=UI_Show, _
                WorkbookTabs:=UI_Show, _
                Gridlines:=UI_Show, _
                TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Show), _
                FailureCount:=FailureCount, _
                FailureList:=FailureList)

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT RESULT BUFFERS
'------------------------------------------------------------------------------
    'Assert that the returned result buffers represent a clean success path
        TST_AssertResultSuccess OK, FailureCount, FailureList, "WithResult.AllSuccess.Result"

'------------------------------------------------------------------------------
' ASSERT APPLIED UI STATE
'------------------------------------------------------------------------------
    'Assert Ribbon visible
        TST_AssertRibbonVisible True, "WithResult.AllSuccess.Ribbon"

    'Assert StatusBar visible
        TST_AssertApplicationProperty True, "DisplayStatusBar", "WithResult.AllSuccess.StatusBar"

    'Assert ScrollBars visible
        TST_AssertApplicationProperty True, "DisplayScrollBars", "WithResult.AllSuccess.ScrollBars"

    'Assert FormulaBar visible
        TST_AssertApplicationProperty True, "DisplayFormulaBar", "WithResult.AllSuccess.FormulaBar"

    'Assert Headings visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayHeadings", "WithResult.AllSuccess.Headings"

    'Assert WorkbookTabs visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayWorkbookTabs", "WithResult.AllSuccess.WorkbookTabs"

    'Assert Gridlines visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayGridlines", "WithResult.AllSuccess.Gridlines"

    'Assert TitleBar visible when title-bar testing is enabled
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "WithResult.AllSuccess.TitleBar"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_WithResult_AllSuccess", "PASS", "Structured-result success path behaved as expected"

End Sub

Public Sub TST_Case_WithResult_NoOpSuccess(ByVal IncludeTitleBarTests As Boolean)

'
'------------------------------------------------------------------------------
'                       VERIFY STRUCTURED NO-OP SUCCESS
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that UI_SetExcelUI_WithResult returns a clean-success outcome when
'   invoked as a no-op with all arguments omitted or left unchanged
'
' WHY THIS EXISTS
'   The structured-result path should preserve the leave-unchanged contract
'   while still reporting clean success
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => include TitleBar in the baseline and assertion
'     FALSE => skip TitleBar assertions in this case
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Establishes a mixed baseline with UI_SetExcelUI: Ribbon, ScrollBars,
'     Headings and Gridlines shown; StatusBar, FormulaBar and WorkbookTabs
'     hidden; TitleBar shown when title-bar tests are included.
'   - Calls UI_SetExcelUI_WithResult with only FailureCount and FailureList,
'     and waits.
'   - Asserts a clean success, and every element still has its baseline
'     state.
'
' ERROR POLICY
'   - Raises on assertion failure
'
' DEPENDENCIES
'   - UI_SetExcelUI
'   - UI_SetExcelUI_WithResult
'   - TST_TitleBarMode
'   - TST_WaitUI
'   - TST_AssertResultSuccess
'   - TST_AssertRibbonVisible
'   - TST_AssertApplicationProperty
'   - TST_AssertAllWindowsProperty
'   - TST_AssertTitleBarVisible
'   - TST_Log
'
' NOTES
'   Changes the managed UI state of the host and does not restore it; the
'   pack that runs the case restores its entry state afterwards.
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_RUNTIME.TST_Case_WithResult_NoOpSuccess"

    Dim OK              As Boolean  'Boolean success flag returned by the API
    Dim FailureCount    As Long     'Number of recorded failures
    Dim FailureList     As Variant  'Optional array of recorded failures

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log "TST_Case_WithResult_NoOpSuccess", "START", "Validating structured-result no-op path"

'------------------------------------------------------------------------------
' ESTABLISH MIXED BASELINE
'------------------------------------------------------------------------------
    'Establish a mixed baseline that should remain unchanged
        UI_SetExcelUI _
            Ribbon:=UI_Show, _
            StatusBar:=UI_Hide, _
            ScrollBars:=UI_Show, _
            FormulaBar:=UI_Hide, _
            Headings:=UI_Show, _
            WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Show, _
            TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Show)

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' APPLY NO-OP THROUGH STRUCTURED-RESULT API
'------------------------------------------------------------------------------
    'Invoke the structured-result API with no arguments so every element is
    'LeaveUnchanged
        OK = UI_SetExcelUI_WithResult( _
                FailureCount:=FailureCount, _
                FailureList:=FailureList)

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT RESULT BUFFERS
'------------------------------------------------------------------------------
    'Assert that the returned result buffers represent a clean success path
        TST_AssertResultSuccess OK, FailureCount, FailureList, "WithResult.NoOp.Result"

'------------------------------------------------------------------------------
' ASSERT NO-OP UI STATE
'------------------------------------------------------------------------------
    'Assert Ribbon remained visible
        TST_AssertRibbonVisible True, "WithResult.NoOp.Ribbon"

    'Assert StatusBar remained hidden
        TST_AssertApplicationProperty False, "DisplayStatusBar", "WithResult.NoOp.StatusBar"

    'Assert ScrollBars remained visible
        TST_AssertApplicationProperty True, "DisplayScrollBars", "WithResult.NoOp.ScrollBars"

    'Assert FormulaBar remained hidden
        TST_AssertApplicationProperty False, "DisplayFormulaBar", "WithResult.NoOp.FormulaBar"

    'Assert Headings remained visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayHeadings", "WithResult.NoOp.Headings"

    'Assert WorkbookTabs remained hidden across all windows
        TST_AssertAllWindowsProperty False, "DisplayWorkbookTabs", "WithResult.NoOp.WorkbookTabs"

    'Assert Gridlines remained visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayGridlines", "WithResult.NoOp.Gridlines"

    'Assert TitleBar remained visible and unchanged when requested
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "WithResult.NoOp.TitleBar"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_WithResult_NoOpSuccess", "PASS", "Structured-result no-op path behaved as expected"

End Sub

Public Sub TST_Case_WithResult_SuccessWithoutFailureList(ByVal IncludeTitleBarTests As Boolean)

'
'------------------------------------------------------------------------------
'                     VERIFY SUCCESS WITHOUT FAILURE LIST
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that UI_SetExcelUI_WithResult succeeds cleanly when the caller omits
'   the optional FailureList output
'
' WHY THIS EXISTS
'   FailureList is intentionally optional so callers that only need the Boolean
'   result and FailureCount do not need to manage an array
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => include TitleBar in the success assertion
'     FALSE => leave TitleBar unchanged in this case
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Calls UI_SetExcelUI_WithResult with UI_Show for every managed
'     element and FailureCount, but without the FailureList argument, and
'     waits.
'   - Asserts the Boolean result is True, FailureCount is 0, and the local
'     FailureList variable is still Empty.
'   - Asserts every element visible, TitleBar included when title-bar tests
'     are included.
'
' ERROR POLICY
'   - Raises on assertion failure
'
' DEPENDENCIES
'   - UI_SetExcelUI_WithResult
'   - TST_TitleBarMode
'   - TST_WaitUI
'   - TST_AssertRibbonVisible
'   - TST_AssertApplicationProperty
'   - TST_AssertAllWindowsProperty
'   - TST_AssertTitleBarVisible
'   - TST_Log
'
' NOTES
'   The three result assertions are written inline and raise
'   TEST_ERR_BASE + 30, + 31 and + 32.
'
' HISTORY
'   2026-11-06 - Failure messages carry the structured result (#135).
'   2026-07-25
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_RUNTIME.TST_Case_WithResult_SuccessWithoutFailureList"

    Dim OK              As Boolean  'Boolean success flag returned by the API
    Dim FailureCount    As Long     'Number of recorded failures
    Dim FailureList     As Variant  'Local untouched Variant proving omission path

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log "TST_Case_WithResult_SuccessWithoutFailureList", "START", _
            "Validating structured-result success path without FailureList capture"

'------------------------------------------------------------------------------
' APPLY REQUESTED UI STATE
'------------------------------------------------------------------------------
    'Apply a deterministic visible state while omitting the optional
    'FailureList output
        OK = UI_SetExcelUI_WithResult( _
                Ribbon:=UI_Show, _
                StatusBar:=UI_Show, _
                ScrollBars:=UI_Show, _
                FormulaBar:=UI_Show, _
                Headings:=UI_Show, _
                WorkbookTabs:=UI_Show, _
                Gridlines:=UI_Show, _
                TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Show), _
                FailureCount:=FailureCount)

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT RESULT BUFFERS
'------------------------------------------------------------------------------
    'Assert that the Boolean result reports success
        If Not OK Then
            Err.Raise TEST_ERR_BASE + 30, _
                      "WithResult.NoFailureList.Result", _
                      "WithResult.NoFailureList.Result expected=True actual=False" & _
                      " | FailureCount=" & CStr(FailureCount) & _
                      " | FailureList=not requested"
        End If

    'Assert that no failures were recorded
        If FailureCount <> 0 Then
            Err.Raise TEST_ERR_BASE + 31, _
                      "WithResult.NoFailureList.Result", _
                      "WithResult.NoFailureList.Result expected FailureCount=0 actual=" & CStr(FailureCount)
        End If

    'Assert the local untouched Variant remains Empty because it was not passed
    'to the API call
        If Not IsEmpty(FailureList) Then
            Err.Raise TEST_ERR_BASE + 32, _
                      "WithResult.NoFailureList.Result", _
                      "WithResult.NoFailureList.Result expected local FailureList to remain Empty"
        End If

'------------------------------------------------------------------------------
' ASSERT APPLIED UI STATE
'------------------------------------------------------------------------------
    'Assert Ribbon visible
        TST_AssertRibbonVisible True, "WithResult.NoFailureList.Ribbon"

    'Assert StatusBar visible
        TST_AssertApplicationProperty True, "DisplayStatusBar", "WithResult.NoFailureList.StatusBar"

    'Assert ScrollBars visible
        TST_AssertApplicationProperty True, "DisplayScrollBars", "WithResult.NoFailureList.ScrollBars"

    'Assert FormulaBar visible
        TST_AssertApplicationProperty True, "DisplayFormulaBar", "WithResult.NoFailureList.FormulaBar"

    'Assert Headings visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayHeadings", "WithResult.NoFailureList.Headings"

    'Assert WorkbookTabs visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayWorkbookTabs", "WithResult.NoFailureList.WorkbookTabs"

    'Assert Gridlines visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayGridlines", "WithResult.NoFailureList.Gridlines"

    'Assert TitleBar visible when title-bar testing is enabled
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "WithResult.NoFailureList.TitleBar"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_WithResult_SuccessWithoutFailureList", "PASS", _
            "Structured-result success path without FailureList behaved as expected"

End Sub

Public Sub TST_Case_WithResult_InvalidVisibility()

'
'------------------------------------------------------------------------------
'                      VERIFY INVALID VISIBILITY FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that UI_SetExcelUI_WithResult reports a structured failure when an
'   invalid UIVisibility value is supplied
'
' WHY THIS EXISTS
'   The structured-result path should be regression-tested not only for clean
'   success but also for deterministic failure reporting when callers pass an
'   invalid tri-state value
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Calls UI_SetExcelUI_WithResult with Ribbon:=999, an invalid
'     UIVisibility value, and no other element.
'   - Asserts the Boolean result is False.
'   - Asserts FailureCount is at least 1.
'   - Asserts FailureList is not Empty.
'
' ERROR POLICY
'   - Raises on assertion failure
'
' DEPENDENCIES
'   - UI_SetExcelUI_WithResult
'   - TST_DescribeResultBuffers
'   - TST_Log
'
' NOTES
'   The three assertions are written inline and raise TEST_ERR_BASE + 40,
'   + 41 and + 42, each with the result buffers in the message. An invalid
'   UIVisibility value is rejected before any host write.
'
' HISTORY
'   2026-11-06 - Failure messages carry the structured result (#135).
'   2026-07-25
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_RUNTIME.TST_Case_WithResult_InvalidVisibility"

    Dim OK              As Boolean  'Boolean success flag returned by the API
    Dim FailureCount    As Long     'Number of recorded failures
    Dim FailureList     As Variant  'Recorded structured failures

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log "TST_Case_WithResult_InvalidVisibility", "START", _
            "Validating structured failure reporting for invalid UIVisibility input"

'------------------------------------------------------------------------------
' APPLY INVALID INPUT
'------------------------------------------------------------------------------
    'Pass an invalid tri-state value through the structured-result API
        OK = UI_SetExcelUI_WithResult( _
                Ribbon:=999, _
                FailureCount:=FailureCount, _
                FailureList:=FailureList)

'------------------------------------------------------------------------------
' ASSERT FAILURE RESULT
'------------------------------------------------------------------------------
    'Assert that the Boolean result reports failure
        If OK Then
            Err.Raise TEST_ERR_BASE + 40, _
                      "WithResult.InvalidVisibility", _
                      "WithResult.InvalidVisibility expected=False actual=True | " & _
                      TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

    'Assert that one or more failures were recorded
        If FailureCount < 1 Then
            Err.Raise TEST_ERR_BASE + 41, _
                      "WithResult.InvalidVisibility", _
                      "WithResult.InvalidVisibility expected FailureCount>=1 actual=" & CStr(FailureCount) & _
                      " | " & TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

    'Assert that FailureList was populated
        If IsEmpty(FailureList) Then
            Err.Raise TEST_ERR_BASE + 42, _
                      "WithResult.InvalidVisibility", _
                      "WithResult.InvalidVisibility expected FailureList to be populated | " & _
                      TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_WithResult_InvalidVisibility", "PASS", _
            "Structured failure reporting for invalid UIVisibility behaved as expected"

End Sub

Public Sub TST_Case_TargetScope_ActiveWindow()

'
'------------------------------------------------------------------------------
'                          VERIFY ACTIVE-WINDOW SCOPE
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that UI_TargetActiveWindow changes only the active Excel Window while
'   application-level requests still apply normally.
'
' WHY THIS EXISTS
'   UI_TargetActiveWindow limits window-level writes to the active window.
'   This case pins that limit, and that an application-level request in the
'   same call is not narrowed by it.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Activates ThisWorkbook, keeps its active window as the anchor, and
'     saves the anchor's Headings, WorkbookTabs and Gridlines and the
'     StatusBar.
'   - Opens a second window of ThisWorkbook, shows Headings, WorkbookTabs
'     and Gridlines on both windows and the StatusBar, and activates the
'     second window.
'   - Calls UI_SetExcelUI_WithResult hiding those four elements with
'     TargetScope:=UI_TargetActiveWindow, over stale buffers.
'   - Asserts a clean success, the StatusBar hidden, the second window
'     hidden, and the anchor window still showing all three.
'   - Cleanup: closes the second window, restores the saved values and
'     reactivates the anchor, then raises the failure it kept.
'
' ERROR POLICY
'   - Raises after best-effort cleanup.
'
' DEPENDENCIES
'   - UI_SetExcelUI_WithResult
'   - TST_AssertResultSuccess
'   - TST_AssertApplicationProperty
'   - TST_AssertSnapshotWindowState
'   - TST_SafeCloseWindow
'   - TST_Log
'   - ThisWorkbook.NewWindow and the window display properties
'
' NOTES
'   Raises TEST_TARGET_ERR_BASE + 1 when ThisWorkbook gives no active
'   window, and + 2 when NewWindow returns no window.
'
' UPDATED
'   2026-08-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_TargetScope_ActiveWindow"

    Dim AnchorWindow        As Window
    Dim TargetWindow        As Window
    Dim SavedHeadings       As Boolean
    Dim SavedWorkbookTabs   As Boolean
    Dim SavedGridlines      As Boolean
    Dim SavedStatusBar      As Boolean
    Dim OK                  As Boolean
    Dim FailureCount        As Long
    Dim FailureList         As Variant
    Dim HasFailure          As Boolean
    Dim FailNumber          As Long
    Dim FailSource          As String
    Dim FailDescription     As String

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        ThisWorkbook.Activate
        Set AnchorWindow = Application.ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise TEST_TARGET_ERR_BASE + 1, PROC_NAME, _
                "ThisWorkbook could not provide an active Excel window"
        End If

        SavedHeadings = AnchorWindow.DisplayHeadings
        SavedWorkbookTabs = AnchorWindow.DisplayWorkbookTabs
        SavedGridlines = AnchorWindow.DisplayGridlines
        SavedStatusBar = Application.DisplayStatusBar

        Set TargetWindow = ThisWorkbook.NewWindow

        If TargetWindow Is Nothing Then
            Err.Raise TEST_TARGET_ERR_BASE + 2, PROC_NAME, _
                "ThisWorkbook.NewWindow did not return a target window"
        End If

'------------------------------------------------------------------------------
' ESTABLISH BASELINE
'------------------------------------------------------------------------------
        AnchorWindow.DisplayHeadings = True
        AnchorWindow.DisplayWorkbookTabs = True
        AnchorWindow.DisplayGridlines = True

        TargetWindow.DisplayHeadings = True
        TargetWindow.DisplayWorkbookTabs = True
        TargetWindow.DisplayGridlines = True
        TargetWindow.Activate

        Application.DisplayStatusBar = True

'------------------------------------------------------------------------------
' APPLY ACTIVE-WINDOW SCOPE
'------------------------------------------------------------------------------
        FailureCount = 99
        FailureList = Array("stale target failure")

        OK = UI_SetExcelUI_WithResult( _
            StatusBar:=UI_Hide, _
            Headings:=UI_Hide, _
            WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Hide, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            TargetScope:=UI_TargetActiveWindow)

        TST_AssertResultSuccess _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            AssertionName:=PROC_NAME & ".Result"

'------------------------------------------------------------------------------
' ASSERT SCOPE
'------------------------------------------------------------------------------
        TST_AssertApplicationProperty _
            Expected:=False, _
            PropertyName:="DisplayStatusBar", _
            AssertionName:=PROC_NAME & ".ApplicationLevelUnaffectedByScope"

        TST_AssertSnapshotWindowState _
            TargetWindow:=TargetWindow, _
            ExpectedHeadings:=False, _
            ExpectedWorkbookTabs:=False, _
            ExpectedGridlines:=False, _
            AssertionName:=PROC_NAME & ".TargetWindow"

        TST_AssertSnapshotWindowState _
            TargetWindow:=AnchorWindow, _
            ExpectedHeadings:=True, _
            ExpectedWorkbookTabs:=True, _
            ExpectedGridlines:=True, _
            AssertionName:=PROC_NAME & ".NonTargetWindow"

        TST_Log PROC_NAME, "PASS", _
            "Only the active window received window-level changes"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        On Error Resume Next

        TST_SafeCloseWindow TargetWindow

        If Not AnchorWindow Is Nothing Then
            AnchorWindow.DisplayHeadings = SavedHeadings
            AnchorWindow.DisplayWorkbookTabs = SavedWorkbookTabs
            AnchorWindow.DisplayGridlines = SavedGridlines
            AnchorWindow.Activate
        End If

        Application.DisplayStatusBar = SavedStatusBar

        On Error GoTo 0

        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        Resume SafeExit

End Sub

Public Sub TST_Case_TargetScope_ActiveWorkbookWindows()

'
'------------------------------------------------------------------------------
'                         VERIFY ACTIVE-WORKBOOK SCOPE
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that UI_TargetActiveWorkbookWindows changes every Window belonging to
'   the active workbook and leaves another workbook's Window unchanged.
'
' WHY THIS EXISTS
'   UI_TargetActiveWorkbookWindows applies window-level writes to every
'   window of the active workbook. This case pins that a second window of
'   that workbook changes, and that a window of another workbook does not.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Remembers the active window, then creates a target workbook with a
'     second window and a separate other workbook.
'   - Shows Headings, WorkbookTabs and Gridlines on the three windows and
'     activates the target workbook's first window.
'   - Calls UI_SetExcelUI_WithResult hiding the three elements with
'     TargetScope:=UI_TargetActiveWorkbookWindows, over stale buffers.
'   - Asserts a clean success, both target windows hidden, and the other
'     workbook's window still showing all three.
'   - Cleanup: closes both temporary workbooks and reactivates the
'     original window, then raises the failure it kept.
'
' ERROR POLICY
'   - Raises after best-effort cleanup.
'
' DEPENDENCIES
'   - UI_SetExcelUI_WithResult
'   - TST_AssertResultSuccess
'   - TST_AssertSnapshotWindowState
'   - TST_SafeCloseWorkbook
'   - TST_Log
'   - Workbooks.Add and Workbook.NewWindow
'
' NOTES
'   Raises TEST_TARGET_ERR_BASE + 3 when a temporary window cannot be
'   created, and + 4 when the target workbook cannot be activated.
'
' UPDATED
'   2026-08-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_TargetScope_ActiveWorkbookWindows"

    Dim OriginalWindow  As Window
    Dim TargetBook      As Workbook
    Dim OtherBook       As Workbook
    Dim TargetWindowOne As Window
    Dim TargetWindowTwo As Window
    Dim OtherWindow     As Window
    Dim OK              As Boolean
    Dim FailureCount    As Long
    Dim FailureList     As Variant
    Dim HasFailure      As Boolean
    Dim FailNumber      As Long
    Dim FailSource      As String
    Dim FailDescription As String

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        Set OriginalWindow = Application.ActiveWindow

        Set TargetBook = Application.Workbooks.Add
        Set TargetWindowOne = TargetBook.Windows(1)
        Set TargetWindowTwo = TargetBook.NewWindow

        Set OtherBook = Application.Workbooks.Add
        Set OtherWindow = OtherBook.Windows(1)

        If TargetWindowOne Is Nothing Or _
           TargetWindowTwo Is Nothing Or _
           OtherWindow Is Nothing Then

            Err.Raise TEST_TARGET_ERR_BASE + 3, PROC_NAME, _
                "temporary workbook windows could not be created"
        End If

'------------------------------------------------------------------------------
' ESTABLISH BASELINE
'------------------------------------------------------------------------------
        TargetWindowOne.DisplayHeadings = True
        TargetWindowOne.DisplayWorkbookTabs = True
        TargetWindowOne.DisplayGridlines = True

        TargetWindowTwo.DisplayHeadings = True
        TargetWindowTwo.DisplayWorkbookTabs = True
        TargetWindowTwo.DisplayGridlines = True

        OtherWindow.DisplayHeadings = True
        OtherWindow.DisplayWorkbookTabs = True
        OtherWindow.DisplayGridlines = True

        TargetWindowOne.Activate

        If Not (Application.ActiveWorkbook Is TargetBook) Then
            Err.Raise TEST_TARGET_ERR_BASE + 4, PROC_NAME, _
                "temporary target workbook could not be activated"
        End If

'------------------------------------------------------------------------------
' APPLY ACTIVE-WORKBOOK SCOPE
'------------------------------------------------------------------------------
        FailureCount = 99
        FailureList = Array("stale target failure")

        OK = UI_SetExcelUI_WithResult( _
            Headings:=UI_Hide, _
            WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Hide, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            TargetScope:=UI_TargetActiveWorkbookWindows)

        TST_AssertResultSuccess _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            AssertionName:=PROC_NAME & ".Result"

'------------------------------------------------------------------------------
' ASSERT SCOPE
'------------------------------------------------------------------------------
        TST_AssertSnapshotWindowState _
            TargetWindow:=TargetWindowOne, _
            ExpectedHeadings:=False, _
            ExpectedWorkbookTabs:=False, _
            ExpectedGridlines:=False, _
            AssertionName:=PROC_NAME & ".TargetWindowOne"

        TST_AssertSnapshotWindowState _
            TargetWindow:=TargetWindowTwo, _
            ExpectedHeadings:=False, _
            ExpectedWorkbookTabs:=False, _
            ExpectedGridlines:=False, _
            AssertionName:=PROC_NAME & ".TargetWindowTwo"

        TST_AssertSnapshotWindowState _
            TargetWindow:=OtherWindow, _
            ExpectedHeadings:=True, _
            ExpectedWorkbookTabs:=True, _
            ExpectedGridlines:=True, _
            AssertionName:=PROC_NAME & ".OtherWorkbookWindow"

        TST_Log PROC_NAME, "PASS", _
            "Only windows belonging to the active workbook were changed"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        On Error Resume Next

        TST_SafeCloseWorkbook TargetBook
        TST_SafeCloseWorkbook OtherBook

        If Not OriginalWindow Is Nothing Then
            OriginalWindow.Activate
        End If

        On Error GoTo 0

        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        Resume SafeExit

End Sub

Public Sub TST_Case_TargetScope_InvalidValue()

'
'------------------------------------------------------------------------------
'                         VERIFY INVALID TARGET SCOPE
'------------------------------------------------------------------------------
' PURPOSE
'   Verify ordered invalid-scope diagnostics, application-level continuation,
'   and suppression of window-level writes when TargetScope is invalid.
'
' WHY THIS EXISTS
'   An unknown UIWindowTargetScope value must be reported without losing
'   the work that does not depend on it. This case pins the diagnostic,
'   the continuation of application-level work, and the suppression of
'   window-level writes.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Activates ThisWorkbook, keeps its active window as the anchor, saves
'     the anchor's Headings and the StatusBar, and shows both.
'   - Calls UI_SetExcelUI_WithResult hiding StatusBar and Headings with
'     TargetScope 999, over stale buffers.
'   - Asserts a False result with exactly one failure, whose entry starts
'     with "TargetScope | invalid UIWindowTargetScope value: 999".
'   - Asserts the StatusBar hidden and the anchor's Headings still shown.
'   - Cleanup: restores the saved values and reactivates the anchor, then
'     raises the failure it kept.
'
' ERROR POLICY
'   - Raises after best-effort cleanup.
'
' DEPENDENCIES
'   - UI_SetExcelUI_WithResult
'   - TST_AssertSingleFailurePrefix
'   - TST_AssertApplicationProperty
'   - TST_AssertBooleanEquals
'   - TST_Log
'
' NOTES
'   Raises TEST_TARGET_ERR_BASE + 5 when ThisWorkbook gives no active
'   window.
'
' UPDATED
'   2026-08-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_TargetScope_InvalidValue"

    Dim AnchorWindow    As Window
    Dim InvalidScope    As UIWindowTargetScope
    Dim SavedHeadings   As Boolean
    Dim SavedStatusBar  As Boolean
    Dim OK              As Boolean
    Dim FailureCount    As Long
    Dim FailureList     As Variant
    Dim HasFailure      As Boolean
    Dim FailNumber      As Long
    Dim FailSource      As String
    Dim FailDescription As String

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        ThisWorkbook.Activate
        Set AnchorWindow = Application.ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise TEST_TARGET_ERR_BASE + 5, PROC_NAME, _
                "ThisWorkbook could not provide an active Excel window"
        End If

        SavedHeadings = AnchorWindow.DisplayHeadings
        SavedStatusBar = Application.DisplayStatusBar

        AnchorWindow.DisplayHeadings = True
        Application.DisplayStatusBar = True
        InvalidScope = 999

'------------------------------------------------------------------------------
' APPLY INVALID SCOPE
'------------------------------------------------------------------------------
        FailureCount = 99
        FailureList = Array("stale target failure")

        OK = UI_SetExcelUI_WithResult( _
            StatusBar:=UI_Hide, _
            Headings:=UI_Hide, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            TargetScope:=InvalidScope)

        TST_AssertSingleFailurePrefix _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedPrefix:="TargetScope | invalid UIWindowTargetScope value: 999", _
            AssertionName:=PROC_NAME & ".Result"

'------------------------------------------------------------------------------
' ASSERT CONTINUATION AND WINDOW SUPPRESSION
'------------------------------------------------------------------------------
        TST_AssertApplicationProperty _
            Expected:=False, _
            PropertyName:="DisplayStatusBar", _
            AssertionName:=PROC_NAME & ".ApplicationLevelContinued"

        TST_AssertBooleanEquals _
            Expected:=True, _
            Actual:=AnchorWindow.DisplayHeadings, _
            AssertionName:=PROC_NAME & ".WindowLevelSkipped"

        TST_Log PROC_NAME, "PASS", _
            "Invalid scope was reported while application-level work continued"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        On Error Resume Next

        If Not AnchorWindow Is Nothing Then
            AnchorWindow.DisplayHeadings = SavedHeadings
            AnchorWindow.Activate
        End If

        Application.DisplayStatusBar = SavedStatusBar

        On Error GoTo 0

        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        Resume SafeExit

End Sub

Public Sub TST_Case_ScreenUpdatingPreserved()

'
'------------------------------------------------------------------------------
'                       VERIFY SCREENUPDATING PRESERVED
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that the EXCEL_UI apply path restores Application.ScreenUpdating to
'   its prior state after execution
'
' WHY THIS EXISTS
'   The core module uses a quiet-update scope with ScreenUpdating to reduce
'   worksheet redraw flicker where possible, and that behavior must remain
'   invisible to callers from a state-management perspective
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Saves Application.ScreenUpdating.
'   - Sets it True, calls UI_SetExcelUI showing StatusBar and Gridlines,
'     and asserts it is still True.
'   - Sets it False, calls UI_SetExcelUI hiding StatusBar and Gridlines,
'     and asserts it is still False.
'   - Restores the saved value.
'
' ERROR POLICY
'   - Raises on assertion failure
'
' DEPENDENCIES
'   - UI_SetExcelUI
'   - TST_AssertBooleanEquals
'   - TST_Log
'   - Application.ScreenUpdating
'
' NOTES
'   The saved value is put back on the success path only: the case has no
'   handler, so an assertion failure leaves ScreenUpdating as the case set
'   it.
'
' UPDATED
'   2026-07-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_RUNTIME.TST_Case_ScreenUpdatingPreserved"

    Dim SavedScreenUpdating As Boolean  'Caller-visible ScreenUpdating baseline

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log "TST_Case_ScreenUpdatingPreserved", "START", _
            "Validating ScreenUpdating preservation across EXCEL_UI calls"

    'Capture the caller-visible baseline so it can be restored at the end
        SavedScreenUpdating = Application.ScreenUpdating

'------------------------------------------------------------------------------
' ASSERT TRUE => TRUE
'------------------------------------------------------------------------------
    'Set ScreenUpdating to True explicitly
        Application.ScreenUpdating = True

    'Call the public API through a small deterministic mutation
        UI_SetExcelUI _
            StatusBar:=UI_Show, _
            Gridlines:=UI_Show

    'Assert that ScreenUpdating remained True from the caller's perspective
        TST_AssertBooleanEquals True, Application.ScreenUpdating, "ScreenUpdatingPreserved.TruePath"

'------------------------------------------------------------------------------
' ASSERT FALSE => FALSE
'------------------------------------------------------------------------------
    'Set ScreenUpdating to False explicitly
        Application.ScreenUpdating = False

    'Call the public API through a small deterministic mutation
        UI_SetExcelUI _
            StatusBar:=UI_Hide, _
            Gridlines:=UI_Hide

    'Assert that ScreenUpdating remained False from the caller's perspective
        TST_AssertBooleanEquals False, Application.ScreenUpdating, "ScreenUpdatingPreserved.FalsePath"

'------------------------------------------------------------------------------
' RESTORE CALLER BASELINE
'------------------------------------------------------------------------------
    'Restore the original caller-visible ScreenUpdating baseline
        Application.ScreenUpdating = SavedScreenUpdating

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_ScreenUpdatingPreserved", "PASS", _
            "ScreenUpdating preservation behaved as expected"

End Sub

Public Sub TST_Case_FailureAccumulatorDegradesSafely()

'
'------------------------------------------------------------------------------
'                     VERIFY FAILURE LIST DEGRADES SAFELY
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a failure list which cannot be grown degrades visibly instead of
'   raising, and that the authoritative status outputs survive intact.
'
' WHY THIS EXISTS
'   This is the regression for ICR-UI-P2-02. The accumulator is reached FROM
'   error handlers, so an allocation failure inside it used to replace the very
'   failure it was invoked to record, and could abort a pass designed to
'   continue.
'
'   The case drives UI_RuntimeHandleFailure directly with local buffers rather
'   than through the public facade. That is deliberate: the facade clears its
'   result buffers at the start of every call, so a facade-level test cannot
'   arrange for one entry to be recorded successfully and the NEXT one to fail,
'   which is the sequence that exercises the truncation marker.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Records one failure normally and asserts it was listed.
'   - Arms a one-shot growth failure and records a second failure.
'   - Asserts the call did not raise, the count still advanced, the list did
'     not grow, and a truncation marker was written into the existing slot.
'   - Repeats the growth failure against an empty buffer and asserts the count
'     still advances, the list stays exactly Empty and nothing raises.
'
' ERROR POLICY
'   - Raises on assertion failure, after disarming the seam.
'
' DEPENDENCIES
'   - UI_RuntimeHandleFailure
'   - UI_RuntimeClearResultBuffer
'   - UI_InternalInjectFailureListGrowthFailure
'   - TST_AssertBooleanEquals
'   - TST_AssertTrue
'   - TST_Log
'
' NOTES
'   Uses the internal seam UI_InternalInjectFailureListGrowthFailure. It is
'   Public only for same-project regression access; Option Private Module keeps
'   it out of the external automation namespace.
'
' HISTORY
'   2026-11-14 - Asserts the exact marker text and an Empty list (#38).
'   2026-08-19
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_FailureAccumulatorDegradesSafely"

    Dim Succeeded       As Boolean  'Result contract success flag
    Dim FailureCount    As Long     'Result contract failure count
    Dim FailureList     As Variant  'Result contract failure list
    Dim MarkerText      As String   'Final entry after truncation

    Dim HasFailure      As Boolean  'TRUE when a test failure occurred
    Dim FailNumber      As Long     'Captured failure number
    Dim FailSource      As String   'Captured failure source
    Dim FailDescription As String   'Captured failure description

    Const MARKER_TEXT As String = _
        "Diagnostics | failure list could not be grown; " & _
        "2 failures were counted and fewer are listed"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_Log PROC_NAME, "START", _
            "Validating that failure accumulation degrades instead of raising"

    'Start from the documented empty result contract
        Succeeded = True
        UI_RuntimeClearResultBuffer FailureCount, FailureList, True

'------------------------------------------------------------------------------
' RECORD ONE FAILURE NORMALLY
'------------------------------------------------------------------------------
    'A healthy append is the precondition for the truncation case below
        UI_RuntimeHandleFailure _
            PROC_NAME, False, Succeeded, FailureCount, FailureList, True, _
            "StageOne", "first failure detail"

        TST_AssertBooleanEquals _
            False, Succeeded, "FailureAccumulator.SucceededCleared"

        TST_AssertTrue _
            (FailureCount = 1), "FailureAccumulator.FirstCount"

        TST_AssertTrue _
            IsArray(FailureList), "FailureAccumulator.FirstListIsArray"

        TST_AssertTrue _
            (UBound(FailureList) = 1), "FailureAccumulator.FirstListBound"

'------------------------------------------------------------------------------
' FAIL THE NEXT GROWTH
'------------------------------------------------------------------------------
    'Arm the one-shot seam, then record a second failure. Nothing here may
    'raise: the accumulator runs inside error handlers.
        UI_InternalInjectFailureListGrowthFailure True

        UI_RuntimeHandleFailure _
            PROC_NAME, False, Succeeded, FailureCount, FailureList, True, _
            "StageTwo", "second failure detail"

'------------------------------------------------------------------------------
' ASSERT THE COUNT SURVIVED
'------------------------------------------------------------------------------
    'The count is authoritative and must advance even when nothing was listed
        TST_AssertTrue _
            (FailureCount = 2), "FailureAccumulator.CountAdvancedDespiteGrowthFailure"

'------------------------------------------------------------------------------
' ASSERT THE LIST DID NOT GROW
'------------------------------------------------------------------------------
    'A silently short list is the outcome this case exists to rule out
        TST_AssertTrue _
            (UBound(FailureList) = 1), "FailureAccumulator.ListDidNotGrow"

'------------------------------------------------------------------------------
' ASSERT THE TRUNCATION WAS REPORTED
'------------------------------------------------------------------------------
    'The marker must occupy a slot that already existed, so that reporting the
    'allocation failure did not itself require an allocation
        MarkerText = CStr(FailureList(UBound(FailureList)))

        TST_AssertTrue _
            (StrComp(MarkerText, MARKER_TEXT, vbBinaryCompare) = 0), _
            "FailureAccumulator.TruncationMarkerWritten"

'------------------------------------------------------------------------------
' FAIL A GROWTH WITH NO EXISTING SLOT
'------------------------------------------------------------------------------
    'With an empty buffer there is nowhere to write a marker without
    'allocating. The count must still advance and nothing may raise.
        Succeeded = True
        UI_RuntimeClearResultBuffer FailureCount, FailureList, True

        UI_InternalInjectFailureListGrowthFailure True

        UI_RuntimeHandleFailure _
            PROC_NAME, False, Succeeded, FailureCount, FailureList, True, _
            "StageThree", "third failure detail"

        TST_AssertBooleanEquals _
            False, Succeeded, "FailureAccumulator.EmptyBufferSucceededCleared"

        TST_AssertTrue _
            (FailureCount = 1), "FailureAccumulator.EmptyBufferCountAdvanced"

        TST_AssertTrue _
            (VarType(FailureList) = vbEmpty), "FailureAccumulator.EmptyBufferListEmpty"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC_NAME, "PASS", _
            "Accumulation degraded visibly and never raised"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Disarm the seam whatever happened above
        On Error Resume Next
        UI_InternalInjectFailureListGrowthFailure False
        On Error GoTo 0

    'Raise the captured failure after cleanup when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume SafeExit

End Sub

Public Sub TST_Case_FirstFailureAllocationKeepsContract()

'
'------------------------------------------------------------------------------
'                   VERIFY FIRST ALLOCATION FAILURE CONTRACT
'------------------------------------------------------------------------------
' PURPOSE
'   Verify the documented result contract when the FIRST failure-list entry
'   cannot be allocated: the call still fails, the count stays authoritative,
'   and the list is exactly Empty.
'
' WHY THIS EXISTS
'   A failed first allocation leaves no slot in which a truncation marker could
'   be written, so its only visible form is a positive FailureCount with an
'   Empty FailureList (#38). This case pins that form, through the public
'   facade as callers see it and through the accumulator for buffers the
'   facade never passes.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Fails the first allocation of an invalid-Ribbon call made over stale
'     buffers, and asserts False, a count of 1 and an Empty list.
'   - Repeats the call without the seam over the same buffers, and asserts the
'     exact entry text, which also shows the seam was one-shot.
'   - Fails the first allocation of a two-failure call, and asserts a count of
'     2 with only the second entry listed.
'   - Shows that a call without a list attempts no allocation.
'   - Reuses the buffers for a clean call, and asserts True, 0 and Empty.
'   - Drives the accumulator over a non-array buffer, with and without the
'     seam, and asserts an Empty list and then a fresh exact entry.
'   - Shows that growing an existing list neither fails nor consumes the seam.
'
' ERROR POLICY
'   - Raises on assertion failure, after disarming the seams.
'
' DEPENDENCIES
'   - UI_SetExcelUI_WithResult
'   - UI_RuntimeHandleFailure
'   - UI_RuntimeClearResultBuffer
'   - UI_InternalInjectFailureListFirstAllocationFailure
'   - UI_InternalInjectFailureListGrowthFailure
'   - TST_AssertFailureResultExact
'   - TST_AssertResultSuccess
'   - TST_AssertBooleanEquals
'   - TST_AssertTrue
'   - TST_Log
'
' NOTES
'   Uses UI_InternalInjectFailureListFirstAllocationFailure. Invalid
'   UIVisibility values are rejected before any host write, so the facade
'   calls change no Excel state.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_FirstFailureAllocationKeepsContract"

    Dim OK                  As Boolean      'Result returned by the facade
    Dim Succeeded           As Boolean      'Accumulator success flag
    Dim FailureCount        As Long         'Result contract failure count
    Dim FailureList         As Variant      'Result contract failure list
    Dim InvalidVisibility   As UIVisibility 'Value the facade must reject

    Dim HasFailure      As Boolean  'TRUE when a test failure occurred
    Dim FailNumber      As Long     'Captured failure number
    Dim FailSource      As String   'Captured failure source
    Dim FailDescription As String   'Captured failure description

    Const RIBBON_ENTRY  As String = "Ribbon | invalid UIVisibility value: 999"
    Const STATUSBAR_ENTRY As String = _
        "StatusBar | invalid UIVisibility value: 999"
    Const CORRUPT_ENTRY As String = "StageCorrupt | corrupt buffer detail"
    Const GROW_ENTRY    As String = "StageGrow | grow detail"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_Log PROC_NAME, "START", _
            "Validating the result contract when the first list entry fails"

        InvalidVisibility = 999

'------------------------------------------------------------------------------
' FAIL THE FIRST ALLOCATION THROUGH THE FACADE
'------------------------------------------------------------------------------
    'Stale buffers prove the call clears them before it records anything
        FailureCount = 99
        FailureList = Array("stale entry")

        UI_InternalInjectFailureListFirstAllocationFailure True

        OK = UI_SetExcelUI_WithResult( _
            Ribbon:=InvalidVisibility, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertFailureResultExact _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedCount:=1, _
            ExpectedEntry:=vbNullString, _
            AssertionName:=PROC_NAME & ".FirstAllocationFailed"

'------------------------------------------------------------------------------
' REPEAT WITHOUT THE SEAM
'------------------------------------------------------------------------------
    'The seam was consumed, so the same call over the same buffers now lists
    'the failure it reports
        OK = UI_SetExcelUI_WithResult( _
            Ribbon:=InvalidVisibility, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertFailureResultExact _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedCount:=1, _
            ExpectedEntry:=RIBBON_ENTRY, _
            AssertionName:=PROC_NAME & ".SeamWasOneShot"

'------------------------------------------------------------------------------
' FAIL THE FIRST OF TWO ENTRIES
'------------------------------------------------------------------------------
    'Recording continues after the failed first allocation: the count covers
    'both failures and the list holds the second, so the count exceeds it
        UI_InternalInjectFailureListFirstAllocationFailure True

        OK = UI_SetExcelUI_WithResult( _
            Ribbon:=InvalidVisibility, _
            StatusBar:=InvalidVisibility, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertFailureResultExact _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedCount:=2, _
            ExpectedEntry:=STATUSBAR_ENTRY, _
            AssertionName:=PROC_NAME & ".LaterEntryRecorded"

'------------------------------------------------------------------------------
' OMIT THE LIST
'------------------------------------------------------------------------------
    'Without a list argument nothing is allocated, so the armed seam survives
    'the call and the caller's variable is untouched
        UI_InternalInjectFailureListFirstAllocationFailure True

        OK = UI_SetExcelUI_WithResult( _
            Ribbon:=InvalidVisibility, _
            FailureCount:=FailureCount)

        TST_AssertBooleanEquals _
            False, OK, PROC_NAME & ".NoListResult"

        TST_AssertTrue _
            (FailureCount = 1), PROC_NAME & ".NoListCount"

        TST_AssertTrue _
            (UBound(FailureList) = 1), PROC_NAME & ".NoListLeftUntouched"

        TST_AssertTrue _
            (StrComp(FailureList(1), STATUSBAR_ENTRY, vbBinaryCompare) = 0), _
            PROC_NAME & ".NoListEntryUntouched"

    'The next call that does capture a list consumes the seam still armed
        OK = UI_SetExcelUI_WithResult( _
            Ribbon:=InvalidVisibility, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertFailureResultExact _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedCount:=1, _
            ExpectedEntry:=vbNullString, _
            AssertionName:=PROC_NAME & ".SeamSurvivedNoListCall"

'------------------------------------------------------------------------------
' REUSE THE BUFFERS FOR A CLEAN CALL
'------------------------------------------------------------------------------
    'A later success must not be read as carrying the degraded result
        OK = UI_SetExcelUI_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertResultSuccess _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            AssertionName:=PROC_NAME & ".LaterSuccessCleared"

'------------------------------------------------------------------------------
' FAIL THE FIRST ALLOCATION OVER A NON-ARRAY BUFFER
'------------------------------------------------------------------------------
    'The facade always starts from Empty, so a buffer holding a foreign scalar
    'can reach only the accumulator. It must come back Empty, not unchanged.
        Succeeded = True
        FailureCount = 0
        FailureList = "corrupt buffer"

        UI_InternalInjectFailureListFirstAllocationFailure True

        UI_RuntimeHandleFailure _
            PROC_NAME, False, Succeeded, FailureCount, FailureList, True, _
            "StageCorrupt", "corrupt buffer detail"

        TST_AssertFailureResultExact _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedCount:=1, _
            ExpectedEntry:=vbNullString, _
            AssertionName:=PROC_NAME & ".CorruptBufferLeftEmpty"

    'Without the seam the same buffer is replaced by a fresh list
        Succeeded = True
        FailureCount = 0
        FailureList = "corrupt buffer"

        UI_RuntimeHandleFailure _
            PROC_NAME, False, Succeeded, FailureCount, FailureList, True, _
            "StageCorrupt", "corrupt buffer detail"

        TST_AssertFailureResultExact _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedCount:=1, _
            ExpectedEntry:=CORRUPT_ENTRY, _
            AssertionName:=PROC_NAME & ".CorruptBufferReplaced"

'------------------------------------------------------------------------------
' GROW AN EXISTING LIST WHILE ARMED
'------------------------------------------------------------------------------
    'Growth is not a first allocation: it must succeed and leave the seam armed
        UI_InternalInjectFailureListFirstAllocationFailure True

        UI_RuntimeHandleFailure _
            PROC_NAME, False, Succeeded, FailureCount, FailureList, True, _
            "StageGrow", "grow detail"

        TST_AssertTrue _
            (FailureCount = 2), PROC_NAME & ".GrowthCount"

        TST_AssertTrue _
            (UBound(FailureList) = 2), PROC_NAME & ".GrowthNotFailed"

        TST_AssertTrue _
            (StrComp(FailureList(2), GROW_ENTRY, vbBinaryCompare) = 0), _
            PROC_NAME & ".GrowthEntryText"

    'Cleared buffers make the next append a first allocation, which consumes it
        Succeeded = True
        UI_RuntimeClearResultBuffer FailureCount, FailureList, True

        UI_RuntimeHandleFailure _
            PROC_NAME, False, Succeeded, FailureCount, FailureList, True, _
            "StageCorrupt", "corrupt buffer detail"

        TST_AssertFailureResultExact _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedCount:=1, _
            ExpectedEntry:=vbNullString, _
            AssertionName:=PROC_NAME & ".SeamSurvivedGrowth"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC_NAME, "PASS", _
            "A failed first allocation left a counted failure and an Empty list"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Disarm both list seams whatever happened above
        On Error Resume Next
        UI_InternalInjectFailureListFirstAllocationFailure False
        UI_InternalInjectFailureListGrowthFailure False
        On Error GoTo 0

    'Raise the captured failure after cleanup when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume SafeExit

End Sub

Public Sub TST_Case_QuietScopeOwnershipIsVerified()

'
'------------------------------------------------------------------------------
'                         VERIFY QUIET SCOPE OWNERSHIP
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that the quiet-update scope claims ownership only of a suppression
'   it confirmed, and restores only what it owns.
'
' WHY THIS EXISTS
'   The scope used to set its ownership flag immediately after attempting the
'   write. Under On Error Resume Next a refused or ignored write left no trace,
'   so the flag recorded an attempt and the matching End wrote a value back for
'   a transition that never happened - over a caller's own setting.
'
'   Every branch that must refuse ownership is a host refusal, so each is
'   reached through the one-shot seam in M_EXCEL_UI_RUNTIME. The two branches
'   that need no seam - the ordinary path and an already-suppressed caller -
'   are exercised directly, because they are the ones a change to this logic
'   would break first.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Six branches. Each asserts the ownership flag, the reported entry value
'   where it is meaningful, and the host value left behind after End.
'
' ERROR POLICY
'   - Raises on assertion failure
'   - Restores ScreenUpdating itself; the scope is not trusted to
'
' DEPENDENCIES
'   - UI_InternalInjectQuietUpdateFault
'   - UI_RuntimeBeginQuietUpdate
'   - UI_RuntimeEndQuietUpdate
'
' NOTES
'   ScreenUpdating is restored explicitly after every branch rather than by the
'   scope under test. A test that relied on the code it is testing to clean up
'   would pass for the wrong reason on exactly the branches that matter.
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' UPDATED
'   2026-08-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_QuietScopeOwnershipIsVerified"

    Dim OldValue    As Boolean  'Entry value the scope reports
    Dim Changed     As Boolean  'Ownership the scope claims

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log PROC_NAME, "START", _
            "Validating that quiet-scope ownership follows a confirmed readback"

        UI_InternalInjectQuietUpdateFault TEST_QUIET_FAULT_NONE
        Application.ScreenUpdating = True

'------------------------------------------------------------------------------
' BRANCH 1 - ORDINARY PATH OWNS THE TRANSITION
'------------------------------------------------------------------------------
        OldValue = False
        Changed = False

        UI_RuntimeBeginQuietUpdate OldValue, Changed

        TST_AssertBooleanEquals True, OldValue, PROC_NAME & ".Ordinary.EntryValue"
        TST_AssertBooleanEquals True, Changed, PROC_NAME & ".Ordinary.Owned"
        TST_AssertApplicationProperty _
            False, "ScreenUpdating", PROC_NAME & ".Ordinary.Suppressed"

        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC_NAME & ".Ordinary.Restored"

'------------------------------------------------------------------------------
' BRANCH 2 - AN ALREADY-SUPPRESSED CALLER IS NOT OWNED
'------------------------------------------------------------------------------
        Application.ScreenUpdating = False

        OldValue = True
        Changed = True

        UI_RuntimeBeginQuietUpdate OldValue, Changed

        TST_AssertBooleanEquals False, OldValue, PROC_NAME & ".AlreadyOff.EntryValue"
        TST_AssertBooleanEquals False, Changed, PROC_NAME & ".AlreadyOff.NotOwned"

        UI_RuntimeEndQuietUpdate OldValue, Changed

    'End must not re-enable a caller's suppression
        TST_AssertApplicationProperty _
            False, "ScreenUpdating", PROC_NAME & ".AlreadyOff.Untouched"

        Application.ScreenUpdating = True

'------------------------------------------------------------------------------
' BRANCH 3 - A FAILED ENTRY READ OWNS NOTHING
'------------------------------------------------------------------------------
        UI_InternalInjectQuietUpdateFault TEST_QUIET_FAULT_ENTRY

        OldValue = False
        Changed = True

        UI_RuntimeBeginQuietUpdate OldValue, Changed

    'The entry value is defined but not an observation, so only its
    'determinism is asserted
        TST_AssertBooleanEquals True, OldValue, PROC_NAME & ".EntryFails.Deterministic"
        TST_AssertBooleanEquals False, Changed, PROC_NAME & ".EntryFails.NotOwned"
        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC_NAME & ".EntryFails.HostUntouched"

        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC_NAME & ".EntryFails.NoRestore"

'------------------------------------------------------------------------------
' BRANCH 4 - AN INEFFECTIVE WRITE OWNS NOTHING
'------------------------------------------------------------------------------
        UI_InternalInjectQuietUpdateFault TEST_QUIET_FAULT_WRITE

        OldValue = False
        Changed = True

        UI_RuntimeBeginQuietUpdate OldValue, Changed

        TST_AssertBooleanEquals True, OldValue, PROC_NAME & ".WriteIgnored.EntryValue"
        TST_AssertBooleanEquals False, Changed, PROC_NAME & ".WriteIgnored.NotOwned"

    'The readback reported the host still enabled, which is what refused
    'ownership; the host really is still enabled
        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC_NAME & ".WriteIgnored.HostUnchanged"

        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC_NAME & ".WriteIgnored.NoRestore"

'------------------------------------------------------------------------------
' BRANCH 5 - A FAILED READBACK OWNS NOTHING, AND NOTHING IS ROLLED BACK
'------------------------------------------------------------------------------
    'Here the write really happens and only verification fails. The host is
    'therefore left suppressed, deliberately: rolling back would be a second
    'unverified write. This case records that outcome rather than hiding it,
    'and restores the baseline itself afterwards.
        UI_InternalInjectQuietUpdateFault TEST_QUIET_FAULT_READBACK

        OldValue = False
        Changed = True

        UI_RuntimeBeginQuietUpdate OldValue, Changed

        TST_AssertBooleanEquals True, OldValue, PROC_NAME & ".ReadbackFails.EntryValue"
        TST_AssertBooleanEquals False, Changed, PROC_NAME & ".ReadbackFails.NotOwned"

    'The documented consequence: suppressed, unowned, not rolled back
        TST_AssertApplicationProperty _
            False, "ScreenUpdating", PROC_NAME & ".ReadbackFails.LeftSuppressed"

        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            False, "ScreenUpdating", PROC_NAME & ".ReadbackFails.NoRestore"

    'The scope did not own it, so this case restores the baseline
        Application.ScreenUpdating = True

'------------------------------------------------------------------------------
' BRANCH 6 - THE SEAM IS ONE-SHOT
'------------------------------------------------------------------------------
    'Armed once, consumed once. A fault that survived would break every later
    'scope in the session, and the symptom would look like this defect.
        UI_InternalInjectQuietUpdateFault TEST_QUIET_FAULT_WRITE

        OldValue = False
        Changed = False
        UI_RuntimeBeginQuietUpdate OldValue, Changed
        UI_RuntimeEndQuietUpdate OldValue, Changed

        OldValue = False
        Changed = False

        UI_RuntimeBeginQuietUpdate OldValue, Changed

        TST_AssertBooleanEquals True, Changed, PROC_NAME & ".OneShot.SecondScopeOwns"

        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC_NAME & ".OneShot.Restored"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
        TST_Log PROC_NAME, "PASS", _
            "Quiet-scope ownership followed the readback on every branch"

End Sub
