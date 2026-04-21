Attribute VB_Name = "M_EXCEL_UI_REGRESSION_TESTS"
'==============================================================================
'                    MODULE: EXCEL_UI_REGRESSION_TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Provide a regression-test harness for the EXCEL_UI module
'
' WHY THIS EXISTS
'   UI-control code is easy to break accidentally when refining:
'     - tri-state behavior
'     - selective application
'     - leave-unchanged semantics
'     - convenience wrappers
'     - WinAPI-based title-bar control
'     - the structured-result path
'     - the explicit snapshot / reset lifecycle
'
'   A repeatable regression harness reduces the risk of silent regressions and
'   makes the repository more maintainable and release-ready
'
' PUBLIC SURFACE
'   - Test_EXCEL_UI_RunAll
'   - Test_EXCEL_UI_RunCore
'   - Test_EXCEL_UI_RunTitleBarOnly
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
'   Snapshot / restore tests
'     - explicit snapshot lifecycle
'     - reset without snapshot is a no-op
'     - skipped when an explicit EXCEL_UI snapshot already existed before the
'       regression run because the harness cannot safely restore that prior
'       snapshot object
'
'   Environment-preservation tests
'     - ScreenUpdating restored to prior state
'
'   Title-bar tests
'     - hide / show round-trip
'
' STATE MANAGEMENT
'   - The harness snapshots the current Excel UI state before testing
'   - The harness attempts to restore that state at the end, even if a test
'     fails
'   - The harness does not overwrite a pre-existing explicit EXCEL_UI snapshot
'     lifecycle unless the snapshot-related cases are allowed to run
'
' LIMITATIONS
'   - Ribbon visibility is read using best-effort logic
'   - Window-level assertions use the current Application.Windows collection at
'     runtime
'   - Title-bar behavior remains the most OS / Excel-version-sensitive area
'
' COMPATIBILITY
'   - Windows only for title-bar validation
'   - Assumes the EXCEL_UI module is present in the same VBA project
'
' UPDATED
'   2026-04-19
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.0.0
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit declaration of all variables
    Option Private Module
    
'------------------------------------------------------------------------------
' TEST CONFIGURATION
'------------------------------------------------------------------------------
    Private Const TEST_WAIT_SECONDS   As Double = 0.15                'Small UI settle delay after each state change
    Private Const TEST_ERR_BASE       As Long = vbObjectError + 4700  'Base custom error number for test assertions
    Private Const TST_SECONDS_PER_DAY As Double = 86400#              'Timer rollover interval in seconds

'------------------------------------------------------------------------------
' WIN32 / WIN64 API FOR TITLE-BAR STATE READ
'------------------------------------------------------------------------------
    #If VBA7 Then
        #If Win64 Then
            Private Declare PtrSafe Function TST_GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
                ByVal hWnd As LongPtr, _
                ByVal nIndex As Long) _
                As LongPtr
        #Else
            Private Declare PtrSafe Function TST_GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As LongPtr, _
                ByVal nIndex As Long) _
                As Long
        #End If
        Private Declare PtrSafe Function TST_GetLastError Lib "kernel32" Alias "GetLastError" () As Long
        Private Declare PtrSafe Sub TST_SetLastError Lib "kernel32" Alias "SetLastError" ( _
            ByVal dwErrCode As Long)
    #Else
        Private Declare Function TST_GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
            ByVal hWnd As Long, _
            ByVal nIndex As Long) _
            As Long
        Private Declare Function TST_GetLastError Lib "kernel32" Alias "GetLastError" () As Long
        Private Declare Sub TST_SetLastError Lib "kernel32" Alias "SetLastError" ( _
            ByVal dwErrCode As Long)
    #End If

'------------------------------------------------------------------------------
' API CONSTANTS FOR TITLE-BAR STATE READ
'------------------------------------------------------------------------------
    Private Const TST_GWL_STYLE  As Long = -16       'Window style index
    Private Const TST_WS_CAPTION As Long = &HC00000  'Caption / title-bar style bit


'
'------------------------------------------------------------------------------
'
'                              PUBLIC RUNNERS
'
'------------------------------------------------------------------------------
'

Public Sub Test_EXCEL_UI_RunAll()

'
'==============================================================================
'                         Test_EXCEL_UI_RunAll
'------------------------------------------------------------------------------
' PURPOSE
'   Run the full regression-test pack for EXCEL_UI, including title-bar tests
'
' WHY THIS EXISTS
'   A single entry point is useful when validating the whole module before a
'   release, refactor, or repository update
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Snapshots current state
'   - Runs the core regression cases
'   - Runs wrapper and title-bar regression cases
'   - Attempts to restore the original state
'
' ERROR POLICY
'   - Raises on assertion failure after attempting restoration
'
' DEPENDENCIES
'   - TST_RunRegressionPack
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' APPLY FULL PACK
'------------------------------------------------------------------------------
    'Run the full regression pack including title-bar tests
        TST_RunRegressionPack IncludeTitleBarTests:=True, CallerProc:="Test_EXCEL_UI_RunAll"

End Sub

Public Sub Test_EXCEL_UI_RunCore()

'
'==============================================================================
'                         Test_EXCEL_UI_RunCore
'------------------------------------------------------------------------------
' PURPOSE
'   Run the core regression-test pack for EXCEL_UI, excluding the dedicated
'   title-bar cases and the wrapper case that also toggles TitleBar
'
' WHY THIS EXISTS
'   Core UI-state tests are useful when faster and less intrusive validation is
'   preferred
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Snapshots current state
'   - Runs the core regression cases
'   - Skips the convenience-wrapper case because the wrappers also affect
'     TitleBar
'   - Skips the dedicated title-bar round-trip case
'   - Attempts to restore the original state
'
' ERROR POLICY
'   - Raises on assertion failure after attempting restoration
'
' DEPENDENCIES
'   - TST_RunRegressionPack
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' APPLY CORE PACK
'------------------------------------------------------------------------------
    'Run the core regression pack without title-bar-specific cases
        TST_RunRegressionPack IncludeTitleBarTests:=False, CallerProc:="Test_EXCEL_UI_RunCore"

End Sub

Public Sub Test_EXCEL_UI_RunTitleBarOnly()

'
'==============================================================================
'                      Test_EXCEL_UI_RunTitleBarOnly
'------------------------------------------------------------------------------
' PURPOSE
'   Run only the dedicated title-bar regression case
'
' WHY THIS EXISTS
'   Title-bar behavior is the most WinAPI-sensitive area and benefits from a
'   focused runner that can be executed independently
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Snapshots current state
'   - Runs only the title-bar round-trip case
'   - Attempts to restore the original state
'
' ERROR POLICY
'   - Raises on assertion failure after attempting restoration
'
' DEPENDENCIES
'   - TST_RunTitleBarOnlyPack
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' APPLY TITLE-BAR-ONLY PACK
'------------------------------------------------------------------------------
    'Run the title-bar-only regression pack
        TST_RunTitleBarOnlyPack CallerProc:="Test_EXCEL_UI_RunTitleBarOnly"

End Sub


'
'------------------------------------------------------------------------------
'
'                          PRIVATE PACK RUNNERS
'
'------------------------------------------------------------------------------
'

Private Sub TST_RunRegressionPack( _
    ByVal IncludeTitleBarTests As Boolean, _
    ByVal CallerProc As String)

'
'==============================================================================
'                         TST_RunRegressionPack
'------------------------------------------------------------------------------
' PURPOSE
'   Execute the requested regression-test pack and restore the pre-test UI
'   state afterward
'
' WHY THIS EXISTS
'   The public runners differ mainly by whether title-bar tests are included,
'   so the main harness logic is centralized here
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => include wrapper and title-bar round-trip cases
'     FALSE => skip the wrapper case and the dedicated title-bar round-trip case
'
'   CallerProc
'     Public caller procedure name used for diagnostics
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Snapshots current UI state
'   - Runs the requested regression cases
'   - Skips snapshot-destructive lifecycle cases when an explicit EXCEL_UI
'     snapshot already existed before the run
'   - Attempts to restore the original UI state at the end
'
' ERROR POLICY
'   - Raises after restoration on assertion failure or unexpected error
'
' DEPENDENCIES
'   - TST_SnapshotState
'   - TST_RestoreState
'   - regression case routines
'   - TST_Log
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedRibbonKnown            As Boolean   'TRUE when pre-test Ribbon state was read successfully
    Dim SavedRibbonVisible          As Boolean   'Pre-test Ribbon visibility
    Dim SavedStatusBarVisible       As Boolean   'Pre-test StatusBar visibility
    Dim SavedScrollBarsVisible      As Boolean   'Pre-test ScrollBars visibility
    Dim SavedFormulaBarVisible      As Boolean   'Pre-test FormulaBar visibility

    Dim SavedWindowCount            As Long      'Pre-test Application.Windows.Count
    Dim SavedHeadingsVisible()      As Boolean   'Pre-test per-window Headings visibility
    Dim SavedWorkbookTabsVisible()  As Boolean   'Pre-test per-window WorkbookTabs visibility
    Dim SavedGridlinesVisible()     As Boolean   'Pre-test per-window Gridlines visibility

    Dim SavedTitleBarKnown          As Boolean   'TRUE when pre-test title-bar state was read successfully
    Dim SavedTitleBarVisible        As Boolean   'Pre-test title-bar visibility

    Dim HadExplicitSnapshot         As Boolean   'TRUE when an explicit EXCEL_UI snapshot already existed before the run
    Dim OldScreenUpdating           As Boolean   'Cached ScreenUpdating state
    Dim HasFailure                  As Boolean   'TRUE when a test failure occurred
    Dim FailNumber                  As Long      'Captured failure number
    Dim FailSource                  As String    'Captured failure source
    Dim FailDescription             As String    'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Capture whether an explicit EXCEL_UI snapshot already exists before the run
        HadExplicitSnapshot = UI_HasExcelUIStateSnapshot

    'Cache and suppress screen updates during the regression run
        OldScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False

    'Log the start of the requested regression pack
        TST_Log CallerProc, "START", "Regression pack started"

'------------------------------------------------------------------------------
' SNAPSHOT CURRENT STATE
'------------------------------------------------------------------------------
    'Snapshot the current Excel UI state before the tests mutate it
        TST_SnapshotState _
            RibbonKnown:=SavedRibbonKnown, _
            RibbonVisible:=SavedRibbonVisible, _
            StatusBarVisible:=SavedStatusBarVisible, _
            ScrollBarsVisible:=SavedScrollBarsVisible, _
            FormulaBarVisible:=SavedFormulaBarVisible, _
            WindowCount:=SavedWindowCount, _
            HeadingsVisible:=SavedHeadingsVisible, _
            WorkbookTabsVisible:=SavedWorkbookTabsVisible, _
            GridlinesVisible:=SavedGridlinesVisible, _
            TitleBarKnown:=SavedTitleBarKnown, _
            TitleBarVisible:=SavedTitleBarVisible

'------------------------------------------------------------------------------
' RUN CORE REGRESSION CASES
'------------------------------------------------------------------------------
    'Run the show-all baseline case
        TST_Case_ShowAllBaseline IncludeTitleBarTests

    'Run the selective-hide case
        TST_Case_SelectiveHide IncludeTitleBarTests

    'Run the selective-show case
        TST_Case_SelectiveShow IncludeTitleBarTests

    'Run the no-op / leave-unchanged case
        TST_Case_NoOpLeaveUnchanged IncludeTitleBarTests

    'Run the structured-result success case
        TST_Case_WithResult_AllSuccess IncludeTitleBarTests

    'Run the structured-result no-op case
        TST_Case_WithResult_NoOpSuccess IncludeTitleBarTests

    'Run the structured-result success case without FailureList capture
        TST_Case_WithResult_SuccessWithoutFailureList IncludeTitleBarTests

    'Run the structured-result invalid-visibility failure case
        TST_Case_WithResult_InvalidVisibility

    'Run the ScreenUpdating preservation case
        TST_Case_ScreenUpdatingPreserved

'------------------------------------------------------------------------------
' RUN OPTIONAL SNAPSHOT CASES
'------------------------------------------------------------------------------
    'Run snapshot-related cases only when no explicit EXCEL_UI snapshot already
    'existed before the run because the harness cannot restore that prior
    'snapshot object safely
        If HadExplicitSnapshot Then

            'Log that snapshot-destructive cases were skipped
                TST_Log CallerProc, "SKIP", _
                    "Snapshot lifecycle cases skipped because an explicit EXCEL_UI snapshot already existed before the run"

        Else

            'Run the explicit snapshot lifecycle case
                TST_Case_SnapshotLifecycle IncludeTitleBarTests

            'Run the reset-without-snapshot no-op case
                TST_Case_ResetWithoutSnapshot_NoOp IncludeTitleBarTests

        End If

'------------------------------------------------------------------------------
' RUN OPTIONAL TITLE-BAR / WRAPPER CASES
'------------------------------------------------------------------------------
    'Run the convenience-wrapper case only when title-bar testing is enabled
    'because the wrappers also affect TitleBar
        If IncludeTitleBarTests Then
            TST_Case_ConvenienceWrappers True
        Else
            TST_Log CallerProc, "SKIP", _
                "Convenience-wrapper case skipped in core mode because the wrappers also toggle TitleBar"
        End If

    'Run the dedicated title-bar case when requested
        If IncludeTitleBarTests Then
            TST_Case_TitleBarRoundTrip
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
    'Log successful completion before restoration
        TST_Log CallerProc, "PASS", "All requested regression cases passed"

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Attempt to restore the original pre-test UI state
        On Error Resume Next
            TST_RestoreState _
                RibbonKnown:=SavedRibbonKnown, _
                RibbonVisible:=SavedRibbonVisible, _
                StatusBarVisible:=SavedStatusBarVisible, _
                ScrollBarsVisible:=SavedScrollBarsVisible, _
                FormulaBarVisible:=SavedFormulaBarVisible, _
                WindowCount:=SavedWindowCount, _
                HeadingsVisible:=SavedHeadingsVisible, _
                WorkbookTabsVisible:=SavedWorkbookTabsVisible, _
                GridlinesVisible:=SavedGridlinesVisible, _
                TitleBarKnown:=SavedTitleBarKnown, _
                TitleBarVisible:=SavedTitleBarVisible
        On Error GoTo 0

    'Restore ScreenUpdating before leaving the harness
        Application.ScreenUpdating = OldScreenUpdating

    'Raise the captured failure after restoration when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Capture failure information so it can be re-raised after restoration
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description & _
                          IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

    'Log the failure immediately
        TST_Log CallerProc, "FAIL", TST_BuildRuntimeErrorText

    'Proceed to restoration / re-raise path
        Resume SafeExit

End Sub

Private Sub TST_RunTitleBarOnlyPack(ByVal CallerProc As String)

'
'==============================================================================
'                        TST_RunTitleBarOnlyPack
'------------------------------------------------------------------------------
' PURPOSE
'   Execute only the dedicated title-bar regression case and restore the
'   pre-test UI state afterward
'
' WHY THIS EXISTS
'   Title-bar behavior is the most environment-sensitive area and benefits from
'   a focused execution path
'
' INPUTS
'   CallerProc
'     Public caller procedure name used for diagnostics
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises after restoration on assertion failure or unexpected error
'
' DEPENDENCIES
'   - TST_SnapshotState
'   - TST_RestoreState
'   - TST_Case_TitleBarRoundTrip
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedRibbonKnown            As Boolean   'TRUE when pre-test Ribbon state was read successfully
    Dim SavedRibbonVisible          As Boolean   'Pre-test Ribbon visibility
    Dim SavedStatusBarVisible       As Boolean   'Pre-test StatusBar visibility
    Dim SavedScrollBarsVisible      As Boolean   'Pre-test ScrollBars visibility
    Dim SavedFormulaBarVisible      As Boolean   'Pre-test FormulaBar visibility

    Dim SavedWindowCount            As Long      'Pre-test Application.Windows.Count
    Dim SavedHeadingsVisible()      As Boolean   'Pre-test per-window Headings visibility
    Dim SavedWorkbookTabsVisible()  As Boolean   'Pre-test per-window WorkbookTabs visibility
    Dim SavedGridlinesVisible()     As Boolean   'Pre-test per-window Gridlines visibility

    Dim SavedTitleBarKnown          As Boolean   'TRUE when pre-test title-bar state was read successfully
    Dim SavedTitleBarVisible        As Boolean   'Pre-test title-bar visibility

    Dim OldScreenUpdating           As Boolean   'Cached ScreenUpdating state
    Dim HasFailure                  As Boolean   'TRUE when a test failure occurred
    Dim FailNumber                  As Long      'Captured failure number
    Dim FailSource                  As String    'Captured failure source
    Dim FailDescription             As String    'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Cache and suppress screen updates during the regression run
        OldScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False

    'Log the start of the requested regression pack
        TST_Log CallerProc, "START", "Title-bar-only regression pack started"

'------------------------------------------------------------------------------
' SNAPSHOT CURRENT STATE
'------------------------------------------------------------------------------
    'Snapshot the current Excel UI state before the test mutates it
        TST_SnapshotState _
            RibbonKnown:=SavedRibbonKnown, _
            RibbonVisible:=SavedRibbonVisible, _
            StatusBarVisible:=SavedStatusBarVisible, _
            ScrollBarsVisible:=SavedScrollBarsVisible, _
            FormulaBarVisible:=SavedFormulaBarVisible, _
            WindowCount:=SavedWindowCount, _
            HeadingsVisible:=SavedHeadingsVisible, _
            WorkbookTabsVisible:=SavedWorkbookTabsVisible, _
            GridlinesVisible:=SavedGridlinesVisible, _
            TitleBarKnown:=SavedTitleBarKnown, _
            TitleBarVisible:=SavedTitleBarVisible

'------------------------------------------------------------------------------
' RUN REGRESSION CASE
'------------------------------------------------------------------------------
    'Run the dedicated title-bar round-trip case
        TST_Case_TitleBarRoundTrip

    'Log successful completion before restoration
        TST_Log CallerProc, "PASS", "Title-bar round-trip case passed"

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Attempt to restore the original pre-test UI state
        On Error Resume Next
            TST_RestoreState _
                RibbonKnown:=SavedRibbonKnown, _
                RibbonVisible:=SavedRibbonVisible, _
                StatusBarVisible:=SavedStatusBarVisible, _
                ScrollBarsVisible:=SavedScrollBarsVisible, _
                FormulaBarVisible:=SavedFormulaBarVisible, _
                WindowCount:=SavedWindowCount, _
                HeadingsVisible:=SavedHeadingsVisible, _
                WorkbookTabsVisible:=SavedWorkbookTabsVisible, _
                GridlinesVisible:=SavedGridlinesVisible, _
                TitleBarKnown:=SavedTitleBarKnown, _
                TitleBarVisible:=SavedTitleBarVisible
        On Error GoTo 0

    'Restore ScreenUpdating before leaving the harness
        Application.ScreenUpdating = OldScreenUpdating

    'Raise the captured failure after restoration when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Capture failure information so it can be re-raised after restoration
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description & _
                          IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

    'Log the failure immediately
        TST_Log CallerProc, "FAIL", TST_BuildRuntimeErrorText

    'Proceed to restoration / re-raise path
        Resume SafeExit

End Sub


'
'------------------------------------------------------------------------------
'
'                           REGRESSION CASES
'
'------------------------------------------------------------------------------
'

Private Sub TST_Case_ShowAllBaseline(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
'                        TST_Case_ShowAllBaseline
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
'   None
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
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
    'Log successful completion of the case
        TST_Log "TST_Case_ShowAllBaseline", "PASS", "All requested elements are visible"

End Sub

Private Sub TST_Case_SelectiveHide(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
'                          TST_Case_SelectiveHide
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
'   None
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
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
    'Log successful completion of the case
        TST_Log "TST_Case_SelectiveHide", "PASS", "Selective hide behaved as expected"

End Sub

Private Sub TST_Case_SelectiveShow(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
'                          TST_Case_SelectiveShow
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
'   None
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
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
    'Log successful completion of the case
        TST_Log "TST_Case_SelectiveShow", "PASS", "Selective show behaved as expected"

End Sub

Private Sub TST_Case_NoOpLeaveUnchanged(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
'                        TST_Case_NoOpLeaveUnchanged
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
'   None
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
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
    'Log successful completion of the case
        TST_Log "TST_Case_NoOpLeaveUnchanged", "PASS", "No-op behavior behaved as expected"

End Sub

Private Sub TST_Case_ConvenienceWrappers(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
'                       TST_Case_ConvenienceWrappers
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
'   None
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
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
    'Log successful completion of the case
        TST_Log "TST_Case_ConvenienceWrappers", "PASS", "Convenience wrappers behaved as expected"

End Sub

Private Sub TST_Case_WithResult_AllSuccess(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
'                       TST_Case_WithResult_AllSuccess
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
'   None
'
' ERROR POLICY
'   - Raises on assertion failure
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean   'Boolean success flag returned by the API
    Dim FailureCount        As Long      'Number of recorded failures
    Dim FailureList         As Variant   'Optional array of recorded failures

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
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
    'Log successful completion of the case
        TST_Log "TST_Case_WithResult_AllSuccess", "PASS", "Structured-result success path behaved as expected"

End Sub

Private Sub TST_Case_WithResult_NoOpSuccess(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
'                       TST_Case_WithResult_NoOpSuccess
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
'   None
'
' ERROR POLICY
'   - Raises on assertion failure
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean   'Boolean success flag returned by the API
    Dim FailureCount        As Long      'Number of recorded failures
    Dim FailureList         As Variant   'Optional array of recorded failures

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
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
    'Log successful completion of the case
        TST_Log "TST_Case_WithResult_NoOpSuccess", "PASS", "Structured-result no-op path behaved as expected"

End Sub

Private Sub TST_Case_WithResult_SuccessWithoutFailureList(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
'              TST_Case_WithResult_SuccessWithoutFailureList
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
'   None
'
' ERROR POLICY
'   - Raises on assertion failure
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean   'Boolean success flag returned by the API
    Dim FailureCount        As Long      'Number of recorded failures
    Dim FailureList         As Variant   'Local untouched Variant proving omission path

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
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
                      "WithResult.NoFailureList.Result expected=True actual=False"
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
    'Log successful completion of the case
        TST_Log "TST_Case_WithResult_SuccessWithoutFailureList", "PASS", _
            "Structured-result success path without FailureList behaved as expected"

End Sub

Private Sub TST_Case_WithResult_InvalidVisibility()

'
'==============================================================================
'                   TST_Case_WithResult_InvalidVisibility
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
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on assertion failure
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean   'Boolean success flag returned by the API
    Dim FailureCount        As Long      'Number of recorded failures
    Dim FailureList         As Variant   'Recorded structured failures

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
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
                      "WithResult.InvalidVisibility expected=False actual=True"
        End If

    'Assert that one or more failures were recorded
        If FailureCount < 1 Then
            Err.Raise TEST_ERR_BASE + 41, _
                      "WithResult.InvalidVisibility", _
                      "WithResult.InvalidVisibility expected FailureCount>=1 actual=" & CStr(FailureCount)
        End If

    'Assert that FailureList was populated
        If IsEmpty(FailureList) Then
            Err.Raise TEST_ERR_BASE + 42, _
                      "WithResult.InvalidVisibility", _
                      "WithResult.InvalidVisibility expected FailureList to be populated"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
    'Log successful completion of the case
        TST_Log "TST_Case_WithResult_InvalidVisibility", "PASS", _
            "Structured failure reporting for invalid UIVisibility behaved as expected"

End Sub

Private Sub TST_Case_SnapshotLifecycle(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
'                      TST_Case_SnapshotLifecycle
'------------------------------------------------------------------------------
' PURPOSE
'   Verify the explicit snapshot and reset lifecycle exposed by the core module
'
' WHY THIS EXISTS
'   The core module separates UI_ShowExcelUI from explicit snapshot and reset
'   semantics, so the explicit lifecycle deserves direct regression coverage
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => include TitleBar in the capture and reset assertions
'     FALSE => skip TitleBar assertions in this case
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on assertion failure
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
        TST_Log "TST_Case_SnapshotLifecycle", "START", "Validating explicit snapshot and reset lifecycle"

'------------------------------------------------------------------------------
' CLEAR ANY PRIOR SNAPSHOT
'------------------------------------------------------------------------------
    'Clear any prior explicit snapshot before starting the lifecycle test
        UI_ClearExcelUIStateSnapshot

    'Assert that no snapshot is now available
        TST_AssertSnapshotAvailability False, "SnapshotLifecycle.InitialClear"

'------------------------------------------------------------------------------
' ESTABLISH MIXED BASELINE
'------------------------------------------------------------------------------
    'Establish a mixed baseline that will be captured explicitly
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
' CAPTURE SNAPSHOT
'------------------------------------------------------------------------------
    'Capture the current mixed baseline explicitly
        UI_CaptureExcelUIState

    'Assert that a snapshot is now available
        TST_AssertSnapshotAvailability True, "SnapshotLifecycle.AfterCapture"

'------------------------------------------------------------------------------
' MUTATE AWAY FROM CAPTURED BASELINE
'------------------------------------------------------------------------------
    'Drive the managed UI to a materially different state so the reset path has
    'something meaningful to restore
        UI_SetExcelUI _
            Ribbon:=UI_Hide, _
            StatusBar:=UI_Show, _
            ScrollBars:=UI_Hide, _
            FormulaBar:=UI_Show, _
            Headings:=UI_Hide, _
            WorkbookTabs:=UI_Show, _
            Gridlines:=UI_Hide, _
            TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Hide)

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' RESET TO CAPTURED SNAPSHOT
'------------------------------------------------------------------------------
    'Restore the explicitly captured baseline
        UI_ResetExcelUIToSnapshot

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT RESET RESULT
'------------------------------------------------------------------------------
    'Assert Ribbon restored to the captured baseline
        TST_AssertRibbonVisible True, "SnapshotLifecycle.Ribbon"

    'Assert StatusBar restored to the captured baseline
        TST_AssertApplicationProperty False, "DisplayStatusBar", "SnapshotLifecycle.StatusBar"

    'Assert ScrollBars restored to the captured baseline
        TST_AssertApplicationProperty True, "DisplayScrollBars", "SnapshotLifecycle.ScrollBars"

    'Assert FormulaBar restored to the captured baseline
        TST_AssertApplicationProperty False, "DisplayFormulaBar", "SnapshotLifecycle.FormulaBar"

    'Assert Headings restored to the captured baseline across all windows
        TST_AssertAllWindowsProperty True, "DisplayHeadings", "SnapshotLifecycle.Headings"

    'Assert WorkbookTabs restored to the captured baseline across all windows
        TST_AssertAllWindowsProperty False, "DisplayWorkbookTabs", "SnapshotLifecycle.WorkbookTabs"

    'Assert Gridlines restored to the captured baseline across all windows
        TST_AssertAllWindowsProperty True, "DisplayGridlines", "SnapshotLifecycle.Gridlines"

    'Assert TitleBar restored to the captured baseline when title-bar testing is
    'enabled
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "SnapshotLifecycle.TitleBar"
        End If

    'Assert the snapshot still remains available after reset
        TST_AssertSnapshotAvailability True, "SnapshotLifecycle.AfterReset"

'------------------------------------------------------------------------------
' CLEAR SNAPSHOT AGAIN
'------------------------------------------------------------------------------
    'Clear the explicit snapshot and assert it is gone
        UI_ClearExcelUIStateSnapshot
        TST_AssertSnapshotAvailability False, "SnapshotLifecycle.FinalClear"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
    'Log successful completion of the case
        TST_Log "TST_Case_SnapshotLifecycle", "PASS", "Explicit snapshot and reset lifecycle behaved as expected"

End Sub

Private Sub TST_Case_ResetWithoutSnapshot_NoOp(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
'                 TST_Case_ResetWithoutSnapshot_NoOp
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that UI_ResetExcelUIToSnapshot behaves as a no-op when no explicit
'   snapshot is currently available
'
' WHY THIS EXISTS
'   The reset API is intentionally explicit and should not fabricate a baseline
'   when none was captured
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => include TitleBar in the unchanged-state assertion
'     FALSE => skip TitleBar assertions in this case
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on assertion failure
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
        TST_Log "TST_Case_ResetWithoutSnapshot_NoOp", "START", _
            "Validating reset-to-snapshot no-op behavior when no snapshot exists"

'------------------------------------------------------------------------------
' CLEAR ANY PRIOR SNAPSHOT
'------------------------------------------------------------------------------
    'Clear any prior explicit snapshot before starting the no-snapshot test
        UI_ClearExcelUIStateSnapshot

    'Assert that no snapshot is available
        TST_AssertSnapshotAvailability False, "ResetWithoutSnapshot.NoSnapshot"

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
' APPLY RESET WITHOUT SNAPSHOT
'------------------------------------------------------------------------------
    'Invoke reset without any explicit snapshot being available
        UI_ResetExcelUIToSnapshot

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT UNCHANGED STATE
'------------------------------------------------------------------------------
    'Assert Ribbon remained visible
        TST_AssertRibbonVisible True, "ResetWithoutSnapshot.Ribbon"

    'Assert StatusBar remained hidden
        TST_AssertApplicationProperty False, "DisplayStatusBar", "ResetWithoutSnapshot.StatusBar"

    'Assert ScrollBars remained visible
        TST_AssertApplicationProperty True, "DisplayScrollBars", "ResetWithoutSnapshot.ScrollBars"

    'Assert FormulaBar remained hidden
        TST_AssertApplicationProperty False, "DisplayFormulaBar", "ResetWithoutSnapshot.FormulaBar"

    'Assert Headings remained visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayHeadings", "ResetWithoutSnapshot.Headings"

    'Assert WorkbookTabs remained hidden across all windows
        TST_AssertAllWindowsProperty False, "DisplayWorkbookTabs", "ResetWithoutSnapshot.WorkbookTabs"

    'Assert Gridlines remained visible across all windows
        TST_AssertAllWindowsProperty True, "DisplayGridlines", "ResetWithoutSnapshot.Gridlines"

    'Assert TitleBar remained visible when title-bar testing is enabled
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "ResetWithoutSnapshot.TitleBar"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
    'Log successful completion of the case
        TST_Log "TST_Case_ResetWithoutSnapshot_NoOp", "PASS", _
            "Reset-to-snapshot no-op behavior without snapshot behaved as expected"

End Sub

Private Sub TST_Case_ScreenUpdatingPreserved()

'
'==============================================================================
'                    TST_Case_ScreenUpdatingPreserved
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
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on assertion failure
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedScreenUpdating As Boolean   'Caller-visible ScreenUpdating baseline

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
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
    'Log successful completion of the case
        TST_Log "TST_Case_ScreenUpdatingPreserved", "PASS", _
            "ScreenUpdating preservation behaved as expected"

End Sub

Private Sub TST_Case_TitleBarRoundTrip()

'
'==============================================================================
'                        TST_Case_TitleBarRoundTrip
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that the title bar can be hidden and then shown again through the
'   public API
'
' WHY THIS EXISTS
'   Title-bar control is the most WinAPI-sensitive part of the module and
'   benefits from a dedicated regression case
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Log the start of the case
        TST_Log "TST_Case_TitleBarRoundTrip", "START", "Validating title-bar hide and show round-trip"

'------------------------------------------------------------------------------
' APPLY TITLE-BAR HIDE
'------------------------------------------------------------------------------
    'Hide only the title bar
        UI_SetExcelUI TitleBar:=UI_Hide

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

    'Assert TitleBar hidden
        TST_AssertTitleBarVisible False, "TitleBarRoundTrip.Hide"

'------------------------------------------------------------------------------
' APPLY TITLE-BAR SHOW
'------------------------------------------------------------------------------
    'Show only the title bar
        UI_SetExcelUI TitleBar:=UI_Show

    'Allow the UI a short time to settle
        TST_WaitUI TEST_WAIT_SECONDS

    'Assert TitleBar visible
        TST_AssertTitleBarVisible True, "TitleBarRoundTrip.Show"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
    'Log successful completion of the case
        TST_Log "TST_Case_TitleBarRoundTrip", "PASS", "Title-bar round-trip behaved as expected"

End Sub


'
'------------------------------------------------------------------------------
'
'                      SNAPSHOT / RESTORE HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub TST_SnapshotState( _
    ByRef RibbonKnown As Boolean, _
    ByRef RibbonVisible As Boolean, _
    ByRef StatusBarVisible As Boolean, _
    ByRef ScrollBarsVisible As Boolean, _
    ByRef FormulaBarVisible As Boolean, _
    ByRef WindowCount As Long, _
    ByRef HeadingsVisible() As Boolean, _
    ByRef WorkbookTabsVisible() As Boolean, _
    ByRef GridlinesVisible() As Boolean, _
    ByRef TitleBarKnown As Boolean, _
    ByRef TitleBarVisible As Boolean)

'
'==============================================================================
'                           TST_SnapshotState
'------------------------------------------------------------------------------
' PURPOSE
'   Capture the current Excel UI state before the regression harness mutates it
'
' WHY THIS EXISTS
'   Regression tests should attempt to return the user's environment to its
'   prior state after execution
'
' INPUTS / OUTPUTS
'   [ByRef arguments]
'     Receive the captured application-level, window-level, and title-bar state
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Best-effort capture; unknown Ribbon or TitleBar state is marked via flags
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long      'Current window index during snapshot
    Dim Msg                 As String    'Diagnostic message from reader helpers

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Capture application-level state directly from Excel
        StatusBarVisible = Application.DisplayStatusBar
        ScrollBarsVisible = Application.DisplayScrollBars
        FormulaBarVisible = Application.DisplayFormulaBar

    'Capture Ribbon state through the best-effort reader
        RibbonKnown = TST_TryGetRibbonVisible(RibbonVisible, Msg)
        If Not RibbonKnown Then
            TST_Log "TST_SnapshotState", "Ribbon", Msg
        End If

    'Capture TitleBar state through the best-effort reader
        TitleBarKnown = TST_TryGetTitleBarVisible(TitleBarVisible, Msg)
        If Not TitleBarKnown Then
            TST_Log "TST_SnapshotState", "TitleBar", Msg
        End If

'------------------------------------------------------------------------------
' SNAPSHOT WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Capture the current Application.Windows count
        WindowCount = Application.Windows.Count

    'Allocate per-window snapshot arrays when at least one window exists
        If WindowCount > 0 Then

            'Size the Headings state array
                ReDim HeadingsVisible(1 To WindowCount)

            'Size the WorkbookTabs state array
                ReDim WorkbookTabsVisible(1 To WindowCount)

            'Size the Gridlines state array
                ReDim GridlinesVisible(1 To WindowCount)

            'Capture each window's relevant state
                For i = 1 To WindowCount

                    'Capture the current window's Headings visibility
                        HeadingsVisible(i) = Application.Windows(i).DisplayHeadings

                    'Capture the current window's WorkbookTabs visibility
                        WorkbookTabsVisible(i) = Application.Windows(i).DisplayWorkbookTabs

                    'Capture the current window's Gridlines visibility
                        GridlinesVisible(i) = Application.Windows(i).DisplayGridlines

                Next i

        End If

End Sub

Private Sub TST_RestoreState( _
    ByVal RibbonKnown As Boolean, _
    ByVal RibbonVisible As Boolean, _
    ByVal StatusBarVisible As Boolean, _
    ByVal ScrollBarsVisible As Boolean, _
    ByVal FormulaBarVisible As Boolean, _
    ByVal WindowCount As Long, _
    ByRef HeadingsVisible() As Boolean, _
    ByRef WorkbookTabsVisible() As Boolean, _
    ByRef GridlinesVisible() As Boolean, _
    ByVal TitleBarKnown As Boolean, _
    ByVal TitleBarVisible As Boolean)

'
'==============================================================================
'                            TST_RestoreState
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to restore the pre-test Excel UI state after the regression run
'
' WHY THIS EXISTS
'   Regression tests should clean up after themselves as much as possible
'
' INPUTS
'   [Captured snapshot values]
'     Pre-test UI state captured by TST_SnapshotState
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Best-effort restore only
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long      'Current window index during restore
    Dim WindowLimit         As Long      'Minimum of saved and current window counts
    Dim Msg                 As String    'Diagnostic message from helper routines

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress local restore failures so all restore steps are attempted
        On Error Resume Next

'------------------------------------------------------------------------------
' RESTORE TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Restore TitleBar first when its original state was captured successfully
        If TitleBarKnown Then

            'Restore TitleBar via the public API using explicit enum values
                If TitleBarVisible Then
                    UI_SetExcelUI TitleBar:=UI_Show
                Else
                    UI_SetExcelUI TitleBar:=UI_Hide
                End If

            'Allow the UI a short time to settle
                TST_WaitUI TEST_WAIT_SECONDS

        End If

'------------------------------------------------------------------------------
' RESTORE RIBBON STATE
'------------------------------------------------------------------------------
    'Restore Ribbon when its original state was captured successfully
        If RibbonKnown Then

            'Attempt Ribbon restore through the test helper
                If Not TST_TrySetRibbonVisible(RibbonVisible, Msg) Then
                    TST_Log "TST_RestoreState", "Ribbon", Msg
                End If

        End If

'------------------------------------------------------------------------------
' RESTORE APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Restore StatusBar visibility directly
        Application.DisplayStatusBar = StatusBarVisible

    'Restore ScrollBars visibility directly
        Application.DisplayScrollBars = ScrollBarsVisible

    'Restore FormulaBar visibility directly
        Application.DisplayFormulaBar = FormulaBarVisible

'------------------------------------------------------------------------------
' RESTORE WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Compute the number of windows that can be restored safely by index
        WindowLimit = Application.Windows.Count
        If WindowCount < WindowLimit Then WindowLimit = WindowCount

    'Restore each saved window state up to the common window count
        For i = 1 To WindowLimit

            'Restore the current window's Headings visibility
                TST_TryRestoreWindowProp Application.Windows(i), "DisplayHeadings", HeadingsVisible(i)

            'Restore the current window's WorkbookTabs visibility
                TST_TryRestoreWindowProp Application.Windows(i), "DisplayWorkbookTabs", WorkbookTabsVisible(i)

            'Restore the current window's Gridlines visibility
                TST_TryRestoreWindowProp Application.Windows(i), "DisplayGridlines", GridlinesVisible(i)

        Next i

'------------------------------------------------------------------------------
' SETTLE UI
'------------------------------------------------------------------------------
    'Allow the UI a short time to settle after restoration
        TST_WaitUI TEST_WAIT_SECONDS

End Sub

Private Sub TST_WaitUI(ByVal SecondsToWait As Double)

'
'==============================================================================
'                               TST_WaitUI
'------------------------------------------------------------------------------
' PURPOSE
'   Give Excel and Windows a short opportunity to settle after a UI state change
'
' WHY THIS EXISTS
'   Some UI changes, especially Ribbon and TitleBar changes, can be slightly
'   asynchronous from the perspective of immediate assertions
'
' INPUTS
'   SecondsToWait
'     Requested wait duration in seconds
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise
'
' NOTES
'   - Handles Timer rollover at midnight
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim t0                  As Double    'Timer baseline

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Exit immediately when no positive wait duration was requested
        If SecondsToWait <= 0# Then
            Exit Sub
        End If

    'Capture the timer baseline
        t0 = Timer

'------------------------------------------------------------------------------
' WAIT LOOP
'------------------------------------------------------------------------------
    'Yield to Excel until the requested duration has elapsed, handling midnight
    'rollover safely
        Do While TST_TimerElapsedSeconds(t0) < SecondsToWait
            DoEvents
        Loop

End Sub


'
'------------------------------------------------------------------------------
'
'                           ASSERTION HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub TST_AssertBooleanEquals( _
    ByVal Expected As Boolean, _
    ByVal Actual As Boolean, _
    ByVal AssertionName As String)

'
'==============================================================================
'                         TST_AssertBooleanEquals
'------------------------------------------------------------------------------
' PURPOSE
'   Raise a descriptive assertion failure when two Boolean values differ
'
' WHY THIS EXISTS
'   Regression tests need explicit, readable failures instead of silent
'   mismatches
'
' INPUTS
'   Expected
'     Expected Boolean state
'
'   Actual
'     Actual Boolean state
'
'   AssertionName
'     Human-readable assertion identifier
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on mismatch
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Raise an assertion failure when the Boolean values differ
        If Expected <> Actual Then
            Err.Raise TEST_ERR_BASE + 1, _
                      AssertionName, _
                      AssertionName & " expected=" & CStr(Expected) & " actual=" & CStr(Actual)
        End If

End Sub

Private Sub TST_AssertApplicationProperty( _
    ByVal Expected As Boolean, _
    ByVal PropertyName As String, _
    ByVal AssertionName As String)

'
'==============================================================================
'                     TST_AssertApplicationProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Assert the current Boolean value of an Application-level property
'
' WHY THIS EXISTS
'   The public UI API controls several Application-level Boolean properties
'   that need regression assertions
'
' INPUTS
'   Expected
'     Expected property value
'
'   PropertyName
'     Application property name to read
'
'   AssertionName
'     Human-readable assertion identifier
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on read failure or mismatch
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Actual              As Boolean   'Actual property value
    Dim Msg                 As String    'Diagnostic message from the reader helper

'------------------------------------------------------------------------------
' READ PROPERTY
'------------------------------------------------------------------------------
    'Attempt to read the requested Application property
        If Not TST_TryGetBooleanProperty(Application, PropertyName, Actual, Msg) Then
            Err.Raise TEST_ERR_BASE + 2, AssertionName, AssertionName & " read failed | " & Msg
        End If

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Assert the read value against the expectation
        TST_AssertBooleanEquals Expected, Actual, AssertionName

End Sub

Private Sub TST_AssertAllWindowsProperty( _
    ByVal Expected As Boolean, _
    ByVal PropertyName As String, _
    ByVal AssertionName As String)

'
'==============================================================================
'                       TST_AssertAllWindowsProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Assert the current Boolean value of a Window-level property across all open
'   Excel windows
'
' WHY THIS EXISTS
'   The public UI API applies several properties to each open Excel window, not
'   just the active one
'
' INPUTS
'   Expected
'     Expected property value
'
'   PropertyName
'     Window property name to read
'
'   AssertionName
'     Human-readable assertion identifier
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on read failure or mismatch
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim W                   As Window    'Current Excel window during assertion
    Dim Actual              As Boolean   'Actual property value
    Dim Msg                 As String    'Diagnostic message from the reader helper

'------------------------------------------------------------------------------
' ASSERT EACH WINDOW
'------------------------------------------------------------------------------
    'Assert the requested property on every open Excel window
        For Each W In Application.Windows

            'Attempt to read the requested Window property
                If Not TST_TryGetBooleanProperty(W, PropertyName, Actual, Msg) Then
                    Err.Raise TEST_ERR_BASE + 3, _
                              AssertionName, _
                              AssertionName & " read failed on window [" & W.Caption & "] | " & Msg
                End If

            'Assert the read value against the expectation
                TST_AssertBooleanEquals Expected, Actual, AssertionName & " [" & W.Caption & "]"

        Next W

End Sub

Private Sub TST_AssertRibbonVisible( _
    ByVal Expected As Boolean, _
    ByVal AssertionName As String)

'
'==============================================================================
'                           TST_AssertRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Assert the current Ribbon visibility
'
' WHY THIS EXISTS
'   Ribbon state is not best treated as a plain direct property read, so it has
'   a dedicated assertion helper
'
' INPUTS
'   Expected
'     Expected Ribbon visibility
'
'   AssertionName
'     Human-readable assertion identifier
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on read failure or mismatch
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Actual              As Boolean   'Actual Ribbon visibility
    Dim Msg                 As String    'Diagnostic message from the reader helper

'------------------------------------------------------------------------------
' READ RIBBON STATE
'------------------------------------------------------------------------------
    'Attempt to read the current Ribbon visibility
        If Not TST_TryGetRibbonVisible(Actual, Msg) Then
            Err.Raise TEST_ERR_BASE + 4, AssertionName, AssertionName & " read failed | " & Msg
        End If

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Assert the read value against the expectation
        TST_AssertBooleanEquals Expected, Actual, AssertionName

End Sub

Private Sub TST_AssertTitleBarVisible( _
    ByVal Expected As Boolean, _
    ByVal AssertionName As String)

'
'==============================================================================
'                         TST_AssertTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Assert the current title-bar visibility for the Excel window represented by
'   Application.Hwnd
'
' WHY THIS EXISTS
'   Title-bar state is WinAPI-based and benefits from a dedicated assertion
'   helper
'
' INPUTS
'   Expected
'     Expected title-bar visibility
'
'   AssertionName
'     Human-readable assertion identifier
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on read failure or mismatch
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Actual              As Boolean   'Actual title-bar visibility
    Dim Msg                 As String    'Diagnostic message from the reader helper

'------------------------------------------------------------------------------
' READ TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Attempt to read the current title-bar visibility
        If Not TST_TryGetTitleBarVisible(Actual, Msg) Then
            Err.Raise TEST_ERR_BASE + 5, AssertionName, AssertionName & " read failed | " & Msg
        End If

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Assert the read value against the expectation
        TST_AssertBooleanEquals Expected, Actual, AssertionName

End Sub

Private Sub TST_AssertResultSuccess( _
    ByVal Succeeded As Boolean, _
    ByVal FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal AssertionName As String)

'
'==============================================================================
'                         TST_AssertResultSuccess
'------------------------------------------------------------------------------
' PURPOSE
'   Assert that the standard-module result buffers represent a clean success
'   path
'
' WHY THIS EXISTS
'   Structured-result regressions need a shared assertion helper so the tests
'   validate the same core contract consistently
'
' INPUTS
'   Succeeded
'     Boolean success flag returned by UI_SetExcelUI_WithResult
'
'   FailureCount
'     Failure-count output returned by UI_SetExcelUI_WithResult
'
'   FailureList
'     Failure-list output returned by UI_SetExcelUI_WithResult
'
'   AssertionName
'     Human-readable assertion identifier
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on mismatch
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' ASSERT SUCCESS FLAG
'------------------------------------------------------------------------------
    'Assert that the Boolean result reports overall success
        If Not Succeeded Then
            Err.Raise TEST_ERR_BASE + 20, _
                      AssertionName, _
                      AssertionName & " expected=True actual=False"
        End If

'------------------------------------------------------------------------------
' ASSERT FAILURE COUNT
'------------------------------------------------------------------------------
    'Assert that the result buffers recorded no failures
        If FailureCount <> 0 Then
            Err.Raise TEST_ERR_BASE + 21, _
                      AssertionName, _
                      AssertionName & " expected FailureCount=0 actual=" & CStr(FailureCount)
        End If

'------------------------------------------------------------------------------
' ASSERT FAILURE LIST STATE
'------------------------------------------------------------------------------
    'Assert that the captured failure list remains Empty for a clean success
    'path
        If Not IsEmpty(FailureList) Then
            Err.Raise TEST_ERR_BASE + 22, _
                      AssertionName, _
                      AssertionName & " expected FailureList=Empty for clean success path"
        End If

End Sub

Private Sub TST_AssertSnapshotAvailability( _
    ByVal Expected As Boolean, _
    ByVal AssertionName As String)

'
'==============================================================================
'                      TST_AssertSnapshotAvailability
'------------------------------------------------------------------------------
' PURPOSE
'   Assert the availability flag returned by UI_HasExcelUIStateSnapshot
'
' WHY THIS EXISTS
'   The explicit snapshot and reset lifecycle introduced by the core module
'   needs direct regression coverage on the public snapshot-availability
'   contract
'
' INPUTS
'   Expected
'     Expected snapshot-availability state
'
'   AssertionName
'     Human-readable assertion identifier
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on mismatch
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Assert the snapshot-availability flag against the expectation
        TST_AssertBooleanEquals Expected, UI_HasExcelUIStateSnapshot, AssertionName

End Sub


'
'------------------------------------------------------------------------------
'
'                         STATE READ/WRITE HELPERS
'
'------------------------------------------------------------------------------
'

Private Function TST_TryGetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByRef ValueOut As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                       TST_TryGetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read a Boolean property from an object using CallByName
'
' WHY THIS EXISTS
'   Application-level and Window-level assertions need a shared property reader
'   to avoid duplicated boilerplate
'
' INPUTS
'   Target
'     Object exposing the target Boolean property
'
'   PropertyName
'     Name of the Boolean property to read
'
'   ValueOut
'     Receives the property value on success
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => property read succeeded
'   FALSE => property read failed
'
' ERROR POLICY
'   - Does NOT raise to callers
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant   'Late-bound property value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize outputs and default result
        TST_TryGetBooleanProperty = False
        ValueOut = False
        FailMsg = vbNullString

    'Reject invalid object input deterministically
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo SafeExit
        End If

    'Reject empty property name deterministically
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ PROPERTY
'------------------------------------------------------------------------------
    'Read the requested property using late-bound property access
        V = CallByName(Target, PropertyName, VbGet)

    'Convert the result to a Boolean
        ValueOut = CBool(V)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after property access completes
        TST_TryGetBooleanProperty = True

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
    'Return a descriptive failure string without raising
        FailMsg = TST_BuildRuntimeErrorText

End Function

Private Function TST_TrySetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByVal NewValue As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                       TST_TrySetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to assign a Boolean property on an object using a common,
'   best-effort helper
'
' WHY THIS EXISTS
'   The regression harness needs the same kind of shared Boolean property-write
'   logic used in the main UI module, especially during restore
'
' INPUTS
'   Target
'     Object exposing the target Boolean property
'
'   PropertyName
'     Name of the Boolean property to assign
'
'   NewValue
'     Boolean value to write to the target property
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => property write succeeded
'   FALSE => property write failed
'
' BEHAVIOR
'   - Uses CallByName with VbLet to assign the property
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' NOTES
'   - Intended for Application and Window Boolean property writes in this
'     module
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize default failure result
        TST_TrySetBooleanProperty = False
        FailMsg = vbNullString

    'Reject invalid object input deterministically
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo SafeExit
        End If

    'Reject empty property name deterministically
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' APPLY PROPERTY WRITE
'------------------------------------------------------------------------------
    'Assign the requested Boolean value using late-bound property assignment
        CallByName Target, PropertyName, VbLet, NewValue

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after property assignment completes
        TST_TrySetBooleanProperty = True

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
    'Return a descriptive failure string without raising
        FailMsg = TST_BuildRuntimeErrorText

End Function

Private Function TST_TryGetRibbonVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                         TST_TryGetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read current Ribbon visibility
'
' WHY THIS EXISTS
'   The Ribbon is not best treated as a simple direct property read, so the
'   regression harness uses a dedicated best-effort reader
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
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant   'Fallback Excel4 macro result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize outputs and default result
        TST_TryGetRibbonVisible = False
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
            TST_TryGetRibbonVisible = True
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
            TST_TryGetRibbonVisible = True
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
        FailMsg = TST_BuildRuntimeErrorText

End Function

Private Function TST_TrySetRibbonVisible( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                         TST_TrySetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to set Ribbon visibility from the regression harness
'
' WHY THIS EXISTS
'   State restoration needs a local Ribbon setter because Ribbon control is not
'   exposed through a simple Application Boolean property
'
' INPUTS
'   IsVisible
'     Requested Ribbon visibility
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => Ribbon update succeeded
'   FALSE => Ribbon update failed
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MacroText           As String    'Excel4 macro text for Ribbon visibility

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize default failure result
        TST_TrySetRibbonVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' BUILD MACRO
'------------------------------------------------------------------------------
    'Build the Ribbon visibility macro text explicitly
        If IsVisible Then
            MacroText = "Show.TOOLBAR(""Ribbon"",True)"
        Else
            MacroText = "Show.TOOLBAR(""Ribbon"",False)"
        End If

'------------------------------------------------------------------------------
' EXECUTE MACRO
'------------------------------------------------------------------------------
    'Execute the Ribbon visibility macro
        Application.ExecuteExcel4Macro MacroText

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after macro execution completes
        TST_TrySetRibbonVisible = True

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
    'Return a descriptive failure string without raising
        FailMsg = TST_BuildRuntimeErrorText

End Function

Private Function TST_TryGetTitleBarVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                      TST_TryGetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read current title-bar visibility for the Excel window
'   represented by Application.Hwnd
'
' WHY THIS EXISTS
'   Title-bar state is controlled through WinAPI in EXCEL_UI, so the regression
'   harness uses a corresponding WinAPI-based read helper
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
        TST_TryGetTitleBarVisible = False
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
        TST_SetLastError 0

#If VBA7 Then
    #If Win64 Then

        'Read the current window style using the 64-bit API
            StyleValue = TST_GetWindowLongPtr(xlHnd, TST_GWL_STYLE)

    #Else

        'Read the current window style using the 32-bit API under VBA7
            StyleValue = TST_GetWindowLong(xlHnd, TST_GWL_STYLE)

    #End If
#Else

    'Read the current window style using the legacy 32-bit API
        StyleValue = TST_GetWindowLong(xlHnd, TST_GWL_STYLE)

#End If

    'Read the Win32 last-error value immediately after the API call
        LastErr = TST_GetLastError

    'Treat zero plus nonzero last error as failure
        If StyleValue = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' MAP STYLE TO TITLE-BAR VISIBILITY
'------------------------------------------------------------------------------
    'Treat the caption style bit as the title-bar visibility signal
        IsVisible = ((StyleValue And TST_WS_CAPTION) <> 0)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after a valid style read
        TST_TryGetTitleBarVisible = True

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
        FailMsg = TST_BuildRuntimeErrorText

End Function

Private Sub TST_TryRestoreWindowProp( _
    ByVal W As Window, _
    ByVal PropName As String, _
    ByVal Value As Boolean)

'
'==============================================================================
'                         TST_TryRestoreWindowProp
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to restore a specific Boolean Window property during test cleanup
'
' WHY THIS EXISTS
'   The restore path must be best-effort and should log window-specific restore
'   failures without interrupting later cleanup steps
'
' INPUTS
'   W
'     Target Excel window
'
'   PropName
'     Window Boolean property name to restore
'
'   Value
'     Boolean value to assign
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Logs any failure to the Immediate Window
'
' DEPENDENCIES
'   - TST_TrySetBooleanProperty
'   - TST_Log
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Msg                 As String    'Diagnostic message from the property-write helper

'------------------------------------------------------------------------------
' APPLY PROPERTY RESTORE
'------------------------------------------------------------------------------
    'Attempt to restore the requested Window property and log any failure
        If Not TST_TrySetBooleanProperty(W, PropName, Value, Msg) Then
            TST_Log "TST_RestoreState", "Restore." & PropName & " [" & W.Caption & "]", Msg
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                         DIAGNOSTICS AND TIMING
'
'------------------------------------------------------------------------------
'

Private Sub TST_Log( _
    ByVal ProcName As String, _
    ByVal Stage As String, _
    ByVal Detail As String)

'
'==============================================================================
'                                TST_Log
'------------------------------------------------------------------------------
' PURPOSE
'   Write a consistent diagnostic line to the Immediate Window for the
'   regression harness
'
' WHY THIS EXISTS
'   The harness needs readable progress and failure logging
'
' INPUTS
'   ProcName
'     Procedure name associated with the log line
'
'   Stage
'     Logical stage associated with the log line
'
'   Detail
'     Message detail to append
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Suppresses any unexpected logging failure locally
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
        Debug.Print ProcName & " @ " & Stage & " | " & Detail

End Sub

Private Function TST_TimerElapsedSeconds(ByVal TimerStart As Double) As Double

'
'==============================================================================
'                         TST_TimerElapsedSeconds
'------------------------------------------------------------------------------
' PURPOSE
'   Return elapsed seconds since a Timer baseline, handling midnight rollover
'
' WHY THIS EXISTS
'   VBA Timer resets at midnight, so direct subtraction can become negative in
'   long-running sessions or when tests span midnight
'
' INPUTS
'   TimerStart
'     Baseline Timer value captured earlier
'
' RETURNS
'   Elapsed seconds since TimerStart, adjusted for midnight rollover when
'   necessary
'
' ERROR POLICY
'   - Does NOT raise
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TimerNow            As Double    'Current Timer reading

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Read the current Timer value
        TimerNow = Timer

'------------------------------------------------------------------------------
' RETURN ELAPSED SECONDS
'------------------------------------------------------------------------------
    'Adjust for midnight rollover when the current Timer value is less than the
    'baseline
        If TimerNow >= TimerStart Then
            TST_TimerElapsedSeconds = TimerNow - TimerStart
        Else
            TST_TimerElapsedSeconds = (TST_SECONDS_PER_DAY - TimerStart) + TimerNow
        End If

End Function

Private Function TST_TitleBarMode( _
    ByVal IncludeTitleBarTests As Boolean, _
    ByVal RequestedMode As UIVisibility) As UIVisibility

'
'==============================================================================
'                            TST_TitleBarMode
'------------------------------------------------------------------------------
' PURPOSE
'   Return the effective TitleBar mode for a test case based on whether the
'   current pack includes title-bar assertions
'
' WHY THIS EXISTS
'   Many test cases need the same small policy:
'     - when title-bar testing is enabled, apply the requested TitleBar mode
'     - otherwise leave TitleBar unchanged
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE  => use RequestedMode
'     FALSE => return UI_LeaveUnchanged
'
'   RequestedMode
'     TitleBar visibility mode to apply when title-bar testing is enabled
'
' RETURNS
'   Effective TitleBar mode for the test case
'
' ERROR POLICY
'   - Does NOT raise
'
' UPDATED
'   2026-04-19
'==============================================================================
'
'------------------------------------------------------------------------------
' RETURN EFFECTIVE MODE
'------------------------------------------------------------------------------
    'Use the requested mode only when title-bar testing is enabled
        If IncludeTitleBarTests Then
            TST_TitleBarMode = RequestedMode
        Else
            TST_TitleBarMode = UI_LeaveUnchanged
        End If

End Function

Private Function TST_BuildRuntimeErrorText() As String

'
'==============================================================================
'                       TST_BuildRuntimeErrorText
'------------------------------------------------------------------------------
' PURPOSE
'   Build a consistent runtime diagnostic string from the active Err object
'
' WHY THIS EXISTS
'   Several helpers in this module need identical fail-soft diagnostic text
'   Centralizing the formatting keeps the harness consistent and easier to
'   maintain
'
' RETURNS
'   A formatted diagnostic string including:
'     - Err.Number
'     - Err.Description
'     - Err.Source, when available
'     - Erl, when available
'
' ERROR POLICY
'   - Does NOT raise
'   - Returns best-effort text
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
' BUILD RUNTIME ERROR TEXT
'------------------------------------------------------------------------------
    'Build a consistent diagnostic string from the current Err state
        TST_BuildRuntimeErrorText = _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

End Function


