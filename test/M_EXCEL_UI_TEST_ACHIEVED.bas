Attribute VB_Name = "M_EXCEL_UI_TEST_ACHIEVED"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: M_EXCEL_UI_TEST_ACHIEVED
'==============================================================================
' PURPOSE
'   Verify, for every managed element, the state that a write through the
'   public API actually achieves: write, settle, read back and compare, and
'   record the outcome of each observation in the certification record
'
' WHY THIS EXISTS
'   The object-model and Ribbon setters are best effort. A write that raises no
'   runtime error, and a setter that returns True, do not show that the
'   requested state was reached. Only reading the state back, independently of
'   the setter, does. This unit is that check, and it is a mandatory unit of
'   release certification: it is a release-certification feature, not a new
'   guarantee of the fire-and-forget API
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   - Test_EXCEL_UI_RunAchievedStateVerification runs the eight cases of the
'     unit through TST_CertRunCase and restores the host afterwards.
'   - One observation is one write: it reads the state, writes it through
'     UI_SetExcelUI_WithResult, settles within a bound, reads it back, compares
'     it with what was requested and records one outcome: verified-no-op,
'     verified-write, write-failed, readback-failed, state-mismatch,
'     restore-failed or not-verifiable, with the reason.
'   - A scenario is four observations of one element under one scope (hide,
'     hide again, show, show again), so that the path that changes the state
'     and the path that finds it already correct are both verified, and then a
'     restoration of the entry state, itself read back.
'   - The readers and the entry restoration never use the API under test.
'
' ERROR POLICY
'   - A case raises, after restoration, when an observation failed, whatever the
'     setter returned. An observation that is not-verifiable is listed in the
'     evidence and does not fail the case.
'   - The recording procedures do not raise: an observation that cannot be
'     stored flags the record, which forces FAIL | INCOMPLETE.
'
' DEPENDENCIES
'   - M_EXCEL_UI_TEST_HARNESS
'   - M_EXCEL_UI_TEST_CERTIFICATION, for TST_CertRunCase
'   - M_EXCEL_UI: UI_SetExcelUI_WithResult
'   - M_EXCEL_UI_TITLEBAR: UI_TryGetActiveFramePair
'
' NOTES
'   None
'
' PUBLIC SURFACE
'   - Test_EXCEL_UI_RunAchievedStateVerification
'
'   The cases are Public so that the regression dispatcher in
'   M_EXCEL_UI_TEST_CERTIFICATION can run them
'
' TEST SCOPE
'   - Status Bar, Scroll Bars and Formula Bar, which are application level
'   - Headings, Workbook Tabs and Gridlines, for each UIWindowTargetScope: the
'     windows in the scope are read back at the requested state, and every other
'     window is read back unchanged
'   - the title bar, read back from the frame the active frame pair named
'     before the write, which is visible only with the complete caption mask; an
'     entry frame without it is not the baseline this unit verifies, and fails
'   - the Ribbon, read back only while the window that owned it before the write
'     is still the active window; a Ribbon that cannot be addressed or read is
'     not-verifiable, the only surface that is
'   - the fault seams: an ignored write, a failed write, a failed readback, a
'     failed restoration and a title bar with no caption on entry each give their own outcome and their own text, and a
'     state mismatch fails the verdict whatever the setter returned
'
' STATE MANAGEMENT
'   - Every scenario ends by restoring the entry state through the harness
'     writers and reading it back; a restoration that did not read back is an
'     observation of its own, restore-failed, and fails the case
'   - The runner snapshots the managed state before the cases and restores it
'     afterwards, as the regression pack does
'   - A one-shot fault applies to the next step or restoration it matches
'   - Waits are bounded and centralized: every settle is one poll loop of at most
'     TEST_ACH_POLL_LIMIT reads, TEST_ACH_POLL_SECONDS apart, and no case waits
'     on its own
'
' COMPATIBILITY
'   - Windows only for the title-bar surface
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.2
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const TEST_ACH_ERR_BASE     As Long = vbObjectError + 5400  'Base custom error number for the achieved-state unit
    Private Const TEST_ACH_ERR_UNCLEAN  As Long = 1                     'Offset: an observation failed
    Private Const TEST_ACH_ERR_SEAM     As Long = 2                     'Offset: a fault seam did not behave as designed
    Private Const TEST_ACH_ERR_SETUP    As Long = 3                     'Offset: a fixture could not be built

    Private Const TEST_ACH_POLL_SECONDS     As Double = 0.05    'Interval between two reads while a state settles
    Private Const TEST_ACH_POLL_LIMIT       As Long = 20        'Most reads of one settle; with the interval it is the bound
    Private Const TEST_ACH_SLACK_SECONDS    As Double = 2#      'Allowance, on top of the bound, for the reads and the write of one step
    Private Const TEST_ACH_SECONDS_PER_DAY  As Double = 86400#  'Timer rollover interval in seconds

    Private Const TEST_ACH_SURFACE_NONE     As Long = 0 'Surface: none selected
    Private Const TEST_ACH_SURFACE_APP      As Long = 1 'Surface: one Boolean property of the Application
    Private Const TEST_ACH_SURFACE_WINDOW   As Long = 2 'Surface: one Boolean property of each Window
    Private Const TEST_ACH_SURFACE_TITLEBAR As Long = 3 'Surface: the frame style of the active frame pair
    Private Const TEST_ACH_SURFACE_RIBBON   As Long = 4 'Surface: the Ribbon of the active window

    Private Const TEST_ACH_FAULT_NONE           As Long = 0 'Fault: none armed
    Private Const TEST_ACH_FAULT_IGNORED_WRITE  As Long = 1 'Fault: the setter is asked for nothing, so it returns True and writes nothing
    Private Const TEST_ACH_FAULT_WRITE_FAILED   As Long = 2 'Fault: the setter is given a value no member names, so it reports a failure
    Private Const TEST_ACH_FAULT_READBACK       As Long = 3 'Fault: the readback after the write fails
    Private Const TEST_ACH_FAULT_RESTORE        As Long = 4 'Fault: the restoration write is not made
    Private Const TEST_ACH_FAULT_NO_CAPTION     As Long = 5 'Fault: the entry frame of a title-bar scenario reads as having no complete caption
    Private Const TEST_ACH_FAULT_RESTORE_READ   As Long = 6 'Fault: the readback after the restoration fails

    'What the contract allows. A verified outcome is allowed at the operations of
    'a scenario. not-verifiable is allowed only at the entry of the surfaces the
    'manifest exempts: the Ribbon, which is unverifiable for every observation of
    'a scenario or for none, and the windows that could not be read on entry.
    'tools/check_excel_evidence.py accepts exactly these, and
    'tests/test_achieved_contract.py holds the two to each other.
    Private Const TEST_ACH_VERIFIED_OPERATIONS  As String = "|hide|show|restore|"                                               'Operations at which a verified outcome is allowed
    Private Const TEST_ACH_UNVERIFIABLE_ALLOWED As String = "|Ribbon.entry|Headings.entry|WorkbookTabs.entry|Gridlines.entry|"  'Element.operation pairs at which not-verifiable is allowed

    Private Const TEST_ACH_INVALID_VISIBILITY   As Long = 99    'A UIVisibility value no member names

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'The scenario under observation. Only TST_AchBegin sets it, and the steps and
    'the restoration of that scenario read it, so the cases are not reentrant.
    Private m_AchFault          As Long         'One-shot fault, taken by the next step or restoration it applies to
    Private m_AchElement        As String       'Element under observation
    Private m_AchScope          As String       'Scope it is written under
    Private m_AchSurface        As Long         'TEST_ACH_SURFACE_* of the element
    Private m_AchProperty       As String       'Boolean property of the Application or of each Window
    Private m_AchScopeValue     As Long         'UIWindowTargetScope passed to the setter
    Private m_AchCount          As Long         'Entries of the state vector
    Private m_AchEntry()        As Boolean      'State of each entry when the scenario began
    Private m_AchLabel()        As String       'Label of each entry
    Private m_AchTarget()       As Boolean      'TRUE when the scope writes the entry
    Private m_AchWin()          As Object       'Window of each entry of a window surface
    Private m_AchSkipped        As String       'Windows that could not be read when the scenario began
    Private m_AchUnverifiable   As Boolean      'TRUE when a failure to begin names a surface that cannot be verified
    Private m_AchFrameCaptured  As Boolean      'TRUE once the entry frame style was read
    Private m_AchNote           As String       'What the last read adds to the observed text
    Private m_AchSetterReturned As Boolean      'What the setter of the last step returned
    Private m_AchFailCount      As Long         'Failed observations of the scenario
    Private m_AchFailText       As String       'The failed observations, as text
    Private m_AchLocal          As tCertRecord  'Record of a case that runs outside certification, and of the fault seams

#If VBA7 Then
    Private m_AchHwnd       As LongPtr  'Frame or window the observation was established on
    Private m_AchEntryStyle As LongPtr  'Frame style when the scenario began
    Private m_AchStyle      As LongPtr  'Frame style at the last read
#Else
    Private m_AchHwnd       As Long 'Frame or window the observation was established on
    Private m_AchEntryStyle As Long 'Frame style when the scenario began
    Private m_AchStyle      As Long 'Frame style at the last read
#End If


'
'------------------------------------------------------------------------------
'
'                                PUBLIC RUNNER
'
'------------------------------------------------------------------------------
'

Public Sub Test_EXCEL_UI_RunAchievedStateVerification()

'
'------------------------------------------------------------------------------
'                       RUN ACHIEVED-STATE VERIFICATION
'------------------------------------------------------------------------------
' PURPOSE
'   Run the eight cases that read back the state each managed element achieves
'   after a write, and put the host back as it was found.
'
' WHY THIS EXISTS
'   A setter that returns True does not show that the requested state was
'   reached (#28). Release certification runs this runner as the
'   AchievedStateVerification unit, so a mismatch fails the unit and the
'   release, and the evidence lists what was requested and what was observed.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Snapshots the managed state through the harness, which does not use the
'     API under test.
'   - Runs the application-level case, the three window-scope cases, the
'     title-bar case, the Ribbon case, the fault-seam case and the
'     mismatch-verdict case through TST_CertRunCase.
'   - Restores the managed state and fails when a restoration step did not
'     read back as written.
'
' ERROR POLICY
'   - Raises after restoration on an assertion failure, an unexpected error or
'     an incomplete restoration.
'
' DEPENDENCIES
'   - TST_CertRunCase
'   - TST_SnapshotState
'   - TST_RestoreState
'   - TST_RecordRestoreFailure
'
' NOTES
'   Destructive: it creates and closes temporary workbooks and toggles every
'   managed UI element, so save unsaved work before running it. Each case runs
'   through TST_CertRunCase, so release certification, which runs this runner as
'   the AchievedStateVerification unit, records its start and its one terminal
'   status. Run directly, the runner records nothing in a certification record.
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "Test_EXCEL_UI_RunAchievedStateVerification"

    Dim SavedRibbonKnown        As Boolean  'TRUE when the Ribbon was read before the run
    Dim SavedRibbonVisible      As Boolean  'Ribbon visibility before the run
    Dim SavedStatusBarVisible   As Boolean  'Status Bar visibility before the run
    Dim SavedScrollBarsVisible  As Boolean  'Scroll Bars visibility before the run
    Dim SavedFormulaBarVisible  As Boolean  'Formula Bar visibility before the run

    Dim SavedWindowCount            As Long     'Application.Windows.Count before the run
    Dim SavedHeadingsVisible()      As Boolean  'Headings of each window before the run
    Dim SavedWorkbookTabsVisible()  As Boolean  'Workbook Tabs of each window before the run
    Dim SavedGridlinesVisible()     As Boolean  'Gridlines of each window before the run

    Dim SavedTitleBarKnown      As Boolean  'TRUE when the title bar was read before the run
    Dim SavedTitleBarVisible    As Boolean  'Title-bar visibility before the run

    Dim RestoreOK       As Boolean      'TRUE once every restoration step completed
    Dim RestoreFindings As tCertRecord  'Steps the restoration could not complete

    Dim OldScreenUpdating   As Boolean  'ScreenUpdating before the run
    Dim HasFailure          As Boolean  'TRUE when a failure was captured
    Dim FailNumber          As Long     'Captured failure number
    Dim FailSource          As String   'Captured failure source
    Dim FailDescription     As String   'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_Log PROC_NAME, "START", "Achieved-state verification started"

    'No scenario of an earlier run may leak into this one
        m_AchFault = TEST_ACH_FAULT_NONE
        m_AchSurface = TEST_ACH_SURFACE_NONE
        m_AchFrameCaptured = False

        OldScreenUpdating = Application.ScreenUpdating

'------------------------------------------------------------------------------
' SNAPSHOT CURRENT STATE
'------------------------------------------------------------------------------
    'The harness captures the managed state without the API under test
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
' RUN THE CASES
'------------------------------------------------------------------------------
    'Status Bar, Scroll Bars and Formula Bar
        TST_CertRunCase "TST_Case_AchievedAppLevelState", True

    'Headings, Workbook Tabs and Gridlines, once for each target scope, with the
    'windows outside the scope read back unchanged
        TST_CertRunCase "TST_Case_AchievedWindowStateAllWindows", True
        TST_CertRunCase "TST_Case_AchievedWindowStateActiveWindow", True
        TST_CertRunCase "TST_Case_AchievedWindowStateActiveWorkbookWindows", True

    'The title bar, read back from the frame the active frame pair named
        TST_CertRunCase "TST_Case_AchievedTitleBarFrameState", True

    'The Ribbon, only against the window that owned it
        TST_CertRunCase "TST_Case_AchievedRibbonOwningWindow", True

    'The fault seams, and the verdict of a mismatch
        TST_CertRunCase "TST_Case_AchievedOutcomesAreDistinct", True
        TST_CertRunCase "TST_Case_AchievedMismatchFailsTheUnit", True

    'Log successful completion before restoration
        TST_Log PROC_NAME, "PASS", _
            "Every managed element was read back at the state it was asked for"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Attempt to restore the pre-run state. The result starts FALSE, so a step
    'that raised past this boundary counts as failed.
        RestoreOK = False

        On Error Resume Next
        RestoreOK = TST_RestoreState( _
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
            TitleBarVisible:=SavedTitleBarVisible, _
            Findings:=RestoreFindings)

        Application.ScreenUpdating = OldScreenUpdating
        On Error GoTo 0

    'A restoration step that did not complete fails the run, after any case
    'failure
        TST_RecordRestoreFailure _
            RestoreOK, RestoreFindings, PROC_NAME, _
            FailNumber, FailSource, FailDescription

        If FailNumber <> 0 Then
            HasFailure = True
        End If

    'Raise the captured failure after restoration when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the failure so it can be re-raised after restoration
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        TST_Log PROC_NAME, "FAIL", _
            CStr(FailNumber) & ": " & FailDescription & _
            IIf(Len(FailSource) > 0, " | Source: " & FailSource, vbNullString)

        Resume SafeExit

End Sub

'
'------------------------------------------------------------------------------
'
'                               REGRESSION CASES
'
'------------------------------------------------------------------------------
'

Public Sub TST_Case_AchievedAppLevelState()

'
'------------------------------------------------------------------------------
'                   VERIFY APPLICATION-LEVEL ACHIEVED STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that the Status Bar, the Scroll Bars and the Formula Bar are written,
'   settled, read back and compared, on the path that changes them and on the
'   path that finds them already correct.
'
' WHY THIS EXISTS
'   These three properties belong to the Application. A write that returned
'   True is no proof that Excel kept it (#28), so each is read back from the
'   Application and compared with the request, and the entry state is put back.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - For each element, hides it, hides it again, shows it and shows it again
'     through UI_SetExcelUI_WithResult, reading the state back after each
'     write, then restores the entry state and reads that back too.
'   - Raises when any observation is a failed outcome.
'
' ERROR POLICY
'   - Raises after restoration, with every failed observation in the text.
'
' DEPENDENCIES
'   - TST_AchRunScenario
'   - TST_AchAssertClean
'
' NOTES
'   Each observation is one entry of the certification record while a
'   certification run is active, and is logged in every case.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunAchievedStateVerification
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_AchievedAppLevelState"

    Dim FailText        As String   'Failed observations of the three elements
    Dim HasFailure      As Boolean  'TRUE when a failure was captured
    Dim FailNumber      As Long     'Captured failure number
    Dim FailSource      As String   'Captured failure source
    Dim FailDescription As String   'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_Log PROC_NAME, "START", _
            "Reading back the Status Bar, Scroll Bars and Formula Bar after a write"

'------------------------------------------------------------------------------
' OBSERVE THE THREE ELEMENTS
'------------------------------------------------------------------------------
    'Every element is observed before any failure is raised, so one run names
    'every element that did not read back
        FailText = TST_AchRunScenario("StatusBar", "Application", _
            TEST_ACH_SURFACE_APP, "DisplayStatusBar", UI_TargetAllExcelWindows)

        FailText = TST_CertAppendDetail(FailText, _
            TST_AchRunScenario("ScrollBars", "Application", _
            TEST_ACH_SURFACE_APP, "DisplayScrollBars", UI_TargetAllExcelWindows))

        FailText = TST_CertAppendDetail(FailText, _
            TST_AchRunScenario("FormulaBar", "Application", _
            TEST_ACH_SURFACE_APP, "DisplayFormulaBar", UI_TargetAllExcelWindows))

'------------------------------------------------------------------------------
' ASSERT THE READBACK
'------------------------------------------------------------------------------
        TST_AchAssertClean FailText, PROC_NAME

        TST_Log PROC_NAME, "PASS", _
            "Status Bar, Scroll Bars and Formula Bar read back as requested"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Raise the captured failure after the scenarios restored the host
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

Public Sub TST_Case_AchievedWindowStateAllWindows()

'
'------------------------------------------------------------------------------
'                     VERIFY ACHIEVED STATE OF ALL WINDOWS
'------------------------------------------------------------------------------
' PURPOSE
'   Verify Headings, Workbook Tabs and Gridlines under
'   UI_TargetAllExcelWindows.
'
' WHY THIS EXISTS
'   Every window of the Excel instance is in this scope, so every window must be
'   read back at the requested state (#28).
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Builds two temporary workbooks, one with a second window, and runs the
'     window scenario for the three elements.
'
' ERROR POLICY
'   - Raises after the temporary workbooks are closed and the host is restored.
'
' DEPENDENCIES
'   - TST_AchRunWindowCase
'
' NOTES
'   None
'
' CALLED FROM
'   - Test_EXCEL_UI_RunAchievedStateVerification
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN THE WINDOW CASE
'------------------------------------------------------------------------------
        TST_AchRunWindowCase _
            "TST_Case_AchievedWindowStateAllWindows", _
            "AllWindows", UI_TargetAllExcelWindows

End Sub

Public Sub TST_Case_AchievedWindowStateActiveWindow()

'
'------------------------------------------------------------------------------
'                  VERIFY ACHIEVED STATE OF THE ACTIVE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Verify Headings, Workbook Tabs and Gridlines under UI_TargetActiveWindow,
'   and that every other window stays as it was.
'
' WHY THIS EXISTS
'   Only the active window is in this scope. A second window of the same
'   workbook, a window of another workbook and every window already open must be
'   read back unchanged (#28).
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Builds two temporary workbooks, one with a second window, activates the
'     first window of the first and runs the window scenario for the three
'     elements.
'
' ERROR POLICY
'   - Raises after the temporary workbooks are closed and the host is restored.
'
' DEPENDENCIES
'   - TST_AchRunWindowCase
'
' NOTES
'   None
'
' CALLED FROM
'   - Test_EXCEL_UI_RunAchievedStateVerification
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN THE WINDOW CASE
'------------------------------------------------------------------------------
        TST_AchRunWindowCase _
            "TST_Case_AchievedWindowStateActiveWindow", _
            "ActiveWindow", UI_TargetActiveWindow

End Sub

Public Sub TST_Case_AchievedWindowStateActiveWorkbookWindows()

'
'------------------------------------------------------------------------------
'             VERIFY ACHIEVED STATE OF THE ACTIVE WORKBOOK WINDOWS
'------------------------------------------------------------------------------
' PURPOSE
'   Verify Headings, Workbook Tabs and Gridlines under
'   UI_TargetActiveWorkbookWindows, and that the windows of other workbooks stay
'   as they were.
'
' WHY THIS EXISTS
'   Both windows of the active workbook are in this scope. The window of
'   another workbook and every window already open must be read back unchanged
'   (#28).
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Builds two temporary workbooks, one with a second window, activates the
'     first window of the first and runs the window scenario for the three
'     elements.
'
' ERROR POLICY
'   - Raises after the temporary workbooks are closed and the host is restored.
'
' DEPENDENCIES
'   - TST_AchRunWindowCase
'
' NOTES
'   None
'
' CALLED FROM
'   - Test_EXCEL_UI_RunAchievedStateVerification
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN THE WINDOW CASE
'------------------------------------------------------------------------------
        TST_AchRunWindowCase _
            "TST_Case_AchievedWindowStateActiveWorkbookWindows", _
            "ActiveWorkbookWindows", UI_TargetActiveWorkbookWindows

End Sub

Public Sub TST_Case_AchievedTitleBarFrameState()

'
'------------------------------------------------------------------------------
'                    VERIFY ACHIEVED STATE OF THE TITLE BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a title-bar hide and a title-bar show are read back from the
'   frame the active frame pair named before the write, and that a show needs
'   the complete caption mask.
'
' WHY THIS EXISTS
'   The title bar is a native frame style, written through WinAPI. Reading it
'   back from another frame, or accepting a partial caption mask as a title bar,
'   would confirm a write that did not happen (#45, #6, #28). The frame is the
'   one UI_TryGetActiveFramePair resolved before the write, and the read is
'   refused when the active frame is no longer that frame.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Resolves the active frame pair and reads the frame's style. A frame
'     without the complete caption, WS_BORDER and WS_DLGFRAME both set, is not
'     the baseline this case verifies a hide and a show against, so it fails
'     with readback-failed and a text of its own. Only the Ribbon may be
'     not-verifiable, which is all the manifest allows.
'   - Hides, hides again, shows and shows again, reading the style of that same
'     frame after each write and treating the frame as visible only with both
'     component bits of the caption.
'   - Puts the entry style back through the harness and reads it back.
'
' ERROR POLICY
'   - Raises after the entry style is put back, with every failed observation.
'   - Puts the entry style back again on every exit, and fails when it cannot.
'
' DEPENDENCIES
'   - TST_AchRunScenario
'   - TST_AchAssertClean
'   - TST_RestoreTitleBarStyle
'   - TST_RecordCleanupFailure
'
' NOTES
'   None
'
' CALLED FROM
'   - Test_EXCEL_UI_RunAchievedStateVerification
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_AchievedTitleBarFrameState"

    Dim FailText        As String   'Failed observations of the title bar
    Dim FrameRestored   As Boolean  'TRUE once the frame is back, or was never touched
    Dim FrameMsg        As String   'Diagnostic from the restoration

    Dim FailNumber      As Long     'Captured failure number
    Dim FailSource      As String   'Captured failure source
    Dim FailDescription As String   'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        FrameRestored = True

        m_AchFrameCaptured = False

        TST_Log PROC_NAME, "START", _
            "Reading back the title bar from the frame the write targeted"

'------------------------------------------------------------------------------
' OBSERVE THE TITLE BAR
'------------------------------------------------------------------------------
        FailText = TST_AchRunScenario("TitleBar", "ActiveFrame", _
            TEST_ACH_SURFACE_TITLEBAR, vbNullString, UI_TargetAllExcelWindows)

'------------------------------------------------------------------------------
' ASSERT THE READBACK
'------------------------------------------------------------------------------
        TST_AchAssertClean FailText, PROC_NAME

        TST_Log PROC_NAME, "PASS", _
            "The title bar read back as requested from the frame that was written"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Disable the handler first, so the raise below leaves this procedure
        On Error GoTo 0

    'Put the frame back through the harness whatever happened above
        If m_AchFrameCaptured Then
            FrameRestored = TST_RestoreTitleBarStyle( _
                m_AchHwnd, m_AchEntryStyle, True, FrameMsg)
        End If

        TST_RecordCleanupFailure _
            FrameRestored, FrameMsg, PROC_NAME, _
            FailNumber, FailSource, FailDescription

        If FailNumber <> 0 Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume SafeExit

End Sub

Public Sub TST_Case_AchievedRibbonOwningWindow()

'
'------------------------------------------------------------------------------
'                     VERIFY ACHIEVED STATE OF THE RIBBON
'------------------------------------------------------------------------------
' PURPOSE
'   Verify the Ribbon only against the window that owned it before the write.
'
' WHY THIS EXISTS
'   The Ribbon belongs to a window, and the setter and the reader both act on
'   the active window and accept no target. A value read back is attributable
'   only to the window that was active when it was read, and only if that window
'   did not change during the read. A Ribbon that cannot be addressed or read
'   from the start is not-verifiable: it is listed in the evidence and is never a
'   silent pass. A Ribbon that was readable and then is not, after a write or
'   after the restoration, is a failure, because the manifest exempts only a
'   Ribbon that is unverifiable for every observation (#28).
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Identifies the active window and its handle before the write.
'   - Hides, hides again, shows and shows again, reading the Ribbon back after
'     each write and accepting the value only while the active window is still
'     that window.
'   - Puts the Ribbon back through the harness and reads it back.
'
' ERROR POLICY
'   - Raises after restoration when an observation is a failed outcome. A Ribbon
'     that is not-verifiable from the start does not fail the case.
'
' DEPENDENCIES
'   - TST_AchRunScenario
'   - TST_AchAssertClean
'
' NOTES
'   None
'
' CALLED FROM
'   - Test_EXCEL_UI_RunAchievedStateVerification
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_AchievedRibbonOwningWindow"

    Dim FailText        As String   'Failed observations of the Ribbon
    Dim HasFailure      As Boolean  'TRUE when a failure was captured
    Dim FailNumber      As Long     'Captured failure number
    Dim FailSource      As String   'Captured failure source
    Dim FailDescription As String   'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_Log PROC_NAME, "START", _
            "Reading back the Ribbon of the window that owned it"

'------------------------------------------------------------------------------
' OBSERVE THE RIBBON
'------------------------------------------------------------------------------
        FailText = TST_AchRunScenario("Ribbon", "ActiveWindow", _
            TEST_ACH_SURFACE_RIBBON, vbNullString, UI_TargetAllExcelWindows)

'------------------------------------------------------------------------------
' ASSERT THE READBACK
'------------------------------------------------------------------------------
        TST_AchAssertClean FailText, PROC_NAME

        TST_Log PROC_NAME, "PASS", _
            "The Ribbon read back as requested, or was named not-verifiable"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
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

Public Sub TST_Case_AchievedOutcomesAreDistinct()

'
'------------------------------------------------------------------------------
'                        VERIFY EACH OUTCOME IS ITS OWN
'------------------------------------------------------------------------------
' PURPOSE
'   Verify, with fault seams and without relying on a naturally failing host,
'   that an ignored write, a failed write, a failed readback, a failed
'   restoration and a surface that cannot be verified each give a different
'   outcome and a different diagnostic text, and that a wait that never settles
'   ends within its bound.
'
' WHY THIS EXISTS
'   A host that behaves does not show that a failure would be seen (#28). Each
'   seam makes one of the failures happen on a real surface, so the case
'   asserts the outcome the observation records, the text it gives, and that the
'   host is back as it was.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Ignored write: the setter is asked to change nothing, so it returns True
'     and the state does not change. Expects state-mismatch, then a restoration
'     that finds nothing to put back.
'   - Failed write: the setter is given a value no member names, so it reports
'     a failure. Expects write-failed.
'   - Failed readback: the state is read back as a failure after a write that
'     took effect. Expects readback-failed, then a restoration that writes.
'   - Failed restoration: the restoration write is not made while the state
'     differs from the entry state. Expects restore-failed, then, once the seam
'     is spent, a restoration that writes.
'   - Unverifiable surface: the Ribbon is unreadable from the start. Expects
'     not-verifiable, which fails nothing and is listed in both documents.
'   - Ribbon that stops being readable: the readback fails after a write and
'     after the restoration. Expects readback-failed and restore-failed, never
'     not-verifiable, because the manifest exempts only a Ribbon that is
'     unverifiable for every observation.
'   - Title bar with no caption on entry: the entry frame reads as captionless.
'     Expects readback-failed, not an exemption, because only the Ribbon may be
'     not-verifiable.
'   - Asserts that the five texts differ, each naming its own outcome, and that
'     the ignored write ended within the bound of the settle.
'   - Asserts that none of this reached the record of the run executing the
'     case.
'
' ERROR POLICY
'   - Raises on assertion failure, after the host is restored.
'
' DEPENDENCIES
'   - TST_AchSeam
'   - TST_AchRestoreSeam
'   - TST_AchUnverifiableSeam
'   - TST_AchRibbonMidSequenceSeam
'   - TST_AssertTrue
'
' NOTES
'   Every record here is local, so the run executing the case records nothing
'   from the failures it provokes.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunAchievedStateVerification
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_AchievedOutcomesAreDistinct"

    Dim IgnoredText As String   'Failure text of the ignored write
    Dim WriteText   As String   'Failure text of the failed write
    Dim ReadText    As String   'Failure text of the failed readback
    Dim RestoreText As String   'Failure text of the failed restoration
    Dim CaptionText As String   'Failure text of a title bar with no caption on entry
    Dim MidText     As String   'Failure text of a Ribbon that stops being readable
    Dim Elapsed     As Double   'Seconds the ignored write took to end

    Dim RunEntries  As Long     'Run record: entries on entry
    Dim RunAchieved As Long     'Run record: observations on entry
    Dim RunFailed   As Boolean  'Run record: fatal flag on entry

    Dim HasFailure      As Boolean  'TRUE when a failure was captured
    Dim FailNumber      As Long     'Captured failure number
    Dim FailSource      As String   'Captured failure source
    Dim FailDescription As String   'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_Log PROC_NAME, "START", _
            "Validating that each way a write can fail has its own outcome"

    'The record of the run executing this case, compared again at the end
        RunEntries = m_CertRecord.EntryCount
        RunAchieved = m_CertRecord.AchievedCount
        RunFailed = m_CertRecord.RecordFailed

'------------------------------------------------------------------------------
' AN IGNORED WRITE
'------------------------------------------------------------------------------
    'The setter returns True and writes nothing; only the readback finds it
        IgnoredText = TST_AchSeam( _
            TEST_ACH_FAULT_IGNORED_WRITE, _
            TEST_ACH_OUTCOME_MISMATCH, _
            TEST_ACH_OUTCOME_NOOP, _
            Elapsed)

    'A state that never settles ends the poll at its bound, not later
        TST_AssertTrue _
            (Elapsed <= (TEST_ACH_POLL_LIMIT * TEST_ACH_POLL_SECONDS + _
                TEST_ACH_SLACK_SECONDS)), _
            PROC_NAME & ".IgnoredWrite.TheWaitIsBounded"

'------------------------------------------------------------------------------
' A FAILED WRITE
'------------------------------------------------------------------------------
    'The setter reports the failure itself, so the outcome is not a mismatch
        WriteText = TST_AchSeam( _
            TEST_ACH_FAULT_WRITE_FAILED, _
            TEST_ACH_OUTCOME_WRITE_FAILED, _
            TEST_ACH_OUTCOME_NOOP, _
            Elapsed)

'------------------------------------------------------------------------------
' A FAILED READBACK
'------------------------------------------------------------------------------
    'The write took effect and the state could not be read, so the restoration
    'has something to put back
        ReadText = TST_AchSeam( _
            TEST_ACH_FAULT_READBACK, _
            TEST_ACH_OUTCOME_READ_FAILED, _
            TEST_ACH_OUTCOME_WRITE, _
            Elapsed)

'------------------------------------------------------------------------------
' A FAILED RESTORATION
'------------------------------------------------------------------------------
        RestoreText = TST_AchRestoreSeam()

'------------------------------------------------------------------------------
' A TITLE BAR WITH NO CAPTION ON ENTRY
'------------------------------------------------------------------------------
    'Only the Ribbon may be not-verifiable, so this fails the run
        CaptionText = TST_AchCaptionSeam()

'------------------------------------------------------------------------------
' A SURFACE THAT CANNOT BE VERIFIED
'------------------------------------------------------------------------------
        TST_AchUnverifiableSeam

'------------------------------------------------------------------------------
' A SURFACE THAT STOPS BEING READABLE
'------------------------------------------------------------------------------
    'Only a Ribbon unreadable from the start is exempt; half-way it is a failure
        MidText = TST_AchRibbonMidSequenceSeam()

'------------------------------------------------------------------------------
' THE FOUR FAILURES ARE FOUR DIFFERENT TEXTS
'------------------------------------------------------------------------------
    'Each text names its own outcome and the reason that belongs to it
        TST_AssertTrue _
            (InStr(1, IgnoredText, TEST_ACH_OUTCOME_MISMATCH, vbBinaryCompare) > 0) And _
            (InStr(1, IgnoredText, "setter returned True", vbBinaryCompare) > 0), _
            PROC_NAME & ".Text.StateMismatch"

        TST_AssertTrue _
            (InStr(1, WriteText, TEST_ACH_OUTCOME_WRITE_FAILED, vbBinaryCompare) > 0) And _
            (InStr(1, WriteText, "setter reported a failure", vbBinaryCompare) > 0), _
            PROC_NAME & ".Text.WriteFailed"

        TST_AssertTrue _
            (InStr(1, ReadText, TEST_ACH_OUTCOME_READ_FAILED, vbBinaryCompare) > 0) And _
            (InStr(1, ReadText, "could not be read after the write", vbBinaryCompare) > 0), _
            PROC_NAME & ".Text.ReadbackFailed"

        TST_AssertTrue _
            (InStr(1, RestoreText, TEST_ACH_OUTCOME_RESTORE_FAILED, vbBinaryCompare) > 0) And _
            (InStr(1, RestoreText, "entry state was not put back", vbBinaryCompare) > 0), _
            PROC_NAME & ".Text.RestoreFailed"

        TST_AssertTrue _
            (Len(MidText) = 0) Or _
            ((InStr(1, MidText, TEST_ACH_OUTCOME_READ_FAILED, vbBinaryCompare) > 0) And _
            (InStr(1, MidText, TEST_ACH_OUTCOME_RESTORE_FAILED, vbBinaryCompare) > 0) And _
            (InStr(1, MidText, "could not be read after the write", vbBinaryCompare) > 0) And _
            (InStr(1, MidText, "could not be read after it was put back", vbBinaryCompare) > 0)), _
            PROC_NAME & ".Text.RibbonMidSequence"

        TST_AssertTrue _
            (InStr(1, CaptionText, TEST_ACH_OUTCOME_READ_FAILED, vbBinaryCompare) > 0) And _
            (InStr(1, CaptionText, "has no complete caption", vbBinaryCompare) > 0) And _
            (InStr(1, CaptionText, "not the baseline", vbBinaryCompare) > 0), _
            PROC_NAME & ".Text.CaptionlessEntry"

        TST_AssertTrue _
            (IgnoredText <> WriteText) And (IgnoredText <> ReadText) And _
            (IgnoredText <> RestoreText) And (WriteText <> ReadText) And _
            (WriteText <> RestoreText) And (ReadText <> RestoreText) And _
            (CaptionText <> IgnoredText) And (CaptionText <> WriteText) And _
            (CaptionText <> ReadText) And (CaptionText <> RestoreText), _
            PROC_NAME & ".Text.AllDiffer"

'------------------------------------------------------------------------------
' THE RUN'S OWN RECORD IS UNCHANGED
'------------------------------------------------------------------------------
    'Nothing above may reach the record of the run executing this case
        TST_AssertTrue _
            (m_CertRecord.EntryCount = RunEntries) And _
            (m_CertRecord.AchievedCount = RunAchieved) And _
            (m_CertRecord.RecordFailed = RunFailed), _
            PROC_NAME & ".RunRecordUnchanged"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC_NAME, "PASS", _
            "Ignored write, failed write, failed readback, failed restoration and " & _
            "an unverifiable surface each gave their own outcome and text"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Disarm a seam that was not reached, so it cannot fault the next scenario
        m_AchFault = TEST_ACH_FAULT_NONE

    'Raise the captured failure after the scenarios restored the host
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

Public Sub TST_Case_AchievedMismatchFailsTheUnit()

'
'------------------------------------------------------------------------------
'                       VERIFY A MISMATCH FAILS THE UNIT
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a state mismatch fails the case and the verdict even when the
'   production setter returned True, and that both evidence documents carry
'   what was requested, what was observed and the outcome.
'
' WHY THIS EXISTS
'   The defect this unit exists to catch is a setter that succeeds and an
'   Excel that keeps the old state (#28). A mismatch that only appeared in a log
'   line would leave the unit passing, so this case asserts the raise, the
'   verdict and the documents.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Provokes an ignored write on the Status Bar, so the setter returns True
'     and the state does not change, and asserts that the setter did return
'     True.
'   - Asserts that the assertion every case uses raises, with the number of an
'     unclean observation and the outcome state-mismatch in its text.
'   - Composes the evidence of a record that holds that observation and a
'     passing unit, and asserts FAIL, achieved=FAILED, and the requested and
'     observed state, the outcome, the MISMATCH line and the OBSERVATIONS line.
'   - Runs the same scenario with no fault and asserts PASS, achieved=OK and an
'     empty mismatch list.
'
' ERROR POLICY
'   - Raises on assertion failure, after the host is restored.
'
' DEPENDENCIES
'   - TST_AchBegin
'   - TST_AchStep
'   - TST_AchRestore
'   - TST_AchRunOn
'   - TST_AchCompose
'   - TST_AchAssertClean
'
' NOTES
'   Every record here is local, so the run executing the case records nothing
'   from the mismatch it provokes.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunAchievedStateVerification
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    'The value is deliberate: it stays the bare procedure name, without the module prefix
    Const PROC_NAME As String = "TST_Case_AchievedMismatchFailsTheUnit"

    Dim Why             As String   'Why the scenario could not begin
    Dim Requested       As Boolean  'State asked for by the ignored write
    Dim Operation       As String   'hide or show
    Dim RequestedText   As String   'Requested state as the evidence says it
    Dim ObservedText    As String   'Observed state as the evidence says it
    Dim SetterReturned  As Boolean  'What the setter of the ignored write returned
    Dim Passed          As Boolean  'Composed verdict
    Dim Verdict         As String   'Composed RESULT line
    Dim JsonText        As String   'Composed JSON evidence
    Dim ReportText      As String   'Composed text evidence
    Dim FailText        As String   'Failed observations of the scenario

    Dim RaisedNumber        As Long     'Number the assertion raised
    Dim RaisedDescription   As String   'Text the assertion raised

    Dim HasFailure      As Boolean  'TRUE when a failure was captured
    Dim FailNumber      As Long     'Captured failure number
    Dim FailSource      As String   'Captured failure source
    Dim FailDescription As String   'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_Log PROC_NAME, "START", _
            "Validating that a mismatch fails the unit although the setter returned True"

'------------------------------------------------------------------------------
' PROVOKE AN IGNORED WRITE
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertResetRecord(m_AchLocal), _
            PROC_NAME & ".Reset"

        If Not TST_AchBegin("StatusBar", "Application", TEST_ACH_SURFACE_APP, _
            "DisplayStatusBar", UI_TargetAllExcelWindows, Why) Then

            Err.Raise TEST_ACH_ERR_BASE + TEST_ACH_ERR_SEAM, PROC_NAME, _
                "the Status Bar could not be read: " & Why
        End If

    'Ask for the state the Status Bar is not in, so that a change is owed
        Requested = Not m_AchEntry(1)
        Operation = IIf(Requested, "show", "hide")

        RequestedText = IIf(Requested, "visible", "hidden")
        ObservedText = IIf(Requested, "hidden", "visible")

        m_AchFault = TEST_ACH_FAULT_IGNORED_WRITE

        TST_AchStep m_AchLocal, Operation, Requested

        SetterReturned = m_AchSetterReturned

        TST_AchRestore m_AchLocal

        FailText = m_AchFailText

        TST_AssertBooleanEquals True, SetterReturned, _
            PROC_NAME & ".TheSetterReturnedTrue"

'------------------------------------------------------------------------------
' THE ASSERTION THAT EVERY CASE USES RAISES
'------------------------------------------------------------------------------
    'Capture immediately, before anything else can clear Err
        On Error Resume Next

        TST_AchAssertClean FailText, PROC_NAME

        RaisedNumber = Err.Number
        RaisedDescription = Err.Description

        Err.Clear

        On Error GoTo ErrorHandler

        TST_AssertTrue _
            (RaisedNumber = TEST_ACH_ERR_BASE + TEST_ACH_ERR_UNCLEAN), _
            PROC_NAME & ".TheUnitFails.Number"

        TST_AssertTrue _
            (InStr(1, RaisedDescription, TEST_ACH_OUTCOME_MISMATCH, vbBinaryCompare) > 0), _
            PROC_NAME & ".TheUnitFails.Text"

'------------------------------------------------------------------------------
' THE VERDICT AND BOTH DOCUMENTS
'------------------------------------------------------------------------------
        TST_AchCompose Passed, Verdict, JsonText, ReportText

        TST_AssertBooleanEquals False, Passed, PROC_NAME & ".Verdict.NotPassed"

        TST_AssertTrue _
            (InStr(1, Verdict, "RESULT: FAIL | COMPLETE | units=1 failed=0 skipped=0 cleanup=OK", _
                vbBinaryCompare) = 1), _
            PROC_NAME & ".Verdict.Line"

        TST_AssertTrue _
            (InStr(1, JsonText, """achieved"":""FAILED""", vbBinaryCompare) > 0) And _
            (InStr(1, JsonText, """operation"":""" & Operation & """," & _
                """requested"":""" & RequestedText & """," & _
                """observed"":""" & ObservedText & """," & _
                """outcome"":""" & TEST_ACH_OUTCOME_MISMATCH & """", vbBinaryCompare) > 0), _
            PROC_NAME & ".Json.CarriesTheObservation"

        TST_AssertTrue _
            (InStr(1, ReportText, vbNewLine & "  MISMATCH  StatusBar.Application." & _
                Operation & " | " & TEST_ACH_OUTCOME_MISMATCH & ": requested=" & _
                RequestedText & " observed=" & ObservedText, vbBinaryCompare) > 0) And _
            (InStr(1, ReportText, vbNewLine & "  OBSERVATIONS  checked=2 verified=1 " & _
                "failed=1 unverifiable=0 achieved=FAILED" & vbNewLine, vbBinaryCompare) > 0), _
            PROC_NAME & ".Report.ListsTheMismatch"

'------------------------------------------------------------------------------
' THE SAME SCENARIO WITHOUT A FAULT PASSES
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertResetRecord(m_AchLocal), _
            PROC_NAME & ".Clean.Reset"

        FailText = TST_AchRunOn(m_AchLocal, "StatusBar", "Application", _
            TEST_ACH_SURFACE_APP, "DisplayStatusBar", UI_TargetAllExcelWindows)

        TST_AssertTrue (Len(FailText) = 0), PROC_NAME & ".Clean.NoFailure"

        TST_AchCompose Passed, Verdict, JsonText, ReportText

        TST_AssertBooleanEquals True, Passed, PROC_NAME & ".Clean.Passed"

        TST_AssertTrue _
            (InStr(1, JsonText, """achieved"":""OK""", vbBinaryCompare) > 0) And _
            (InStr(1, JsonText, """achievedChecked"":5", vbBinaryCompare) > 0) And _
            (InStr(1, JsonText, """achievedMismatches"":[]", vbBinaryCompare) > 0), _
            PROC_NAME & ".Clean.Json"

        TST_AssertTrue _
            (InStr(1, ReportText, vbNewLine & "  OBSERVATIONS  checked=5 verified=5 " & _
                "failed=0 unverifiable=0 achieved=OK" & vbNewLine, vbBinaryCompare) > 0), _
            PROC_NAME & ".Clean.Report"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC_NAME, "PASS", _
            "A state mismatch failed the unit and the verdict although the " & _
            "setter returned True"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Disarm a seam that was not reached, so it cannot fault the next scenario
        m_AchFault = TEST_ACH_FAULT_NONE

    'Raise the captured failure after the scenarios restored the host
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

'
'------------------------------------------------------------------------------
'
'                             FAULT SEAM SCENARIOS
'
'------------------------------------------------------------------------------
'

Private Function TST_AchSeam( _
    ByVal Fault As Long, _
    ByVal FirstOutcome As String, _
    ByVal RestoreOutcome As String, _
    ByRef Elapsed As Double) _
    As String
'
'------------------------------------------------------------------------------
'                     RUN ONE FAULTED STEP AND ITS RESTORE
'------------------------------------------------------------------------------
' PURPOSE
'   Arm one fault, make one step that owes a change on the Status Bar, restore
'   the entry state and assert the outcome of each.
'
' WHY THIS EXISTS
'   The ignored-write, failed-write and failed-readback seams differ only in the
'   fault and in the outcomes they expect, so one procedure holds the sequence
'   (#28).
'
' INPUTS
'   Fault
'     The TEST_ACH_FAULT_* the first step takes.
'
'   FirstOutcome / RestoreOutcome
'     The outcomes the step and the restoration must record.
'
'   Elapsed
'     ByRef. Receives the seconds the step took, with the timer rollover handled.
'
' RETURNS
'   String
'     The failed observations of the scenario, as text.
'
' BEHAVIOR
'   - Resets the local record and begins a Status Bar scenario.
'   - Arms the fault, asks for the state the Status Bar is not in, times the
'     step, restores and asserts the two outcomes and that the fault was used.
'
' ERROR POLICY
'   - Raises on assertion failure, or when the Status Bar cannot be read.
'
' DEPENDENCIES
'   - TST_AchBegin
'   - TST_AchStep
'   - TST_AchRestore
'   - TST_AchOutcomeAt
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_Case_AchievedOutcomesAreDistinct
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchSeam"

    Dim Why         As String   'Why the scenario could not begin
    Dim Requested   As Boolean  'State asked for by the step
    Dim Started     As Double   'Timer when the step began

'------------------------------------------------------------------------------
' BEGIN
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertResetRecord(m_AchLocal), _
            PROC_NAME & ".Reset"

        If Not TST_AchBegin("StatusBar", "Application", TEST_ACH_SURFACE_APP, _
            "DisplayStatusBar", UI_TargetAllExcelWindows, Why) Then

            Err.Raise TEST_ACH_ERR_BASE + TEST_ACH_ERR_SEAM, PROC_NAME, _
                "the Status Bar could not be read: " & Why
        End If

'------------------------------------------------------------------------------
' STEP WITH THE FAULT
'------------------------------------------------------------------------------
    'Ask for the state the Status Bar is not in, so that a change is owed
        Requested = Not m_AchEntry(1)

        m_AchFault = Fault

        Started = Timer

        TST_AchStep m_AchLocal, IIf(Requested, "show", "hide"), Requested

        Elapsed = Timer - Started

        If Elapsed < 0# Then
            Elapsed = Elapsed + TEST_ACH_SECONDS_PER_DAY
        End If

'------------------------------------------------------------------------------
' RESTORE AND ASSERT
'------------------------------------------------------------------------------
        TST_AchRestore m_AchLocal

        TST_AssertTrue (m_AchFault = TEST_ACH_FAULT_NONE), _
            PROC_NAME & ".TheFaultWasUsed"

        TST_AssertTrue _
            (TST_AchOutcomeAt(m_AchLocal, 1) = FirstOutcome), _
            PROC_NAME & ".FirstOutcome." & FirstOutcome

        TST_AssertTrue _
            (TST_AchOutcomeAt(m_AchLocal, 2) = RestoreOutcome), _
            PROC_NAME & ".RestoreOutcome." & RestoreOutcome

        TST_AchSeam = m_AchFailText

End Function

Private Function TST_AchRestoreSeam() _
    As String
'
'------------------------------------------------------------------------------
'                        RUN A FAILED RESTORATION SEAM
'------------------------------------------------------------------------------
' PURPOSE
'   Leave the Status Bar changed, make the restoration write fail, and assert
'   restore-failed; then restore for real.
'
' WHY THIS EXISTS
'   A restoration that is not read back is a restoration nobody can fail (#28).
'   After a normal step the state differs from the entry state, which is what
'   gives the skipped restoration write something to leave undone.
'
' INPUTS
'   None
'
' RETURNS
'   String
'     The failed observations of the scenario, as text: the restore-failed one.
'
' BEHAVIOR
'   - Begins a Status Bar scenario and makes one normal step, which changes it.
'   - Arms the restoration fault and restores: the write is not made, the
'     state read back differs, and the outcome is restore-failed.
'   - Restores again, with the seam spent, and asserts the state is back.
'
' ERROR POLICY
'   - Raises on assertion failure, or when the Status Bar cannot be read.
'
' DEPENDENCIES
'   - TST_AchBegin
'   - TST_AchStep
'   - TST_AchRestore
'   - TST_AchOutcomeAt
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_Case_AchievedOutcomesAreDistinct
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchRestoreSeam"

    Dim Why         As String   'Why the scenario could not begin
    Dim Requested   As Boolean  'State asked for by the step
    Dim FailedText  As String   'The restore-failed observation, as text

'------------------------------------------------------------------------------
' BEGIN AND CHANGE THE STATE
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertResetRecord(m_AchLocal), _
            PROC_NAME & ".Reset"

        If Not TST_AchBegin("StatusBar", "Application", TEST_ACH_SURFACE_APP, _
            "DisplayStatusBar", UI_TargetAllExcelWindows, Why) Then

            Err.Raise TEST_ACH_ERR_BASE + TEST_ACH_ERR_SEAM, PROC_NAME, _
                "the Status Bar could not be read: " & Why
        End If

        Requested = Not m_AchEntry(1)

        TST_AchStep m_AchLocal, IIf(Requested, "show", "hide"), Requested

'------------------------------------------------------------------------------
' RESTORE WITH THE FAULT
'------------------------------------------------------------------------------
        m_AchFault = TEST_ACH_FAULT_RESTORE

        TST_AchRestore m_AchLocal

        FailedText = m_AchFailText

        TST_AssertTrue (m_AchFault = TEST_ACH_FAULT_NONE), _
            PROC_NAME & ".TheFaultWasUsed"

        TST_AssertTrue _
            (TST_AchOutcomeAt(m_AchLocal, 1) = TEST_ACH_OUTCOME_WRITE), _
            PROC_NAME & ".TheStepWasVerified"

        TST_AssertTrue _
            (TST_AchOutcomeAt(m_AchLocal, 2) = TEST_ACH_OUTCOME_RESTORE_FAILED), _
            PROC_NAME & ".RestoreOutcome"

'------------------------------------------------------------------------------
' RESTORE FOR REAL
'------------------------------------------------------------------------------
    'The seam is spent: this restoration writes, and reads the entry state back
        TST_AchRestore m_AchLocal

        TST_AssertTrue _
            (TST_AchOutcomeAt(m_AchLocal, 3) = TEST_ACH_OUTCOME_WRITE), _
            PROC_NAME & ".TheHostIsBack"

        TST_AchRestoreSeam = FailedText

End Function

Private Function TST_AchCaptionSeam() _
    As String
'
'------------------------------------------------------------------------------
'                      RUN A CAPTIONLESS ENTRY FRAME SEAM
'------------------------------------------------------------------------------
' PURPOSE
'   Make the entry frame of a title-bar scenario read as having no complete
'   caption, and assert that the unit fails instead of exempting it.
'
' WHY THIS EXISTS
'   A title bar that is hidden when certification starts means the host is not
'   the baseline the unit verifies. The manifest and the evidence validator
'   exempt only the Ribbon from verification, so a live verdict that passed
'   here would pass a run whose bundle the validator rejects (#28).
'
' INPUTS
'   None
'
' RETURNS
'   String
'     The failed observations of the scenario, as text.
'
' BEHAVIOR
'   - Arms TEST_ACH_FAULT_NO_CAPTION and begins a title-bar scenario, which must
'     not begin and must not be marked unverifiable.
'   - Records the begin failure and asserts that its outcome is readback-failed.
'
' ERROR POLICY
'   - Raises on assertion failure, or when the scenario begins.
'
' DEPENDENCIES
'   - TST_AchBegin
'   - TST_AchRecordBeginFailure
'   - TST_AchOutcomeAt
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_Case_AchievedOutcomesAreDistinct
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchCaptionSeam"

    Dim Why As String   'Why the scenario could not begin

'------------------------------------------------------------------------------
' BEGIN WITH A CAPTIONLESS ENTRY FRAME
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertResetRecord(m_AchLocal), _
            PROC_NAME & ".Reset"

        m_AchFault = TEST_ACH_FAULT_NO_CAPTION

        If TST_AchBegin("TitleBar", "ActiveFrame", TEST_ACH_SURFACE_TITLEBAR, _
            vbNullString, UI_TargetAllExcelWindows, Why) Then

            Err.Raise TEST_ACH_ERR_BASE + TEST_ACH_ERR_SEAM, PROC_NAME, _
                "the title-bar scenario began on a frame with no caption"
        End If

'------------------------------------------------------------------------------
' ASSERT THE FAILURE
'------------------------------------------------------------------------------
        TST_AssertTrue (m_AchFault = TEST_ACH_FAULT_NONE), _
            PROC_NAME & ".TheFaultWasUsed"

        TST_AssertBooleanEquals False, m_AchUnverifiable, _
            PROC_NAME & ".IsNotExempt"

        TST_AchRecordBeginFailure m_AchLocal, Why

        TST_AssertTrue _
            (TST_AchOutcomeAt(m_AchLocal, 1) = TEST_ACH_OUTCOME_READ_FAILED), _
            PROC_NAME & ".Outcome"

        TST_AchCaptionSeam = m_AchFailText

End Function

Private Sub TST_AchUnverifiableSeam()

'
'------------------------------------------------------------------------------
'                       RUN AN UNVERIFIABLE-SURFACE SEAM
'------------------------------------------------------------------------------
' PURPOSE
'   Make the Ribbon unreadable from the start and assert not-verifiable, that it
'   fails nothing, and that both documents list it.
'
' WHY THIS EXISTS
'   A surface that cannot be addressed or read must be named, never passed in
'   silence, and never counted as a failure of the code under test (#28). The
'   Ribbon is the one surface the manifest exempts, and only when it is
'   unverifiable for every observation of a scenario; a Ribbon that was readable
'   and then is not is TST_AchRibbonMidSequenceSeam.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Arms the readback fault and begins a Ribbon scenario, which reads the
'     entry state as a failure and so cannot begin.
'   - Records the begin failure, which is the only observation of the scenario,
'     and expects not-verifiable.
'   - Asserts that no failure was recorded and that the composed documents list
'     the Ribbon as UNVERIFIABLE and as achievedUnverifiable.
'
' ERROR POLICY
'   - Raises on assertion failure, or when the scenario begins.
'
' DEPENDENCIES
'   - TST_AchBegin
'   - TST_AchRecordBeginFailure
'   - TST_AchCompose
'
' NOTES
'   On a host whose Ribbon cannot be addressed at all, the scenario cannot begin
'   either, and records the same outcome of its own accord.
'
' CALLED FROM
'   - TST_Case_AchievedOutcomesAreDistinct
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchUnverifiableSeam"

    Dim Why         As String   'Why the scenario could not begin
    Dim Passed      As Boolean  'Composed verdict
    Dim Verdict     As String   'Composed RESULT line
    Dim JsonText    As String   'Composed JSON evidence
    Dim ReportText  As String   'Composed text evidence

'------------------------------------------------------------------------------
' BEGIN WITH AN UNREADABLE RIBBON
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertResetRecord(m_AchLocal), _
            PROC_NAME & ".Reset"

        m_AchFault = TEST_ACH_FAULT_READBACK

        If TST_AchBegin("Ribbon", "ActiveWindow", TEST_ACH_SURFACE_RIBBON, _
            vbNullString, UI_TargetAllExcelWindows, Why) Then

            Err.Raise TEST_ACH_ERR_BASE + TEST_ACH_ERR_SEAM, PROC_NAME, _
                "the Ribbon scenario began although its entry state was unreadable"
        End If

    'The fault stays armed when the Ribbon could not be addressed at all
        m_AchFault = TEST_ACH_FAULT_NONE

        TST_AssertBooleanEquals True, m_AchUnverifiable, _
            PROC_NAME & ".BeginFailureIsUnverifiable"

        TST_AchRecordBeginFailure m_AchLocal, Why

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        TST_AssertTrue _
            (TST_AchOutcomeAt(m_AchLocal, 1) = TEST_ACH_OUTCOME_UNVERIFIABLE), _
            PROC_NAME & ".Outcome"

    'Not verifiable is not a failure of the code under test
        TST_AssertTrue (Len(m_AchFailText) = 0), PROC_NAME & ".FailsNothing"

        TST_AchCompose Passed, Verdict, JsonText, ReportText

        TST_AssertBooleanEquals True, Passed, PROC_NAME & ".Verdict.Passed"

        TST_AssertTrue _
            (InStr(1, ReportText, vbNewLine & "  UNVERIFIABLE  Ribbon.ActiveWindow.entry", _
                vbBinaryCompare) > 0), _
            PROC_NAME & ".Report.ListsTheSurface"

        TST_AssertTrue _
            (InStr(1, JsonText, """achievedUnverifiable"":[{""seq"":1,", _
                vbBinaryCompare) > 0) And _
            (InStr(1, JsonText, """outcome"":""" & TEST_ACH_OUTCOME_UNVERIFIABLE & """", _
                vbBinaryCompare) > 0), _
            PROC_NAME & ".Json.ListsTheSurface"

End Sub

Private Function TST_AchRibbonMidSequenceSeam() _
    As String
'
'------------------------------------------------------------------------------
'                    RUN A RIBBON THAT STOPS BEING READABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Make the Ribbon unreadable after the scenario began, after a write and after
'   the restoration, and assert that each is a failure and not an exemption.
'
' WHY THIS EXISTS
'   The manifest exempts a Ribbon that is unverifiable for every observation,
'   and tools/check_excel_evidence.py rejects a mixed result. A live verdict that
'   passed a Ribbon that was readable and then was not would pass a run whose
'   bundle the validator rejects (#28).
'
' INPUTS
'   None
'
' RETURNS
'   String
'     The failed observations of the scenario, as text; empty when the Ribbon
'     cannot be read from the start, which is the exempt case and has nothing to
'     cut short.
'
' BEHAVIOR
'   - Begins a Ribbon scenario. When it cannot begin it records the begin
'     failure, expects not-verifiable and returns an empty text.
'   - Otherwise fails the readback after a write (readback-failed), restores,
'     makes a normal step, fails the readback after the restoration
'     (restore-failed) and restores again, and asserts that no observation is
'     not-verifiable and that the composed verdict fails.
'
' ERROR POLICY
'   - Raises on assertion failure.
'
' DEPENDENCIES
'   - TST_AchBegin
'   - TST_AchStep
'   - TST_AchRestore
'   - TST_AchRecordBeginFailure
'   - TST_AchCompose
'   - TST_AchOutcomeAt
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_Case_AchievedOutcomesAreDistinct
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchRibbonMidSequenceSeam"

    Dim Why         As String   'Why the scenario could not begin
    Dim Position    As Long     'Cursor over the observations
    Dim Requested   As Boolean  'State asked for by a step
    Dim Passed      As Boolean  'Composed verdict
    Dim Verdict     As String   'Composed RESULT line
    Dim JsonText    As String   'Composed JSON evidence
    Dim ReportText  As String   'Composed text evidence

'------------------------------------------------------------------------------
' BEGIN
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertResetRecord(m_AchLocal), _
            PROC_NAME & ".Reset"

        TST_AchRibbonMidSequenceSeam = vbNullString

        If Not TST_AchBegin("Ribbon", "ActiveWindow", TEST_ACH_SURFACE_RIBBON, _
            vbNullString, UI_TargetAllExcelWindows, Why) Then

        'Unreadable from the start is the exemption itself
            TST_AssertBooleanEquals True, m_AchUnverifiable, _
                PROC_NAME & ".BeginFailureIsUnverifiable"

            TST_AchRecordBeginFailure m_AchLocal, Why

            TST_AssertTrue _
                (TST_AchOutcomeAt(m_AchLocal, 1) = TEST_ACH_OUTCOME_UNVERIFIABLE), _
                PROC_NAME & ".Unreadable.Outcome"

        Else

        'Readable on entry, unreadable after a write
            Requested = Not m_AchEntry(1)

            m_AchFault = TEST_ACH_FAULT_READBACK

            TST_AchStep m_AchLocal, IIf(Requested, "show", "hide"), Requested

            TST_AchRestore m_AchLocal

        'Readable on entry, unreadable after the restoration
            Requested = Not m_AchEntry(1)

            TST_AchStep m_AchLocal, IIf(Requested, "show", "hide"), Requested

            m_AchFault = TEST_ACH_FAULT_RESTORE_READ

            TST_AchRestore m_AchLocal

            TST_AchRestore m_AchLocal

            TST_AssertTrue (m_AchFault = TEST_ACH_FAULT_NONE), _
                PROC_NAME & ".TheFaultsWereUsed"

            TST_AssertTrue _
                (TST_AchOutcomeAt(m_AchLocal, 1) = TEST_ACH_OUTCOME_READ_FAILED), _
                PROC_NAME & ".AfterTheWrite"

            TST_AssertTrue _
                (TST_AchOutcomeAt(m_AchLocal, 4) = TEST_ACH_OUTCOME_RESTORE_FAILED), _
                PROC_NAME & ".AfterTheRestoration"

        'Nothing was exempted halfway through the scenario
            For Position = 1 To 5

                TST_AssertTrue _
                    (TST_AchOutcomeAt(m_AchLocal, Position) <> TEST_ACH_OUTCOME_UNVERIFIABLE), _
                    PROC_NAME & ".NotExempt." & CStr(Position)

            Next Position

            TST_AssertTrue (m_AchFailCount = 2), PROC_NAME & ".TwoFailures"

            TST_AchCompose Passed, Verdict, JsonText, ReportText

            TST_AssertBooleanEquals False, Passed, PROC_NAME & ".Verdict.NotPassed"

            TST_AssertTrue _
                (InStr(1, ReportText, vbNewLine & "  UNVERIFIABLE  ", _
                    vbBinaryCompare) = 0), _
                PROC_NAME & ".Report.ListsNoExemption"

            TST_AchRibbonMidSequenceSeam = m_AchFailText

        End If

End Function

Private Sub TST_AchCompose( _
    ByRef Passed As Boolean, _
    ByRef Verdict As String, _
    ByRef JsonText As String, _
    ByRef ReportText As String)
'
'------------------------------------------------------------------------------
'                     COMPOSE EVIDENCE OF THE LOCAL RECORD
'------------------------------------------------------------------------------
' PURPOSE
'   Add one passing unit to the local record, compare its empty case inventory
'   and compose both evidence documents from it.
'
' WHY THIS EXISTS
'   A seam asserts what certification would emit for the observations it
'   provoked, so it composes through the same TST_CertComposeEvidence. The unit
'   and the compared inventory leave the observations as the only thing that can
'   fail the run (#28).
'
' INPUTS
'   Passed / Verdict / JsonText / ReportText
'     ByRef outputs: the composed verdict, the RESULT line and the two
'     documents.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Stores a passing unit U in the local record and compares its case
'     inventory, which is empty and therefore exact.
'   - Composes the evidence with cleanup passed as OK.
'
' ERROR POLICY
'   - Raises on assertion failure, when the unit cannot be stored.
'
' DEPENDENCIES
'   - TST_CertTryAppendEntry
'   - TST_CertEvaluateInventory
'   - TST_CertComposeEvidence
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_Case_AchievedMismatchFailsTheUnit
'   - TST_AchUnverifiableSeam
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchCompose"

    Dim UnitsFailed     As Long     'Composed failed-unit count
    Dim Complete        As Boolean  'Composed completeness
    Dim RecordingDetail As String   'Composed recording failure

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
    'One passing unit, so that only the observations can fail the run
        TST_AssertBooleanEquals True, _
            TST_CertTryAppendEntry(m_AchLocal, TEST_CERT_ENTRY_UNIT, "U", True, _
                vbNullString), _
            PROC_NAME & ".UnitStored"

        TST_AssertBooleanEquals True, TST_CertEvaluateInventory(m_AchLocal), _
            PROC_NAME & ".InventoryExact"

        TST_CertComposeEvidence _
            Rec:=m_AchLocal, _
            CleanupOK:=True, _
            CleanupDetail:=vbNullString, _
            UnitsFailed:=UnitsFailed, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            RecordingDetail:=RecordingDetail, _
            JsonText:=JsonText, _
            ReportText:=ReportText

End Sub

Private Function TST_AchOutcomeAt( _
    ByRef Rec As tCertRecord, _
    ByVal Position As Long) _
    As String
'
'------------------------------------------------------------------------------
'                      READ THE OUTCOME OF AN OBSERVATION
'------------------------------------------------------------------------------
' PURPOSE
'   Return the outcome of the Position-th achieved-state observation of a
'   record, in recording order.
'
' WHY THIS EXISTS
'   A seam asserts the outcome of each step and restoration it made, and the
'   outcome is one field of the packed detail of the entry (#28).
'
' INPUTS
'   Rec
'     ByRef. The record to read. It is not changed.
'
'   Position
'     1-based position among the achieved-state observations.
'
' RETURNS
'   String
'     The outcome text, or "none" when the record holds fewer observations or
'     cannot be read.
'
' BEHAVIOR
'   - Walks the entries once and counts the achieved ones.
'   - Splits the detail of the Position-th and returns its fifth field.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertRecordIsConsistent
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchSeam
'   - TST_AchRestoreSeam
'   - TST_AchUnverifiableSeam
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchOutcomeAt"

    Dim ScanIdx     As Long     'Cursor over the entries
    Dim Seen        As Long     'Achieved entries counted
    Dim Fields()    As String   'Fields of one observation

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_AchOutcomeAt = "none"

        If Not TST_CertRecordIsConsistent(Rec) Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' FIND THE OBSERVATION
'------------------------------------------------------------------------------
        For ScanIdx = 1 To Rec.EntryCount

            If Rec.EntryKind(ScanIdx) = TEST_CERT_ENTRY_ACHIEVED Then

                Seen = Seen + 1

                If Seen = Position Then

                    Fields = Split(Rec.EntryDetail(ScanIdx), vbTab)

                    If UBound(Fields) = 5 Then
                        TST_AchOutcomeAt = Fields(4)
                    End If

                    Exit For

                End If

            End If

        Next ScanIdx

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'An unreadable record is reported as one, never as a verified outcome
        TST_AchOutcomeAt = "none"

        Resume SafeExit

End Function

'
'------------------------------------------------------------------------------
'
'                            SCENARIOS AND FIXTURES
'
'------------------------------------------------------------------------------
'

Private Sub TST_AchRunWindowCase( _
    ByVal CaseName As String, _
    ByVal Scope As String, _
    ByVal ScopeValue As Long)
'
'------------------------------------------------------------------------------
'                   OBSERVE THE WINDOW ELEMENTS OF ONE SCOPE
'------------------------------------------------------------------------------
' PURPOSE
'   Build the window fixture, observe Headings, Workbook Tabs and Gridlines
'   under one target scope, assert the readback and put the host back.
'
' WHY THIS EXISTS
'   The three scope cases differ only in the scope, so the fixture, the
'   observation and the cleanup are written once (#28).
'
' INPUTS
'   CaseName
'     Name of the calling case, used as the log prefix and the error source.
'
'   Scope
'     The scope label recorded with each observation.
'
'   ScopeValue
'     The UIWindowTargetScope the setter is called with.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Creates a workbook with a second window and another workbook, shows the
'     three elements on their three windows and activates the first window of
'     the first workbook.
'   - Observes the three elements: hide, hide again, show, show again and the
'     restoration. Every window already open is read too, so a window outside
'     the scope that changed is a mismatch.
'   - Closes both workbooks and activates the window that was active.
'
' ERROR POLICY
'   - Raises after the cleanup, with every failed observation.
'
' DEPENDENCIES
'   - TST_AchRunScenario
'   - TST_AchAssertClean
'   - TST_AchShowWindow
'   - TST_SafeCloseWorkbook
'
' NOTES
'   Raises TEST_ACH_ERR_BASE + TEST_ACH_ERR_SETUP when a fixture window cannot
'   be created or the first workbook cannot be made active.
'
' CALLED FROM
'   - TST_Case_AchievedWindowStateAllWindows
'   - TST_Case_AchievedWindowStateActiveWindow
'   - TST_Case_AchievedWindowStateActiveWorkbookWindows
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchRunWindowCase"

    Dim OriginalWindow  As Window   'Window active before the case
    Dim TargetBook      As Workbook 'Workbook with two windows
    Dim OtherBook       As Workbook 'Workbook with one window
    Dim TargetWindowOne As Window   'First window of the target workbook
    Dim TargetWindowTwo As Window   'Second window of the target workbook
    Dim OtherWindow     As Window   'Window of the other workbook
    Dim FailText        As String   'Failed observations of the three elements

    Dim HasFailure      As Boolean  'TRUE when a failure was captured
    Dim FailNumber      As Long     'Captured failure number
    Dim FailSource      As String   'Captured failure source
    Dim FailDescription As String   'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_Log CaseName, "START", _
            "Reading back Headings, Workbook Tabs and Gridlines under " & Scope

        Set OriginalWindow = Application.ActiveWindow

'------------------------------------------------------------------------------
' BUILD THE FIXTURE
'------------------------------------------------------------------------------
    'A workbook with two windows, and a workbook of its own, next to every
    'window already open
        Set TargetBook = Application.Workbooks.Add
        Set TargetWindowOne = TargetBook.Windows(1)
        Set TargetWindowTwo = TargetBook.NewWindow

        Set OtherBook = Application.Workbooks.Add
        Set OtherWindow = OtherBook.Windows(1)

        If TargetWindowOne Is Nothing Or _
            TargetWindowTwo Is Nothing Or _
            OtherWindow Is Nothing Then

            Err.Raise TEST_ACH_ERR_BASE + TEST_ACH_ERR_SETUP, CaseName, _
                "temporary workbook windows could not be created"
        End If

    'Every fixture window starts with the three elements shown, so that the
    'first hide owes a change
        TST_AchShowWindow TargetWindowOne, CaseName
        TST_AchShowWindow TargetWindowTwo, CaseName
        TST_AchShowWindow OtherWindow, CaseName

        TargetWindowOne.Activate

        If Not (Application.ActiveWorkbook Is TargetBook) Then
            Err.Raise TEST_ACH_ERR_BASE + TEST_ACH_ERR_SETUP, CaseName, _
                "temporary target workbook could not be activated"
        End If

'------------------------------------------------------------------------------
' OBSERVE THE THREE ELEMENTS
'------------------------------------------------------------------------------
    'Every element is observed before any failure is raised, so one run names
    'every element that did not read back
        FailText = TST_AchRunScenario("Headings", Scope, _
            TEST_ACH_SURFACE_WINDOW, "DisplayHeadings", ScopeValue)

        FailText = TST_CertAppendDetail(FailText, _
            TST_AchRunScenario("WorkbookTabs", Scope, _
            TEST_ACH_SURFACE_WINDOW, "DisplayWorkbookTabs", ScopeValue))

        FailText = TST_CertAppendDetail(FailText, _
            TST_AchRunScenario("Gridlines", Scope, _
            TEST_ACH_SURFACE_WINDOW, "DisplayGridlines", ScopeValue))

'------------------------------------------------------------------------------
' ASSERT THE READBACK
'------------------------------------------------------------------------------
        TST_AchAssertClean FailText, CaseName

        TST_Log CaseName, "PASS", _
            "Headings, Workbook Tabs and Gridlines read back as requested under " & _
            Scope & ", and every window outside it read back unchanged"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Best-effort cleanup of the temporary workbooks and the active window
        On Error Resume Next

        TST_SafeCloseWorkbook TargetBook
        TST_SafeCloseWorkbook OtherBook

        If Not OriginalWindow Is Nothing Then
            OriginalWindow.Activate
        End If

        On Error GoTo 0

    'Raise the captured failure after cleanup
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

Private Sub TST_AchShowWindow( _
    ByVal TargetWindow As Object, _
    ByVal CaseName As String)
'
'------------------------------------------------------------------------------
'                            SHOW A FIXTURE WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Show Headings, Workbook Tabs and Gridlines on one fixture window, without
'   the API under test.
'
' WHY THIS EXISTS
'   A window opened from another window can start with an element hidden, and
'   then the first hide of a scenario would owe no change (#28).
'
' INPUTS
'   TargetWindow
'     The fixture window.
'
'   CaseName
'     Error source.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Writes True to the three properties through the harness writer.
'
' ERROR POLICY
'   - Raises TEST_ACH_ERR_BASE + TEST_ACH_ERR_SETUP when a write fails.
'
' DEPENDENCIES
'   - TST_TrySetBooleanProperty
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchRunWindowCase
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchShowWindow"

    Dim PropertyName    As Variant  'One of the three properties
    Dim Msg             As String   'Diagnostic from the writer

'------------------------------------------------------------------------------
' SHOW THE THREE ELEMENTS
'------------------------------------------------------------------------------
        For Each PropertyName In Array("DisplayHeadings", _
            "DisplayWorkbookTabs", "DisplayGridlines")

            If Not TST_TrySetBooleanProperty(TargetWindow, CStr(PropertyName), _
                True, Msg) Then

                Err.Raise TEST_ACH_ERR_BASE + TEST_ACH_ERR_SETUP, CaseName, _
                    "a fixture window could not be shown (" & _
                    CStr(PropertyName) & "): " & Msg
            End If

        Next PropertyName

End Sub

Private Function TST_AchRunScenario( _
    ByVal Element As String, _
    ByVal Scope As String, _
    ByVal Surface As Long, _
    ByVal PropertyName As String, _
    ByVal ScopeValue As Long) _
    As String
'
'------------------------------------------------------------------------------
'                     OBSERVE ONE ELEMENT UNDER ONE SCOPE
'------------------------------------------------------------------------------
' PURPOSE
'   Run one scenario into the record of the certification run when one is
'   active, and into a local record otherwise.
'
' WHY THIS EXISTS
'   The regression packs are shared with the legacy runners, so recording is
'   inert outside a certification run: no count may leak from one invocation
'   into the next (#28). The choice is made here once, for every case.
'
' INPUTS
'   Element / Scope / Surface / PropertyName / ScopeValue
'     As TST_AchRunOn.
'
' RETURNS
'   String
'     The failed observations of the scenario, as text; empty when none failed.
'
' BEHAVIOR
'   - While a certification run is active, runs the scenario into
'     m_CertRecord.
'   - Otherwise empties the local record and runs the scenario into it.
'
' ERROR POLICY
'   - Does not raise, except what TST_AchRunOn raises.
'
' DEPENDENCIES
'   - TST_AchRunOn
'   - TST_CertResetRecord
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_Case_AchievedAppLevelState
'   - TST_Case_AchievedTitleBarFrameState
'   - TST_Case_AchievedRibbonOwningWindow
'   - TST_AchRunWindowCase
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN THE SCENARIO
'------------------------------------------------------------------------------
        If m_CertActive Then

            TST_AchRunScenario = TST_AchRunOn(m_CertRecord, Element, Scope, _
                Surface, PropertyName, ScopeValue)

        Else

            TST_CertResetRecord m_AchLocal

            TST_AchRunScenario = TST_AchRunOn(m_AchLocal, Element, Scope, _
                Surface, PropertyName, ScopeValue)

        End If

End Function

Private Function TST_AchRunOn( _
    ByRef Rec As tCertRecord, _
    ByVal Element As String, _
    ByVal Scope As String, _
    ByVal Surface As Long, _
    ByVal PropertyName As String, _
    ByVal ScopeValue As Long) _
    As String
'
'------------------------------------------------------------------------------
'                        RUN ONE SCENARIO INTO A RECORD
'------------------------------------------------------------------------------
' PURPOSE
'   Observe one element under one scope: four writes and the restoration of
'   the entry state, each read back and recorded.
'
' WHY THIS EXISTS
'   Hiding twice and showing twice verifies the path that changes the state and
'   the path that finds it already correct, whatever the entry state was, and
'   the restoration puts the host back and proves it (#28).
'
' INPUTS
'   Rec
'     ByRef. The record that receives the observations.
'
'   Element
'     StatusBar, ScrollBars, FormulaBar, Headings, WorkbookTabs, Gridlines,
'     TitleBar or Ribbon.
'
'   Scope
'     The scope label recorded with each observation.
'
'   Surface
'     The TEST_ACH_SURFACE_* the element is read and restored through.
'
'   PropertyName
'     The Boolean property of an application or window surface; empty otherwise.
'
'   ScopeValue
'     The UIWindowTargetScope the setter is called with.
'
' RETURNS
'   String
'     The failed observations of the scenario, as text; empty when none failed.
'
' BEHAVIOR
'   - Begins the scenario. When it cannot begin, records that as one
'     observation: not-verifiable for a surface that cannot be addressed or
'     read, readback-failed otherwise.
'   - Records one not-verifiable observation for the windows that could not be
'     read on entry.
'   - Hides, hides again, shows and shows again, then restores.
'
' ERROR POLICY
'   - Does not raise for an observation that fails: it is recorded and returned.
'
' DEPENDENCIES
'   - TST_AchBegin
'   - TST_AchStep
'   - TST_AchRestore
'   - TST_AchRecordBeginFailure
'   - TST_AchRecordObservation
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchRunScenario
'   - TST_Case_AchievedMismatchFailsTheUnit
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchRunOn"

    Dim Why As String   'Why the scenario could not begin

'------------------------------------------------------------------------------
' RUN THE SCENARIO
'------------------------------------------------------------------------------
        If TST_AchBegin(Element, Scope, Surface, PropertyName, ScopeValue, Why) Then

        'Windows that could not be read are named, not left out in silence
            If Len(m_AchSkipped) > 0 Then
                TST_AchRecordObservation Rec, "entry", "not attempted", _
                    "unreadable", TEST_ACH_OUTCOME_UNVERIFIABLE, _
                    "windows that could not be read on entry were not verified: " & _
                    m_AchSkipped
            End If

            TST_AchStep Rec, "hide", False
            TST_AchStep Rec, "hide", False
            TST_AchStep Rec, "show", True
            TST_AchStep Rec, "show", True

            TST_AchRestore Rec

        Else

            TST_AchRecordBeginFailure Rec, Why

        End If

        TST_AchRunOn = m_AchFailText

End Function

Private Sub TST_AchRecordBeginFailure( _
    ByRef Rec As tCertRecord, _
    ByVal Why As String)
'
'------------------------------------------------------------------------------
'                    RECORD A SCENARIO THAT COULD NOT BEGIN
'------------------------------------------------------------------------------
' PURPOSE
'   Record one observation that says why a scenario could not begin.
'
' WHY THIS EXISTS
'   A surface that cannot be addressed or read before the first write is
'   named in the evidence, not skipped. The Ribbon alone is not-verifiable, as
'   the manifest allows; any other surface that cannot be read, a title bar
'   whose entry frame has no complete caption included, is readback-failed and
'   fails the run (#28).
'
' INPUTS
'   Rec
'     ByRef. The record that receives the observation.
'
'   Why
'     The reason TST_AchBegin gave.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Records an "entry" observation, with the outcome not-verifiable when
'     m_AchUnverifiable is set and readback-failed otherwise.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_AchRecordObservation
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchRunOn
'   - TST_AchUnverifiableSeam
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchRecordBeginFailure"

    Dim Outcome As String   'Outcome of the observation

'------------------------------------------------------------------------------
' RECORD THE OBSERVATION
'------------------------------------------------------------------------------
    'Not verifiable is not a failure of the code under test; anything else that
    'cannot be read is
        If m_AchUnverifiable Then
            Outcome = TEST_ACH_OUTCOME_UNVERIFIABLE
        Else
            Outcome = TEST_ACH_OUTCOME_READ_FAILED
        End If

        TST_AchRecordObservation Rec, "entry", "not attempted", "unreadable", _
            Outcome, Why

End Sub

Private Function TST_AchBegin( _
    ByVal Element As String, _
    ByVal Scope As String, _
    ByVal Surface As Long, _
    ByVal PropertyName As String, _
    ByVal ScopeValue As Long, _
    ByRef Why As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                         BEGIN OBSERVING ONE ELEMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Establish the identity of what will be written, read its entry state and
'   keep both for the steps and the restoration of the scenario.
'
' WHY THIS EXISTS
'   A readback means something only when it reads what the write targeted. The
'   windows are retained as objects and the frame and the Ribbon owner are
'   named by their handle before the first write, so a later read can be
'   refused when it would read something else (#45, #28).
'
' INPUTS
'   Element / Scope / Surface / PropertyName / ScopeValue
'     As TST_AchRunOn.
'
'   Why
'     ByRef. Receives the reason the scenario cannot begin; empty on success.
'
' RETURNS
'   Boolean
'     TRUE  => the surface was resolved and its entry state read.
'     FALSE => it was not; m_AchUnverifiable says whether that is a surface that
'              cannot be verified.
'
' BEHAVIOR
'   - Clears the scenario state, then resolves the surface:
'     application   one entry, the Application property.
'     window        every readable window as an entry, retained as an object;
'                   an entry is in the scope when it is the active window, a
'                   window of the active workbook, or any window.
'     title bar     the active frame pair, its handle and its entry style; a
'                   frame with no complete caption is not the baseline and
'                   fails (the fault TEST_ACH_FAULT_NO_CAPTION makes any frame
'                   read so).
'     Ribbon        the active window and its handle; none cannot be verified.
'   - Reads the entry state of every entry.
'
' ERROR POLICY
'   - Does not raise. An error is the reason it returns FALSE.
'
' DEPENDENCIES
'   - TST_TryGetBooleanProperty
'   - TST_TryGetWindowStyle
'   - TST_StyleHasCompleteCaption
'   - UI_TryGetActiveFramePair
'   - TST_AchRead
'
' NOTES
'   Windows are matched by the handle each object reports, never with Is,
'   because Excel returns a new wrapper from every read of
'   Application.ActiveWindow.
'
' CALLED FROM
'   - TST_AchRunOn
'   - TST_AchSeam
'   - TST_AchRestoreSeam
'   - TST_AchUnverifiableSeam
'   - TST_Case_AchievedMismatchFailsTheUnit
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchBegin"

    Dim WindowIdx   As Long     'Cursor over Application.Windows
    Dim WindowTotal As Long     'Windows open
    Dim Candidate   As Object   'Window under consideration
    Dim BookWindow  As Object   'Window of the active workbook
    Dim ActiveWin   As Object   'The active window
    Dim PairWindow  As Object   'Window of the active frame pair
    Dim ActiveKey   As String   'Key of the active window
    Dim BookKeys    As String   'Keys of the windows of the active workbook, between pipes
    Dim WindowKey   As String   'Key of one window
    Dim Probe       As Boolean  'Value read while probing a window
    Dim AnyTarget   As Boolean  'TRUE when the scope holds a readable window
    Dim Captionless As Boolean  'TRUE when the entry frame has no complete caption
    Dim EntryReadOK As Boolean  'TRUE when the entry state was read
    Dim Detail      As String   'Reason from a helper

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_AchBegin = False
        Why = vbNullString

        m_AchElement = Element
        m_AchScope = Scope
        m_AchSurface = Surface
        m_AchProperty = PropertyName
        m_AchScopeValue = ScopeValue
        m_AchCount = 0
        m_AchSkipped = vbNullString
        m_AchUnverifiable = False
        m_AchFrameCaptured = False
        m_AchNote = vbNullString
        m_AchSetterReturned = False
        m_AchFailCount = 0
        m_AchFailText = vbNullString
        m_AchHwnd = 0
        m_AchStyle = 0
        m_AchEntryStyle = 0

'------------------------------------------------------------------------------
' RESOLVE THE SURFACE
'------------------------------------------------------------------------------
        Select Case Surface

            Case TEST_ACH_SURFACE_APP
                m_AchCount = 1

                ReDim m_AchLabel(1 To 1)
                ReDim m_AchTarget(1 To 1)

                m_AchLabel(1) = Element
                m_AchTarget(1) = True

            Case TEST_ACH_SURFACE_WINDOW

            'Every window open is part of the observation, in scope or not
                WindowTotal = Application.Windows.Count

                If WindowTotal = 0 Then
                    Why = "no window is open"
                    GoTo SafeExit
                End If

                ReDim m_AchWin(1 To WindowTotal)
                ReDim m_AchLabel(1 To WindowTotal)
                ReDim m_AchTarget(1 To WindowTotal)

                Select Case ScopeValue

                    Case UI_TargetActiveWindow
                        Set ActiveWin = Application.ActiveWindow
                        ActiveKey = TST_AchWindowKey(ActiveWin)

                        If Len(ActiveKey) = 0 Then
                            Why = "the active window cannot be identified"
                            GoTo SafeExit
                        End If

                    Case UI_TargetActiveWorkbookWindows
                        If Application.ActiveWorkbook Is Nothing Then
                            Why = "the active workbook is unavailable"
                            GoTo SafeExit
                        End If

                        BookKeys = "|"

                        For Each BookWindow In Application.ActiveWorkbook.Windows
                            BookKeys = BookKeys & TST_AchWindowKey(BookWindow) & "|"
                        Next BookWindow

                    Case Else
                    'Every window is in the scope

                End Select

                For WindowIdx = 1 To WindowTotal

                    Set Candidate = Application.Windows(WindowIdx)

                    If TST_TryGetBooleanProperty(Candidate, PropertyName, Probe, Detail) Then

                        m_AchCount = m_AchCount + 1

                        Set m_AchWin(m_AchCount) = Candidate

                        WindowKey = TST_AchWindowKey(Candidate)

                        m_AchLabel(m_AchCount) = "#" & CStr(WindowIdx) & " " & _
                            TST_AchWindowCaption(Candidate)

                        Select Case ScopeValue

                            Case UI_TargetActiveWindow
                                m_AchTarget(m_AchCount) = _
                                    (Len(WindowKey) > 0) And (WindowKey = ActiveKey)

                            Case UI_TargetActiveWorkbookWindows
                                m_AchTarget(m_AchCount) = _
                                    (Len(WindowKey) > 0) And _
                                    (InStr(1, BookKeys, "|" & WindowKey & "|", _
                                        vbBinaryCompare) > 0)

                            Case Else
                                m_AchTarget(m_AchCount) = True

                        End Select

                        If m_AchTarget(m_AchCount) Then
                            AnyTarget = True
                        End If

                    Else

                        m_AchSkipped = TST_CertAppendDetail(m_AchSkipped, _
                            "#" & CStr(WindowIdx) & " " & Detail)

                    End If

                Next WindowIdx

                If m_AchCount = 0 Then
                    Why = "no window could be read: " & m_AchSkipped
                    GoTo SafeExit
                End If

                If Not AnyTarget Then
                    Why = "the scope holds no window that could be read"
                    GoTo SafeExit
                End If

                ReDim Preserve m_AchWin(1 To m_AchCount)
                ReDim Preserve m_AchLabel(1 To m_AchCount)
                ReDim Preserve m_AchTarget(1 To m_AchCount)

            Case TEST_ACH_SURFACE_TITLEBAR
                m_AchCount = 1

                ReDim m_AchLabel(1 To 1)
                ReDim m_AchTarget(1 To 1)

                m_AchLabel(1) = Element
                m_AchTarget(1) = True

            'The frame the write will target, named before it: the production pair
                If Not UI_TryGetActiveFramePair(PairWindow, m_AchHwnd, Detail) Then
                    Why = "the active frame pair could not be resolved | " & Detail
                    GoTo SafeExit
                End If

                If Not TST_TryGetWindowStyle(m_AchHwnd, m_AchEntryStyle, Detail) Then
                    Why = "GWL_STYLE of frame " & CStr(m_AchHwnd) & _
                        " could not be read | " & Detail
                    GoTo SafeExit
                End If

                m_AchFrameCaptured = True

            'A show restores a frame that had a caption. An entry frame without one
            'is not the baseline the unit verifies against, and the manifest allows
            'only the Ribbon to be not-verifiable, so it fails the observation
                Captionless = Not TST_StyleHasCompleteCaption(m_AchEntryStyle)

                If m_AchFault = TEST_ACH_FAULT_NO_CAPTION Then
                    m_AchFault = TEST_ACH_FAULT_NONE
                    Captionless = True
                End If

                If Captionless Then
                    Why = "frame " & CStr(m_AchHwnd) & " has no complete caption " & _
                        "(WS_BORDER and WS_DLGFRAME): its title bar is hidden or " & _
                        "was never a caption, so the host is not the baseline " & _
                        "that a hide and a show are verified against"
                    GoTo SafeExit
                End If

            Case TEST_ACH_SURFACE_RIBBON
                m_AchCount = 1

                ReDim m_AchLabel(1 To 1)
                ReDim m_AchTarget(1 To 1)

                m_AchLabel(1) = Element
                m_AchTarget(1) = True

            'The Ribbon acts on the active window only, so that window is its owner
                Set ActiveWin = Application.ActiveWindow

                If Not TST_AchTryReadHwnd(ActiveWin, m_AchHwnd, Detail) Then
                    m_AchUnverifiable = True
                    Why = "no active window can be identified as the owner of " & _
                        "the Ribbon | " & Detail
                    GoTo SafeExit
                End If

            Case Else
                Why = "unknown surface " & CStr(Surface)
                GoTo SafeExit

        End Select

'------------------------------------------------------------------------------
' READ THE ENTRY STATE
'------------------------------------------------------------------------------
        ReDim m_AchEntry(1 To m_AchCount)

    'The readback fault reaches the Ribbon here, so that it is unreadable from
    'the start: the one case in which the Ribbon is exempt
        If m_AchFault = TEST_ACH_FAULT_READBACK And _
            Surface = TEST_ACH_SURFACE_RIBBON Then

            m_AchFault = TEST_ACH_FAULT_NONE
            EntryReadOK = False
            Detail = "injected readback failure"
        Else
            EntryReadOK = TST_AchRead(m_AchEntry, Detail)
        End If

        If Not EntryReadOK Then
            m_AchUnverifiable = (Surface = TEST_ACH_SURFACE_RIBBON)
            Why = "the state could not be read before the first write | " & Detail
            GoTo SafeExit
        End If

        TST_AchBegin = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        Why = TST_BuildRuntimeErrorText

        TST_AchBegin = False

        Resume SafeExit

End Function

Private Sub TST_AchStep( _
    ByRef Rec As tCertRecord, _
    ByVal Operation As String, _
    ByVal Requested As Boolean)
'
'------------------------------------------------------------------------------
'                     WRITE, SETTLE, READ BACK AND COMPARE
'------------------------------------------------------------------------------
' PURPOSE
'   Make one write through the public API and record what was achieved.
'
' WHY THIS EXISTS
'   This is the verification model of #28: write, settle, read, compare. The
'   comparison is against what the scope owes, not against what the setter
'   said: every entry in the scope must read back at the requested state and
'   every other entry must read back as it was before the write.
'
' INPUTS
'   Rec
'     ByRef. The record that receives the observation.
'
'   Operation
'     hide or show, recorded with the observation.
'
'   Requested
'     TRUE to show, FALSE to hide.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Takes the armed fault, when it applies to a step.
'   - Reads the state before the write. When that fails the outcome is
'     readback-failed: the surface was readable when the scenario began, so
'     there is no exemption mid-sequence, the Ribbon included.
'   - Works out the state each entry must read back at, and whether any entry
'     owes a change.
'   - Writes through TST_AchWrite: the requested visibility, nothing for an
'     ignored write, or a value no member names for a failed write.
'   - Settles once when the setter reported success, then reads back, or fails
'     the read for a readback fault.
'   - Records write-failed when the setter reported a failure; otherwise
'     readback-failed when the read failed;
'     otherwise state-mismatch when any entry differs from what it owes;
'     otherwise verified-write when a change was owed and verified-no-op when
'     none was.
'
' ERROR POLICY
'   - Does not raise. A failed observation is recorded, counted and kept as
'     text for the caller.
'
' DEPENDENCIES
'   - TST_AchTakeFault
'   - TST_AchRead
'   - TST_AchWrite
'   - TST_AchSettle
'   - TST_AchText
'   - TST_AchDifferences
'   - TST_AchRecordObservation
'
' NOTES
'   The setter is the production one, called with exactly one element and the
'   scope under test. Nothing here changes it.
'
' CALLED FROM
'   - TST_AchRunOn
'   - TST_AchSeam
'   - TST_AchRestoreSeam
'   - TST_AchUnverifiableSeam
'   - TST_Case_AchievedMismatchFailsTheUnit
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchStep"

    Dim Before()    As Boolean  'State of every entry before the write
    Dim Expected()  As Boolean  'State every entry must read back at
    Dim After()     As Boolean  'State of every entry after the write
    Dim EntryIdx    As Long     'Cursor over the entries
    Dim Fault       As Long     'Fault taken for this step
    Dim Visibility  As Long     'UIVisibility passed to the setter
    Dim Changed     As Boolean  'TRUE when an entry in the scope owes a change
    Dim WriteOK     As Boolean  'What the setter returned
    Dim WriteText   As String   'The setter's failure text
    Dim ReadOK      As Boolean  'TRUE when the readback succeeded
    Dim Why         As String   'Why a read failed
    Dim Differing   As String   'Entries that differ from what they owe
    Dim Outcome     As String   'Outcome of the observation
    Dim Reason      As String   'Why the outcome is not a verified one
    Dim ReqText     As String   'Requested state as text
    Dim ObsText     As String   'Observed state as text

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        Fault = TST_AchTakeFault(False)

        ReqText = IIf(Requested, "visible", "hidden") & _
            IIf(m_AchCount > 1, " in the scope, the other windows unchanged", _
                vbNullString)

        ObsText = "unreadable"

        ReDim Before(1 To m_AchCount)
        ReDim Expected(1 To m_AchCount)
        ReDim After(1 To m_AchCount)

'------------------------------------------------------------------------------
' WRITE, SETTLE, READ AND COMPARE
'------------------------------------------------------------------------------
        If TST_AchRead(Before, Why) Then

        'What every entry owes: the request in the scope, the entry state outside
            For EntryIdx = 1 To m_AchCount

                If m_AchTarget(EntryIdx) Then
                    Expected(EntryIdx) = Requested

                    If Before(EntryIdx) <> Requested Then
                        Changed = True
                    End If
                Else
                    Expected(EntryIdx) = Before(EntryIdx)
                End If

            Next EntryIdx

        'The write, through the public API
            If Requested Then
                Visibility = UI_Show
            Else
                Visibility = UI_Hide
            End If

            Select Case Fault

                Case TEST_ACH_FAULT_IGNORED_WRITE
                    Visibility = UI_LeaveUnchanged

                Case TEST_ACH_FAULT_WRITE_FAILED
                    Visibility = TEST_ACH_INVALID_VISIBILITY

                Case Else
                'Write the requested visibility

            End Select

            WriteOK = TST_AchWrite(Visibility, WriteText)

            m_AchSetterReturned = WriteOK

        'Settle within the bound, then read back once more and decide
            If Fault = TEST_ACH_FAULT_READBACK Then
                ReadOK = False
                Why = "injected readback failure"
            Else
                If WriteOK Then
                    TST_AchSettle Expected
                End If

                ReadOK = TST_AchRead(After, Why)
            End If

            If ReadOK Then
                ObsText = TST_AchText(After, True)
            End If

        'Each failure is its own outcome with its own reason
            If Not WriteOK Then

                Outcome = TEST_ACH_OUTCOME_WRITE_FAILED
                Reason = "the setter reported a failure: " & WriteText

            ElseIf Not ReadOK Then

            'The surface was readable when the scenario began, so this is a
            'failure of the readback, whatever the surface
                Outcome = TEST_ACH_OUTCOME_READ_FAILED
                Reason = "the state could not be read after the write | " & Why

            Else

                Differing = TST_AchDifferences(Expected, After)

                If Len(Differing) > 0 Then
                    Outcome = TEST_ACH_OUTCOME_MISMATCH
                    Reason = "the setter returned True, but the state read back " & _
                        "differs from the request: " & Differing
                ElseIf Changed Then
                    Outcome = TEST_ACH_OUTCOME_WRITE
                Else
                    Outcome = TEST_ACH_OUTCOME_NOOP
                End If

            End If

        Else

        'Nothing can be said about a write whose entry state is not known, and
        'the surface was readable when the scenario began
            Outcome = TEST_ACH_OUTCOME_READ_FAILED
            Reason = "the state could not be read before the write | " & Why

        End If

'------------------------------------------------------------------------------
' RECORD THE OBSERVATION
'------------------------------------------------------------------------------
        TST_AchRecordObservation Rec, Operation, ReqText, ObsText, Outcome, Reason

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'An unexpected error is an observation that could not be made
        Why = TST_BuildRuntimeErrorText

        TST_AchRecordObservation Rec, Operation, ReqText, ObsText, _
            TEST_ACH_OUTCOME_READ_FAILED, "an unexpected error ended the step | " & Why

        Resume SafeExit

End Sub

Private Sub TST_AchRestore( _
    ByRef Rec As tCertRecord)
'
'------------------------------------------------------------------------------
'                   RESTORE THE ENTRY STATE AND READ IT BACK
'------------------------------------------------------------------------------
' PURPOSE
'   Put the entry state back through the harness writers, never through the API
'   under test, and record whether it read back as it was.
'
' WHY THIS EXISTS
'   Every test restores its entry state and reports a restoration that did not
'   take (#28). A restoration through the API under test would hide the very
'   defect the scenario may have found.
'
' INPUTS
'   Rec
'     ByRef. The record that receives the observation.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Takes the armed fault, when it is a restoration fault: the restoring write
'     is not made, or the readback after it fails.
'   - Reads the state. When it cannot be read the outcome is restore-failed,
'     the Ribbon included: it was readable when the scenario began.
'   - When the state differs from the entry state, or the owned frame bits do,
'     writes the entry state back unless the restoration fault is armed, which
'     makes no write.
'   - Settles, reads back, and records restore-failed when anything still
'     differs, verified-write when a write was owed and verified-no-op when
'     nothing was.
'
' ERROR POLICY
'   - Does not raise. A restoration that did not take is recorded and kept as
'     text for the caller.
'
' DEPENDENCIES
'   - TST_AchTakeFault
'   - TST_AchRead
'   - TST_AchNeedsRestore
'   - TST_AchRawRestore
'   - TST_AchSettle
'   - TST_AchRecordObservation
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchRunOn
'   - TST_AchSeam
'   - TST_AchRestoreSeam
'   - TST_AchUnverifiableSeam
'   - TST_Case_AchievedMismatchFailsTheUnit
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchRestore"

    Dim Current()   As Boolean  'State before the restoration
    Dim Final()     As Boolean  'State after the restoration
    Dim Fault       As Long     'Fault taken for this restoration
    Dim Needed      As Boolean  'TRUE when the host differs from the entry state
    Dim FinalOK     As Boolean  'TRUE when the readback after the restoration succeeded
    Dim WriteOK     As Boolean  'TRUE when every restoring write succeeded
    Dim Why         As String   'Why a read failed
    Dim WriteText   As String   'Why a restoring write failed
    Dim Differing   As String   'Entries that differ from the entry state
    Dim Outcome     As String   'Outcome of the observation
    Dim Reason      As String   'Why the outcome is not a verified one
    Dim ReqText     As String   'Entry state as text
    Dim ObsText     As String   'State read back as text

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        Fault = TST_AchTakeFault(True)

        ReqText = TST_AchText(m_AchEntry, False)
        ObsText = "unreadable"

        ReDim Current(1 To m_AchCount)
        ReDim Final(1 To m_AchCount)

'------------------------------------------------------------------------------
' RESTORE AND READ BACK
'------------------------------------------------------------------------------
        If TST_AchRead(Current, Why) Then

            Needed = TST_AchNeedsRestore(Current)
            WriteOK = True

        'The restoration write is not made under the restoration fault
            If Needed And Fault <> TEST_ACH_FAULT_RESTORE Then
                WriteOK = TST_AchRawRestore(Current, WriteText)
            End If

            TST_AchSettle m_AchEntry

            If Fault = TEST_ACH_FAULT_RESTORE_READ Then
                FinalOK = False
                Why = "injected readback failure"
            Else
                FinalOK = TST_AchRead(Final, Why)
            End If

            If FinalOK Then

                ObsText = TST_AchText(Final, True)

                If TST_AchNeedsRestore(Final) Then

                    Differing = TST_AchDifferences(m_AchEntry, Final)

                    If Len(Differing) = 0 Then
                        Differing = "the owned frame style bits differ from the entry style"
                    End If

                    Outcome = TEST_ACH_OUTCOME_RESTORE_FAILED
                    Reason = "the entry state was not put back: " & Differing & _
                        IIf(Len(WriteText) > 0, " | " & WriteText, vbNullString)

                ElseIf Needed Then
                    Outcome = TEST_ACH_OUTCOME_WRITE
                Else
                    Outcome = TEST_ACH_OUTCOME_NOOP
                End If

            Else

            'A surface that was readable and no longer is cannot be confirmed
            'restored, and the Ribbon is no exception
                Outcome = TEST_ACH_OUTCOME_RESTORE_FAILED
                Reason = "the entry state could not be read after it was put back | " & Why

            End If

        Else

            Outcome = TEST_ACH_OUTCOME_RESTORE_FAILED
            Reason = "the entry state could not be read to put it back | " & Why

        End If

'------------------------------------------------------------------------------
' RECORD THE OBSERVATION
'------------------------------------------------------------------------------
        TST_AchRecordObservation Rec, "restore", ReqText, ObsText, Outcome, Reason

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'An unexpected error is a restoration that cannot be confirmed
        Why = TST_BuildRuntimeErrorText

        TST_AchRecordObservation Rec, "restore", ReqText, ObsText, _
            TEST_ACH_OUTCOME_RESTORE_FAILED, _
            "an unexpected error ended the restoration | " & Why

        Resume SafeExit

End Sub

'
'------------------------------------------------------------------------------
'
'                            READ, WRITE AND SETTLE
'
'------------------------------------------------------------------------------
'

Private Function TST_AchRead( _
    ByRef Values() As Boolean, _
    ByRef Why As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                        READ THE STATE OF THE SURFACE
'------------------------------------------------------------------------------
' PURPOSE
'   Read the state of every entry of the scenario, independently of the setter.
'
' WHY THIS EXISTS
'   The readback is the evidence. It reads through the harness readers and the
'   object model, never through the production getters or setters, and it
'   refuses a read that does not belong to what the write targeted (#45, #28).
'
' INPUTS
'   Values
'     ByRef. 1 To m_AchCount. Receives the state of each entry: TRUE when it is
'     visible.
'
'   Why
'     ByRef. Receives the reason a read failed; empty on success.
'
' RETURNS
'   Boolean
'     TRUE  => every entry was read and belongs to the target.
'     FALSE => at least one was not; Why says which and why.
'
' BEHAVIOR
'   - application    reads the Application property.
'   - window         reads the property of each retained window.
'   - title bar      resolves the active frame pair again and refuses the read
'                    when its handle is not the one established before the
'                    write, reads that frame's GWL_STYLE, and takes the frame as
'                    visible only when WS_BORDER and WS_DLGFRAME are both set.
'   - Ribbon         reads the Ribbon, then the active window's handle, and
'                    refuses the value when that is not the window that owned the
'                    Ribbon before the write.
'
' ERROR POLICY
'   - Does not raise. An error is the reason it returns FALSE.
'
' DEPENDENCIES
'   - TST_TryGetBooleanProperty
'   - TST_TryGetWindowStyle
'   - TST_StyleHasCompleteCaption
'   - TST_TryGetRibbonVisible
'   - UI_TryGetActiveFramePair
'   - TST_AchTryReadHwnd
'
' NOTES
'   A title-bar read leaves the owned bits of the style in m_AchNote, so that
'   the observed text shows them.
'
' CALLED FROM
'   - TST_AchBegin
'   - TST_AchStep
'   - TST_AchRestore
'   - TST_AchSettle
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchRead"

    Dim EntryIdx    As Long     'Cursor over the entries
    Dim Detail      As String   'Reason from a helper
    Dim PairWindow  As Object   'Window of the active frame pair
    Dim ActiveWin   As Object   'Active window read after the Ribbon

#If VBA7 Then
    Dim PairHwnd    As LongPtr  'Handle of the active frame pair
    Dim ActiveHwnd  As LongPtr  'Handle of the active window
    Dim StyleValue  As LongPtr  'GWL_STYLE of the frame
#Else
    Dim PairHwnd    As Long 'Handle of the active frame pair
    Dim ActiveHwnd  As Long 'Handle of the active window
    Dim StyleValue  As Long 'GWL_STYLE of the frame
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_AchRead = False
        Why = vbNullString
        m_AchNote = vbNullString

'------------------------------------------------------------------------------
' READ THE SURFACE
'------------------------------------------------------------------------------
        Select Case m_AchSurface

            Case TEST_ACH_SURFACE_APP
                If Not TST_TryGetBooleanProperty(Application, m_AchProperty, _
                    Values(1), Detail) Then

                    Why = "Application." & m_AchProperty & " could not be read | " & Detail
                    GoTo SafeExit
                End If

            Case TEST_ACH_SURFACE_WINDOW
                For EntryIdx = 1 To m_AchCount

                    If Not TST_TryGetBooleanProperty(m_AchWin(EntryIdx), _
                        m_AchProperty, Values(EntryIdx), Detail) Then

                        Why = m_AchLabel(EntryIdx) & "." & m_AchProperty & _
                            " could not be read | " & Detail
                        GoTo SafeExit
                    End If

                Next EntryIdx

            Case TEST_ACH_SURFACE_TITLEBAR

            'The frame must still be the one the write targeted
                If Not UI_TryGetActiveFramePair(PairWindow, PairHwnd, Detail) Then
                    Why = "the active frame pair could not be resolved | " & Detail
                    GoTo SafeExit
                End If

                If PairHwnd <> m_AchHwnd Then
                    Why = "the active frame is " & CStr(PairHwnd) & ", not frame " & _
                        CStr(m_AchHwnd) & " that the write targeted"
                    GoTo SafeExit
                End If

                If Not TST_TryGetWindowStyle(m_AchHwnd, StyleValue, Detail) Then
                    Why = "GWL_STYLE of frame " & CStr(m_AchHwnd) & _
                        " could not be read | " & Detail
                    GoTo SafeExit
                End If

            'Visible only with WS_BORDER and WS_DLGFRAME both set
                Values(1) = TST_StyleHasCompleteCaption(StyleValue)

                m_AchStyle = StyleValue

                m_AchNote = " [owned style bits &H" & _
                    Hex$(StyleValue And TST_TITLEBAR_OWNED_MASK) & "]"

            Case TEST_ACH_SURFACE_RIBBON

                If Not TST_TryGetRibbonVisible(Values(1), Detail) Then
                    Why = "the Ribbon could not be read | " & Detail
                    GoTo SafeExit
                End If

            'The value is attributable only to the window that was active when it
            'was read, and only if that window did not change during the read
                Set ActiveWin = Application.ActiveWindow

                If Not TST_AchTryReadHwnd(ActiveWin, ActiveHwnd, Detail) Then
                    Why = "the active window could not be identified after the " & _
                        "read | " & Detail
                    GoTo SafeExit
                End If

                If ActiveHwnd <> m_AchHwnd Then
                    Why = "the active window changed while the Ribbon was read, " & _
                        "so the value is not attributable to window " & CStr(m_AchHwnd)
                    GoTo SafeExit
                End If

            Case Else
                Why = "no surface is selected"
                GoTo SafeExit

        End Select

        TST_AchRead = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        Why = TST_BuildRuntimeErrorText

        TST_AchRead = False

        Resume SafeExit

End Function

Private Function TST_AchWrite( _
    ByVal Visibility As Long, _
    ByRef Detail As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                         WRITE THROUGH THE PUBLIC API
'------------------------------------------------------------------------------
' PURPOSE
'   Ask UI_SetExcelUI_WithResult for one element, under the scope of the
'   scenario, and report what it returned.
'
' WHY THIS EXISTS
'   The write under verification is the production one, called as a caller
'   calls it. Passing exactly one element keeps every other element out of the
'   comparison (#28).
'
' INPUTS
'   Visibility
'     The UIVisibility value for the element, or a value no member names for the
'     failed-write seam.
'
'   Detail
'     ByRef. Receives the failure list the setter returned when it reported a
'     failure; empty otherwise.
'
' RETURNS
'   Boolean
'     What UI_SetExcelUI_WithResult returned.
'
' BEHAVIOR
'   - Calls UI_SetExcelUI_WithResult with Visibility for the element of the
'     scenario, and UI_LeaveUnchanged for every other element.
'   - When it returns False, describes its failure count and list in Detail.
'
' ERROR POLICY
'   - Does not raise. An error is reported as a failure of the write.
'
' DEPENDENCIES
'   - UI_SetExcelUI_WithResult
'   - TST_DescribeResultBuffers
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchStep
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchWrite"

    Dim Succeeded       As Boolean  'What the setter returned
    Dim FailureCount    As Long     'Structured failure count
    Dim FailureList     As Variant  'Structured failure list

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_AchWrite = False
        Detail = vbNullString

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        Select Case m_AchElement

            Case "Ribbon"
                Succeeded = UI_SetExcelUI_WithResult( _
                    Ribbon:=Visibility, _
                    FailureCount:=FailureCount, _
                    FailureList:=FailureList, _
                    TargetScope:=m_AchScopeValue)

            Case "StatusBar"
                Succeeded = UI_SetExcelUI_WithResult( _
                    StatusBar:=Visibility, _
                    FailureCount:=FailureCount, _
                    FailureList:=FailureList, _
                    TargetScope:=m_AchScopeValue)

            Case "ScrollBars"
                Succeeded = UI_SetExcelUI_WithResult( _
                    ScrollBars:=Visibility, _
                    FailureCount:=FailureCount, _
                    FailureList:=FailureList, _
                    TargetScope:=m_AchScopeValue)

            Case "FormulaBar"
                Succeeded = UI_SetExcelUI_WithResult( _
                    FormulaBar:=Visibility, _
                    FailureCount:=FailureCount, _
                    FailureList:=FailureList, _
                    TargetScope:=m_AchScopeValue)

            Case "Headings"
                Succeeded = UI_SetExcelUI_WithResult( _
                    Headings:=Visibility, _
                    FailureCount:=FailureCount, _
                    FailureList:=FailureList, _
                    TargetScope:=m_AchScopeValue)

            Case "WorkbookTabs"
                Succeeded = UI_SetExcelUI_WithResult( _
                    WorkbookTabs:=Visibility, _
                    FailureCount:=FailureCount, _
                    FailureList:=FailureList, _
                    TargetScope:=m_AchScopeValue)

            Case "Gridlines"
                Succeeded = UI_SetExcelUI_WithResult( _
                    Gridlines:=Visibility, _
                    FailureCount:=FailureCount, _
                    FailureList:=FailureList, _
                    TargetScope:=m_AchScopeValue)

            Case "TitleBar"
                Succeeded = UI_SetExcelUI_WithResult( _
                    TitleBar:=Visibility, _
                    FailureCount:=FailureCount, _
                    FailureList:=FailureList, _
                    TargetScope:=m_AchScopeValue)

            Case Else
                Detail = "unknown element " & m_AchElement
                GoTo SafeExit

        End Select

        If Not Succeeded Then
            Detail = TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

        TST_AchWrite = Succeeded

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        Detail = TST_BuildRuntimeErrorText

        TST_AchWrite = False

        Resume SafeExit

End Function

Private Sub TST_AchSettle( _
    ByRef Expected() As Boolean)
'
'------------------------------------------------------------------------------
'                      WAIT, WITHIN A BOUND, FOR A STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Give Excel and Windows the chance to reach a state, for at most a fixed
'   number of short waits.
'
' WHY THIS EXISTS
'   Some changes, the Ribbon and the title bar above all, are a little
'   asynchronous. Every wait of this unit goes through here, so the bound is one
'   number, and no case hides a race with a delay of its own (#28).
'
' INPUTS
'   Expected
'     ByRef. The state each entry should read back at.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Yields once, then reads the state up to TEST_ACH_POLL_LIMIT times,
'     TEST_ACH_POLL_SECONDS apart, and stops at the first read that equals
'     Expected or that fails.
'
' ERROR POLICY
'   - Does not raise. Reaching the end of the bound is never a pass: the caller
'     reads the state again and decides.
'
' DEPENDENCIES
'   - TST_AchRead
'   - TST_AchSame
'   - TST_WaitUI
'
' NOTES
'   The longest settle is TEST_ACH_POLL_LIMIT times TEST_ACH_POLL_SECONDS.
'
' CALLED FROM
'   - TST_AchStep
'   - TST_AchRestore
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchSettle"

    Dim Current()   As Boolean  'State read on one poll
    Dim PollIdx     As Long     'Cursor over the polls
    Dim Why         As String   'Why a read failed

'------------------------------------------------------------------------------
' POLL
'------------------------------------------------------------------------------
        ReDim Current(1 To m_AchCount)

        DoEvents

        For PollIdx = 1 To TEST_ACH_POLL_LIMIT

            If Not TST_AchRead(Current, Why) Then
                Exit For
            End If

            If TST_AchSame(Expected, Current) Then
                Exit For
            End If

            TST_WaitUI TEST_ACH_POLL_SECONDS

        Next PollIdx

End Sub

Private Function TST_AchTakeFault( _
    ByVal ForRestore As Boolean) _
    As Long
'
'------------------------------------------------------------------------------
'                             TAKE THE ARMED FAULT
'------------------------------------------------------------------------------
' PURPOSE
'   Return the armed fault when it applies to what is about to run, and disarm
'   it.
'
' WHY THIS EXISTS
'   A fault is one-shot, and a restoration fault must not be spent by the step
'   that runs before it (#28).
'
' INPUTS
'   ForRestore
'     TRUE for a restoration, FALSE for a step.
'
' RETURNS
'   Long
'     The TEST_ACH_FAULT_* taken, or TEST_ACH_FAULT_NONE.
'
' BEHAVIOR
'   - Takes the restoration fault for a restoration and any other armed fault
'     for a step; leaves the other kind armed.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   None.
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchStep
'   - TST_AchRestore
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' TAKE THE FAULT
'------------------------------------------------------------------------------
        TST_AchTakeFault = TEST_ACH_FAULT_NONE

    'A fault of the beginning of a scenario is taken by TST_AchBegin, not here
        If m_AchFault <> TEST_ACH_FAULT_NO_CAPTION Then

            If ForRestore = ((m_AchFault = TEST_ACH_FAULT_RESTORE) Or _
                (m_AchFault = TEST_ACH_FAULT_RESTORE_READ)) Then

                TST_AchTakeFault = m_AchFault
                m_AchFault = TEST_ACH_FAULT_NONE
            End If

        End If

End Function

'
'------------------------------------------------------------------------------
'
'                        COMPARISON, TEXT AND RECORDING
'
'------------------------------------------------------------------------------
'

Private Function TST_AchSame( _
    ByRef StateA() As Boolean, _
    ByRef StateB() As Boolean) _
    As Boolean
'
'------------------------------------------------------------------------------
'                          COMPARE TWO STATE VECTORS
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether two state vectors are equal entry for entry.
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchSame"

    Dim EntryIdx    As Long 'Cursor over the entries

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        TST_AchSame = True

        For EntryIdx = 1 To m_AchCount

            If StateA(EntryIdx) <> StateB(EntryIdx) Then
                TST_AchSame = False
                Exit For
            End If

        Next EntryIdx

End Function

Private Function TST_AchNeedsRestore( _
    ByRef Current() As Boolean) _
    As Boolean
'
'------------------------------------------------------------------------------
'                   TELL WHETHER THE HOST IS AWAY FROM ENTRY
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether a state differs from the entry state of the scenario.
'
' WHY THIS EXISTS
'   The restoration and its readback both ask the same question: any entry
'   differs, or, for the title bar, the owned bits of the frame style differ
'   from the entry style even though the caption reads the same (#28).
'
' INPUTS
'   Current
'     ByRef. The state to compare with the entry state.
'
' RETURNS
'   Boolean
'     TRUE when it differs.
'
' BEHAVIOR
'   - Compares every entry with m_AchEntry.
'   - For the title bar, also compares the owned bits of the style read last
'     with the entry style.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_AchSame
'
' NOTES
'   The style compared is the one the latest TST_AchRead stored.
'
' CALLED FROM
'   - TST_AchRestore
'   - TST_AchRawRestore
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        TST_AchNeedsRestore = Not TST_AchSame(m_AchEntry, Current)

        If m_AchSurface = TEST_ACH_SURFACE_TITLEBAR Then

            If (m_AchStyle And TST_TITLEBAR_OWNED_MASK) <> _
                (m_AchEntryStyle And TST_TITLEBAR_OWNED_MASK) Then

                TST_AchNeedsRestore = True
            End If

        End If

End Function

Private Function TST_AchRawRestore( _
    ByRef Current() As Boolean, _
    ByRef Detail As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                   PUT THE ENTRY STATE BACK WITHOUT THE API
'------------------------------------------------------------------------------
' PURPOSE
'   Write the entry state back through the harness writers.
'
' WHY THIS EXISTS
'   The restoration must not rely on the API under test, so that a defect there
'   cannot leave the host changed (#66, #28).
'
' INPUTS
'   Current
'     ByRef. The state read just before, so that only what differs is written.
'
'   Detail
'     ByRef. Receives the reason a write failed; empty on success.
'
' RETURNS
'   Boolean
'     TRUE when every write that was owed succeeded.
'
' BEHAVIOR
'   - application and window: writes the entry value of each entry that
'     differs, through TST_TrySetBooleanProperty.
'   - title bar: writes the entry frame style back, refreshes the frame and
'     reads the owned bits back, through TST_RestoreTitleBarStyle.
'   - Ribbon: writes the entry visibility, through TST_TrySetRibbonVisible,
'     only while the window that owned it is still the active window.
'
' ERROR POLICY
'   - Does not raise. A write that fails is reported through Detail.
'
' DEPENDENCIES
'   - TST_TrySetBooleanProperty
'   - TST_RestoreTitleBarStyle
'   - TST_TrySetRibbonVisible
'   - TST_AchNeedsRestore
'   - TST_AchTryReadHwnd
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchRestore
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchRawRestore"

    Dim EntryIdx    As Long     'Cursor over the entries
    Dim Written     As Boolean  'TRUE while every write succeeded
    Dim Msg         As String   'Diagnostic from a writer
    Dim ActiveWin   As Object   'The active window, for the Ribbon

#If VBA7 Then
    Dim ActiveHwnd  As LongPtr  'Handle of the active window
#Else
    Dim ActiveHwnd  As Long 'Handle of the active window
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_AchRawRestore = False
        Detail = vbNullString
        Written = True

'------------------------------------------------------------------------------
' WRITE THE ENTRY STATE
'------------------------------------------------------------------------------
        Select Case m_AchSurface

            Case TEST_ACH_SURFACE_APP
                If Current(1) <> m_AchEntry(1) Then
                    If Not TST_TrySetBooleanProperty(Application, m_AchProperty, _
                        m_AchEntry(1), Msg) Then

                        Written = False
                        Detail = TST_CertAppendDetail(Detail, _
                            "Application." & m_AchProperty & " | " & Msg)
                    End If
                End If

            Case TEST_ACH_SURFACE_WINDOW
                For EntryIdx = 1 To m_AchCount

                    If Current(EntryIdx) <> m_AchEntry(EntryIdx) Then
                        If Not TST_TrySetBooleanProperty(m_AchWin(EntryIdx), _
                            m_AchProperty, m_AchEntry(EntryIdx), Msg) Then

                            Written = False
                            Detail = TST_CertAppendDetail(Detail, _
                                m_AchLabel(EntryIdx) & "." & m_AchProperty & " | " & Msg)
                        End If
                    End If

                Next EntryIdx

            Case TEST_ACH_SURFACE_TITLEBAR
                If Not TST_RestoreTitleBarStyle(m_AchHwnd, m_AchEntryStyle, _
                    m_AchFrameCaptured, Msg) Then

                    Written = False
                    Detail = TST_CertAppendDetail(Detail, "frame " & _
                        CStr(m_AchHwnd) & " | " & Msg)
                End If

            Case TEST_ACH_SURFACE_RIBBON

            'Only the window that owned the Ribbon may be written back
                Set ActiveWin = Application.ActiveWindow

                If Not TST_AchTryReadHwnd(ActiveWin, ActiveHwnd, Msg) Then
                    Written = False
                    Detail = "the active window could not be identified | " & Msg
                ElseIf ActiveHwnd <> m_AchHwnd Then
                    Written = False
                    Detail = "window " & CStr(m_AchHwnd) & _
                        " is no longer the active window, so the Ribbon is not written"
                ElseIf Current(1) <> m_AchEntry(1) Then
                    If Not TST_TrySetRibbonVisible(m_AchEntry(1), Msg) Then
                        Written = False
                        Detail = "the Ribbon | " & Msg
                    End If
                End If

            Case Else
                Written = False
                Detail = "no surface is selected"

        End Select

        TST_AchRawRestore = Written

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        Detail = TST_BuildRuntimeErrorText

        TST_AchRawRestore = False

        Resume SafeExit

End Function

Private Function TST_AchText( _
    ByRef Values() As Boolean, _
    ByVal WithNote As Boolean) _
    As String
'
'------------------------------------------------------------------------------
'                        RENDER A STATE VECTOR AS TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Render the state of the entries as visible or hidden, for the evidence.
'
' WHY THIS EXISTS
'   The evidence says what was requested and what was observed in words a
'   person can compare (#28). A window outside the scope is marked by its
'   label, and a window in the scope by an asterisk after it.
'
' INPUTS
'   Values
'     ByRef. The state of each entry.
'
'   WithNote
'     TRUE to add what the latest read noted, the owned bits of a frame style.
'
' RETURNS
'   String
'     "visible" or "hidden" for one entry; "label*=visible, label=hidden" for
'     several, where * marks a window in the scope.
'
' BEHAVIOR
'   - Joins the entries in order.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   None.
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchStep
'   - TST_AchRestore
'   - TST_AchDifferences
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchText"

    Dim EntryIdx    As Long     'Cursor over the entries
    Dim Rendered    As String   'Text under construction

'------------------------------------------------------------------------------
' RENDER
'------------------------------------------------------------------------------
        If m_AchCount = 1 Then

            Rendered = IIf(Values(1), "visible", "hidden")

        Else

            For EntryIdx = 1 To m_AchCount

                If EntryIdx > 1 Then
                    Rendered = Rendered & ", "
                End If

                Rendered = Rendered & m_AchLabel(EntryIdx) & _
                    IIf(m_AchTarget(EntryIdx), "*", vbNullString) & "=" & _
                    IIf(Values(EntryIdx), "visible", "hidden")

            Next EntryIdx

        End If

        If WithNote Then
            Rendered = Rendered & m_AchNote
        End If

        TST_AchText = Rendered

End Function

Private Function TST_AchDifferences( _
    ByRef Expected() As Boolean, _
    ByRef Observed() As Boolean) _
    As String
'
'------------------------------------------------------------------------------
'                         NAME THE ENTRIES THAT DIFFER
'------------------------------------------------------------------------------
' PURPOSE
'   Name every entry whose observed state is not the one expected, with both.
'
' WHY THIS EXISTS
'   A mismatch that says only that something differs cannot be diagnosed; one
'   that names the window and both states can (#28).
'
' INPUTS
'   Expected / Observed
'     ByRef. The two state vectors.
'
' RETURNS
'   String
'     "label expected X, observed Y" for each entry that differs, joined by "; ";
'     empty when none does.
'
' BEHAVIOR
'   - Compares the entries in order.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertAppendDetail
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchStep
'   - TST_AchRestore
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchDifferences"

    Dim EntryIdx    As Long     'Cursor over the entries
    Dim Named       As String   'Entries named so far

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        For EntryIdx = 1 To m_AchCount

            If Expected(EntryIdx) <> Observed(EntryIdx) Then

                If Len(Named) > 0 Then
                    Named = Named & "; "
                End If

                Named = Named & m_AchLabel(EntryIdx) & _
                    IIf(m_AchTarget(EntryIdx), "*", vbNullString) & _
                    " expected " & IIf(Expected(EntryIdx), "visible", "hidden") & _
                    ", observed " & IIf(Observed(EntryIdx), "visible", "hidden")

            End If

        Next EntryIdx

        TST_AchDifferences = Named

End Function

Private Function TST_AchIsAllowed( _
    ByVal Element As String, _
    ByVal Operation As String, _
    ByVal Outcome As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                 TELL WHETHER THE CONTRACT ALLOWS AN OUTCOME
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether an outcome may stand at an element and an operation without
'   failing the run.
'
' WHY THIS EXISTS
'   The live verdict and tools/check_excel_evidence.py must accept exactly the
'   same observations. A verified outcome stands at a hide, a show and a
'   restoration. not-verifiable stands only where the manifest exempts a surface
'   (the Ribbon, and windows that could not be read on entry), and only at its
'   entry, so a Ribbon that was readable and then is not, or a title bar that
'   cannot be verified, is a failure and not an exemption (#28).
'
' INPUTS
'   Element / Operation / Outcome
'     What the observation names.
'
' RETURNS
'   Boolean
'     TRUE  => the outcome is a failed one, which fails the run on its own, or a
'              verified outcome at a hide, show or restore, or a not-verifiable
'              outcome at an element and operation the contract exempts.
'     FALSE => any other outcome, which the caller turns into a failure.
'
' BEHAVIOR
'   - Classifies the outcome with TST_CertAchievedClass and looks the operation
'     or the element and operation up in the constants of the module.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertAchievedClass
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchRecordObservation
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
        Select Case TST_CertAchievedClass(Outcome)

            Case TEST_ACH_CLASS_VERIFIED
                TST_AchIsAllowed = (InStr(1, TEST_ACH_VERIFIED_OPERATIONS, _
                    "|" & Operation & "|", vbBinaryCompare) > 0)

            Case TEST_ACH_CLASS_UNVERIFIABLE
                TST_AchIsAllowed = (InStr(1, TEST_ACH_UNVERIFIABLE_ALLOWED, _
                    "|" & Element & "." & Operation & "|", vbBinaryCompare) > 0)

            Case Else
            'A failed outcome fails the run by itself
                TST_AchIsAllowed = True

        End Select

End Function

Private Sub TST_AchRecordObservation( _
    ByRef Rec As tCertRecord, _
    ByVal Operation As String, _
    ByVal RequestedText As String, _
    ByVal ObservedText As String, _
    ByVal Outcome As String, _
    ByVal Reason As String)
'
'------------------------------------------------------------------------------
'                            RECORD ONE OBSERVATION
'------------------------------------------------------------------------------
' PURPOSE
'   Store one observation in the record, log it, and keep a failed one as text
'   for the case.
'
' WHY THIS EXISTS
'   Recording, logging and counting a failure are the same step for every
'   observation, so they are written once (#28).
'
' INPUTS
'   Rec
'     ByRef. The record that receives the observation.
'
'   Operation / RequestedText / ObservedText / Outcome / Reason
'     What TST_CertAddAchieved stores, with the element and scope of the
'     scenario.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Stores the observation through TST_CertAddAchieved.
'   - Logs it under its outcome.
'   - A failed outcome, and an observation that could not be stored, add to
'     m_AchFailCount and m_AchFailText.
'
' ERROR POLICY
'   - Does not raise. A record that cannot store flags itself.
'
' DEPENDENCIES
'   - TST_CertAddAchieved
'   - TST_CertAchievedClass
'   - TST_AchIsAllowed
'   - TST_CertAppendDetail
'   - TST_Log
'
' NOTES
'   An outcome the contract does not allow at this element and operation is
'   recorded as readback-failed, or restore-failed for a restoration, with the
'   reason it names, so that no observation can pass the live verdict that the
'   evidence validator would reject.
'
' CALLED FROM
'   - TST_AchRunOn
'   - TST_AchRecordBeginFailure
'   - TST_AchStep
'   - TST_AchRestore
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchRecordObservation"

    Dim Code    As String   'Element, scope and operation
    Dim Text    As String   'The observation as a finding reads
    Dim Stored  As Boolean  'TRUE once the record holds the observation

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
    'An outcome the contract does not allow here is a failure, not an exemption:
    'what the live verdict passes, the evidence validator must accept
        If Not TST_AchIsAllowed(m_AchElement, Operation, Outcome) Then

            Reason = "the outcome " & Outcome & " is not allowed for " & _
                m_AchElement & " at " & Operation & " | " & Reason

            If Operation = "restore" Then
                Outcome = TEST_ACH_OUTCOME_RESTORE_FAILED
            Else
                Outcome = TEST_ACH_OUTCOME_READ_FAILED
            End If

        End If

        Code = m_AchElement & "." & m_AchScope & "." & Operation

        Text = Outcome & ": requested=" & RequestedText & _
            " observed=" & ObservedText & _
            IIf(Len(Reason) > 0, " | " & Reason, vbNullString)

        Stored = TST_CertAddAchieved(Rec, m_AchElement, m_AchScope, Operation, _
            RequestedText, ObservedText, Outcome, Reason)

        TST_Log "AchievedState", Outcome, Code & " | " & Text

    'A failed outcome fails the scenario, and so does a record that lost it
        If TST_CertAchievedClass(Outcome) = TEST_ACH_CLASS_FAILED Then
            m_AchFailCount = m_AchFailCount + 1
            m_AchFailText = TST_CertAppendDetail(m_AchFailText, Code & ": " & Text)
        End If

        If Not Stored Then
            m_AchFailCount = m_AchFailCount + 1
            m_AchFailText = TST_CertAppendDetail(m_AchFailText, Code & _
                ": the observation could not be recorded")
        End If

End Sub

Private Sub TST_AchAssertClean( _
    ByVal FailText As String, _
    ByVal CaseName As String)
'
'------------------------------------------------------------------------------
'                    FAIL THE CASE ON A FAILED OBSERVATION
'------------------------------------------------------------------------------
' PURPOSE
'   Raise when any observation of a scenario was a failed outcome.
'
' WHY THIS EXISTS
'   A mismatch fails the mandatory unit even when the production setter
'   returned True (#28), so the failure travels as an error through
'   TST_CertRunCase and not as a log line.
'
' INPUTS
'   FailText
'     The failed observations, as text; empty when none failed.
'
'   CaseName
'     Error source.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Raises TEST_ACH_ERR_BASE + TEST_ACH_ERR_UNCLEAN with FailText when it is
'     not empty.
'
' ERROR POLICY
'   - Raises on a failed observation.
'
' DEPENDENCIES
'   None.
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_Case_AchievedAppLevelState
'   - TST_Case_AchievedTitleBarFrameState
'   - TST_Case_AchievedRibbonOwningWindow
'   - TST_Case_AchievedMismatchFailsTheUnit
'   - TST_AchRunWindowCase
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchAssertClean"

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        If Len(FailText) > 0 Then
            Err.Raise _
                TEST_ACH_ERR_BASE + TEST_ACH_ERR_UNCLEAN, _
                CaseName, _
                "achieved state not verified: " & FailText
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                           WINDOW IDENTITY HELPERS
'
'------------------------------------------------------------------------------
'

#If VBA7 Then
Private Function TST_AchTryReadHwnd( _
    ByVal TargetWindow As Object, _
    ByRef HwndOut As LongPtr, _
    ByRef FailMsg As String) _
    As Boolean
#Else
Private Function TST_AchTryReadHwnd( _
    ByVal TargetWindow As Object, _
    ByRef HwndOut As Long, _
    ByRef FailMsg As String) _
    As Boolean
#End If
'
'------------------------------------------------------------------------------
'                         READ THE HANDLE OF A WINDOW
'------------------------------------------------------------------------------
' PURPOSE
'   Read the handle a Window reports, without raising.
'
' WHY THIS EXISTS
'   Windows are matched by the handle each object reports, never with Is,
'   because Excel returns a new wrapper from every read of
'   Application.ActiveWindow (#28).
'
' INPUTS
'   TargetWindow
'     The Window, or Nothing.
'
'   HwndOut
'     ByRef. Receives the handle, and 0 on failure.
'
'   FailMsg
'     ByRef. Receives the reason on failure; empty on success.
'
' RETURNS
'   Boolean
'     TRUE when a non-zero handle was read.
'
' BEHAVIOR
'   - Fails for Nothing and for a window that reports handle 0.
'
' ERROR POLICY
'   - Does not raise. An error is the reason it returns FALSE.
'
' DEPENDENCIES
'   - TST_BuildRuntimeErrorText
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchRead
'   - TST_AchRawRestore
'   - TST_AchBegin
'   - TST_AchWindowKey
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchTryReadHwnd"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        TST_AchTryReadHwnd = False
        HwndOut = 0
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' READ THE HANDLE
'------------------------------------------------------------------------------
        If TargetWindow Is Nothing Then
            FailMsg = "the Window is Nothing"
            GoTo SafeExit
        End If

        HwndOut = TargetWindow.hWnd

        If HwndOut = 0 Then
            FailMsg = "the Window reports hWnd 0"
            GoTo SafeExit
        End If

        TST_AchTryReadHwnd = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        FailMsg = TST_BuildRuntimeErrorText

        HwndOut = 0

        TST_AchTryReadHwnd = False

        Resume SafeExit

End Function

Private Function TST_AchWindowKey( _
    ByVal TargetWindow As Object) _
    As String
'
'------------------------------------------------------------------------------
'                         NAME A WINDOW BY ITS HANDLE
'------------------------------------------------------------------------------
' PURPOSE
'   Return the handle a Window reports as text, to match windows by.
'
' WHY THIS EXISTS
'   The scope of a write is the set of windows the production code reaches
'   through its own enumeration, so the observation names them by handle
'   (#28).
'
' INPUTS
'   TargetWindow
'     The Window, or Nothing.
'
' RETURNS
'   String
'     The handle as text; empty when it cannot be read.
'
' BEHAVIOR
'   - Reads the handle through TST_AchTryReadHwnd.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_AchTryReadHwnd
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchBegin
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchWindowKey"

    Dim Msg As String   'Unused reason from the read

#If VBA7 Then
    Dim WindowHwnd  As LongPtr  'Handle the window reports
#Else
    Dim WindowHwnd  As Long 'Handle the window reports
#End If

'------------------------------------------------------------------------------
' RETURN THE KEY
'------------------------------------------------------------------------------
        TST_AchWindowKey = vbNullString

        If TST_AchTryReadHwnd(TargetWindow, WindowHwnd, Msg) Then
            TST_AchWindowKey = CStr(WindowHwnd)
        End If

End Function

Private Function TST_AchWindowCaption( _
    ByVal TargetWindow As Object) _
    As String
'
'------------------------------------------------------------------------------
'                        LABEL A WINDOW BY ITS CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Return the caption of a window for a diagnostic label.
'
' WHY THIS EXISTS
'   A mismatch names the window it found, and a caption is what a person
'   recognizes (#28). Reading it must never end an observation.
'
' INPUTS
'   TargetWindow
'     The Window.
'
' RETURNS
'   String
'     The caption, or "window" when it cannot be read.
'
' BEHAVIOR
'   - Reads Caption under On Error Resume Next.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   None.
'
' NOTES
'   None
'
' CALLED FROM
'   - TST_AchBegin
'
' HISTORY
'   2026-11-28 - Added (#28).
'
' UPDATED
'   2026-11-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_EXCEL_UI_TEST_ACHIEVED.TST_AchWindowCaption"

'------------------------------------------------------------------------------
' READ THE CAPTION
'------------------------------------------------------------------------------
    'A caption that cannot be read is a label, not a failure
        TST_AchWindowCaption = "window"

        On Error Resume Next

        TST_AchWindowCaption = CStr(TargetWindow.Caption)

        On Error GoTo 0

End Function
