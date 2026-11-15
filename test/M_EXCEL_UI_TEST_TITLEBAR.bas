Attribute VB_Name = "M_EXCEL_UI_TEST_TITLEBAR"
Option Explicit
Option Private Module

'==============================================================================
' M_EXCEL_UI_TEST_TITLEBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Hold the title-bar regression cases of EXCEL_UI, the title-bar pack and
'   the SDI title-bar identity runner
'
' WHY THIS EXISTS
'   Title-bar control writes native window styles. It must change only the
'   bits EXCEL_UI owns, write only the frame a snapshot was captured from, and
'   report a WinAPI failure with the error that call left
'
' PUBLIC SURFACE
'   - Test_EXCEL_UI_RunTitleBarOnly
'   - Test_EXCEL_UI_RunTitleBarSdiIdentity
'
'   The cases are Public so that the regression pack in
'   M_EXCEL_UI_TEST_CERTIFICATION can run them
'
' TEST SCOPE
'   - hide / show round-trip
'   - preservation of unrelated GWL_STYLE bits across hide / show
'   - show recovers without a baseline and rejects a captionless baseline
'   - equal style bits cannot authenticate a recycled handle
'   - a failed frame refresh is owed and retried
'   - the frame-state registry refuses an entry it cannot prove
'   - WinAPI error capture: a valid zero, an injected DLL error and an
'     error left by an earlier call, for the GWL_STYLE read and write
'   - every title-bar refusal, through the public snapshot, restore and
'     show/hide entry points, leaves the active frame's complete GWL_STYLE
'     unchanged
'   - SDI identity: restoration writes the captured frame, reports a closed
'     one, and refuses a closed-and-reopened one
'
' STATE MANAGEMENT
'   - Every case that hides a title bar, the title-bar pack and the SDI
'     identity runner put the entry frame style back through the harness
'     WinAPI helpers, never through the API under test, and fail when they
'     cannot
'
' COMPATIBILITY
'   - Windows only
'   - Title-bar behavior remains the most OS / Excel-version-sensitive area
'
' DEPENDENCIES
'   - M_EXCEL_UI_TEST_HARNESS
'   - M_EXCEL_UI_TEST_SNAPSHOT, for the Ribbon cases of the SDI identity
'     runner
'
' UPDATED
'   2026-11-15 - Split out of the single regression module; the runners and
'                cases moved here unchanged apart from visibility (#90).
'   2026-11-06 - Every title-bar case that hides a frame restores the entry
'                style on every exit and fails when it cannot; the pack and
'                the SDI identity runner do the same for the entry frame
'                (#66).
'   2026-11-06 - Title-bar refusal cases compare the active frame's complete
'                style before and after the refused call. Added the
'                public-path pair-refusal and closed-and-reopened-frame cases
'                to the SDI identity runner (#45).
'   2026-11-06 - Added the WinAPI error-capture case to both title-bar runs
'                (#27).
'   2026-08-19 - Added the SDI title-bar identity pack and the frame-refresh-
'                debt case.
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.2
'==============================================================================


'------------------------------------------------------------------------------
' TEST CONFIGURATION
'------------------------------------------------------------------------------
    Private Const TEST_STYLE_OP_NONE As Long = 0  'Window-style seam: no operation
    Private Const TEST_STYLE_OP_READ As Long = 1  'Window-style seam: GWL_STYLE read
    Private Const TEST_STYLE_OP_WRITE As Long = 2  'Window-style seam: GWL_STYLE write
    Private Const TEST_STYLE_FAULT_NONE As Long = 0  'Window-style seam: no fault
    Private Const TEST_STYLE_FAULT_VALID_ZERO As Long = 1  'Zero result, captured error kept
    Private Const TEST_STYLE_FAULT_DLL_ERROR As Long = 2  'Zero result, armed error captured
    Private Const TEST_STYLE_FAULT_STALE_ERROR As Long = 3  'Armed error left before the clear

'------------------------------------------------------------------------------
' API CONSTANTS FOR TITLE-BAR STYLE TESTS
'------------------------------------------------------------------------------
    Private Const TST_SYNTHETIC_UNRELATED_BIT As Long = &H2000000

    'A legal caption frame that is not the full owned set. Written behind the
    'component to contradict an entry claiming the frame is hidden, which is
    'the evidence a reissued handle presents to the registry.
    Private Const TST_TITLEBAR_FOREIGN_FRAME As Long = &HC80000
    Private Const TST_TITLEBAR_CAPTIONLESS_FRAME As Long = &H40000

    'Win32 error codes the error-capture case injects. Any non-zero value would
    'do; real codes keep the expected failure text recognisable.
    Private Const TST_ERROR_FILE_NOT_FOUND  As Long = 2
    Private Const TST_ERROR_ACCESS_DENIED   As Long = 5
    Private Const TST_ERROR_INVALID_WINDOW_HANDLE As Long = 1400


'
'------------------------------------------------------------------------------
'
'                              PUBLIC RUNNERS
'
'------------------------------------------------------------------------------
'


Public Sub Test_EXCEL_UI_RunTitleBarOnly()

'
'==============================================================================
' Test_EXCEL_UI_RunTitleBarOnly
'------------------------------------------------------------------------------
' PURPOSE
'   Run the dedicated title-bar regression cases
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
'   - Runs the title-bar round-trip and owned-style-bit preservation cases
'   - Attempts to restore the original state
'
' ERROR POLICY
'   - Raises on assertion failure after attempting restoration
'
' DEPENDENCIES
'   - TST_RunTitleBarOnlyPack
'
' UPDATED
'   2026-07-29
'==============================================================================
'
'------------------------------------------------------------------------------
' APPLY TITLE-BAR-ONLY PACK
'------------------------------------------------------------------------------
    'Run the title-bar-only regression pack
        TST_RunTitleBarOnlyPack CallerProc:="Test_EXCEL_UI_RunTitleBarOnly"

End Sub


Public Sub Test_EXCEL_UI_RunTitleBarSdiIdentity()

'
'==============================================================================
' Test_EXCEL_UI_RunTitleBarSdiIdentity
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that title-bar snapshot restoration targets the window the state was
'   captured from, and never the window that happens to be active at restore
'   time.
'
' WHY THIS EXISTS
'   Under the Single Document Interface each workbook window owns a separate
'   top-level window, and Application.Hwnd reports whichever of them is active
'   when it is read. A single-window regression run cannot distinguish "wrote
'   to the captured frame" from "wrote to the active frame", because the two are
'   the same window. Only a second window makes the difference observable.
'
'   The two cases below are the ones that would have caught ICR-UI-P1-01 before
'   release.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Refuses to run when an explicit EXCEL_UI snapshot already exists.
'   - Runs the redirect case and the closed-frame case for the title bar, and
'     the other-active-window and closed-window cases for the Ribbon.
'   - Runs the public-path pair-refusal case and the closed-and-reopened-frame
'     case, which assert that a refused title bar leaves every frame exactly
'     as it was.
'   - Puts the anchor frame's entry style back through the harness before
'     returning, and fails when it cannot.
'
' ERROR POLICY
'   - Raises after best-effort cleanup.
'   - Preserves the original test failure through cleanup.
'
' DEPENDENCIES
'   - TST_Case_TitleBarSdiRestoreTargetsCapturedFrame
'   - TST_Case_TitleBarCapturedFrameClosedIsReported
'   - TST_Case_RibbonRestoreRefusesOtherActiveWindow
'   - TST_Case_RibbonRestoreRefusesClosedWindow
'   - TST_Case_TitleBarPairRefusalLeavesFrameUntouched
'   - TST_Case_TitleBarClosedFrameReopenedIsRefused
'
' NOTES
'   Destructive: creates and closes temporary workbooks. It is deliberately a
'   separate runner rather than part of Test_EXCEL_UI_RunAll. RunAll also opens
'   and closes temporary windows and workbooks for its targeting and
'   handle-reuse cases, but it does not run these SDI identity cases. Folding
'   every mandatory case into one runner is
'   tracked separately as the release-certification work.
'
' UPDATED
'   2026-11-06 - Restored the anchor frame through the harness instead of
'                the API under test, and failed when it could not (#66).
'   2026-11-06 - Registered the pair-refusal and closed-and-reopened-frame
'                cases (#45).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window          'Window active before the run
    Dim OldScreenUpdating   As Boolean         'Cached ScreenUpdating state

#If VBA7 Then
    Dim AnchorHwnd          As LongPtr         'Frame of the anchor window
    Dim AnchorFrameStyle    As LongPtr         'Its style on entry
#Else
    Dim AnchorHwnd          As Long            'Frame of the anchor window
    Dim AnchorFrameStyle    As Long            'Its style on entry
#End If

    Dim AnchorFrameCaptured As Boolean         'TRUE once AnchorFrameStyle was read
    Dim FrameRestored       As Boolean         'TRUE once the style is back
    Dim FrameRestoreMsg     As String          'Diagnostic from the restoration

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "Test_EXCEL_UI_RunTitleBarSdiIdentity"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Refuse to destroy a snapshot the caller is relying on, before the handler
    'is armed. Armed first, this refusal travelled Err_Handler, resumed at
    'Safe_Exit, and the clear there deleted the caller snapshot it had just
    'refused to disturb. This is a certification unit, so that reached the
    'release path.
        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 1, _
                PROC, _
                "an explicit EXCEL_UI snapshot already exists; clear or restore it before running this destructive test"
        End If

        On Error GoTo Err_Handler

        TST_Log PROC, "START", "SDI title-bar identity test started"

    'Hold the window to return to, so the run leaves the host as it found it
        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 2, _
                PROC, _
                "no active Excel window is available"
        End If

    'Record its frame's exact style through the harness, so cleanup can put it
    'back without the API under test
        AnchorHwnd = AnchorWindow.hWnd
        AnchorFrameCaptured = TST_TryGetWindowStyle( _
            AnchorHwnd, AnchorFrameStyle, FrameRestoreMsg)

        If Not AnchorFrameCaptured Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 3, _
                PROC, _
                "unable to read the anchor frame style | " & FrameRestoreMsg
        End If

    'These cases activate windows, so screen updates stay on for correctness
        OldScreenUpdating = Application.ScreenUpdating

'------------------------------------------------------------------------------
' RUN REGRESSION CASES
'------------------------------------------------------------------------------
    'Restoration must reach the captured frame, not the active one
        TST_Case_TitleBarSdiRestoreTargetsCapturedFrame

    'A captured frame that has closed must be reported, never redirected
        TST_Case_TitleBarCapturedFrameClosedIsReported

    'The Ribbon acts only on the active window, so restore must refuse it when
    'its captured window is not active, and when that window has closed
        TST_Case_RibbonRestoreRefusesOtherActiveWindow
        TST_Case_RibbonRestoreRefusesClosedWindow

    'A refused frame pair must write nothing through any public entry point,
    'and a captured frame that closed must stay refused once a new window has
    'opened in its place
        TST_Case_TitleBarPairRefusalLeavesFrameUntouched
        TST_Case_TitleBarClosedFrameReopenedIsRefused

    'Log successful completion before cleanup
        TST_Log PROC, "PASS", _
            "Restoration targeted the captured frame; closed, reopened and " & _
            "unpaired frames were refused and left untouched"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Put the anchor frame back exactly, through the harness. No snapshot is
    'cleared here: this runner never captures one. Its child cases create and
    'release their own, so an outer clear could only ever have destroyed a
    'caller-owned snapshot that reached this label through an early failure.
        On Error Resume Next
            If Not AnchorWindow Is Nothing Then
                AnchorWindow.Activate
            End If

            FrameRestored = TST_RestoreTitleBarStyle( _
                AnchorHwnd, AnchorFrameStyle, AnchorFrameCaptured, FrameRestoreMsg)
            Application.ScreenUpdating = OldScreenUpdating
        On Error GoTo 0

    'A frame that cannot be put back fails the unit
        TST_RecordCleanupFailure _
            FrameRestored, FrameRestoreMsg, PROC, _
            FailNumber, FailSource, FailDescription

        If FailNumber <> 0 Then
            HasFailure = True
        End If

    'Raise the captured failure after cleanup when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub


'
'------------------------------------------------------------------------------
'
'                          PRIVATE PACK RUNNERS
'
'------------------------------------------------------------------------------
'


Private Sub TST_RunTitleBarOnlyPack(ByVal CallerProc As String)

'
'==============================================================================
' TST_RunTitleBarOnlyPack
'------------------------------------------------------------------------------
' PURPOSE
'   Execute the dedicated title-bar regression cases and restore the
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
'   - TST_Case_TitleBarOwnedBitPreservation
'   - TST_Case_TitleBarShowRecoversWithoutBaseline
'   - TST_Case_TitleBarFrameRefreshDebtRetried
'   - TST_Case_TitleBarStaleFrameEntryNotReused
'   - TST_Case_TitleBarApiErrorCapture
'   - TST_RestoreTitleBarStyle
'   - TST_RecordCleanupFailure
'
' UPDATED
'   2026-11-06 - Put the entry frame's style back through the harness and
'                failed the pack when it could not (#66).
'   2026-11-06 - Registered the WinAPI error-capture case (#27).
'   2026-08-21
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedRibbonKnown    As Boolean         'TRUE when pre-test Ribbon state was read successfully
    Dim SavedRibbonVisible  As Boolean         'Pre-test Ribbon visibility
    Dim SavedStatusBarVisible As Boolean         'Pre-test StatusBar visibility
    Dim SavedScrollBarsVisible As Boolean         'Pre-test ScrollBars visibility
    Dim SavedFormulaBarVisible As Boolean         'Pre-test FormulaBar visibility

    Dim SavedWindowCount    As Long            'Pre-test Application.Windows.Count
    Dim SavedHeadingsVisible() As Boolean         'Pre-test per-window Headings visibility
    Dim SavedWorkbookTabsVisible() As Boolean         'Pre-test per-window WorkbookTabs visibility
    Dim SavedGridlinesVisible() As Boolean         'Pre-test per-window Gridlines visibility

    Dim SavedTitleBarKnown  As Boolean         'TRUE when pre-test title-bar state was read successfully
    Dim SavedTitleBarVisible As Boolean         'Pre-test title-bar visibility

#If VBA7 Then
    Dim EntryHwnd           As LongPtr         'Frame active before the run
    Dim EntryFrameStyle     As LongPtr         'Its style before the run
#Else
    Dim EntryHwnd           As Long            'Frame active before the run
    Dim EntryFrameStyle     As Long            'Its style before the run
#End If

    Dim EntryFrameCaptured  As Boolean         'TRUE once EntryFrameStyle was read
    Dim FrameRestored       As Boolean         'TRUE once the style is back
    Dim FrameRestoreMsg     As String          'Diagnostic from the restoration

    Dim OldScreenUpdating   As Boolean         'Cached ScreenUpdating state
    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

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

    'Record the active frame's exact style through the harness, so cleanup can
    'put it back without the API under test
        EntryHwnd = Application.hWnd
        EntryFrameCaptured = TST_TryGetWindowStyle( _
            EntryHwnd, EntryFrameStyle, FrameRestoreMsg)

        If Not EntryFrameCaptured Then
            TST_Log CallerProc, "TitleBar", _
                "entry frame style not recorded | " & FrameRestoreMsg
        End If

'------------------------------------------------------------------------------
' RUN REGRESSION CASE
'------------------------------------------------------------------------------
    'Run the dedicated title-bar round-trip case
        TST_Case_TitleBarRoundTrip

    'Verify the exact production merge policy with deterministic style values
        TST_Case_TitleBarOwnedBitPreservation

    'Verify that a show restores the frame with no captured baseline
        TST_Case_TitleBarShowRecoversWithoutBaseline

    'Verify that a non-zero captionless baseline is also unusable for show
        TST_Case_TitleBarShowRejectsCaptionlessBaseline

    'Verify that a failed frame refresh is retried rather than short-circuited
        TST_Case_TitleBarFrameRefreshDebtRetried

    'Verify that frame state which cannot be proved is discarded, not applied
        TST_Case_TitleBarStaleFrameEntryNotReused

    'Verify equal-zero and equal-nonzero handle reuse with object identity
        TST_Case_TitleBarSameStyleHandleReuse

    'Verify that a disagreed active-frame pair is refused by both callers
        TST_Case_ActiveFramePairRefusesMismatch

    'Verify how a zero GWL_STYLE result and its captured DLL error are read
        TST_Case_TitleBarApiErrorCapture

    'Log successful completion before restoration
        TST_Log CallerProc, "PASS", _
            "Title-bar round-trip, owned-bit preservation, show-recovery, refresh-debt and stale-entry cases passed"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
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

        'Put the entry frame's exact style back through the harness. The
        'restore above goes through the API under test, and a defect there
        'must not be able to leave the host with a hidden title bar.
            FrameRestored = TST_RestoreTitleBarStyle( _
                EntryHwnd, EntryFrameStyle, EntryFrameCaptured, FrameRestoreMsg)
        On Error GoTo 0

    'Restore ScreenUpdating before leaving the harness
        Application.ScreenUpdating = OldScreenUpdating

    'A frame that cannot be put back fails the pack, after any case failure
        TST_RecordCleanupFailure _
            FrameRestored, FrameRestoreMsg, CallerProc, _
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
Err_Handler:
    'Capture failure information so it can be re-raised after restoration
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description & _
                          IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

    'Log the captured failure without consulting the mutable Err object
        TST_Log CallerProc, "FAIL", _
            CStr(FailNumber) & ": " & FailDescription & _
            IIf(Len(FailSource) > 0, " | Source: " & FailSource, vbNullString)

        Resume Safe_Exit

End Sub


'
'------------------------------------------------------------------------------
'
'                           REGRESSION CASES
'
'------------------------------------------------------------------------------
'


Public Sub TST_Case_TitleBarShowRejectsCaptionlessBaseline()
'
'==============================================================================
' TST_Case_TitleBarShowRejectsCaptionlessBaseline
'------------------------------------------------------------------------------
' PURPOSE
'   Prove that show cannot adopt a non-zero owned-bit baseline without
'   WS_CAPTION or report success while the caption remains absent.
'
' ERROR POLICY
'   - Restores the entry style through the harness on every exit, with the
'     handler disabled, and raises the saved failure once.
'   - Fails, with the restoration diagnostic, when the entry style cannot be
'     written, refreshed and read back.
'   - Fails at .Precondition when Windows does not keep the captionless
'     owned-bit set, because the case could not then detect the defect.
'
' NOTES
'   This is the case the #6 negative control must fail. With the v1.1.2
'   fallback condition the show worker keeps the captionless baseline, finds
'   nothing to write and reports that no caption is visible, so the first
'   failing assertion is CaptionlessBaseline.ShowReportsSuccess.
'   tools/make_titlebar_mutant.py produces that control; CONTRIBUTING.md
'   describes the run.
'
' UPDATED
'   2026-11-06 - Verified and reported the entry-style restoration, disabled
'                the handler before cleanup, and asserted the captionless
'                precondition (#66).
'   2026-09-10
'==============================================================================

#If VBA7 Then
    Dim TargetHwnd          As LongPtr
    Dim EntryStyle          As LongPtr
    Dim CaptionlessStyle    As LongPtr
    Dim ResultStyle         As LongPtr
#Else
    Dim TargetHwnd          As Long
    Dim EntryStyle          As Long
    Dim CaptionlessStyle    As Long
    Dim ResultStyle         As Long
#End If

    Dim StyleCaptured       As Boolean
    Dim OK                  As Boolean
    Dim Msg                 As String
    Dim RestoreOK           As Boolean
    Dim RestoreMsg          As String
    Dim FailNumber          As Long
    Dim FailSource          As String
    Dim FailDescription     As String

    Const PROC As String = _
        "TST_Case_TitleBarShowRejectsCaptionlessBaseline"

        On Error GoTo Err_Handler
        TST_Log PROC, "START", _
            "Validating recovery from a non-zero captionless baseline"

        TargetHwnd = ActiveWindow.hWnd
        If TargetHwnd = 0 Then
            Err.Raise TEST_ERR_BASE + 75, PROC, _
                "invalid Excel Window hWnd"
        End If

        If Not TST_TryGetWindowStyle(TargetHwnd, EntryStyle, Msg) Then
            Err.Raise TEST_ERR_BASE + 76, PROC, _
                "could not read the entry style | " & Msg
        End If
        StyleCaptured = True

        TST_Log PROC, "TRACE", _
            "entry captured | hwnd=" & CStr(TargetHwnd) & _
            " | style=" & CStr(EntryStyle)

    'Leave one non-zero owned bit but deliberately remove WS_CAPTION.
        CaptionlessStyle = _
            (EntryStyle And Not TST_TITLEBAR_OWNED_MASK) Or _
            TST_TITLEBAR_CAPTIONLESS_FRAME

        TST_Log PROC, "TRACE", "captionless style write: enter"
        If Not TST_TrySetWindowStyle( _
            TargetHwnd, CaptionlessStyle, Msg) Then

            Err.Raise TEST_ERR_BASE + 77, PROC, _
                "could not write the captionless style | " & Msg
        End If

        TST_Log PROC, "TRACE", "style write returned; frame refresh: enter"
        If Not TST_TryRefreshWindowFrame(TargetHwnd, Msg) Then
            Err.Raise TEST_ERR_BASE + 78, PROC, _
                "could not refresh the captionless frame | " & Msg
        End If

    'The case means something only while Windows keeps a non-zero owned-bit
    'set without WS_CAPTION. A host that normalised it to zero would let the
    'v1.1.2 condition pass here, so say so instead of passing.
        If Not TST_TryGetWindowStyle(TargetHwnd, ResultStyle, Msg) Then
            Err.Raise TEST_ERR_BASE + 83, PROC, _
                "could not read the captionless style back | " & Msg
        End If

        TST_Log PROC, "TRACE", _
            "captionless style read back | style=" & CStr(ResultStyle)

        If (ResultStyle And TST_TITLEBAR_OWNED_MASK) = 0 Or _
            (ResultStyle And TST_WS_CAPTION) <> 0 Then

            Err.Raise TEST_ERR_BASE + 84, PROC & ".Precondition", _
                "the frame did not keep a non-zero captionless owned-bit " & _
                "set; style=" & CStr(ResultStyle) & "; the case cannot " & _
                "tell the current condition from the v1.1.2 one"
        End If

        TST_Log PROC, "TRACE", "frame refresh returned; baseline reset: enter"
        UI_InternalResetTitleBarBaseline

        TST_Log PROC, "TRACE", "baseline reset returned; show worker: enter"
        OK = UI_TrySetTitleBarVisibleForHwndIfNeeded( _
            TargetHwnd:=TargetHwnd, _
            IsVisible:=True, _
            FailMsg:=Msg)

        TST_Log PROC, "TRACE", "show worker returned; assertions: enter"
        TST_AssertBooleanEquals True, OK, _
            "CaptionlessBaseline.ShowReportsSuccess"
        TST_AssertTitleBarVisible True, _
            "CaptionlessBaseline.CaptionVisible"

        If Not TST_TryGetWindowStyle(TargetHwnd, ResultStyle, Msg) Then
            Err.Raise TEST_ERR_BASE + 79, PROC, _
                "could not read the resulting style | " & Msg
        End If

        If (ResultStyle And Not TST_TITLEBAR_OWNED_MASK) <> _
            (EntryStyle And Not TST_TITLEBAR_OWNED_MASK) Then

            Err.Raise TEST_ERR_BASE + 80, PROC, _
                "show changed unrelated window-style bits"
        End If

Safe_Exit:
    'Resume ends active handling but leaves the local handler enabled.
    'Disable it before cleanup and before re-raising.
    'Otherwise the saved failure re-enters Err_Handler and repeats cleanup.
        On Error GoTo 0

        TST_Log PROC, "TRACE", "entry style restoration: enter"
        RestoreOK = TST_RestoreTitleBarStyle( _
            TargetHwnd, EntryStyle, StyleCaptured, RestoreMsg)

    'The helper writes, refreshes and reads the owned bits back. A failure at
    'any step is this case's failure too, after the one it follows.
        TST_Log PROC, "TRACE", _
            "restoration helper returned | owned bits read back=" & _
            CStr(RestoreOK)

        TST_RecordCleanupFailure _
            RestoreOK, RestoreMsg, PROC, FailNumber, FailSource, FailDescription

        If FailNumber <> 0 Then
            TST_Log PROC, "FAIL", FailSource & " | " & FailDescription
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        TST_Log PROC, "PASS", _
            "Captionless baseline rejected and visible frame recovered"
        Exit Sub

Err_Handler:
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        Resume Safe_Exit

End Sub


Public Sub TST_Case_TitleBarSameStyleHandleReuse()
'
'==============================================================================
' TST_Case_TitleBarSameStyleHandleReuse
'------------------------------------------------------------------------------
' PURPOSE
'   Prove that equal style bits cannot authenticate a recycled numeric hWnd.
'
' BEHAVIOR
'   Exercises both equal-zero and equal-nonzero owned-bit collisions. The seam
'   changes only the slot's numeric name and claimed last-written bits; its
'   retained Window object remains the original generation token.
'
' ERROR POLICY
'   - Restores both entry styles through the harness on every exit, with the
'     handler disabled, and raises the saved failure once.
'   - Fails, with the restoration diagnostic, when either entry style cannot
'     be written, refreshed and read back.
'
' UPDATED
'   2026-11-06 - Verified and reported both entry-style restorations (#66).
'   2026-09-03
'==============================================================================

    Dim AnchorWindow        As Window
    Dim ReplacementBook     As Workbook
    Dim ReplacementWindow   As Window
    Dim CaseIndex           As Long
    Dim OK                  As Boolean
    Dim Msg                 As String
    Dim FailNumber          As Long
    Dim FailSource          As String
    Dim FailDescription     As String

#If VBA7 Then
    Dim AnchorHwnd          As LongPtr
    Dim ReplacementHwnd     As LongPtr
    Dim AnchorEntryStyle    As LongPtr
    Dim ReplacementEntryStyle As LongPtr
    Dim AnchorVisibleStyle  As LongPtr
    Dim ReplacementStyle    As LongPtr
    Dim ResultStyle         As LongPtr
    Dim CollisionBits       As LongPtr
#Else
    Dim AnchorHwnd          As Long
    Dim ReplacementHwnd     As Long
    Dim AnchorEntryStyle    As Long
    Dim ReplacementEntryStyle As Long
    Dim AnchorVisibleStyle  As Long
    Dim ReplacementStyle    As Long
    Dim ResultStyle         As Long
    Dim CollisionBits       As Long
#End If

    Dim AnchorCaptured      As Boolean
    Dim ReplacementCaptured As Boolean
    Dim AnchorRestored      As Boolean
    Dim ReplacementRestored As Boolean
    Dim AnchorRestoreMsg    As String
    Dim ReplacementRestoreMsg As String

    Const PROC As String = "TST_Case_TitleBarSameStyleHandleReuse"

        On Error GoTo Err_Handler
        TST_Log PROC, "START", _
            "Validating same-style recycled-hWnd rejection"

        Set AnchorWindow = ActiveWindow
        If AnchorWindow Is Nothing Then
            Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 90, PROC, _
                "no anchor Excel Window is available"
        End If
        AnchorHwnd = AnchorWindow.hWnd

        Set ReplacementBook = Workbooks.Add
        Set ReplacementWindow = ReplacementBook.Windows(1)
        ReplacementHwnd = ReplacementWindow.hWnd

        If AnchorHwnd = ReplacementHwnd Then
            Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 91, PROC, _
                "the host did not provide distinct SDI window handles"
        End If

        If Not TST_TryGetWindowStyle( _
            AnchorHwnd, AnchorEntryStyle, Msg) Then

            Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 92, PROC, Msg
        End If
        AnchorCaptured = True

        If Not TST_TryGetWindowStyle( _
            ReplacementHwnd, ReplacementEntryStyle, Msg) Then

            Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 93, PROC, Msg
        End If
        ReplacementCaptured = True

        For CaseIndex = 0 To 1
            UI_InternalResetTitleBarBaseline

        'Give the original generation a distinctive visible baseline before the
        'component hides it and records ownership.
            AnchorVisibleStyle = _
                (AnchorEntryStyle And Not TST_TITLEBAR_OWNED_MASK) Or _
                TST_TITLEBAR_FOREIGN_FRAME

            If Not TST_TrySetWindowStyle( _
                AnchorHwnd, AnchorVisibleStyle, Msg) Then

                Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 94, PROC, Msg
            End If
            If Not TST_TryRefreshWindowFrame(AnchorHwnd, Msg) Then
                Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 95, PROC, Msg
            End If

            OK = UI_TrySetTitleBarVisibleForHwndIfNeeded( _
                TargetHwnd:=AnchorHwnd, _
                IsVisible:=False, _
                FailMsg:=Msg)
            TST_AssertBooleanEquals True, OK, _
                "SameStyleReuse.AnchorHidden"

        'Model a replacement generation whose live style coincides with the
        'stale slot: zero in the first pass, non-zero captionless in the second.
            If CaseIndex = 0 Then
                CollisionBits = 0
            Else
                CollisionBits = TST_TITLEBAR_CAPTIONLESS_FRAME
            End If

            ReplacementStyle = _
                (ReplacementEntryStyle And Not TST_TITLEBAR_OWNED_MASK) Or _
                CollisionBits

            If Not TST_TrySetWindowStyle( _
                ReplacementHwnd, ReplacementStyle, Msg) Then

                Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 96, PROC, Msg
            End If
            If Not TST_TryRefreshWindowFrame(ReplacementHwnd, Msg) Then
                Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 97, PROC, Msg
            End If

            If Not UI_InternalSimulateFrameHandleReuse( _
                OriginalHwnd:=AnchorHwnd, _
                ReplacementHwnd:=ReplacementHwnd, _
                ReplacementOwnedBits:=CollisionBits) Then

                Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 98, PROC, _
                    "unable to arm the recycled-hWnd seam"
            End If

            OK = UI_TrySetTitleBarVisibleForHwndIfNeeded( _
                TargetHwnd:=ReplacementHwnd, _
                IsVisible:=True, _
                FailMsg:=Msg)
            TST_AssertBooleanEquals True, OK, _
                "SameStyleReuse.ReplacementShown"

            If Not TST_TryGetWindowStyle( _
                ReplacementHwnd, ResultStyle, Msg) Then

                Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 99, PROC, Msg
            End If

        'A discarded slot uses the safe visible fallback. The original
        'generation's distinctive baseline must never cross to the replacement.
            If (ResultStyle And TST_TITLEBAR_OWNED_MASK) <> _
                TST_TITLEBAR_OWNED_MASK Then

                Err.Raise TEST_TITLEBAR_SDI_ERR_BASE + 100, PROC, _
                    "stale frame state crossed to the replacement generation"
            End If
        Next CaseIndex

Safe_Exit:
        On Error Resume Next
        UI_InternalResetTitleBarBaseline
        AnchorRestored = TST_RestoreTitleBarStyle( _
            AnchorHwnd, AnchorEntryStyle, AnchorCaptured, AnchorRestoreMsg)
        ReplacementRestored = TST_RestoreTitleBarStyle( _
            ReplacementHwnd, ReplacementEntryStyle, ReplacementCaptured, _
            ReplacementRestoreMsg)
        TST_SafeCloseWorkbook ReplacementBook
        If Not AnchorWindow Is Nothing Then AnchorWindow.Activate
        On Error GoTo 0

    'A frame that cannot be put back fails the case, after any earlier failure
        TST_RecordCleanupFailure _
            AnchorRestored, AnchorRestoreMsg, PROC, _
            FailNumber, FailSource, FailDescription
        TST_RecordCleanupFailure _
            ReplacementRestored, ReplacementRestoreMsg, PROC & ".Replacement", _
            FailNumber, FailSource, FailDescription

        If FailNumber <> 0 Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        TST_Log PROC, "PASS", _
            "Equal-zero and equal-nonzero recycled handles were rejected"
        Exit Sub

Err_Handler:
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        Resume Safe_Exit

End Sub


Public Sub TST_Case_TitleBarRoundTrip()

'
'==============================================================================
' TST_Case_TitleBarRoundTrip
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
' ERROR POLICY
'   - Restores the entry style through the harness on every exit, never
'     through the API under test, and raises the saved failure once.
'   - Fails, with the restoration diagnostic, when the entry style cannot be
'     written, refreshed and read back.
'
' UPDATED
'   2026-11-06 - Restored the entry style through the harness on every exit
'                and failed when it could not (#66).
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim xlHnd               As LongPtr         'Frame active on entry
    Dim EntryStyle          As LongPtr         'Style observed on entry
#Else
    Dim xlHnd               As Long            'Frame active on entry
    Dim EntryStyle          As Long            'Style observed on entry
#End If

    Dim StyleCaptured       As Boolean         'TRUE once EntryStyle is usable
    Dim RestoreOK           As Boolean         'TRUE once the style is back
    Dim RestoreMsg          As String          'Diagnostic from the restoration
    Dim Msg                 As String          'Diagnostic from a harness helper

    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_TitleBarRoundTrip"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", "Validating title-bar hide and show round-trip"

'------------------------------------------------------------------------------
' CAPTURE ENTRY STYLE
'------------------------------------------------------------------------------
    'Retain the entry style, so the frame can be put back on every exit
        xlHnd = Application.hWnd

        If Not TST_TryGetWindowStyle(xlHnd, EntryStyle, Msg) Then
            Err.Raise TEST_ERR_BASE + 82, PROC, _
                "could not read the entry window style | " & Msg
        End If

        StyleCaptured = True

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
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Disable the handler first: Resume ended active handling but left it
    'enabled, and a re-raise through it would repeat this cleanup
        On Error GoTo 0

    'Put the frame back through the harness, whatever happened above
        RestoreOK = TST_RestoreTitleBarStyle( _
            xlHnd, EntryStyle, StyleCaptured, RestoreMsg)

        TST_RecordCleanupFailure _
            RestoreOK, RestoreMsg, PROC, FailNumber, FailSource, FailDescription

        If FailNumber <> 0 Then
            TST_Log PROC, "FAIL", FailSource & " | " & FailDescription
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        TST_Log PROC, "PASS", "Title-bar round-trip behaved as expected"

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub


Public Sub TST_Case_TitleBarOwnedBitPreservation()

'
'==============================================================================
' TST_Case_TitleBarOwnedBitPreservation
'------------------------------------------------------------------------------
' PURPOSE
'   Verify deterministically that the production title-bar merge policy changes
'   only the style bits owned by EXCEL_UI.
'
' WHY THIS EXISTS
'   Windows may normalize or reject individual GWL_STYLE bits on Excel's
'   top-level window. A test that writes an arbitrary sentinel bit to the live
'   window can therefore fail even when the production merge algorithm is
'   correct.
'
'   This case validates the exact production merge helper with synthetic style
'   values, while TST_Case_TitleBarRoundTrip continues to exercise the live
'   WinAPI hide/show path.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Verifies a show merge preserves every unrelated current-style bit.
'   - Verifies a hide merge clears all and only the owned bits.
'   - Verifies unrelated bits supplied through OwnedStyleBits are ignored.
'   - Uses the production UI_InternalMergeTitleBarStyleBits helper.
'
' ERROR POLICY
'   - Logs and raises on assertion failure.
'
' DEPENDENCIES
'   - UI_InternalMergeTitleBarStyleBits
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim CurrentStyle        As LongPtr
    Dim RequestedOwned      As LongPtr
    Dim ExpectedStyle       As LongPtr
    Dim ActualStyle         As LongPtr
    Dim UnrelatedMask       As LongPtr
#Else
    Dim CurrentStyle        As Long
    Dim RequestedOwned      As Long
    Dim ExpectedStyle       As Long
    Dim ActualStyle         As Long
    Dim UnrelatedMask       As Long
#End If

    Dim FailNumber          As Long
    Dim FailSource          As String
    Dim FailDescription     As String

    Const PROC As String = "TST_Case_TitleBarOwnedBitPreservation"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating deterministic title-bar style ownership"

        CurrentStyle = &H10345678
        RequestedOwned = TST_TITLEBAR_OWNED_MASK
        UnrelatedMask = Not TST_TITLEBAR_OWNED_MASK

'------------------------------------------------------------------------------
' ASSERT SHOW MERGE
'------------------------------------------------------------------------------
        ExpectedStyle = _
            (CurrentStyle And UnrelatedMask) Or _
            (RequestedOwned And TST_TITLEBAR_OWNED_MASK)

        ActualStyle = UI_InternalMergeTitleBarStyleBits( _
            CurrentStyle:=CurrentStyle, _
            OwnedStyleBits:=RequestedOwned)

        If ActualStyle <> ExpectedStyle Then
            Err.Raise _
                TEST_ERR_BASE + 50, _
                PROC, _
                "show merge returned an unexpected style"
        End If

        If (ActualStyle And UnrelatedMask) <> _
            (CurrentStyle And UnrelatedMask) Then

            Err.Raise _
                TEST_ERR_BASE + 51, _
                PROC, _
                "show merge changed unrelated current-style bits"
        End If

'------------------------------------------------------------------------------
' ASSERT HIDE MERGE
'------------------------------------------------------------------------------
        ActualStyle = UI_InternalMergeTitleBarStyleBits( _
            CurrentStyle:=CurrentStyle, _
            OwnedStyleBits:=0)

        If (ActualStyle And TST_TITLEBAR_OWNED_MASK) <> 0 Then
            Err.Raise _
                TEST_ERR_BASE + 52, _
                PROC, _
                "hide merge did not clear every owned style bit"
        End If

        If (ActualStyle And UnrelatedMask) <> _
            (CurrentStyle And UnrelatedMask) Then

            Err.Raise _
                TEST_ERR_BASE + 53, _
                PROC, _
                "hide merge changed unrelated current-style bits"
        End If

'------------------------------------------------------------------------------
' ASSERT DEFENSIVE MASKING
'------------------------------------------------------------------------------
    'Supply one unrelated bit through OwnedStyleBits. The helper must ignore it
    'and continue to source unrelated bits exclusively from CurrentStyle.
        RequestedOwned = _
            TST_TITLEBAR_OWNED_MASK Or TST_SYNTHETIC_UNRELATED_BIT

        ActualStyle = UI_InternalMergeTitleBarStyleBits( _
            CurrentStyle:=CurrentStyle, _
            OwnedStyleBits:=RequestedOwned)

        If (ActualStyle And TST_SYNTHETIC_UNRELATED_BIT) <> _
            (CurrentStyle And TST_SYNTHETIC_UNRELATED_BIT) Then

            Err.Raise _
                TEST_ERR_BASE + 54, _
                PROC, _
                "unrelated bits from OwnedStyleBits were not ignored"
        End If

        TST_Log PROC, "PASS", _
            "Unrelated bits preserved and owned bits merged correctly"

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

        TST_Log PROC, "FAIL", _
            CStr(FailNumber) & ": " & FailDescription & _
            IIf(Len(FailSource) > 0, " | Source: " & FailSource, vbNullString)

        Err.Raise _
            Number:=FailNumber, _
            Source:=FailSource, _
            Description:=FailDescription

End Sub


Public Sub TST_Case_TitleBarShowRecoversWithoutBaseline()
'
'==============================================================================
' TST_Case_TitleBarShowRecoversWithoutBaseline
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a show request restores the title bar even when the title-bar
'   subsystem holds no captured owned-bit baseline and the frame is already
'   hidden on entry.
'
' WHY THIS EXISTS
'   This is the regression guard for the defect where UI_ShowExcelUI silently
'   failed to restore the title bar after a VBA project reset.
'
'   The window style survives a project reset because it belongs to the running
'   Excel process, while M_EXCEL_UI_TITLEBAR module state does not. If the first
'   title-bar call after such a reset was a show while the frame was already
'   hidden, the subsystem captured an all-zero owned-bit baseline, merged it,
'   found nothing to change, short-circuited, and returned TRUE. The title bar
'   stayed hidden and no failure was reported through either diagnostic path.
'
'   TST_Case_TitleBarRoundTrip cannot detect this. It always begins from a
'   visible frame, so the baseline it causes to be captured is never zero. This
'   case deliberately reproduces the reset condition instead: it hides the frame
'   through the harness WinAPI helpers, which leaves M_EXCEL_UI_TITLEBAR module
'   state untouched, and only then asks the public API to show it.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Reads and retains the entry style so the frame can be restored exactly.
'   - Clears every owned bit directly through the harness WinAPI helpers, which
'     bypasses M_EXCEL_UI entirely and therefore captures no baseline.
'   - Confirms the frame really is hidden before the assertion of interest, so
'     a pass cannot be produced by a frame that was never hidden.
'   - Calls UI_ShowExcelUI and requires the caption bit to return.
'   - Restores the entry style on every exit path, including failure.
'
' ERROR POLICY
'   - Raises a TEST_ERR_BASE assertion error on failure, for the pack handler.
'   - Restores the entry style through the harness before re-raising, with
'     the handler disabled, so one failure cannot leave the host with a
'     hidden title bar.
'   - Fails, with the restoration diagnostic, when the entry style cannot be
'     written, refreshed and read back.
'
' DEPENDENCIES
'   - TST_TryGetWindowStyle
'   - TST_TrySetWindowStyle
'   - TST_TryRefreshWindowFrame
'   - TST_AssertTitleBarVisible
'   - TST_WaitUI
'   - TST_Log
'   - UI_InternalResetTitleBarBaseline
'   - UI_ShowExcelUI
'
' CALLED FROM
'   - TST_RunTitleBarOnlyPack
'
' NOTES
'   Both halves of the precondition are established explicitly. The frame is
'   hidden through the harness WinAPI helpers, which M_EXCEL_UI_TITLEBAR never
'   observes, and the captured baseline is discarded through
'   UI_InternalResetTitleBarBaseline.
'
'   The second step is not optional. An earlier version of this case relied on
'   the subsystem happening to hold no baseline, and passed with the production
'   fix reverted, because TST_Case_TitleBarRoundTrip runs first and captures one
'   from a visible frame. The case is now deterministic regardless of ordering
'   or of what ran earlier in the session.
'
' UPDATED
'   2026-11-06 - Restored the entry style in one cleanup path with the
'                handler disabled, and failed when it could not (#66).
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim xlHnd               As LongPtr         'Excel main-window handle
    Dim EntryStyle          As LongPtr         'Style observed on entry
    Dim HiddenStyle         As LongPtr         'Style with owned bits cleared
#Else
    Dim xlHnd               As Long            'Excel main-window handle
    Dim EntryStyle          As Long            'Style observed on entry
    Dim HiddenStyle         As Long            'Style with owned bits cleared
#End If

    Dim StyleCaptured       As Boolean         'TRUE once EntryStyle is usable
    Dim FailMsg             As String          'Diagnostic from a WinAPI helper
    Dim SavedErrNumber      As Long            'Captured assertion error number
    Dim SavedErrSource      As String          'Captured assertion error source
    Dim SavedErrDescription As String          'Captured assertion description
    Dim RestoreOK           As Boolean         'TRUE once the style is back
    Dim RestoreMsg          As String          'Diagnostic from the restoration

    Const PROC As String = "TST_Case_TitleBarShowRecoversWithoutBaseline"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route assertion and runtime errors to the restoring handler
        On Error GoTo Err_Handler

    'Announce the case
        TST_Log PROC, "START", _
            "Validating title-bar show recovery without a captured baseline"

    'Resolve the Excel main-window handle
        xlHnd = Application.hWnd

        If xlHnd = 0 Then
            Err.Raise TEST_ERR_BASE + 55, PROC, _
                "invalid Excel window handle"
        End If

'------------------------------------------------------------------------------
' CAPTURE ENTRY STYLE
'------------------------------------------------------------------------------
    'Retain the entry style so the frame can be restored exactly afterwards
        If Not TST_TryGetWindowStyle(xlHnd, EntryStyle, FailMsg) Then
            Err.Raise TEST_ERR_BASE + 56, PROC, _
                "could not read the entry window style | " & FailMsg
        End If

        StyleCaptured = True

'------------------------------------------------------------------------------
' ESTABLISH HIDDEN FRAME WITHOUT TOUCHING MODULE STATE
'------------------------------------------------------------------------------
    'Clear the owned bits directly. Writing through the harness rather than
    'through UI_SetExcelUI is the whole point: M_EXCEL_UI_TITLEBAR never sees
    'this change and therefore captures no baseline from it.
        HiddenStyle = EntryStyle And Not TST_TITLEBAR_OWNED_MASK

        If Not TST_TrySetWindowStyle(xlHnd, HiddenStyle, FailMsg) Then
            Err.Raise TEST_ERR_BASE + 57, PROC, _
                "could not clear the owned title-bar style bits | " & FailMsg
        End If

    'Recalculate the non-client frame so the change is observable
        If Not TST_TryRefreshWindowFrame(xlHnd, FailMsg) Then
            Err.Raise TEST_ERR_BASE + 58, PROC, _
                "could not refresh the non-client frame | " & FailMsg
        End If

        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' CONFIRM THE PRECONDITION
'------------------------------------------------------------------------------
    'A frame that was never hidden would let the real assertion pass for the
    'wrong reason, so the precondition is asserted rather than assumed
        TST_AssertTitleBarVisible False, "TitleBarShowRecovery.Precondition"

    'Discard any baseline captured earlier in this session. Without this the
    'case cannot reach the branch it exists to guard: TST_Case_TitleBarRoundTrip
    'runs first and captures a good baseline from a visible frame, so the show
    'below would succeed by the ordinary path whether or not the defect is
    'present.
        UI_InternalResetTitleBarBaseline

'------------------------------------------------------------------------------
' REQUEST RECOVERY THROUGH THE PUBLIC API
'------------------------------------------------------------------------------
    'This is the documented emergency recovery path
        UI_ShowExcelUI

        TST_WaitUI TEST_WAIT_SECONDS

    'The caption must return even though no baseline was ever captured
        TST_AssertTitleBarVisible True, "TitleBarShowRecovery.Show"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Disable the handler first: Resume ended active handling but left it
    'enabled, and a re-raise through it would repeat this cleanup
        On Error GoTo 0

    'Leave the host exactly as the case found it, through the harness and
    'never through the API under test, on the success and failure paths alike
        RestoreOK = TST_RestoreTitleBarStyle( _
            xlHnd, EntryStyle, StyleCaptured, RestoreMsg)

        TST_RecordCleanupFailure _
            RestoreOK, RestoreMsg, PROC, _
            SavedErrNumber, SavedErrSource, SavedErrDescription

    'Hand the failure, if any, to the pack handler
        If SavedErrNumber <> 0 Then
            TST_Log PROC, "FAIL", SavedErrSource & " | " & SavedErrDescription
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

        TST_Log PROC, "PASS", _
            "Show restored the title bar with no captured baseline"

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Retain the failure so it can be re-raised after the host is restored
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description

        Resume Safe_Exit

End Sub


Private Sub TST_Case_TitleBarSdiRestoreTargetsCapturedFrame()

'
'==============================================================================
' TST_Case_TitleBarSdiRestoreTargetsCapturedFrame
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that restoring a snapshot writes the captured title-bar state to the
'   window it was captured from, and leaves a different active window alone.
'
' WHY THIS EXISTS
'   This is the direct regression for ICR-UI-P1-01. Before the fix the snapshot
'   stored a Boolean with no record of its window and re-resolved
'   Application.Hwnd on restore, so this sequence applied the anchor window's
'   captured frame to the second window and reported success for it.
'
'   The assertion that matters is the negative one: the second window must be
'   untouched. A test that only checked the anchor window would have passed
'   against the defective build.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a second workbook window.
'   - Captures with the anchor window active and its title bar visible.
'   - Hides the anchor title bar, then activates the second window.
'   - Hides the second window's frame through the harness, so a restore of
'     the anchor's visible state redirected to it would be observable.
'   - Records the second window's complete frame style, then restores.
'   - Asserts the anchor frame is restored and the second frame is unchanged,
'     by visibility and by its complete style.
'
' ERROR POLICY
'   - Raises after cleanup of the temporary workbook, with both frames' entry
'     styles restored through the harness.
'   - Fails, with the restoration diagnostic, when either entry style cannot
'     be written, refreshed and read back.
'
' UPDATED
'   2026-11-06 - Restored both frames' entry styles through the harness and
'                failed when it could not (#66).
'   2026-11-06 - Hid the active frame before restore and compared its
'                complete style (#45).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window          'Window whose frame is captured
    Dim SecondWorkbook      As Workbook        'Temporary second workbook
    Dim SecondWindow        As Window          'Window of the second workbook

#If VBA7 Then
    Dim AnchorHwnd          As LongPtr         'Top-level frame of the anchor
    Dim SecondHwnd          As LongPtr         'Top-level frame of the second
    Dim SecondEntryStyle    As LongPtr         'Second frame style on entry
    Dim SecondFrameBefore   As LongPtr         'Second frame style before restore
    Dim AnchorEntryStyle    As LongPtr         'Anchor frame style on entry
#Else
    Dim AnchorHwnd          As Long            'Top-level frame of the anchor
    Dim SecondHwnd          As Long            'Top-level frame of the second
    Dim SecondEntryStyle    As Long            'Second frame style on entry
    Dim SecondFrameBefore   As Long            'Second frame style before restore
    Dim AnchorEntryStyle    As Long            'Anchor frame style on entry
#End If

    Dim SecondCaptured      As Boolean         'TRUE once SecondEntryStyle was read
    Dim AnchorCaptured      As Boolean         'TRUE once AnchorEntryStyle was read
    Dim SecondRestored      As Boolean         'TRUE once the second frame is back
    Dim AnchorRestored      As Boolean         'TRUE once the anchor frame is back
    Dim SecondRestoreMsg    As String          'Diagnostic from its restoration
    Dim AnchorRestoreMsg    As String          'Diagnostic from its restoration
    Dim Msg                 As String          'Diagnostic from a harness helper

    Dim SecondVisibleBefore As Boolean         'Second frame state before restore
    Dim SecondVisibleAfter  As Boolean         'Second frame state after restore
    Dim AnchorVisibleAfter  As Boolean         'Anchor frame state after restore

    Dim OK                  As Boolean         'Structured result flag
    Dim FailureCount        As Long            'Structured result failure count
    Dim FailureList         As Variant         'Structured result failure list

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_TitleBarSdiRestoreTargetsCapturedFrame"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that restore targets the captured frame under SDI"

'------------------------------------------------------------------------------
' PREPARE TWO TOP-LEVEL FRAMES
'------------------------------------------------------------------------------
    'Start from the current window and make sure its frame is visible
        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 10, _
                PROC, _
                "no active Excel window is available"
        End If

    'Keep the anchor's entry style before anything changes it, so the frame
    'can be put back through the harness on every exit
        AnchorHwnd = AnchorWindow.hWnd

        If Not TST_TryGetWindowStyle(AnchorHwnd, AnchorEntryStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 15, _
                PROC, _
                "unable to read the anchor window style | " & Msg
        End If

        AnchorCaptured = True

        UI_SetExcelUI TitleBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

    'The frame the title bar is captured from must be the anchor's own
        If Application.hWnd <> AnchorHwnd Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 16, _
                PROC, _
                "the anchor Window and Application.hWnd do not agree; the " & _
                "case cannot be evaluated"
        End If

    'A second workbook gives a second top-level window under SDI
        Set SecondWorkbook = Workbooks.Add
        Set SecondWindow = SecondWorkbook.Windows(1)

        SecondWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        SecondHwnd = Application.hWnd

    'Without two distinct frames the case proves nothing, so say so rather than
    'passing vacuously. A host that reports one handle for both windows is not
    'running the interface this case exists to exercise.
        If SecondHwnd = AnchorHwnd Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 11, _
                PROC, _
                "both workbook windows report the same top-level handle; " & _
                "this host is not running the Single Document Interface " & _
                "and the case cannot be evaluated"
        End If

'------------------------------------------------------------------------------
' CAPTURE WITH THE ANCHOR ACTIVE
'------------------------------------------------------------------------------
    'Return to the anchor and capture its frame while it is visible
        AnchorWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        OK = UI_CaptureExcelUIState_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertResultSuccess OK, FailureCount, FailureList, _
            "SdiRestoreTargetsCapturedFrame.Capture"

'------------------------------------------------------------------------------
' DIVERGE THE TWO FRAMES
'------------------------------------------------------------------------------
    'Hide the anchor frame, so restoration has something to put back
        UI_SetExcelUI TitleBar:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertBooleanEquals _
            False, _
            TST_TitleBarVisibleForHwndOrRaise(AnchorHwnd, PROC), _
            "SdiRestoreTargetsCapturedFrame.AnchorHiddenBeforeRestore"

'------------------------------------------------------------------------------
' ACTIVATE THE OTHER FRAME
'------------------------------------------------------------------------------
    'Make a different window active, which is the whole point of the case
        SecondWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

    'Hide the active frame through the harness. The anchor's captured state is
    'visible, so a restore redirected to this frame would show it; a visible
    'frame could not reveal that write.
        If Not TST_TryGetWindowStyle(SecondHwnd, SecondEntryStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 12, _
                PROC, _
                "unable to read the second window style | " & Msg
        End If

        SecondCaptured = True

        If Not TST_TrySetWindowStyle( _
            SecondHwnd, _
            SecondEntryStyle And Not TST_TITLEBAR_OWNED_MASK, _
            Msg) Then

            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 13, _
                PROC, _
                "unable to hide the second frame | " & Msg
        End If

        If Not TST_TryRefreshWindowFrame(SecondHwnd, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 14, _
                PROC, _
                "unable to refresh the second frame | " & Msg
        End If

        TST_WaitUI TEST_WAIT_SECONDS

    'Record the second frame exactly as it stands before restoration
        SecondFrameBefore = TST_FrameStyleOrRaise(SecondHwnd, PROC)
        SecondVisibleBefore = TST_TitleBarVisibleForHwndOrRaise(SecondHwnd, PROC)

        TST_AssertBooleanEquals _
            False, _
            SecondVisibleBefore, _
            "SdiRestoreTargetsCapturedFrame.SecondHiddenBeforeRestore"

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Restore while the WRONG window is active
        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_WaitUI TEST_WAIT_SECONDS

    'The title bar is written through the captured frame, so it is restored
    'while the second window is active. The Ribbon is not: it acts only on the
    'active window, so restore refuses it and reports exactly one Ribbon
    'failure instead of applying the anchor's value to the second window (#23).
        TST_AssertSingleFailurePrefix _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedPrefix:="Ribbon | captured Ribbon window is not the active window", _
            AssertionName:="SdiRestoreTargetsCapturedFrame.Restore"

'------------------------------------------------------------------------------
' ASSERT THE CAPTURED FRAME WAS RESTORED
'------------------------------------------------------------------------------
    'The anchor frame must be visible again, even though it was not active
        AnchorVisibleAfter = TST_TitleBarVisibleForHwndOrRaise(AnchorHwnd, PROC)

        TST_AssertBooleanEquals _
            True, _
            AnchorVisibleAfter, _
            "SdiRestoreTargetsCapturedFrame.AnchorRestored"

'------------------------------------------------------------------------------
' ASSERT THE ACTIVE FRAME WAS NOT TOUCHED
'------------------------------------------------------------------------------
    'This is the assertion the defective build failed: the active window must
    'not receive another window's captured state
        SecondVisibleAfter = TST_TitleBarVisibleForHwndOrRaise(SecondHwnd, PROC)

        TST_AssertBooleanEquals _
            SecondVisibleBefore, _
            SecondVisibleAfter, _
            "SdiRestoreTargetsCapturedFrame.SecondWindowUntouched"

    'No owned or unrelated bit of the active frame may change either
        TST_AssertFrameUntouched _
            SecondHwnd, _
            SecondFrameBefore, _
            "SdiRestoreTargetsCapturedFrame.SecondFrameUntouched"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Captured frame restored; the active frame was left unchanged"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Release the snapshot, put both frames back through the harness and close
    'the temporary workbook before leaving
        On Error Resume Next
            UI_ClearExcelUIStateSnapshot
            SecondRestored = TST_RestoreTitleBarStyle( _
                SecondHwnd, SecondEntryStyle, SecondCaptured, SecondRestoreMsg)
            TST_SafeCloseWorkbook SecondWorkbook

            If Not AnchorWindow Is Nothing Then
                AnchorWindow.Activate
            End If

            AnchorRestored = TST_RestoreTitleBarStyle( _
                AnchorHwnd, AnchorEntryStyle, AnchorCaptured, AnchorRestoreMsg)
        On Error GoTo 0

    'A frame that cannot be put back fails the case, after any earlier failure
        TST_RecordCleanupFailure _
            AnchorRestored, AnchorRestoreMsg, PROC, _
            FailNumber, FailSource, FailDescription
        TST_RecordCleanupFailure _
            SecondRestored, SecondRestoreMsg, PROC & ".Second", _
            FailNumber, FailSource, FailDescription

        If FailNumber <> 0 Then
            HasFailure = True
        End If

    'Raise the captured failure after cleanup when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub


Private Sub TST_Case_TitleBarCapturedFrameClosedIsReported()

'
'==============================================================================
' TST_Case_TitleBarCapturedFrameClosedIsReported
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that restoring a snapshot whose captured title-bar window has since
'   closed reports a TitleBar failure and writes nothing.
'
' WHY THIS EXISTS
'   Reporting the miss is the other half of ICR-UI-P1-01. A closed captured
'   frame must not silently fall back to whatever window is active: that is the
'   same misdirection, reached by a different route.
'
'   The surviving window's frame is asserted unchanged for exactly that reason.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a second workbook and captures with its window active.
'   - Closes that workbook, destroying the captured frame, and activates the
'     surviving window.
'   - Restores and asserts a TitleBar failure is reported.
'   - Asserts the surviving active frame is unchanged, by visibility and by
'     its complete style.
'
' ERROR POLICY
'   - Raises after best-effort cleanup of the temporary workbook.
'
' NOTES
'   Closing the captured window also invalidates its captured Window entry, so
'   the structured result legitimately carries more than one failure. The
'   assertion therefore looks for a TitleBar entry within the list rather than
'   requiring it to be the only one.
'
'   The captured and surviving frames are both visible, so a write redirected
'   to the survivor would not change it. The closed-and-reopened case,
'   TST_Case_TitleBarClosedFrameReopenedIsRefused, hides the frame that is
'   active at restore time to make that write observable.
'
' UPDATED
'   2026-11-06 - Activated the survivor before restore and compared its
'                complete style (#45).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window          'Window that survives the case
    Dim DoomedWorkbook      As Workbook        'Temporary workbook to be closed
    Dim DoomedWindow        As Window          'Window whose frame is captured

#If VBA7 Then
    Dim AnchorHwnd          As LongPtr         'Top-level frame of the anchor
    Dim AnchorFrameBefore   As LongPtr         'Anchor style before restore
#Else
    Dim AnchorHwnd          As Long            'Top-level frame of the anchor
    Dim AnchorFrameBefore   As Long            'Anchor style before restore
#End If

    Dim AnchorVisibleBefore As Boolean         'Anchor frame before restore
    Dim AnchorVisibleAfter  As Boolean         'Anchor frame after restore

    Dim OK                  As Boolean         'Structured result flag
    Dim FailureCount        As Long            'Structured result failure count
    Dim FailureList         As Variant         'Structured result failure list

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_TitleBarCapturedFrameClosedIsReported"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a closed captured frame is reported"

'------------------------------------------------------------------------------
' PREPARE THE SURVIVING FRAME
'------------------------------------------------------------------------------
    'Hold the window that must remain untouched
        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 20, _
                PROC, _
                "no active Excel window is available"
        End If

        UI_SetExcelUI TitleBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        AnchorHwnd = Application.hWnd

'------------------------------------------------------------------------------
' CAPTURE THE FRAME THAT WILL BE DESTROYED
'------------------------------------------------------------------------------
    'Capture with the temporary window active, so the snapshot names its frame
        Set DoomedWorkbook = Workbooks.Add
        Set DoomedWindow = DoomedWorkbook.Windows(1)

        DoomedWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        OK = UI_CaptureExcelUIState_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertResultSuccess OK, FailureCount, FailureList, _
            "TitleBarCapturedFrameClosed.Capture"

'------------------------------------------------------------------------------
' DESTROY THE CAPTURED FRAME
'------------------------------------------------------------------------------
    'Close the captured window; the snapshot now names a frame that is gone
        TST_SafeCloseWorkbook DoomedWorkbook

        Set DoomedWindow = Nothing
        TST_WaitUI TEST_WAIT_SECONDS

    'Make the survivor the active frame, so restore runs with the frame a
    'fallback to the active window would reach
        AnchorWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

    'Record the surviving frame exactly as it stands before restoration
        AnchorVisibleBefore = TST_TitleBarVisibleForHwndOrRaise(AnchorHwnd, PROC)
        AnchorFrameBefore = TST_FrameStyleOrRaise(AnchorHwnd, PROC)

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Restoration must refuse the title bar rather than redirect it
        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' ASSERT THE MISS IS REPORTED
'------------------------------------------------------------------------------
    'A TitleBar entry must appear in the ordered failure list
        TST_AssertFailureListContainsPrefix _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedPrefix:="TitleBar", _
            AssertionName:="TitleBarCapturedFrameClosed.Reported"

'------------------------------------------------------------------------------
' ASSERT THE SURVIVING FRAME WAS NOT TOUCHED
'------------------------------------------------------------------------------
    'Nothing may be written when the captured frame cannot be proven present
        AnchorVisibleAfter = TST_TitleBarVisibleForHwndOrRaise(AnchorHwnd, PROC)

        TST_AssertBooleanEquals _
            AnchorVisibleBefore, _
            AnchorVisibleAfter, _
            "TitleBarCapturedFrameClosed.SurvivorUntouched"

        TST_AssertFrameUntouched _
            AnchorHwnd, _
            AnchorFrameBefore, _
            "TitleBarCapturedFrameClosed.SurvivorFrameUntouched"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Closed captured frame was reported and no state was applied"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Release the snapshot and any surviving temporary workbook
        On Error Resume Next
            UI_ClearExcelUIStateSnapshot
            TST_SafeCloseWorkbook DoomedWorkbook

            If Not AnchorWindow Is Nothing Then
                AnchorWindow.Activate
            End If
        On Error GoTo 0

    'Raise the captured failure after cleanup when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub


Private Sub TST_Case_TitleBarPairRefusalLeavesFrameUntouched()

'
'==============================================================================
' TST_Case_TitleBarPairRefusalLeavesFrameUntouched
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a refused active-frame pair leaves the active frame exactly as
'   it was, through every public entry point that resolves the pair.
'
' WHY THIS EXISTS
'   TST_Case_ActiveFramePairRefusesMismatch proves that the helper refuses a
'   disagreed pair. It does not prove that its callers act on the refusal: a
'   caller that ignored it, or wrote before asking, would still pass there.
'   This case arms the same one-shot seam, drives the public entry points
'   instead, and compares the active frame's complete GWL_STYLE before and
'   after each refused call (#45).
'
' RETURNS
'   None.
'
' BEHAVIOR
'   1  UI_SetExcelUI_WithResult      the title-bar request is refused with one
'                                    TitleBar failure; the frame is unchanged
'   2  UI_SetExcelUI                 the fire-and-forget request is refused;
'                                    the frame is unchanged
'   3  UI_CaptureExcelUIState_WithResult
'                                    the title bar is refused with one
'                                    TitleBar failure; the frame is unchanged
'   4  UI_ResetExcelUIToSnapshot_WithResult, frame as on entry
'                                    nothing is written to the frame
'   5  UI_ResetExcelUIToSnapshot_WithResult, owned bits inverted through the
'      harness                       nothing is written to the frame
'
' ERROR POLICY
'   - Raises after disarming the seam, releasing the snapshot it created and
'     restoring the entry style through the harness.
'   - Fails, with the restoration diagnostic, when the entry style cannot be
'     written, refreshed and read back.
'
' DEPENDENCIES
'   - UI_InternalInjectFramePairFault
'   - UI_SetExcelUI_WithResult
'   - UI_SetExcelUI
'   - UI_CaptureExcelUIState_WithResult
'   - UI_ResetExcelUIToSnapshot_WithResult
'   - TST_TryGetWindowStyle
'   - TST_TrySetWindowStyle
'   - TST_TryRefreshWindowFrame
'   - TST_FrameStyleOrRaise
'   - TST_AssertFrameUntouched
'   - TST_AssertSingleFailurePrefix
'   - TST_RestoreTitleBarStyle
'
' CALLED FROM
'   - Test_EXCEL_UI_RunTitleBarSdiIdentity
'
' NOTES
'   The seam is consumed by the first pair resolution, so it is armed
'   immediately before each call and disarmed immediately after it. Nothing in
'   these calls resolves the pair before the title bar does.
'
'   Each request asks for the opposite of the frame's state on entry, so a
'   write that ignored the refusal would change the frame.
'
'   A refused capture records no frame identity and leaves the title-bar state
'   unknown, so restore does not attempt the title bar. Branches 4 and 5 put
'   the frame in each of the two states a wrongful restore could write and
'   assert that neither is written. The restore result is logged, not
'   asserted: what this case asserts is that nothing reached the frame.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window          'Window active on entry

#If VBA7 Then
    Dim TargetHwnd          As LongPtr         'Frame of the active window
    Dim EntryStyle          As LongPtr         'Style observed on entry
    Dim InvertedStyle       As LongPtr         'Entry style, owned bits inverted
    Dim FrameBefore         As LongPtr         'Style before each refused call
#Else
    Dim TargetHwnd          As Long            'Frame of the active window
    Dim EntryStyle          As Long            'Style observed on entry
    Dim InvertedStyle       As Long            'Entry style, owned bits inverted
    Dim FrameBefore         As Long            'Style before each refused call
#End If

    Dim StyleCaptured       As Boolean         'TRUE once EntryStyle is usable
    Dim WrongfulMode        As UIVisibility    'Request a write would act on
    Dim CreatedSnapshot     As Boolean         'TRUE once this case owns one
    Dim Msg                 As String          'Diagnostic from a harness helper
    Dim RestoreOK           As Boolean         'TRUE once the style is back
    Dim RestoreMsg          As String          'Diagnostic from the restoration

    Dim OK                  As Boolean         'Structured result flag
    Dim FailureCount        As Long            'Structured result failure count
    Dim FailureList         As Variant         'Structured result failure list

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_TitleBarPairRefusalLeavesFrameUntouched"
    Const PAIR_REFUSAL As String = _
        "TitleBar | the active Excel Window is not the active native frame"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a refused frame pair writes nothing"

        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_NONE

'------------------------------------------------------------------------------
' RESOLVE AND RECORD THE ACTIVE FRAME
'------------------------------------------------------------------------------
        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 130, _
                PROC, _
                "no active Excel window is available"
        End If

        TargetHwnd = AnchorWindow.hWnd

    'The seam forces a disagreement. Without an agreement first, the refusal
    'below could have a cause other than the seam.
        If TargetHwnd = 0 Or TargetHwnd <> Application.hWnd Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 131, _
                PROC, _
                "the active Window and Application.hWnd do not agree before " & _
                "the seam is armed; the case cannot be evaluated"
        End If

    'Retain the entry style so the frame can be put back exactly
        If Not TST_TryGetWindowStyle(TargetHwnd, EntryStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 132, _
                PROC, _
                "could not read the entry window style | " & Msg
        End If

        StyleCaptured = True

    'Ask for the opposite of the entry state, so a write would show
        If (EntryStyle And TST_WS_CAPTION) <> 0 Then
            WrongfulMode = UI_Hide
        Else
            WrongfulMode = UI_Show
        End If

'------------------------------------------------------------------------------
' BRANCH 1 - THE STRUCTURED SETTER REFUSES AND WRITES NOTHING
'------------------------------------------------------------------------------
        FrameBefore = TST_FrameStyleOrRaise(TargetHwnd, PROC)

        FailureCount = 99
        FailureList = Array("stale set failure")

        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_MISMATCH

        OK = UI_SetExcelUI_WithResult( _
            TitleBar:=WrongfulMode, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_NONE

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertSingleFailurePrefix _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedPrefix:=PAIR_REFUSAL, _
            AssertionName:=PROC & ".Set.Refused"

        TST_AssertFrameUntouched _
            TargetHwnd, _
            FrameBefore, _
            PROC & ".Set.FrameUntouched"

'------------------------------------------------------------------------------
' BRANCH 2 - THE FIRE-AND-FORGET SETTER REFUSES AND WRITES NOTHING
'------------------------------------------------------------------------------
        FrameBefore = TST_FrameStyleOrRaise(TargetHwnd, PROC)

        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_MISMATCH

        UI_SetExcelUI TitleBar:=WrongfulMode

        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_NONE

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertFrameUntouched _
            TargetHwnd, _
            FrameBefore, _
            PROC & ".FireAndForget.FrameUntouched"

'------------------------------------------------------------------------------
' BRANCH 3 - CAPTURE REFUSES THE TITLE BAR AND WRITES NOTHING
'------------------------------------------------------------------------------
        FrameBefore = TST_FrameStyleOrRaise(TargetHwnd, PROC)

        FailureCount = 99
        FailureList = Array("stale capture failure")

        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_MISMATCH

        OK = UI_CaptureExcelUIState_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_NONE

    'Ownership follows an observed snapshot, never the returned flag
        CreatedSnapshot = UI_HasExcelUIStateSnapshot

        TST_AssertSingleFailurePrefix _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedPrefix:=PAIR_REFUSAL, _
            AssertionName:=PROC & ".Capture.Refused"

        TST_AssertTrue _
            ActualValue:=CreatedSnapshot, _
            AssertionName:=PROC & ".Capture.SnapshotAvailable"

        TST_AssertFrameUntouched _
            TargetHwnd, _
            FrameBefore, _
            PROC & ".Capture.FrameUntouched"

'------------------------------------------------------------------------------
' BRANCH 4 - RESTORE WRITES NOTHING TO THE FRAME AS IT WAS ON ENTRY
'------------------------------------------------------------------------------
        FrameBefore = TST_FrameStyleOrRaise(TargetHwnd, PROC)

        FailureCount = 99
        FailureList = Array("stale restore failure")

        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_WaitUI TEST_WAIT_SECONDS

        TST_Log PROC, "OBSERVED", _
            "Restore.EntryFrame | OK=" & CStr(OK) & _
            " | FailureCount=" & CStr(FailureCount)

        TST_AssertFrameUntouched _
            TargetHwnd, _
            FrameBefore, _
            PROC & ".RestoreEntryFrame.FrameUntouched"

'------------------------------------------------------------------------------
' BRANCH 5 - RESTORE WRITES NOTHING TO THE FRAME IN THE OPPOSITE STATE
'------------------------------------------------------------------------------
    'Invert the owned bits through the harness, which the title-bar module
    'never observes. A restore that wrote either visibility would now change
    'the frame in one of the two branches.
        If (EntryStyle And TST_WS_CAPTION) <> 0 Then
            InvertedStyle = EntryStyle And Not TST_TITLEBAR_OWNED_MASK
        Else
            InvertedStyle = _
                (EntryStyle And Not TST_TITLEBAR_OWNED_MASK) Or _
                TST_TITLEBAR_OWNED_MASK
        End If

        If Not TST_TrySetWindowStyle(TargetHwnd, InvertedStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 133, _
                PROC, _
                "could not invert the owned title-bar style bits | " & Msg
        End If

        If Not TST_TryRefreshWindowFrame(TargetHwnd, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 134, _
                PROC, _
                "could not refresh the inverted frame | " & Msg
        End If

        TST_WaitUI TEST_WAIT_SECONDS

        FrameBefore = TST_FrameStyleOrRaise(TargetHwnd, PROC)

    'A frame Windows refused to invert would make this branch prove nothing
        TST_AssertBooleanEquals _
            ((EntryStyle And TST_WS_CAPTION) = 0), _
            ((FrameBefore And TST_WS_CAPTION) <> 0), _
            PROC & ".RestoreInvertedFrame.Precondition"

        FailureCount = 99
        FailureList = Array("stale restore failure")

        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_WaitUI TEST_WAIT_SECONDS

        TST_Log PROC, "OBSERVED", _
            "Restore.InvertedFrame | OK=" & CStr(OK) & _
            " | FailureCount=" & CStr(FailureCount)

        TST_AssertFrameUntouched _
            TargetHwnd, _
            FrameBefore, _
            PROC & ".RestoreInvertedFrame.FrameUntouched"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "A refused frame pair wrote nothing through the setters, capture " & _
            "or restore"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Disarm the seam, release the snapshot this case created and put the frame
    'back through the harness, whatever happened above
        On Error Resume Next
            UI_InternalInjectFramePairFault TEST_PAIR_FAULT_NONE

            If CreatedSnapshot Then
                UI_ClearExcelUIStateSnapshot
            End If

            RestoreOK = TST_RestoreTitleBarStyle( _
                TargetHwnd, EntryStyle, StyleCaptured, RestoreMsg)

            If Not AnchorWindow Is Nothing Then
                AnchorWindow.Activate
            End If
        On Error GoTo 0

    'A frame that cannot be put back fails the case, after any earlier failure
        TST_RecordCleanupFailure _
            RestoreOK, RestoreMsg, PROC, FailNumber, FailSource, FailDescription

        If FailNumber <> 0 Then
            HasFailure = True
        End If

    'Raise the captured failure after cleanup when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub


Private Sub TST_Case_TitleBarClosedFrameReopenedIsRefused()

'
'==============================================================================
' TST_Case_TitleBarClosedFrameReopenedIsRefused
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that restoring a snapshot whose captured frame was closed, after a
'   new window has opened, refuses the title bar and writes nothing to the
'   new window's frame or to the surviving one.
'
' WHY THIS EXISTS
'   Closing a window and opening another is how a captured hWnd can come to
'   name a different window: Windows may issue the closed frame's handle to
'   the next window it creates. Whether it does or not, the retained Window
'   object belongs to the closed window, so restore must refuse the title bar
'   rather than write the captured state to the window now open (#45).
'
'   The new window is active and its frame is hidden through the harness
'   before restore, so a restore of the captured visible state onto it would
'   be observable.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Captures with a temporary window active and its frame visible.
'   - Closes that workbook and opens a replacement workbook, which is active.
'   - Logs whether the replacement received the closed frame's handle.
'   - Hides the replacement's frame through the harness.
'   - Restores and asserts a TitleBar failure is reported.
'   - Asserts the replacement and surviving frames keep their complete style.
'
' ERROR POLICY
'   - Raises after releasing the snapshot it created, restoring the
'     replacement's entry style through the harness and closing the temporary
'     workbooks.
'   - Fails, with the restoration diagnostic, when the replacement's entry
'     style cannot be written, refreshed and read back.
'
' DEPENDENCIES
'   - UI_CaptureExcelUIState_WithResult
'   - UI_ResetExcelUIToSnapshot_WithResult
'   - TST_TryGetWindowStyle
'   - TST_TrySetWindowStyle
'   - TST_TryRefreshWindowFrame
'   - TST_FrameStyleOrRaise
'   - TST_AssertFrameUntouched
'   - TST_AssertFailureListContainsPrefix
'   - TST_RestoreTitleBarStyle
'
' CALLED FROM
'   - Test_EXCEL_UI_RunTitleBarSdiIdentity
'
' NOTES
'   A reissued handle cannot be forced on demand. When it does not occur, the
'   restore is refused by the liveness check; when it does, by the retained
'   Window check. The OBSERVED line records which of the two this run reached.
'
'   Closing the captured window also invalidates its Ribbon window and its
'   Window entry, so the result carries more than one failure. The assertion
'   looks for the TitleBar entry within the list.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window          'Window that survives the case
    Dim DoomedWorkbook      As Workbook        'Workbook whose frame is captured
    Dim DoomedWindow        As Window          'Window whose frame is captured
    Dim ReplacementWorkbook As Workbook        'Workbook opened after the close
    Dim ReplacementWindow   As Window          'Window active at restore time

#If VBA7 Then
    Dim AnchorHwnd          As LongPtr         'Frame of the surviving window
    Dim DoomedHwnd          As LongPtr         'Captured frame, closed later
    Dim ReplacementHwnd     As LongPtr         'Frame of the replacement
    Dim ReplacementEntryStyle As LongPtr         'Replacement style on entry
    Dim ReplacementFrameBefore As LongPtr         'Replacement style before restore
    Dim AnchorFrameBefore   As LongPtr         'Survivor style before restore
#Else
    Dim AnchorHwnd          As Long            'Frame of the surviving window
    Dim DoomedHwnd          As Long            'Captured frame, closed later
    Dim ReplacementHwnd     As Long            'Frame of the replacement
    Dim ReplacementEntryStyle As Long            'Replacement style on entry
    Dim ReplacementFrameBefore As Long            'Replacement style before restore
    Dim AnchorFrameBefore   As Long            'Survivor style before restore
#End If

    Dim ReplacementCaptured As Boolean         'TRUE once its entry style was read
    Dim ReplacementRestored As Boolean         'TRUE once its entry style is back
    Dim RestoreMsg          As String          'Diagnostic from the restoration
    Dim CreatedSnapshot     As Boolean         'TRUE once this case owns one
    Dim Msg                 As String          'Diagnostic from a harness helper

    Dim OK                  As Boolean         'Structured result flag
    Dim FailureCount        As Long            'Structured result failure count
    Dim FailureList         As Variant         'Structured result failure list

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_TitleBarClosedFrameReopenedIsRefused"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a closed and reopened captured frame is refused"

'------------------------------------------------------------------------------
' CAPTURE THE FRAME THAT WILL BE CLOSED
'------------------------------------------------------------------------------
        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 140, _
                PROC, _
                "no active Excel window is available"
        End If

        AnchorHwnd = AnchorWindow.hWnd

        Set DoomedWorkbook = Workbooks.Add
        Set DoomedWindow = DoomedWorkbook.Windows(1)

        DoomedWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        DoomedHwnd = DoomedWindow.hWnd

    'Without two distinct frames the case proves nothing, so say so rather than
    'passing vacuously
        If DoomedHwnd = AnchorHwnd Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 141, _
                PROC, _
                "both workbook windows report the same top-level handle; " & _
                "this host is not running the Single Document Interface " & _
                "and the case cannot be evaluated"
        End If

    'The captured state must be visible, so restoring it onto the hidden
    'replacement below would be observable
        TST_AssertBooleanEquals _
            True, _
            TST_TitleBarVisibleForHwndOrRaise(DoomedHwnd, PROC), _
            "ClosedFrameReopened.CapturedFrameVisible"

        FailureCount = 99
        FailureList = Array("stale capture failure")

        OK = UI_CaptureExcelUIState_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

    'Ownership follows an observed snapshot, never the returned flag
        CreatedSnapshot = UI_HasExcelUIStateSnapshot

        TST_AssertResultSuccess OK, FailureCount, FailureList, _
            "ClosedFrameReopened.Capture"

'------------------------------------------------------------------------------
' CLOSE THE CAPTURED FRAME AND OPEN ANOTHER
'------------------------------------------------------------------------------
        TST_SafeCloseWorkbook DoomedWorkbook

        Set DoomedWindow = Nothing
        TST_WaitUI TEST_WAIT_SECONDS

        Set ReplacementWorkbook = Workbooks.Add
        Set ReplacementWindow = ReplacementWorkbook.Windows(1)

        ReplacementWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        ReplacementHwnd = ReplacementWindow.hWnd

    'Record which refusal this run can reach; a reused handle cannot be forced
        TST_Log PROC, "OBSERVED", _
            "closed hWnd=" & CStr(DoomedHwnd) & _
            " | replacement hWnd=" & CStr(ReplacementHwnd) & _
            " | handle reused=" & CStr(ReplacementHwnd = DoomedHwnd)

'------------------------------------------------------------------------------
' HIDE THE REPLACEMENT FRAME THROUGH THE HARNESS
'------------------------------------------------------------------------------
        If Not TST_TryGetWindowStyle( _
            ReplacementHwnd, ReplacementEntryStyle, Msg) Then

            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 142, _
                PROC, _
                "unable to read the replacement window style | " & Msg
        End If

        ReplacementCaptured = True

        If Not TST_TrySetWindowStyle( _
            ReplacementHwnd, _
            ReplacementEntryStyle And Not TST_TITLEBAR_OWNED_MASK, _
            Msg) Then

            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 143, _
                PROC, _
                "unable to hide the replacement frame | " & Msg
        End If

        If Not TST_TryRefreshWindowFrame(ReplacementHwnd, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 144, _
                PROC, _
                "unable to refresh the replacement frame | " & Msg
        End If

        TST_WaitUI TEST_WAIT_SECONDS

    'Record both frames exactly as they stand before restoration
        ReplacementFrameBefore = TST_FrameStyleOrRaise(ReplacementHwnd, PROC)
        AnchorFrameBefore = TST_FrameStyleOrRaise(AnchorHwnd, PROC)

        TST_AssertBooleanEquals _
            False, _
            ((ReplacementFrameBefore And TST_WS_CAPTION) <> 0), _
            "ClosedFrameReopened.ReplacementHiddenBeforeRestore"

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        FailureCount = 99
        FailureList = Array("stale restore failure")

        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_WaitUI TEST_WAIT_SECONDS

    'Every refusal of the captured frame reports under this prefix
        TST_AssertFailureListContainsPrefix _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedPrefix:="TitleBar | captured title-bar ", _
            AssertionName:="ClosedFrameReopened.Reported"

'------------------------------------------------------------------------------
' ASSERT NEITHER FRAME WAS TOUCHED
'------------------------------------------------------------------------------
        TST_AssertFrameUntouched _
            ReplacementHwnd, _
            ReplacementFrameBefore, _
            "ClosedFrameReopened.ReplacementFrameUntouched"

        TST_AssertFrameUntouched _
            AnchorHwnd, _
            AnchorFrameBefore, _
            "ClosedFrameReopened.SurvivorFrameUntouched"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "A closed and reopened captured frame was refused and no frame " & _
            "was written"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Release the snapshot this case created, put the replacement frame back
    'and close the temporary workbooks
        On Error Resume Next
            If CreatedSnapshot Then
                UI_ClearExcelUIStateSnapshot
            End If

            ReplacementRestored = TST_RestoreTitleBarStyle( _
                ReplacementHwnd, ReplacementEntryStyle, ReplacementCaptured, _
                RestoreMsg)

            TST_SafeCloseWorkbook ReplacementWorkbook
            TST_SafeCloseWorkbook DoomedWorkbook

            If Not AnchorWindow Is Nothing Then
                AnchorWindow.Activate
            End If
        On Error GoTo 0

    'A frame that cannot be put back fails the case, after any earlier failure
        TST_RecordCleanupFailure _
            ReplacementRestored, RestoreMsg, PROC, _
            FailNumber, FailSource, FailDescription

        If FailNumber <> 0 Then
            HasFailure = True
        End If

    'Raise the captured failure after cleanup when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub


Public Sub TST_Case_TitleBarFrameRefreshDebtRetried()

'
'==============================================================================
' TST_Case_TitleBarFrameRefreshDebtRetried
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a style write whose frame refresh failed is recorded as owed and
'   retried on the next call, instead of being short-circuited as a no-op.
'
' WHY THIS EXISTS
'   This is the regression for ICR-UI-P2-03. After a failed refresh the style
'   already matches the request, so the no-op test would otherwise fire and
'   report success over a frame Windows never re-measured. That false success is
'   invisible from the outside, which is why the module exposes a debt-query
'   seam: without it this case could only assert that the second call succeeded,
'   which it would also do if the debt had simply been forgotten.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Clears the frame-state registry so the case starts from a known state.
'   - Arms a one-shot frame-refresh failure.
'   - Requests a hide and asserts it reports failure.
'   - Asserts a refresh is recorded as owed for the window.
'   - Repeats the request and asserts it succeeds and clears the debt.
'
' ERROR POLICY
'   - Raises after restoring the entry style through the harness, never
'     through the API under test, and discarding the frame-state registry.
'   - Fails, with the restoration diagnostic, when the entry style cannot be
'     written, refreshed and read back.
'
' NOTES
'   Uses the internal seams UI_InternalInjectFrameRefreshFailure and
'   UI_InternalIsFrameRefreshPending. They are Public only for same-project
'   regression access; Option Private Module keeps them out of the external
'   automation namespace.
'
' UPDATED
'   2026-11-06 - Captured the entry style before the case changes anything
'                and restored it through the harness instead of the API
'                under test (#66).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim TargetHwnd          As LongPtr         'Frame under test
    Dim EntryStyle          As LongPtr         'Style to restore on the way out
#Else
    Dim TargetHwnd          As Long            'Frame under test
    Dim EntryStyle          As Long            'Style to restore on the way out
#End If

    Dim StyleCaptured       As Boolean         'TRUE once EntryStyle was read
    Dim RestoreOK           As Boolean         'TRUE once the style is back
    Dim RestoreMsg          As String          'Diagnostic from the restoration
    Dim FirstAttemptOK      As Boolean         'Result of the injected-failure call
    Dim SecondAttemptOK     As Boolean         'Result of the retry call
    Dim Msg                 As String          'Diagnostic buffer

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_TitleBarFrameRefreshDebtRetried"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a failed frame refresh is retried"

'------------------------------------------------------------------------------
' CAPTURE THE ENTRY STYLE
'------------------------------------------------------------------------------
        TargetHwnd = Application.hWnd

        If TargetHwnd = 0 Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 30, _
                PROC, _
                "no Excel window handle is available"
        End If

    'Keep the entry style before anything changes it, so the frame can be put
    'back exactly as the case found it
        If Not TST_TryGetWindowStyle(TargetHwnd, EntryStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 31, _
                PROC, _
                "unable to read the entry window style | " & Msg
        End If

        StyleCaptured = True

'------------------------------------------------------------------------------
' PREPARE A KNOWN FRAME STATE
'------------------------------------------------------------------------------
    'Start from a visible frame and an empty registry
        UI_SetExcelUI TitleBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        UI_InternalResetTitleBarBaseline

    'The requests below address the frame the entry style was read from
        If Application.hWnd <> TargetHwnd Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 32, _
                PROC, _
                "the active frame changed while the case was being prepared"
        End If

'------------------------------------------------------------------------------
' FAIL THE FRAME REFRESH
'------------------------------------------------------------------------------
    'Arm the one-shot seam, then request a hide
        UI_InternalInjectFrameRefreshFailure True

        FirstAttemptOK = UI_TrySetTitleBarVisibleForHwndIfNeeded( _
            TargetHwnd:=TargetHwnd, _
            IsVisible:=False, _
            FailMsg:=Msg)

    'The style write succeeded and the repaint did not, so this must report
    'failure rather than success
        TST_AssertBooleanEquals _
            False, _
            FirstAttemptOK, _
            "TitleBarFrameRefreshDebt.FirstAttemptReportsFailure"

'------------------------------------------------------------------------------
' ASSERT THE DEBT WAS RECORDED
'------------------------------------------------------------------------------
    'Without this the next call cannot know a repaint is still owed
        TST_AssertBooleanEquals _
            True, _
            UI_InternalIsFrameRefreshPending(TargetHwnd), _
            "TitleBarFrameRefreshDebt.DebtRecorded"

'------------------------------------------------------------------------------
' RETRY
'------------------------------------------------------------------------------
    'The same request again. The style already matches, so a build without the
    'debt would short-circuit here and report a false success.
        SecondAttemptOK = UI_TrySetTitleBarVisibleForHwndIfNeeded( _
            TargetHwnd:=TargetHwnd, _
            IsVisible:=False, _
            FailMsg:=Msg)

        TST_AssertBooleanEquals _
            True, _
            SecondAttemptOK, _
            "TitleBarFrameRefreshDebt.RetrySucceeds"

'------------------------------------------------------------------------------
' ASSERT THE DEBT WAS SETTLED
'------------------------------------------------------------------------------
    'A confirmed repaint is the only thing that may clear the debt
        TST_AssertBooleanEquals _
            False, _
            UI_InternalIsFrameRefreshPending(TargetHwnd), _
            "TitleBarFrameRefreshDebt.DebtCleared"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Failed refresh was recorded as owed and retried on the next call"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Disarm the seam, put the frame back through the harness and discard the
    'registry, whatever happened above. The API under test takes no part.
        On Error Resume Next
            UI_InternalInjectFrameRefreshFailure False
            RestoreOK = TST_RestoreTitleBarStyle( _
                TargetHwnd, EntryStyle, StyleCaptured, RestoreMsg)
            UI_InternalResetTitleBarBaseline
        On Error GoTo 0

    'A frame that cannot be put back fails the case, after any earlier failure
        TST_RecordCleanupFailure _
            RestoreOK, RestoreMsg, PROC, FailNumber, FailSource, FailDescription

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
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub


Public Sub TST_Case_TitleBarStaleFrameEntryNotReused()

'
'==============================================================================
' TST_Case_TitleBarStaleFrameEntryNotReused
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that the frame-state registry refuses to apply an entry it can no
'   longer prove describes the window its handle names.
'
' WHY THIS EXISTS
'   This guards the different-owned-bits form of ICR-UI-111-P2-01. The
'   equal-bits form, ICR-UI-112-P2-02 (#32), remains open and is exercised by
'   TST_Case_TitleBarSameStyleHandleReuse. Windows reissues a window
'   handle once the window holding it has closed, and IsWindow answers for
'   whichever window holds the handle now, so a handle match was accepted as
'   proof of identity. A show could then restore a closed window's captured
'   frame onto an unrelated window that had merely inherited its handle.
'
'   A reissued handle cannot be forced on demand, so this case reproduces what
'   the registry actually sees: an entry claiming the frame is hidden while the
'   window it names carries owned bits the component never wrote. That is the
'   same evidence a reissued handle with different owned bits presents, and
'   it is the evidence the owned-bits check acts on.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Clears the frame-state registry so the case starts from a known state.
'   - Hides the frame, leaving an entry that claims it.
'   - Writes a different owned frame behind the component's back.
'   - Requests a show and asserts the live frame survives it.
'
' ERROR POLICY
'   - Raises after restoring the entry style through the harness, never
'     through the API under test, and discarding the frame-state registry.
'   - Fails, with the restoration diagnostic, when the entry style cannot be
'     written, refreshed and read back.
'
' NOTES
'   Windows may normalise or reject individual GWL_STYLE bits, so the frame
'   written here is read back and the case asserts it is distinguishable from
'   the stale baseline before drawing any conclusion from it.
'
' UPDATED
'   2026-11-06 - Captured the entry style before the case changes anything
'                and restored it through the harness alone (#66).
'   2026-11-02
'   2026-08-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim TargetHwnd          As LongPtr         'Frame under test
    Dim EntryStyle          As LongPtr         'Style to restore on the way out
    Dim VisibleStyle        As LongPtr         'Style once the frame is shown
    Dim ForeignStyle        As LongPtr         'Style written behind the module
    Dim ResultStyle         As LongPtr         'Style observed after the show
    Dim BaselineOwned       As LongPtr         'Owned bits a stale show writes
    Dim ForeignOwned        As LongPtr         'Owned bits actually accepted
#Else
    Dim TargetHwnd          As Long            'Frame under test
    Dim EntryStyle          As Long            'Style to restore on the way out
    Dim VisibleStyle        As Long            'Style once the frame is shown
    Dim ForeignStyle        As Long            'Style written behind the module
    Dim ResultStyle         As Long            'Style observed after the show
    Dim BaselineOwned       As Long            'Owned bits a stale show writes
    Dim ForeignOwned        As Long            'Owned bits actually accepted
#End If

    Dim StyleCaptured       As Boolean         'TRUE once EntryStyle was read
    Dim RestoreOK           As Boolean         'TRUE once the style is back
    Dim RestoreMsg          As String          'Diagnostic from the restoration
    Dim Msg                 As String          'Diagnostic buffer

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_TitleBarStaleFrameEntryNotReused"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that an unprovable frame entry is not applied"

'------------------------------------------------------------------------------
' CAPTURE THE ENTRY STYLE
'------------------------------------------------------------------------------
        TargetHwnd = Application.hWnd

        If TargetHwnd = 0 Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 60, _
                PROC, _
                "no Excel window handle is available"
        End If

    'Keep the entry style before anything changes it, so the frame can be put
    'back exactly as the case found it
        If Not TST_TryGetWindowStyle(TargetHwnd, EntryStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 61, _
                PROC, _
                "unable to read the entry window style | " & Msg
        End If

        StyleCaptured = True

'------------------------------------------------------------------------------
' PREPARE A KNOWN FRAME STATE
'------------------------------------------------------------------------------
    'Start from a visible frame and an empty registry
        UI_SetExcelUI TitleBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        UI_InternalResetTitleBarBaseline

    'The requests below address the frame the entry style was read from
        If Application.hWnd <> TargetHwnd Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 70, _
                PROC, _
                "the active frame changed while the case was being prepared"
        End If

        If Not TST_TryGetWindowStyle(TargetHwnd, VisibleStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 71, _
                PROC, _
                "unable to read the visible window style | " & Msg
        End If

        BaselineOwned = VisibleStyle And TST_TITLEBAR_OWNED_MASK

'------------------------------------------------------------------------------
' LEAVE AN ENTRY THAT CLAIMS THE FRAME
'------------------------------------------------------------------------------
    'Hide through the component, so the registry records a hidden frame and the
    'owned bits it wrote to achieve it
        If Not UI_TrySetTitleBarVisibleForHwndIfNeeded( _
            TargetHwnd:=TargetHwnd, _
            IsVisible:=False, _
            FailMsg:=Msg) Then

            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 62, _
                PROC, _
                "unable to hide the title bar | " & Msg
        End If

'------------------------------------------------------------------------------
' CONTRADICT THE CLAIM
'------------------------------------------------------------------------------
    'Give the window an owned frame the component never wrote. To the registry
    'this is indistinguishable from a handle Windows has issued to a different
    'window, which is the condition under test.
        ForeignStyle = _
            (VisibleStyle And Not TST_TITLEBAR_OWNED_MASK) Or _
            TST_TITLEBAR_FOREIGN_FRAME

        If Not TST_TrySetWindowStyle(TargetHwnd, ForeignStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 63, _
                PROC, _
                "unable to write the contradicting style | " & Msg
        End If

        If Not TST_TryRefreshWindowFrame(TargetHwnd, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 64, _
                PROC, _
                "unable to refresh the contradicting frame | " & Msg
        End If

    'Read back what Windows actually accepted rather than what was requested
        If Not TST_TryGetWindowStyle(TargetHwnd, ForeignStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 65, _
                PROC, _
                "unable to read the contradicting style back | " & Msg
        End If

        ForeignOwned = ForeignStyle And TST_TITLEBAR_OWNED_MASK

    'The case can only distinguish the two behaviours while the live frame and
    'the stale baseline differ. Say so plainly rather than passing on a
    'comparison that proves nothing.
        If ForeignOwned = BaselineOwned Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 66, _
                PROC, _
                "Windows normalised the contradicting frame back to the " & _
                "baseline; the case cannot distinguish the two outcomes"
        End If

'------------------------------------------------------------------------------
' REQUEST A SHOW
'------------------------------------------------------------------------------
    'A build that trusts the handle match restores the stale baseline over this
    'frame. A build that proves the entry first discards it, adopts the live
    'bits and leaves the visible frame exactly as found.
        If Not UI_TrySetTitleBarVisibleForHwndIfNeeded( _
            TargetHwnd:=TargetHwnd, _
            IsVisible:=True, _
            FailMsg:=Msg) Then

            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 67, _
                PROC, _
                "show reported failure | " & Msg
        End If

'------------------------------------------------------------------------------
' ASSERT THE LIVE FRAME SURVIVED
'------------------------------------------------------------------------------
        If Not TST_TryGetWindowStyle(TargetHwnd, ResultStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 68, _
                PROC, _
                "unable to read the resulting window style | " & Msg
        End If

        If (ResultStyle And TST_TITLEBAR_OWNED_MASK) <> ForeignOwned Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 69, _
                PROC, _
                "the show applied a baseline this window never had; stale " & _
                "frame state was reused"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "An unprovable frame entry was discarded instead of applied"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Put the frame back through the harness and discard the registry, whatever
    'happened above. The API under test takes no part.
        On Error Resume Next
            RestoreOK = TST_RestoreTitleBarStyle( _
                TargetHwnd, EntryStyle, StyleCaptured, RestoreMsg)
            UI_InternalResetTitleBarBaseline
        On Error GoTo 0

    'A frame that cannot be put back fails the case, after any earlier failure
        TST_RecordCleanupFailure _
            RestoreOK, RestoreMsg, PROC, FailNumber, FailSource, FailDescription

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
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub


Public Sub TST_Case_ActiveFramePairRefusesMismatch()

'
'==============================================================================
' TST_Case_ActiveFramePairRefusesMismatch
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that the active-frame pair is refused when the host does not agree
'   the resolved Window is the active native frame, and that the refusal
'   reaches both callers.
'
' WHY THIS EXISTS
'   Resolving Window.hWnd from the Window just returned by
'   Application.ActiveWindow makes the pair internally consistent. It does not
'   make it current: during an SDI activation transition the resolved Window
'   can already be the previously active frame. A fire-and-forget title-bar
'   write would then modify the wrong window, and a snapshot would capture that
'   frame and later restore it reporting success.
'
'   The disagreement cannot be provoked on demand, so the fail-closed branch is
'   reached through the one-shot seam in M_EXCEL_UI_TITLEBAR. Without it the
'   branch would ship having never executed, which is the shape of the defect
'   this pairing exists to prevent.
'
' RETURNS
'   None
'
' BEHAVIOR
'   1  no fault              the pair resolves, Window and handle both set
'   2  fault armed           refused, Window Nothing, handle zero, exact text
'   3  seam is one-shot      the next resolution succeeds again
'   4  refusal delegates     the handle-only wrapper reports the same failure
'
' ERROR POLICY
'   - Raises on assertion failure
'   - Disarms the seam before returning
'
' DEPENDENCIES
'   - UI_InternalInjectFramePairFault
'   - UI_TryGetActiveFramePair
'   - UI_TryGetActiveTitleBarHwnd
'
' CALLED FROM
'   - TST_RunRegressionPack
'   - TST_RunTitleBarOnlyPack
'
' NOTES
'   The snapshot capture path is not asserted separately here. It reaches the
'   same helper through the same call, so a separate assertion would restate
'   this one rather than test anything further; what it would add is coverage
'   of the snapshot's own failure reporting, which belongs with its structured
'   result cases.
'
' UPDATED
'   2026-09-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResolvedWindow      As Object          'Window the helper reports

#If VBA7 Then
    Dim ResolvedHwnd        As LongPtr         'Handle the helper reports
#Else
    Dim ResolvedHwnd        As Long            'Handle the helper reports
#End If

    Dim FailMsg             As String          'Failure text the helper reports
    Dim OK                  As Boolean         'Helper verdict

    Const PROC As String = "TST_Case_ActiveFramePairRefusesMismatch"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log PROC, "START", _
            "Validating that a disagreed active-frame pair is refused"

        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_NONE

'------------------------------------------------------------------------------
' BRANCH 1 - AN AGREED PAIR RESOLVES
'------------------------------------------------------------------------------
        Set ResolvedWindow = Nothing
        ResolvedHwnd = 0
        FailMsg = vbNullString

        OK = UI_TryGetActiveFramePair(ResolvedWindow, ResolvedHwnd, FailMsg)

        TST_AssertBooleanEquals True, OK, PROC & ".Agreed.Resolved"
        TST_AssertBooleanEquals _
            False, (ResolvedWindow Is Nothing), PROC & ".Agreed.WindowSet"
        TST_AssertBooleanEquals _
            False, (ResolvedHwnd = 0), PROC & ".Agreed.HandleSet"
        TST_AssertBooleanEquals _
            True, (Len(FailMsg) = 0), PROC & ".Agreed.NoFailureText"

'------------------------------------------------------------------------------
' BRANCH 2 - A DISAGREED PAIR IS REFUSED, AND NOTHING IS RETURNED
'------------------------------------------------------------------------------
    'Returning a handle alongside a refusal would invite a caller to use it
        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_MISMATCH

        Set ResolvedWindow = Application.ActiveWindow
        ResolvedHwnd = 1
        FailMsg = vbNullString

        OK = UI_TryGetActiveFramePair(ResolvedWindow, ResolvedHwnd, FailMsg)

        TST_AssertBooleanEquals False, OK, PROC & ".Disagreed.Refused"
        TST_AssertBooleanEquals _
            True, (ResolvedWindow Is Nothing), PROC & ".Disagreed.WindowCleared"
        TST_AssertBooleanEquals _
            True, (ResolvedHwnd = 0), PROC & ".Disagreed.HandleCleared"
        TST_AssertBooleanEquals _
            True, _
            (FailMsg = "the active Excel Window is not the active native frame"), _
            PROC & ".Disagreed.FailureText"

'------------------------------------------------------------------------------
' BRANCH 3 - THE SEAM IS ONE-SHOT
'------------------------------------------------------------------------------
    'A fault that survived would refuse every later resolution in the session,
    'and the symptom would look like the defect rather than like the test
        Set ResolvedWindow = Nothing
        ResolvedHwnd = 0
        FailMsg = vbNullString

        OK = UI_TryGetActiveFramePair(ResolvedWindow, ResolvedHwnd, FailMsg)

        TST_AssertBooleanEquals True, OK, PROC & ".OneShot.SecondResolves"

'------------------------------------------------------------------------------
' BRANCH 4 - THE HANDLE-ONLY WRAPPER REPORTS THE SAME REFUSAL
'------------------------------------------------------------------------------
    'The wrapper delegates, so a refusal must reach its callers unchanged
        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_MISMATCH

        ResolvedHwnd = 1
        FailMsg = vbNullString

        OK = UI_TryGetActiveTitleBarHwnd(ResolvedHwnd, FailMsg)

        TST_AssertBooleanEquals False, OK, PROC & ".Delegated.Refused"
        TST_AssertBooleanEquals _
            True, (ResolvedHwnd = 0), PROC & ".Delegated.HandleCleared"
        TST_AssertBooleanEquals _
            True, _
            (FailMsg = "the active Excel Window is not the active native frame"), _
            PROC & ".Delegated.FailureText"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_NONE

        TST_Log PROC, "PASS", _
            "A disagreed active-frame pair was refused by both callers"

End Sub


Public Sub TST_Case_TitleBarApiErrorCapture()

'
'==============================================================================
' TST_Case_TitleBarApiErrorCapture
'------------------------------------------------------------------------------
' PURPOSE
'   Verify how the title-bar module reads a zero GWL_STYLE result: a valid
'   zero is accepted, a captured DLL error is reported with the API,
'   operation, error and hWnd, and an error left by an earlier call is neither
'   mistaken for a failure nor carried into a later report.
'
' WHY THIS EXISTS
'   GetWindowLong[Ptr] and SetWindowLong[Ptr] return zero both for a valid
'   zero and for a failure. Only Err.LastDllError, cleared before the call and
'   captured immediately after it, tells the two apart (ICR-UI-112-P3-05,
'   #27). Windows produces none of the three branches on demand, so they are
'   reached through UI_InternalInjectWindowStyleFault.
'
' RETURNS
'   None
'
' BEHAVIOR
'   1  read, no fault          succeeds with the caption visible
'   2  read, valid zero        succeeds, no caption, no failure text
'   3  seam is one-shot        the next read sees the real style again
'   4  read, DLL error 1400    refused with the exact failure text
'   5  read, stale error 2     the clear removes it: a valid zero again
'   6  read, DLL error 5       a real read first, then the exact text names
'                              5 and nothing an earlier branch left
'   7  write, valid zero       a hide succeeds and the frame is hidden
'   8  write, DLL error 1400   refused with the exact failure text
'
' ERROR POLICY
'   - Raises after disarming the seam, restoring the entry style and
'     discarding the frame-state registry
'   - Fails, with the restoration diagnostic, when the entry style cannot be
'     written, refreshed and read back
'
' DEPENDENCIES
'   - UI_InternalInjectWindowStyleFault
'   - UI_InternalResetTitleBarBaseline
'   - UI_TryGetTitleBarVisibleForHwnd
'   - UI_TrySetTitleBarVisibleForHwndIfNeeded
'   - TST_TryGetWindowStyle
'   - TST_RestoreTitleBarStyle
'   - TST_TitleBarVisibleForHwndOrRaise
'   - TST_BuildApiFailureText
'
' CALLED FROM
'   - TST_RunRegressionPack
'   - TST_RunTitleBarOnlyPack
'
' NOTES
'   The seam replaces only what the capture took, after it took it, so the
'   real API call still runs. Branch 8 has therefore really hidden the frame
'   while reporting a failure, and the registry never recorded that write.
'   The case puts the visible style back through the harness, asserts it,
'   and discards the registry. Cleanup restores the style the frame had
'   before the case began, on the success and failure paths alike.
'
'   The expected text is built here from the harness's own API-name
'   constants, so a module that quoted another function, or the wrong
'   operation, error or handle, fails the exact comparison.
'
' UPDATED
'   2026-11-06 - Checked both entry-style restorations (#66).
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim TargetHwnd          As LongPtr         'Frame under test
    Dim EntryStyle          As LongPtr         'Style observed on entry
    Dim VisibleStyle        As LongPtr         'Style once the frame is shown
#Else
    Dim TargetHwnd          As Long            'Frame under test
    Dim EntryStyle          As Long            'Style observed on entry
    Dim VisibleStyle        As Long            'Style once the frame is shown
#End If

    Dim StyleCaptured       As Boolean         'TRUE once EntryStyle is usable
    Dim IsVisible           As Boolean         'Visibility the read reports
    Dim OK                  As Boolean         'Helper verdict
    Dim Msg                 As String          'Failure text the helper reports
    Dim Expected            As String          'Exact failure text expected
    Dim RestoreOK           As Boolean         'TRUE once the style is back
    Dim RestoreMsg          As String          'Diagnostic from the restoration

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_TitleBarApiErrorCapture"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating WinAPI error capture for the GWL_STYLE read and write"

'------------------------------------------------------------------------------
' CAPTURE THE ENTRY STYLE
'------------------------------------------------------------------------------
        TargetHwnd = Application.hWnd

        If TargetHwnd = 0 Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 80, _
                PROC, _
                "no Excel window handle is available"
        End If

    'Retain the entry style before anything changes it, so cleanup can put
    'the frame back exactly as the case found it
        If Not TST_TryGetWindowStyle(TargetHwnd, EntryStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 81, _
                PROC, _
                "could not read the entry window style | " & Msg
        End If

        StyleCaptured = True

'------------------------------------------------------------------------------
' PREPARE A KNOWN FRAME STATE
'------------------------------------------------------------------------------
    'Start from a visible frame, an empty registry and a disarmed seam
        UI_SetExcelUI TitleBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        UI_InternalResetTitleBarBaseline

    'The calls below address the frame the entry style was read from
        If Application.hWnd <> TargetHwnd Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 83, _
                PROC, _
                "the active frame changed while the case was being prepared"
        End If

    'Retain the visible style so branch 8 can be undone exactly
        If Not TST_TryGetWindowStyle(TargetHwnd, VisibleStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 84, _
                PROC, _
                "could not read the visible window style | " & Msg
        End If

'------------------------------------------------------------------------------
' BRANCH 1 - AN UNFAULTED READ SUCCEEDS
'------------------------------------------------------------------------------
        OK = UI_TryGetTitleBarVisibleForHwnd(TargetHwnd, IsVisible, Msg)

        TST_AssertBooleanEquals True, OK, PROC & ".Real.Read"
        TST_AssertBooleanEquals True, IsVisible, PROC & ".Real.Visible"
        TST_AssertBooleanEquals _
            True, (Len(Msg) = 0), PROC & ".Real.NoFailureText"

'------------------------------------------------------------------------------
' BRANCH 2 - A ZERO STYLE WITH NO CAPTURED ERROR IS VALID
'------------------------------------------------------------------------------
    'Zero with a cleared last-error is a real value under the Windows contract
        UI_InternalInjectWindowStyleFault _
            TEST_STYLE_OP_READ, TEST_STYLE_FAULT_VALID_ZERO, 0

        OK = UI_TryGetTitleBarVisibleForHwnd(TargetHwnd, IsVisible, Msg)

        TST_AssertBooleanEquals True, OK, PROC & ".ValidZero.Read"
        TST_AssertBooleanEquals _
            False, IsVisible, PROC & ".ValidZero.NoCaption"
        TST_AssertBooleanEquals _
            True, (Len(Msg) = 0), PROC & ".ValidZero.NoFailureText"

'------------------------------------------------------------------------------
' BRANCH 3 - THE SEAM IS ONE-SHOT
'------------------------------------------------------------------------------
    'A fault that survived would fail every later read in the session
        OK = UI_TryGetTitleBarVisibleForHwnd(TargetHwnd, IsVisible, Msg)

        TST_AssertBooleanEquals True, OK, PROC & ".OneShot.Read"
        TST_AssertBooleanEquals True, IsVisible, PROC & ".OneShot.Visible"

'------------------------------------------------------------------------------
' BRANCH 4 - A ZERO STYLE WITH A CAPTURED ERROR IS A FAILURE
'------------------------------------------------------------------------------
    'The text must name the API that ran, the operation, the error and the
    'window, and is compared exactly
        UI_InternalInjectWindowStyleFault _
            TEST_STYLE_OP_READ, _
            TEST_STYLE_FAULT_DLL_ERROR, _
            TST_ERROR_INVALID_WINDOW_HANDLE

        OK = UI_TryGetTitleBarVisibleForHwnd(TargetHwnd, IsVisible, Msg)

        Expected = TST_BuildApiFailureText( _
            TST_API_GET_WINDOW_STYLE, _
            "read GWL_STYLE", _
            TST_ERROR_INVALID_WINDOW_HANDLE, _
            TargetHwnd)

        TST_Log PROC, "OBSERVED", "DllError.Read | " & Msg

        TST_AssertBooleanEquals False, OK, PROC & ".DllError.Refused"
        TST_AssertTrue (Msg = Expected), PROC & ".DllError.FailureText"

'------------------------------------------------------------------------------
' BRANCH 5 - AN ERROR LEFT BY AN EARLIER CALL IS CLEARED, NOT REPORTED
'------------------------------------------------------------------------------
    'The seam leaves the error in the thread last-error before the clear. A
    'read that captured without clearing first would refuse this valid zero.
        UI_InternalInjectWindowStyleFault _
            TEST_STYLE_OP_READ, _
            TEST_STYLE_FAULT_STALE_ERROR, _
            TST_ERROR_FILE_NOT_FOUND

        OK = UI_TryGetTitleBarVisibleForHwnd(TargetHwnd, IsVisible, Msg)

        TST_Log PROC, "OBSERVED", "StaleError.Read | " & Msg

        TST_AssertBooleanEquals True, OK, PROC & ".StaleError.Read"
        TST_AssertBooleanEquals _
            False, IsVisible, PROC & ".StaleError.NoCaption"
        TST_AssertBooleanEquals _
            True, (Len(Msg) = 0), PROC & ".StaleError.NoFailureText"

'------------------------------------------------------------------------------
' BRANCH 6 - A LATER FAILURE REPORTS ITS OWN ERROR
'------------------------------------------------------------------------------
    'A real, non-zero read first: it must succeed and report nothing
        OK = UI_TryGetTitleBarVisibleForHwnd(TargetHwnd, IsVisible, Msg)

        TST_AssertBooleanEquals True, OK, PROC & ".LaterFailure.RealRead"
        TST_AssertBooleanEquals _
            True, (Len(Msg) = 0), PROC & ".LaterFailure.NoFailureText"

    'Then a failure with a different error. The exact comparison excludes the
    'errors earlier branches supplied.
        UI_InternalInjectWindowStyleFault _
            TEST_STYLE_OP_READ, _
            TEST_STYLE_FAULT_DLL_ERROR, _
            TST_ERROR_ACCESS_DENIED

        OK = UI_TryGetTitleBarVisibleForHwnd(TargetHwnd, IsVisible, Msg)

        Expected = TST_BuildApiFailureText( _
            TST_API_GET_WINDOW_STYLE, _
            "read GWL_STYLE", _
            TST_ERROR_ACCESS_DENIED, _
            TargetHwnd)

        TST_Log PROC, "OBSERVED", "LaterFailure.Read | " & Msg

        TST_AssertBooleanEquals False, OK, PROC & ".LaterFailure.Refused"
        TST_AssertTrue (Msg = Expected), PROC & ".LaterFailure.FailureText"

'------------------------------------------------------------------------------
' BRANCH 7 - A ZERO PREVIOUS STYLE FROM A WRITE IS VALID
'------------------------------------------------------------------------------
    'SetWindowLong returns the previous style, and zero is a legal one. The
    'write fault waits through the read that precedes the write.
        UI_InternalInjectWindowStyleFault _
            TEST_STYLE_OP_WRITE, TEST_STYLE_FAULT_VALID_ZERO, 0

        OK = UI_TrySetTitleBarVisibleForHwndIfNeeded( _
            TargetHwnd:=TargetHwnd, _
            IsVisible:=False, _
            FailMsg:=Msg)

        TST_Log PROC, "OBSERVED", "ValidZero.Write | " & Msg

        TST_AssertBooleanEquals True, OK, PROC & ".WriteValidZero.Hide"
        TST_AssertBooleanEquals _
            True, (Len(Msg) = 0), PROC & ".WriteValidZero.NoFailureText"

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertBooleanEquals _
            False, _
            TST_TitleBarVisibleForHwndOrRaise(TargetHwnd, PROC), _
            PROC & ".WriteValidZero.Hidden"

    'Show again unfaulted, so branch 8 starts from a visible frame
        OK = UI_TrySetTitleBarVisibleForHwndIfNeeded( _
            TargetHwnd:=TargetHwnd, _
            IsVisible:=True, _
            FailMsg:=Msg)

        TST_AssertBooleanEquals True, OK, PROC & ".WriteValidZero.Show"

        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' BRANCH 8 - A ZERO PREVIOUS STYLE WITH A CAPTURED ERROR IS A FAILURE
'------------------------------------------------------------------------------
        UI_InternalInjectWindowStyleFault _
            TEST_STYLE_OP_WRITE, _
            TEST_STYLE_FAULT_DLL_ERROR, _
            TST_ERROR_INVALID_WINDOW_HANDLE

        OK = UI_TrySetTitleBarVisibleForHwndIfNeeded( _
            TargetHwnd:=TargetHwnd, _
            IsVisible:=False, _
            FailMsg:=Msg)

        Expected = TST_BuildApiFailureText( _
            TST_API_SET_WINDOW_STYLE, _
            "write GWL_STYLE", _
            TST_ERROR_INVALID_WINDOW_HANDLE, _
            TargetHwnd)

        TST_Log PROC, "OBSERVED", "DllError.Write | " & Msg

        TST_AssertBooleanEquals False, OK, PROC & ".WriteDllError.Refused"
        TST_AssertTrue (Msg = Expected), PROC & ".WriteDllError.FailureText"

'------------------------------------------------------------------------------
' RESTORE ENTRY STYLE
'------------------------------------------------------------------------------
    'The faulted write really hid the frame and the registry never recorded
    'it. Put the visible style back and forget the registry.
        If Not TST_RestoreTitleBarStyle( _
            TargetHwnd, VisibleStyle, StyleCaptured, Msg) Then

            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 82, _
                PROC, _
                "could not restore the visible window style | " & Msg
        End If

        UI_InternalResetTitleBarBaseline

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertBooleanEquals _
            True, _
            TST_TitleBarVisibleForHwndOrRaise(TargetHwnd, PROC), _
            PROC & ".Restored.Visible"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Valid zero, captured DLL error and stale error were each read correctly"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Disarm the seam, put the frame back and forget the registry, whatever
    'happened above
        On Error Resume Next
            UI_InternalInjectWindowStyleFault _
                TEST_STYLE_OP_NONE, TEST_STYLE_FAULT_NONE, 0
            RestoreOK = TST_RestoreTitleBarStyle( _
                TargetHwnd, EntryStyle, StyleCaptured, RestoreMsg)
            UI_InternalResetTitleBarBaseline
        On Error GoTo 0

    'A frame that cannot be put back fails the case, after any earlier failure
        TST_RecordCleanupFailure _
            RestoreOK, RestoreMsg, PROC, FailNumber, FailSource, FailDescription

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
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub
