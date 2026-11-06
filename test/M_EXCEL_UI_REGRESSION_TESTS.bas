Attribute VB_Name = "M_EXCEL_UI_REGRESSION_TESTS"
Option Explicit
Option Private Module

'==============================================================================
' M_EXCEL_UI_REGRESSION_TESTS
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
'     - window-target scope behavior
'
'   A repeatable regression harness reduces the risk of silent regressions and
'   makes the repository more maintainable and release-ready
'
' PUBLIC SURFACE
'   - Test_EXCEL_UI_RunAll
'   - Test_EXCEL_UI_RunCore
'   - Test_EXCEL_UI_RunTitleBarOnly
'   - Test_EXCEL_UI_RunSnapshotIdentity
'   - Test_EXCEL_UI_RunTitleBarSdiIdentity
'   - Test_EXCEL_UI_RunReleaseCertification
'   - Test_EXCEL_UI_RunCertificationSelfTest
'   - Test_EXCEL_UI_RunOwnershipCleanupChecks
'   - Test_EXCEL_UI_RunRibbonSdiProbe
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
'     - invalid UIWindowTargetScope structured failure path
'     - snapshot capture clean-success result path
'     - snapshot restoration clean-success result path
'     - snapshot restoration no-snapshot failure path
'     - closed captured-window ordered failure path
'
'   Target-scope tests
'     - active-window-only application
'     - active-workbook-window application
'     - invalid-scope failure with application-level continuation
'
'   Snapshot / restore tests
'     - explicit snapshot lifecycle
'     - reset without snapshot leaves managed UI unchanged and logs a diagnostic
'     - identity-safe restoration of surviving captured windows
'     - closed captured-window handling
'     - replacement-window non-interference
'     - lifecycle cases are skipped when an explicit EXCEL_UI snapshot already
'       existed before the run because the harness cannot reconstruct that prior
'       module-level snapshot object
'     - a certification unit that fails after a capture reports the failure
'       list it asserted on and leaves no snapshot for the next unit
'
'   Environment-preservation tests
'     - ScreenUpdating restored to prior state
'
'   Title-bar tests
'     - hide / show round-trip
'     - preservation of unrelated GWL_STYLE bits across hide / show
'     - WinAPI error capture: a valid zero, an injected DLL error and an
'       error left by an earlier call, for the GWL_STYLE read and write
'     - every title-bar refusal, through the public snapshot, restore and
'       show/hide entry points, leaves the active frame's complete GWL_STYLE
'       unchanged
'
' STATE MANAGEMENT
'   - The harness snapshots the current managed Excel UI state before testing
'   - The harness attempts a best-effort restore after success or failure
'   - Per-window test state is captured and restored by Application.Windows index
'   - A pre-existing explicit EXCEL_UI snapshot is left untouched by skipping
'     the snapshot-destructive lifecycle cases
'   - Every case that hides a title bar, the two packs and the SDI identity
'     runner put the entry frame style back through the harness WinAPI
'     helpers, never through the API under test, and fail when they cannot
'   - The focused snapshot-identity runner creates and closes temporary windows
'     only after confirming that no explicit EXCEL_UI snapshot already exists
'
' LIMITATIONS
'   - Ribbon visibility is read using best-effort logic
'   - Window-level capture and restore use Application.Windows index order; a
'     changed collection can prevent exact restoration to the original windows
'   - The harness cannot preserve and reconstruct a pre-existing core-module
'     snapshot object, so destructive snapshot cases are skipped in that state
'   - Title-bar behavior remains the most OS / Excel-version-sensitive area
'
' COMPATIBILITY
'   - Windows only for title-bar validation
'   - Assumes the EXCEL_UI module is present in the same VBA project
'
' UPDATED
'   2026-11-06 - TST_RestoreTitleBarStyle reports and verifies the restore.
'                Every title-bar case that hides a frame restores the entry
'                style through it on every exit and fails when it cannot;
'                the packs and the SDI identity runner do the same for the
'                entry frame. TST_WaitUI is capped (#66).
'   2026-11-06 - Title-bar refusal cases compare the active frame's complete
'                style before and after the refused call. Added the
'                public-path pair-refusal and closed-and-reopened-frame cases
'                to the SDI identity runner (#45).
'   2026-11-06 - The title-bar WinAPI helpers capture Err.LastDllError
'                immediately after each call and no longer declare
'                GetLastError. Added the WinAPI error-capture case to both
'                title-bar runs (#27).
'   2026-11-06 - Structured-result assertions report the failure count and
'                list they rejected, and certification releases a snapshot a
'                unit left behind before the next unit runs, reporting it as
'                a cleanup failure (#135).
'   2026-11-06 - Header and comments aligned with the code (#114).
'   2026-08-21 - Registered every frame-state case in the pack certification
'                runs, added the stale-frame-entry and certification-cleanup
'                cases, and stopped certification destroying the error it
'                re-raises.
'   2026-08-19 - Added the certification runner, the SDI title-bar identity
'                pack and the frame-refresh-debt case.
'   2026-08-18 - Reformatted to the project house style. No behavior change.
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
    Private Const TEST_WAIT_SECONDS   As Double = 0.15                'Small UI settle delay after each state change
    Private Const TEST_WAIT_MAX_SECONDS As Double = 5#                'Upper bound on any one settle delay
    Private Const TEST_ERR_BASE       As Long = vbObjectError + 4700  'Base custom error number for test assertions
    Private Const TEST_SNAPSHOT_ID_ERR_BASE As Long = vbObjectError + 4810  'Base custom error for snapshot-identity assertions
    Private Const TEST_TARGET_ERR_BASE As Long = vbObjectError + 4900  'Base custom error for target-scope tests
    Private Const TEST_TITLEBAR_SDI_ERR_BASE As Long = vbObjectError + 5000  'Base custom error for title-bar SDI tests
    Private Const TEST_CERT_ERR_BASE  As Long = vbObjectError + 5100  'Base custom error for certification
    Private Const TEST_RIBBON_ERR_BASE As Long = vbObjectError + 5200  'Base custom error for the Ribbon probe
    Private Const TEST_OWNERSHIP_ERR_BASE As Long = vbObjectError + 5300  'Base custom error for the snapshot-ownership case

    Private Const TEST_PAIR_FAULT_NONE As Long = 0  'Frame pair: no fault
    Private Const TEST_PAIR_FAULT_MISMATCH As Long = 1  'Frame pair: sampled handle disagrees

    Private Const TEST_QUIET_FAULT_NONE As Long = 0  'Quiet scope: no fault
    Private Const TEST_QUIET_FAULT_ENTRY As Long = 1  'Quiet scope: entry read fails
    Private Const TEST_QUIET_FAULT_WRITE As Long = 2  'Quiet scope: write ineffective
    Private Const TEST_QUIET_FAULT_READBACK As Long = 3  'Quiet scope: readback fails

    Private Const TEST_STYLE_OP_NONE As Long = 0  'Window-style seam: no operation
    Private Const TEST_STYLE_OP_READ As Long = 1  'Window-style seam: GWL_STYLE read
    Private Const TEST_STYLE_OP_WRITE As Long = 2  'Window-style seam: GWL_STYLE write
    Private Const TEST_STYLE_FAULT_NONE As Long = 0  'Window-style seam: no fault
    Private Const TEST_STYLE_FAULT_VALID_ZERO As Long = 1  'Zero result, captured error kept
    Private Const TEST_STYLE_FAULT_DLL_ERROR As Long = 2  'Zero result, armed error captured
    Private Const TEST_STYLE_FAULT_STALE_ERROR As Long = 3  'Armed error left before the clear

    Private Const TEST_FAULT_NONE As Long = 0  'No post-capture fault armed
    Private Const TEST_FAULT_ASSERTION As Long = 1  'Fail an assertion after capture
    Private Const TEST_FAULT_UNEXPECTED As Long = 2  'Raise an unexpected error after capture
    Private Const TEST_WS_CAPTION     As Long = &HC00000              'Caption bit read by the per-window helper
    Private Const TST_SECONDS_PER_DAY As Double = 86400#              'Timer rollover interval in seconds

'------------------------------------------------------------------------------
' WIN32 / WIN64 API FOR TITLE-BAR STYLE TESTS
'------------------------------------------------------------------------------
    #If VBA7 Then
        #If Win64 Then
            Private Declare PtrSafe Function TST_GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
                ByVal hWnd As LongPtr, _
                ByVal nIndex As Long) _
                As LongPtr

            Private Declare PtrSafe Function TST_SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
                ByVal hWnd As LongPtr, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As LongPtr) _
                As LongPtr
        #Else
            Private Declare PtrSafe Function TST_GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As LongPtr, _
                ByVal nIndex As Long) _
                As Long

            Private Declare PtrSafe Function TST_SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hWnd As LongPtr, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) _
                As Long
        #End If

        Private Declare PtrSafe Function TST_SetWindowPos Lib "user32" Alias _
            "SetWindowPos" ( _
            ByVal hWnd As LongPtr, _
            ByVal hWndInsertAfter As LongPtr, _
            ByVal X As Long, _
            ByVal Y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal uFlags As Long) _
            As Long

        Private Declare PtrSafe Sub TST_SetLastError Lib "kernel32" Alias "SetLastError" ( _
            ByVal dwErrCode As Long)
    #Else
        Private Declare Function TST_GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
            ByVal hWnd As Long, _
            ByVal nIndex As Long) _
            As Long

        Private Declare Function TST_SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
            ByVal hWnd As Long, _
            ByVal nIndex As Long, _
            ByVal dwNewLong As Long) _
            As Long

        Private Declare Function TST_SetWindowPos Lib "user32" Alias "SetWindowPos" ( _
            ByVal hWnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal X As Long, _
            ByVal Y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal uFlags As Long) _
            As Long

        Private Declare Sub TST_SetLastError Lib "kernel32" Alias "SetLastError" ( _
            ByVal dwErrCode As Long)
    #End If

'------------------------------------------------------------------------------
' API CONSTANTS FOR TITLE-BAR STYLE TESTS
'------------------------------------------------------------------------------
    Private Const TST_GWL_STYLE             As Long = -16
    Private Const TST_WS_CAPTION            As Long = &HC00000
    Private Const TST_SYNTHETIC_UNRELATED_BIT As Long = &H2000000
    Private Const TST_TITLEBAR_OWNED_MASK   As Long = &HCF0000

    'A legal caption frame that is not the full owned set. Written behind the
    'component to contradict an entry claiming the frame is hidden, which is
    'the evidence a reissued handle presents to the registry.
    Private Const TST_TITLEBAR_FOREIGN_FRAME As Long = &HC80000
    Private Const TST_TITLEBAR_CAPTIONLESS_FRAME As Long = &H40000

    Private Const TST_SWP_NOSIZE            As Long = &H1
    Private Const TST_SWP_NOMOVE            As Long = &H2
    Private Const TST_SWP_NOZORDER          As Long = &H4
    Private Const TST_SWP_FRAMECHANGED      As Long = &H20
    Private Const TST_SWP_NOOWNERZORDER     As Long = &H200

    'Exported names of the GWL_STYLE read and write in this compilation. The
    'harness quotes them in its own failure text, and the error-capture case
    'expects the title-bar module to quote the same names.
    #If VBA7 Then
        #If Win64 Then
            Private Const TST_API_GET_WINDOW_STYLE As String = "GetWindowLongPtrA"
            Private Const TST_API_SET_WINDOW_STYLE As String = "SetWindowLongPtrA"
        #Else
            Private Const TST_API_GET_WINDOW_STYLE As String = "GetWindowLongA"
            Private Const TST_API_SET_WINDOW_STYLE As String = "SetWindowLongA"
        #End If
    #Else
        Private Const TST_API_GET_WINDOW_STYLE As String = "GetWindowLongA"
        Private Const TST_API_SET_WINDOW_STYLE As String = "SetWindowLongA"
    #End If

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

'==============================================================================
' RELEASE-CERTIFICATION STATE
'==============================================================================

'Accrued only while a certification run is active. The regression packs are
'shared with the legacy runners, so recording must be inert outside a run rather
'than leaking counts from one invocation into the next.
Private m_CertActive                    As Boolean

Private m_CertUnitCount                 As Long
Private m_CertUnitNames()               As String
Private m_CertUnitPassed()              As Boolean
Private m_CertUnitDetail()              As String

Private m_CertSkipCount                 As Long
Private m_CertSkipDetail()              As String

'Snapshots that units left behind, named by unit. Each one is released before
'the next unit runs, so that one failure cannot fail every unit after it. The
'exit state therefore no longer shows the leak, and the cleanup verdict reads
'it from here instead.
Private m_CertLeakDetail                As String

'Ribbon characterization observations, accumulated during a probe run. Rows are
'plain text for reading and JSON for comparison between hosts.
'Post-capture fault injection, consumed by the first runner that reaches its
'injection point. One-shot on purpose: a flag that stayed set would fault every
'later runner in the same session, and the cleanup being proved here is exactly
'what would be running while the operator wondered why.
'
'Only TST_InjectPostCapture sets it, and only the cleanup runner calls that.
'The seam exists because ownership cleanup after a post-capture failure cannot
'be reached any other way: the failures it must survive are ones the runners
'are built not to have.
Private m_InjectFault                   As Long

Private m_RibbonRowCount                As Long
Private m_RibbonRowsText                As String
Private m_RibbonRowsJson                As String


Private Sub TST_InjectPostCapture(ByVal FaultKind As Long)

'
'==============================================================================
' TST_InjectPostCapture
'------------------------------------------------------------------------------
' PURPOSE
'   Arm a one-shot post-capture fault for the next runner that reaches its
'   injection point.
'
' WHY THIS EXISTS
'   A runner that owns a snapshot must release it when something fails between
'   capture and completion. Those failures are the ones the runner is written
'   not to have, so the path is unreachable without injecting one, and an
'   unreachable cleanup path is where the original defect lived.
'
' INPUTS
'   FaultKind
'     TEST_FAULT_NONE, TEST_FAULT_ASSERTION or TEST_FAULT_UNEXPECTED
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Cannot raise
'
' CALLED FROM
'   - Test_EXCEL_UI_RunOwnershipCleanupChecks
'
' UPDATED
'   2026-08-29
'==============================================================================
'

        m_InjectFault = FaultKind

End Sub


Private Function TST_TakePostCaptureFault() As Long

'
'==============================================================================
' TST_TakePostCaptureFault
'------------------------------------------------------------------------------
' PURPOSE
'   Return the armed fault and disarm it in the same step.
'
' WHY THIS EXISTS
'   Reading and clearing separately would leave the flag set on any path that
'   read it and then failed, and every caller reads it immediately before
'   failing on purpose.
'
' RETURNS
'   The armed fault kind, or TEST_FAULT_NONE
'
' ERROR POLICY
'   - Cannot raise
'
' CALLED FROM
'   - Test_EXCEL_UI_RunCertificationSelfTest
'   - Test_EXCEL_UI_RunSnapshotIdentity
'
' UPDATED
'   2026-08-29
'==============================================================================
'

        TST_TakePostCaptureFault = m_InjectFault
        m_InjectFault = TEST_FAULT_NONE

End Function


Public Sub Test_EXCEL_UI_RunCertificationSelfTest()

'
'==============================================================================
' Test_EXCEL_UI_RunCertificationSelfTest
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a failure inside release certification reaches the caller with
'   its error number and description intact.
'
' WHY THIS EXISTS
'   The certification error handler calls TST_Log, which contains
'   On Error Resume Next and therefore clears Err. Reading Err after that call
'   yields zero, and Err.Raise 0 does not raise: a failed certification returned
'   silently, and a programmatic caller saw a normal return.
'
'   The log line still said FAIL, which is why the defect survived a real failed
'   run without being noticed. Only the raise was missing, and nothing asserted
'   the raise.
'
' WHY IT IS NOT PART OF THE PACK
'   The only errors that travel through the certification handler are raised
'   after the handler is armed, which means the run has already reset the
'   counters and set m_CertActive. Triggering one from inside the regression
'   pack would corrupt the accounting of the very run executing it, and the
'   re-entrancy guard deliberately refuses a nested invocation before the
'   handler is reached, so the handler path cannot be exercised from within a
'   certification run at all.
'
'   It is therefore a standalone runner, invoked directly.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Refuses to run while an explicit snapshot exists.
'   - Establishes a snapshot so that certification rejects its precondition.
'   - Invokes certification and captures what it raises.
'   - Asserts the error number and a non-empty description survived.
'   - Clears the snapshot it created.
'
' ERROR POLICY
'   - Raises on assertion failure, after cleanup.
'
' NOTES
'   Asserting a non-empty description matters as much as the number. The
'   original defect produced zero and an empty string together, so a test
'   checking only that something was raised would have passed once the number
'   was fixed and the text still lost.
'
'   The snapshot this runner captures replaces any existing one outright, and
'   the cleanup discards what it captured, so running while the caller held a
'   snapshot would destroy it silently. The precondition is refused instead,
'   which is what every other destructive runner in this module does.
'
' UPDATED
'   2026-08-21 - Refused to run while an explicit snapshot exists, rather than
'                replacing and then discarding it.
'   2026-08-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RaisedNumber        As Long            'Error number reaching the caller
    Dim RaisedDescription   As String          'Error text reaching the caller

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Dim CreatedSnapshot     As Boolean         'TRUE once this run owns one

    Const PROC As String = "Test_EXCEL_UI_RunCertificationSelfTest"

'------------------------------------------------------------------------------
' VALIDATE PRECONDITIONS
'------------------------------------------------------------------------------
    'Refuse to destroy a snapshot the caller is relying on. Capturing below
    'replaces any existing snapshot outright, and the cleanup discards what it
    'captured, so a caller who already held one would lose it with nothing said.
    '
    'This runs before the handler is armed, deliberately. Armed first, the
    'refusal travelled Err_Handler, resumed at Safe_Exit, and the unconditional
    'clear there deleted the very snapshot being refused - the caller received
    'the refusal and lost the state anyway. Raising here propagates the number,
    'source and description to the caller untouched, and nothing has been
    'acquired yet that would need releasing.
        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_CERT_ERR_BASE + 21, _
                PROC, _
                "an explicit EXCEL_UI snapshot already exists; clear or " & _
                "restore it before running this destructive test"
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        CreatedSnapshot = False

        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a certification failure reaches the caller"

'------------------------------------------------------------------------------
' ESTABLISH A REJECTED PRECONDITION
'------------------------------------------------------------------------------
    'Certification refuses to start when an explicit snapshot exists. That
    'refusal is raised after the handler is armed, so it travels the path under
    'test.
        UI_CaptureExcelUIState
        TST_WaitUI TEST_WAIT_SECONDS

    'Ownership is taken from an observed post-capture state, not from the call
    'having returned. A capture that silently left no snapshot must not make
    'this run believe it owns one.
        CreatedSnapshot = UI_HasExcelUIStateSnapshot

        If Not CreatedSnapshot Then
            Err.Raise _
                TEST_CERT_ERR_BASE + 20, _
                PROC, _
                "unable to establish a snapshot; the precondition under test " & _
                "cannot be triggered"
        End If

    'Post-capture fault injection. Nothing arms this during an ordinary run, so
    'the ordinary run is unchanged. Armed, it exercises the one path that
    'cannot otherwise be reached: a failure between owning a snapshot and
    'releasing it.
        Select Case TST_TakePostCaptureFault()

            Case TEST_FAULT_ASSERTION
                TST_AssertTrue _
                    ActualValue:=False, _
                    AssertionName:=PROC & ".InjectedAssertion"

            Case TEST_FAULT_UNEXPECTED
                Err.Raise _
                    TEST_OWNERSHIP_ERR_BASE + 1, _
                    PROC & ".InjectedUnexpected", _
                    "injected unexpected post-capture failure"

        End Select

'------------------------------------------------------------------------------
' INVOKE AND CAPTURE
'------------------------------------------------------------------------------
    'Capture immediately, before anything else can clear Err
        On Error Resume Next

            Test_EXCEL_UI_RunReleaseCertification

            RaisedNumber = Err.Number
            RaisedDescription = Err.Description

            Err.Clear

        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' ASSERT THE FAILURE SURVIVED
'------------------------------------------------------------------------------
    'A silent return is the defect this case exists to catch
        TST_AssertTrue _
            (RaisedNumber = TEST_CERT_ERR_BASE + 1), _
            "CertificationSelfTest.ErrorNumberPreserved"

    'Zero and an empty description were produced together by the original
    'defect, so the text is asserted as well as the number
        TST_AssertTrue _
            (Len(RaisedDescription) > 0), _
            "CertificationSelfTest.ErrorDescriptionPreserved"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Certification failure reached the caller as " & _
            CStr(RaisedNumber)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Release only a snapshot this run created. A caller-owned snapshot reaches
    'this label whenever an early failure resumes here, and clearing it would
    'destroy state this runner never owned.
        On Error Resume Next
            If CreatedSnapshot Then
                UI_ClearExcelUIStateSnapshot
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
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        HasFailure = True

        Resume Safe_Exit

End Sub


Public Sub Test_EXCEL_UI_RunReleaseCertification()

'
'==============================================================================
' Test_EXCEL_UI_RunReleaseCertification
'------------------------------------------------------------------------------
' PURPOSE
'   Execute every mandatory regression unit in one pass and emit an unambiguous
'   complete/incomplete and pass/fail verdict with machine-readable evidence.
'
' WHY THIS EXISTS
'   The existing runners cannot certify a release. Test_EXCEL_UI_RunAll does
'   not run the snapshot-identity and SDI identity units, silently skips the
'   snapshot cases when a snapshot already exists, and reports its outcome
'   only as Immediate Window prose. A
'   reader of that output cannot distinguish
'
'       a complete pass
'       a pass with snapshot cases skipped
'       a pass on one environment only
'
'   which means a green result carries far less information than it appears to.
'
'   This runner makes the difference explicit. Every unit is counted, a skipped
'   mandatory unit is a failure rather than a quiet log line, the host state is
'   verified after the run rather than assumed, and the environment the result
'   was obtained on is recorded alongside it.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Refuses to start when an explicit EXCEL_UI snapshot already exists, rather
'     than degrading into a partial run.
'   - Records the anchor window by object identity and the workbook count, so
'     leaked state is detected rather than tolerated.
'   - Runs the full regression pack, the snapshot-identity runner and the SDI
'     title-bar identity runner, each under its own error boundary so one
'     failing unit does not conceal the rest.
'   - Releases a snapshot a unit left behind before the next unit runs, so the
'     next unit is not refused because of it, and records the leak.
'   - Verifies cleanup: no snapshot left behind by any unit, ScreenUpdating
'     restored, the workbook count back to its starting value, the anchor
'     window still usable.
'   - Emits a JSON evidence document and a text summary, and writes both to the
'     temporary folder on a best-effort basis.
'   - Raises when the verdict is anything other than a complete pass.
'
' ERROR POLICY
'   - Individual unit failures are captured, not propagated.
'   - Raises once at the end when the verdict is not PASS/COMPLETE.
'
' DEPENDENCIES
'   - TST_CertResetCounters
'   - TST_CertRunUnit
'   - TST_CertBuildJsonEvidence
'   - TST_CertBuildTextReport
'   - TST_CertTryWriteEvidence
'
' NOTES
'   Destructive: creates and closes temporary workbooks and toggles every
'   managed UI element. Save unsaved work before running it.
'
' UPDATED
'   2026-11-06 - A snapshot a unit left behind fails cleanup, although it is
'                released before the next unit runs (#135).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window          'Window active before the run
    Dim BaselineBooks       As Long            'Workbooks open before the run
    Dim OldScreenUpdating   As Boolean         'ScreenUpdating before the run

    Dim UnitsFailed         As Long            'Units that reported failure
    Dim CleanupOK           As Boolean         'TRUE when no state leaked
    Dim CleanupDetail       As String          'Reason cleanup was rejected

    Dim Complete            As Boolean         'TRUE when nothing was skipped
    Dim Passed              As Boolean         'TRUE when the verdict is a pass
    Dim Verdict             As String          'Human-readable verdict line

    Dim JsonText            As String          'Machine-readable evidence
    Dim ReportText          As String          'Human-readable evidence
    Dim JsonPath            As String          'Path the JSON was written to
    Dim ReportPath          As String          'Path the report was written to
    Dim ScanIdx             As Long            'Cursor over recorded units

    Dim FailNumber          As Long            'Error number captured on entry
    Dim FailSource          As String          'Error source captured on entry
    Dim FailDescription     As String          'Error text captured on entry
    Dim FailLine            As Long            'Error line captured on entry

    Const PROC As String = "Test_EXCEL_UI_RunReleaseCertification"

'------------------------------------------------------------------------------
' GUARD RE-ENTRY
'------------------------------------------------------------------------------
    'Refuse a nested run BEFORE arming the handler or touching any counter. A
    'nested invocation would reset the outer run's unit records and clear
    'm_CertActive on exit, leaving the outer verdict describing work it never
    'did. Raising here, ahead of On Error, sends the refusal straight to the
    'caller without disturbing anything.
        If m_CertActive Then
            Err.Raise _
                TEST_CERT_ERR_BASE + 4, _
                PROC, _
                "release certification is already running; nested invocation " & _
                "is not supported"
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertResetCounters

        m_CertActive = True

        TST_Log PROC, "START", "Release certification started"

'------------------------------------------------------------------------------
' VALIDATE PRECONDITIONS
'------------------------------------------------------------------------------
    'A pre-existing snapshot cannot be preserved across these units, so it is
    'rejected rather than compared. That asymmetry with the other cleanup checks
    'is deliberate: ScreenUpdating and the workbook count can be observed on
    'entry and restored, whereas a snapshot the caller depends on would be
    'destroyed by the first unit that captures one.
        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_CERT_ERR_BASE + 1, _
                PROC, _
                "an explicit EXCEL_UI snapshot already exists; clear or " & _
                "restore it before certifying a release"
        End If

        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_CERT_ERR_BASE + 2, _
                PROC, _
                "no active Excel window is available"
        End If

    'Identity and counts recorded here are what cleanup is judged against
        BaselineBooks = Workbooks.Count
        OldScreenUpdating = Application.ScreenUpdating

'------------------------------------------------------------------------------
' RUN MANDATORY UNITS
'------------------------------------------------------------------------------
    'Each unit is run under its own boundary so that one failure does not hide
    'the state of everything after it
        TST_CertRunUnit "RegressionPack"
        TST_CertRunUnit "SnapshotIdentity"
        TST_CertRunUnit "TitleBarSdiIdentity"

'------------------------------------------------------------------------------
' COUNT RESULTS
'------------------------------------------------------------------------------
    'Tally the recorded units
        For ScanIdx = 1 To m_CertUnitCount

            If Not m_CertUnitPassed(ScanIdx) Then
                UnitsFailed = UnitsFailed + 1
            End If

        Next ScanIdx

'------------------------------------------------------------------------------
' VERIFY CLEANUP
'------------------------------------------------------------------------------
    'Cleanup failure is a run failure. A suite that leaves a snapshot, a stray
    'workbook or an altered host setting behind has not finished, however many
    'assertions passed on the way.
        CleanupOK = TST_CertEvaluateCleanup( _
            BaselineBooks:=BaselineBooks, _
            BaselineScreenUpdating:=OldScreenUpdating, _
            AnchorWindow:=AnchorWindow, _
            CleanupDetail:=CleanupDetail)

    'A snapshot a unit left behind was released before the next unit ran, so
    'the exit state cannot show it. It is still a cleanup failure: the release
    'stops one defect from failing every later unit, it does not excuse it.
        If Len(m_CertLeakDetail) > 0 Then
            CleanupOK = False
            CleanupDetail = TST_CertAppendDetail(CleanupDetail, m_CertLeakDetail)
        End If

'------------------------------------------------------------------------------
' DETERMINE VERDICT
'------------------------------------------------------------------------------
    'Completeness and correctness are separate questions and are reported as
    'such. A run that skipped a mandatory unit is not a pass, whatever the
    'assertions that did execute reported.
        Complete = (m_CertSkipCount = 0)
        Passed = (UnitsFailed = 0) And CleanupOK And Complete

        Verdict = "RESULT: " & IIf(Passed, "PASS", "FAIL") & _
            " | " & IIf(Complete, "COMPLETE", "INCOMPLETE") & _
            " | units=" & CStr(m_CertUnitCount) & _
            " failed=" & CStr(UnitsFailed) & _
            " skipped=" & CStr(m_CertSkipCount) & _
            " cleanup=" & IIf(CleanupOK, "OK", "FAILED")

'------------------------------------------------------------------------------
' EMIT EVIDENCE
'------------------------------------------------------------------------------
    'Build both forms before writing either, so a file-system failure cannot
    'cost the Immediate Window record as well
        JsonText = TST_CertBuildJsonEvidence( _
            UnitsFailed:=UnitsFailed, _
            CleanupOK:=CleanupOK, _
            CleanupDetail:=CleanupDetail, _
            Complete:=Complete, _
            Passed:=Passed)

        ReportText = TST_CertBuildTextReport( _
            Verdict:=Verdict, _
            CleanupDetail:=CleanupDetail)

        TST_Log PROC, "EVIDENCE", vbNewLine & ReportText
        TST_Log PROC, "JSON", JsonText

    'File output is a convenience, never a gate on the verdict
        If TST_CertTryWriteEvidence(JsonText, "json", JsonPath) Then
            TST_Log PROC, "WROTE", JsonPath
        End If

        If TST_CertTryWriteEvidence(ReportText, "txt", ReportPath) Then
            TST_Log PROC, "WROTE", ReportPath
        End If

        TST_Log PROC, IIf(Passed, "PASS", "FAIL"), Verdict

'------------------------------------------------------------------------------
' RAISE ON A NON-PASSING VERDICT
'------------------------------------------------------------------------------
    'Raise only after the evidence exists, so a failed run is still documented
        If Not Passed Then
            Err.Raise _
                TEST_CERT_ERR_BASE + 3, _
                PROC, _
                Verdict & IIf(Len(CleanupDetail) > 0, " | " & CleanupDetail, vbNullString)
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Certification accounting must not leak into a later legacy run
        m_CertActive = False

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Capture every error field BEFORE calling anything. TST_Log contains
    'On Error Resume Next, which clears Err, so a read taken after it returns
    'yields zero and an empty description. Raising from that reports no failure
    'at all, and the caller sees a normal return from a certification that
    'failed.
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        FailLine = Erl

    'Certification accounting must not leak into a later legacy run
        m_CertActive = False

    'Log from the captured values, never from Err
        TST_Log PROC, "FAIL", _
            CStr(FailNumber) & ": " & FailDescription & _
            IIf(FailLine <> 0, " | Line: " & CStr(FailLine), vbNullString)

    'Re-raise from the captured values, so the failure keeps its identity
        Err.Raise FailNumber, FailSource, FailDescription

End Sub


Private Sub TST_CertResetCounters()

'
'==============================================================================
' TST_CertResetCounters
'------------------------------------------------------------------------------
' PURPOSE
'   Clear every certification counter and buffer before a run.
'
' WHY THIS EXISTS
'   The regression packs are shared with the legacy runners, so skip recording
'   is module state rather than a parameter. Without an explicit reset a second
'   certification run in the same session would inherit the first run's counts
'   and report a verdict about work it did not do.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
'   2026-11-06 - Clears the record of snapshots units left behind (#135).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR COUNTERS
'------------------------------------------------------------------------------
    'Suppress any error; a reset that cannot complete must not abort the run
        On Error Resume Next

        m_CertUnitCount = 0
        m_CertSkipCount = 0
        m_CertLeakDetail = vbNullString

        Erase m_CertUnitNames
        Erase m_CertUnitPassed
        Erase m_CertUnitDetail
        Erase m_CertSkipDetail

End Sub


Private Sub TST_CertRecordSkip( _
    ByVal CallerProc As String, _
    ByVal Reason As String)

'
'==============================================================================
' TST_CertRecordSkip
'------------------------------------------------------------------------------
' PURPOSE
'   Log a skipped case and, during a certification run, count it.
'
' WHY THIS EXISTS
'   A skip used to be an Immediate Window line and nothing else, which made a
'   partial run indistinguishable from a complete one in every artifact that
'   survived the session. Counting the skip is what lets the certification
'   verdict report INCOMPLETE rather than quietly reporting PASS.
'
'   Recording is inert outside a certification run, so the legacy runners keep
'   their existing behavior exactly.
'
' INPUTS
'   CallerProc
'     Procedure that skipped the case, used as the log prefix.
'
'   Reason
'     Why the case was skipped.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' LOG SKIP
'------------------------------------------------------------------------------
    'Preserve the existing diagnostic line for the legacy runners
        TST_Log CallerProc, "SKIP", Reason

'------------------------------------------------------------------------------
' COUNT SKIP
'------------------------------------------------------------------------------
    'Accrue only while certifying, so counts cannot leak between invocations
        On Error Resume Next

        If m_CertActive Then

            m_CertSkipCount = m_CertSkipCount + 1

            ReDim Preserve m_CertSkipDetail(1 To m_CertSkipCount)
            m_CertSkipDetail(m_CertSkipCount) = CallerProc & ": " & Reason

        End If

End Sub


Private Sub TST_CertRunUnit( _
    ByVal UnitName As String)

'
'==============================================================================
' TST_CertRunUnit
'------------------------------------------------------------------------------
' PURPOSE
'   Execute one mandatory certification unit and record its outcome without
'   propagating a failure.
'
' WHY THIS EXISTS
'   The legacy runners raise on the first assertion failure, which is right for
'   interactive debugging and wrong for certification: it reports one defect and
'   conceals the state of everything after it. Trapping per unit means a single
'   run tells you everything that is broken, not merely the first thing.
'
' INPUTS
'   UnitName
'     Identifier of the unit to execute.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Executes the unit through TST_CertExecuteUnit, which traps its failure
'     and releases any snapshot it left behind.
'   - Records the outcome and, when the unit left a snapshot, names the unit in
'     the leak record that the cleanup verdict reads.
'
' ERROR POLICY
'   - Does not raise. A failing unit is recorded and the run continues.
'
' DEPENDENCIES
'   - TST_CertExecuteUnit
'   - TST_CertRecordUnit
'   - TST_CertAppendDetail
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
'   2026-11-06 - Execution moved to TST_CertExecuteUnit; a snapshot the unit
'                left behind is recorded for the cleanup verdict (#135).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Passed              As Boolean         'TRUE when the unit completed
    Dim Detail              As String          'Failure and release detail
    Dim LeftSnapshot        As Boolean         'TRUE when the unit left one

'------------------------------------------------------------------------------
' EXECUTE UNIT
'------------------------------------------------------------------------------
        TST_Log "TST_CertRunUnit", "UNIT", "Running " & UnitName

        TST_CertExecuteUnit _
            UnitName:=UnitName, _
            Passed:=Passed, _
            Detail:=Detail, _
            LeftSnapshot:=LeftSnapshot

'------------------------------------------------------------------------------
' RECORD OUTCOME
'------------------------------------------------------------------------------
    'Recording must never be the thing that breaks a run
        On Error Resume Next

        TST_CertRecordUnit UnitName, Passed, Detail

    'The release hid the leak from the exit state, so it is named here for the
    'cleanup verdict
        If LeftSnapshot Then
            m_CertLeakDetail = TST_CertAppendDetail( _
                m_CertLeakDetail, _
                UnitName & " left an EXCEL_UI snapshot behind")
        End If

End Sub


Private Sub TST_CertExecuteUnit( _
    ByVal UnitName As String, _
    ByRef Passed As Boolean, _
    ByRef Detail As String, _
    ByRef LeftSnapshot As Boolean)

'
'==============================================================================
' TST_CertExecuteUnit
'------------------------------------------------------------------------------
' PURPOSE
'   Run one unit under its own error boundary and release any snapshot it left
'   behind, without recording anything.
'
' WHY THIS EXISTS
'   A unit that fails after a capture can leave the snapshot behind. Every later
'   destructive unit then refuses to run over it, so one defect was reported as
'   every unit after it failing, plus a cleanup failure.
'
'   The release belongs here rather than in each case. Only this level knows
'   that the snapshot is the unit's own: certification refuses to start over a
'   snapshot and every earlier unit is released here, so the slot is empty
'   when a unit starts. A case-level release would need an error handler in
'   every case that captures, and would still not cover the identity runners.
'
'   Recording is left to the caller so that a regression case can drive the
'   same boundary without adding units to the certification run executing it.
'
' INPUTS
'   UnitName
'     Identifier of the unit to execute.
'
'   Passed
'     ByRef. Receives TRUE when the unit completed without raising.
'
'   Detail
'     ByRef. Receives the failure text, followed by the release, if any.
'     Empty when the unit passed and left nothing behind.
'
'   LeftSnapshot
'     ByRef. Receives TRUE when the unit left a snapshot behind.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Records whether a snapshot existed before the unit ran. When one did, it
'     is not the unit's, and it is left alone whatever happens.
'   - Dispatches the unit. An unknown identifier is a failure, not a pass.
'   - Releases a snapshot that exists after the unit when none existed before,
'     whether the unit passed or failed, and says so in Detail.
'
' ERROR POLICY
'   - Does not raise. A failing unit is reported through Passed and Detail.
'
' DEPENDENCIES
'   - TST_RunRegressionPack
'   - Test_EXCEL_UI_RunSnapshotIdentity
'   - Test_EXCEL_UI_RunTitleBarSdiIdentity
'   - TST_CertSelfTestCaptureFailureUnit
'   - TST_CertSelfTestFollowingUnit
'   - TST_CertAppendDetail
'
' CALLED FROM
'   - TST_CertRunUnit
'   - TST_Case_CertUnitIsolatesCaptureFailure
'
' NOTES
'   Dispatch is an explicit Select Case rather than Application.Run because the
'   units are private to this module and a name-based call would fail silently
'   or bind to the wrong project. A compile-time reference cannot rot.
'
'   The SelfTest units are dispatched only by
'   TST_Case_CertUnitIsolatesCaptureFailure. Certification never runs them.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HadSnapshot         As Boolean         'TRUE when the slot was occupied

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        Passed = False
        Detail = vbNullString
        LeftSnapshot = False

    'Assume the slot is not this unit's until it has been read as empty. A read
    'that failed must never license clearing a snapshot someone else owns.
        HadSnapshot = True

        On Error GoTo Err_Handler

        HadSnapshot = UI_HasExcelUIStateSnapshot

'------------------------------------------------------------------------------
' DISPATCH UNIT
'------------------------------------------------------------------------------
    'An unknown identifier is a defect in the registry, not a pass
        Select Case UnitName

            Case "RegressionPack"
                TST_RunRegressionPack _
                    IncludeTitleBarTests:=True, _
                    CallerProc:="Certification.RegressionPack"

            Case "SnapshotIdentity"
                Test_EXCEL_UI_RunSnapshotIdentity

            Case "TitleBarSdiIdentity"
                Test_EXCEL_UI_RunTitleBarSdiIdentity

            Case "SelfTest.CaptureFailure"
                TST_CertSelfTestCaptureFailureUnit

            Case "SelfTest.FollowingUnit"
                TST_CertSelfTestFollowingUnit

            Case Else
                Err.Raise _
                    TEST_CERT_ERR_BASE + 10, _
                    "TST_CertExecuteUnit", _
                    "unknown certification unit: " & UnitName

        End Select

'------------------------------------------------------------------------------
' RECORD PASS
'------------------------------------------------------------------------------
    'The unit completed without raising
        Passed = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Release a snapshot this unit created, on every path. The next unit would
    'otherwise refuse to run over it, and its failure would be this unit's.
        On Error Resume Next

        If Not HadSnapshot Then

            If UI_HasExcelUIStateSnapshot Then

                LeftSnapshot = True

                UI_ClearExcelUIStateSnapshot

                Detail = TST_CertAppendDetail(Detail, _
                    "left an EXCEL_UI snapshot behind; " & _
                    IIf(UI_HasExcelUIStateSnapshot, _
                        "it could not be released", _
                        "released before the next unit"))

            End If

        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Capture the failure before anything else can clear Err
        Detail = CStr(Err.Number) & ": " & Err.Description

        Resume Safe_Exit

End Sub


Private Sub TST_CertSelfTestCaptureFailureUnit()

'
'==============================================================================
' TST_CertSelfTestCaptureFailureUnit
'------------------------------------------------------------------------------
' PURPOSE
'   Fail the way the first real-host certification run failed: an assertion
'   on a capture that reported failure, after that capture had created a
'   snapshot, with nothing on the path to release it.
'
' WHY THIS EXISTS
'   TST_Case_CertUnitIsolatesCaptureFailure needs a unit that fails in that
'   shape on demand. A capture fails only when the host refuses a read, which
'   cannot be provoked, so the active-frame pair seam forces the TitleBar
'   element to fail. The capture still completes and still leaves a snapshot,
'   exactly as a host refusal does.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Raises the TST_AssertResultSuccess failure, on purpose.
'   - Leaves the snapshot behind, on purpose. Releasing it is the job of the
'     boundary under test.
'
' DEPENDENCIES
'   - UI_InternalInjectFramePairFault
'   - UI_CaptureExcelUIState_WithResult
'   - TST_AssertResultSuccess
'
' CALLED FROM
'   - TST_CertExecuteUnit
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean         'Structured capture result
    Dim FailureCount        As Long            'Structured failure count
    Dim FailureList         As Variant         'Structured failure list

    Const PROC As String = "TST_CertSelfTestCaptureFailureUnit"

'------------------------------------------------------------------------------
' CAPTURE WITH AN INJECTED FAILURE
'------------------------------------------------------------------------------
    'The seam is one-shot and the capture resolves the frame pair once, so
    'exactly the TitleBar element fails
        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_MISMATCH

        OK = UI_CaptureExcelUIState_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

    'Disarm in case the capture never reached the resolution
        UI_InternalInjectFramePairFault TEST_PAIR_FAULT_NONE

'------------------------------------------------------------------------------
' ASSERT AS THE CLEAN-SUCCESS CASE DOES
'------------------------------------------------------------------------------
    'Raises, carrying the failure list the capture returned
        TST_AssertResultSuccess _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            AssertionName:=PROC & ".Result"

End Sub


Private Sub TST_CertSelfTestFollowingUnit()

'
'==============================================================================
' TST_CertSelfTestFollowingUnit
'------------------------------------------------------------------------------
' PURPOSE
'   Stand in for the unit that runs after a failed one.
'
' WHY THIS EXISTS
'   The snapshot-identity and SDI identity units begin by refusing to run over
'   an existing snapshot, and that refusal is what turned one failed unit into
'   three. This unit applies the same refusal and nothing else, so it passes
'   exactly when a real following unit would get past its precondition,
'   without running a full destructive unit a second time.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Raises when an explicit EXCEL_UI snapshot exists.
'
' CALLED FROM
'   - TST_CertExecuteUnit
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE PRECONDITIONS
'------------------------------------------------------------------------------
    'The refusal the destructive units apply before anything else
        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_CERT_ERR_BASE + 30, _
                "TST_CertSelfTestFollowingUnit", _
                "an explicit EXCEL_UI snapshot already exists; clear or " & _
                "restore it before running this destructive test"
        End If

        TST_Log "TST_CertSelfTestFollowingUnit", "PASS", _
            "No snapshot was left for this unit to refuse"

End Sub


Private Sub TST_CertRecordUnit( _
    ByVal UnitName As String, _
    ByVal Passed As Boolean, _
    ByVal Detail As String)

'
'==============================================================================
' TST_CertRecordUnit
'------------------------------------------------------------------------------
' PURPOSE
'   Append one unit outcome to the certification record.
'
' WHY THIS EXISTS
'   The verdict, the JSON evidence and the text report are all derived from the
'   same arrays, so there is exactly one place where an outcome enters the run
'   and no opportunity for the three to disagree.
'
' INPUTS
'   UnitName
'     Identifier of the unit.
'
'   Passed
'     TRUE when the unit completed without raising.
'
'   Detail
'     Failure text; empty on success.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - TST_CertRunUnit
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' APPEND OUTCOME
'------------------------------------------------------------------------------
    'Recording must never be the thing that breaks a run
        On Error Resume Next

        m_CertUnitCount = m_CertUnitCount + 1

        ReDim Preserve m_CertUnitNames(1 To m_CertUnitCount)
        ReDim Preserve m_CertUnitPassed(1 To m_CertUnitCount)
        ReDim Preserve m_CertUnitDetail(1 To m_CertUnitCount)

        m_CertUnitNames(m_CertUnitCount) = UnitName
        m_CertUnitPassed(m_CertUnitCount) = Passed
        m_CertUnitDetail(m_CertUnitCount) = Detail

        TST_Log "TST_CertRecordUnit", IIf(Passed, "UNIT PASS", "UNIT FAIL"), _
            UnitName & IIf(Len(Detail) > 0, " | " & Detail, vbNullString)

End Sub


Private Function TST_CertEvaluateCleanup( _
    ByVal BaselineBooks As Long, _
    ByVal BaselineScreenUpdating As Boolean, _
    ByVal AnchorWindow As Object, _
    ByRef CleanupDetail As String) _
    As Boolean

'
'==============================================================================
' TST_CertEvaluateCleanup
'------------------------------------------------------------------------------
' PURPOSE
'   Decide whether a certification run returned the host to the state it found.
'
' WHY THIS EXISTS
'   Every check compares the exit state with the state observed on ENTRY, never
'   with an assumed ideal. Requiring ScreenUpdating to be True reported a false
'   failure for any run started from within a quiet-update scope, where
'   restoring the suppressed value it found is correct behavior rather than
'   leakage.
'
'   A counter that fires when nothing is wrong is a counter a reader learns to
'   discount, which defeats the verdict from the opposite direction to a missed
'   failure.
'
'   Extracting the decision from the runner makes it testable with crafted
'   inputs, without a full destructive run, and gives the additional state
'   comparisons planned for this suite one place to live.
'
' INPUTS
'   BaselineBooks
'     Workbooks.Count observed on entry.
'
'   BaselineScreenUpdating
'     Application.ScreenUpdating observed on entry.
'
'   AnchorWindow
'     The window active on entry, held by object reference.
'
'   CleanupDetail
'     ByRef. Receives every finding, joined in order. Empty when clean.
'
' RETURNS
'   Boolean
'     True  => the host matches its entry state.
'     False => at least one difference was found, and all are in CleanupDetail.
'
' BEHAVIOR
'   - Accumulates findings rather than stopping at the first, so one run reports
'     every way cleanup failed.
'   - Reports the snapshot check absolutely, because a snapshot is rejected as a
'     precondition rather than preserved.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertAppendDetail
'   - TST_CertIsWindowUsable
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'   - TST_Case_CertificationCleanupUsesBaseline
'
' UPDATED
'   2026-08-21
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A cleanup verdict must never itself raise
        On Error Resume Next

        TST_CertEvaluateCleanup = True
        CleanupDetail = vbNullString

'------------------------------------------------------------------------------
' COMPARE AGAINST ENTRY STATE
'------------------------------------------------------------------------------
    'A snapshot is rejected on entry rather than preserved, so any snapshot here
    'was left by the run itself
        If UI_HasExcelUIStateSnapshot Then
            TST_CertEvaluateCleanup = False
            CleanupDetail = "an EXCEL_UI snapshot was left behind"
        End If

        If Workbooks.Count <> BaselineBooks Then
            TST_CertEvaluateCleanup = False
            CleanupDetail = TST_CertAppendDetail(CleanupDetail, _
                "workbook count changed from " & CStr(BaselineBooks) & _
                " to " & CStr(Workbooks.Count))
        End If

        If Application.ScreenUpdating <> BaselineScreenUpdating Then
            TST_CertEvaluateCleanup = False
            CleanupDetail = TST_CertAppendDetail(CleanupDetail, _
                "ScreenUpdating changed from " & _
                CStr(BaselineScreenUpdating) & " to " & _
                CStr(Application.ScreenUpdating))
        End If

        If Not TST_CertIsWindowUsable(AnchorWindow) Then
            TST_CertEvaluateCleanup = False
            CleanupDetail = TST_CertAppendDetail(CleanupDetail, _
                "the anchor window is no longer usable")
        End If

End Function


Private Function TST_CertIsWindowUsable( _
    ByVal TargetWindow As Object) _
    As Boolean

'
'==============================================================================
' TST_CertIsWindowUsable
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether a retained Window reference still responds.
'
' WHY THIS EXISTS
'   Cleanup is judged against the window the run started from, held by object
'   identity rather than by collection index. An index would be satisfied by any
'   window that happens to occupy the same position afterwards, which is exactly
'   the substitution the snapshot identity work exists to prevent.
'
' INPUTS
'   TargetWindow
'     Window to probe. May be Nothing.
'
' RETURNS
'   Boolean
'     TRUE when the reference still responds to a non-mutating read.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - TST_CertEvaluateCleanup
'
' UPDATED
'   2026-08-21
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ProbeValue          As Boolean         'Non-mutating probe output

'------------------------------------------------------------------------------
' PROBE REFERENCE
'------------------------------------------------------------------------------
    'A dead wrapper raises on the read, which is the signal required
        On Error Resume Next

        TST_CertIsWindowUsable = False

        If TargetWindow Is Nothing Then
            Exit Function
        End If

        ProbeValue = TargetWindow.DisplayHeadings

        TST_CertIsWindowUsable = (Err.Number = 0)

        Err.Clear

End Function


Private Function TST_CertAppendDetail( _
    ByVal ExistingDetail As String, _
    ByVal NewDetail As String) _
    As String

'
'==============================================================================
' TST_CertAppendDetail
'------------------------------------------------------------------------------
' PURPOSE
'   Join cleanup findings into one ordered diagnostic string.
'
' WHY THIS EXISTS
'   Cleanup can fail in more than one way at once, and reporting only the first
'   would send the reader back for a second run to discover the second problem.
'
' INPUTS
'   ExistingDetail
'     Findings recorded so far; may be empty.
'
'   NewDetail
'     Finding to append.
'
' RETURNS
'   String
'     The joined findings.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - TST_CertEvaluateCleanup
'   - TST_CertExecuteUnit
'   - TST_CertRunUnit
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
'   2026-11-06 - Also joins unit release detail and the leak record (#135).
'   2026-08-21
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' JOIN FINDINGS
'------------------------------------------------------------------------------
    'Separate with the established diagnostic delimiter
        If Len(ExistingDetail) = 0 Then
            TST_CertAppendDetail = NewDetail
        Else
            TST_CertAppendDetail = ExistingDetail & " | " & NewDetail
        End If

End Function


Private Function TST_CertBuildJsonEvidence( _
    ByVal UnitsFailed As Long, _
    ByVal CleanupOK As Boolean, _
    ByVal CleanupDetail As String, _
    ByVal Complete As Boolean, _
    ByVal Passed As Boolean) _
    As String

'
'==============================================================================
' TST_CertBuildJsonEvidence
'------------------------------------------------------------------------------
' PURPOSE
'   Compose the machine-readable certification result.
'
' WHY THIS EXISTS
'   Prose in the Immediate Window cannot be attached to a release, compared
'   between environments or checked by a workflow. A result that names the exact
'   host it was obtained on is the difference between evidence and an assertion
'   that the tests passed somewhere once.
'
'   The document is assembled by hand rather than through a library because the
'   module carries no dependency, and the field set is small and fixed.
'
' INPUTS
'   UnitsFailed / CleanupOK / CleanupDetail / Complete / Passed
'     Verdict components computed by the caller.
'
' RETURNS
'   String
'     A single-line JSON document.
'
' ERROR POLICY
'   - Does not raise. A field that cannot be read is reported as unknown.
'
' DEPENDENCIES
'   - TST_CertJsonEscape
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'
' NOTES
'   Bitness and VBA generation come from conditional compilation, so they
'   describe the build that is executing rather than what the host reports.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Json                As String          'Document under construction
    Dim ScanIdx             As Long            'Cursor over recorded units
    Dim Bitness             As String          'Office bitness of this build
    Dim VbaGeneration       As String          'VBA generation of this build

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Evidence must be produced even when a field cannot be read
        On Error Resume Next

#If VBA7 Then
        VbaGeneration = "VBA7"
    #If Win64 Then
        Bitness = "x64"
    #Else
        Bitness = "x86"
    #End If
#Else
        VbaGeneration = "pre-VBA7"
        Bitness = "x86"
#End If

'------------------------------------------------------------------------------
' BUILD ENVIRONMENT
'------------------------------------------------------------------------------
    'Name the exact host this verdict describes
        Json = "{""component"":""VBA Excel UI""" & _
            ",""schema"":1" & _
            ",""timestampLocal"":""" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & """" & _
            ",""excelVersion"":""" & TST_CertJsonEscape(Application.Version) & """" & _
            ",""excelBuild"":""" & TST_CertJsonEscape(CStr(Application.Build)) & """" & _
            ",""operatingSystem"":""" & TST_CertJsonEscape(Application.OperatingSystem) & """" & _
            ",""bitness"":""" & Bitness & """" & _
            ",""vbaGeneration"":""" & VbaGeneration & """"

'------------------------------------------------------------------------------
' BUILD COUNTERS
'------------------------------------------------------------------------------
    'Counts first, so a reader sees the shape of the run before its detail
        Json = Json & _
            ",""units"":" & CStr(m_CertUnitCount) & _
            ",""unitsFailed"":" & CStr(UnitsFailed) & _
            ",""skipped"":" & CStr(m_CertSkipCount) & _
            ",""cleanup"":""" & IIf(CleanupOK, "OK", "FAILED") & """" & _
            ",""cleanupDetail"":""" & TST_CertJsonEscape(CleanupDetail) & """" & _
            ",""complete"":" & IIf(Complete, "true", "false") & _
            ",""passed"":" & IIf(Passed, "true", "false")

'------------------------------------------------------------------------------
' BUILD UNIT DETAIL
'------------------------------------------------------------------------------
    'One object per unit, in execution order
        Json = Json & ",""unitResults"":["

        For ScanIdx = 1 To m_CertUnitCount

            If ScanIdx > 1 Then
                Json = Json & ","
            End If

            Json = Json & "{""name"":""" & _
                TST_CertJsonEscape(m_CertUnitNames(ScanIdx)) & """" & _
                ",""passed"":" & IIf(m_CertUnitPassed(ScanIdx), "true", "false") & _
                ",""detail"":""" & _
                TST_CertJsonEscape(m_CertUnitDetail(ScanIdx)) & """}"

        Next ScanIdx

        Json = Json & "]"

'------------------------------------------------------------------------------
' BUILD SKIP DETAIL
'------------------------------------------------------------------------------
    'Skips are listed explicitly; an empty array is a meaningful result
        Json = Json & ",""skipDetail"":["

        For ScanIdx = 1 To m_CertSkipCount

            If ScanIdx > 1 Then
                Json = Json & ","
            End If

            Json = Json & """" & _
                TST_CertJsonEscape(m_CertSkipDetail(ScanIdx)) & """"

        Next ScanIdx

        Json = Json & "]}"

'------------------------------------------------------------------------------
' RETURN DOCUMENT
'------------------------------------------------------------------------------
    'Publish the assembled document
        TST_CertBuildJsonEvidence = Json

End Function


Private Function TST_CertJsonEscape( _
    ByVal Value As String) _
    As String

'
'==============================================================================
' TST_CertJsonEscape
'------------------------------------------------------------------------------
' PURPOSE
'   Escape the characters that would otherwise make the evidence document
'   unparseable.
'
' WHY THIS EXISTS
'   Failure detail is host text and can contain quotes, backslashes and line
'   breaks. Emitting it raw would produce a document that no consumer can read,
'   which defeats the purpose of machine-readable evidence precisely when the
'   run failed and the evidence matters most.
'
' INPUTS
'   Value
'     Text to escape. May be empty.
'
' RETURNS
'   String
'     The escaped text, without surrounding quotes.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - TST_CertBuildJsonEvidence
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result              As String          'Escaped text under construction

'------------------------------------------------------------------------------
' ESCAPE VALUE
'------------------------------------------------------------------------------
    'Escape the backslash first, or the escapes added below are re-escaped
        On Error Resume Next

        Result = Value
        Result = Replace(Result, "\", "\\")
        Result = Replace(Result, """", "\""")
        Result = Replace(Result, vbCrLf, "\n")
        Result = Replace(Result, vbCr, "\n")
        Result = Replace(Result, vbLf, "\n")
        Result = Replace(Result, vbTab, "\t")

        TST_CertJsonEscape = Result

End Function


Private Function TST_CertBuildTextReport( _
    ByVal Verdict As String, _
    ByVal CleanupDetail As String) _
    As String

'
'==============================================================================
' TST_CertBuildTextReport
'------------------------------------------------------------------------------
' PURPOSE
'   Compose the human-readable certification summary.
'
' WHY THIS EXISTS
'   The JSON document is for tooling. A person deciding whether to tag a release
'   needs the same facts in a form they can read at a glance and paste into the
'   changelog validation block.
'
' INPUTS
'   Verdict
'     Pre-composed verdict line.
'
'   CleanupDetail
'     Cleanup findings; empty when cleanup passed.
'
' RETURNS
'   String
'     A multi-line report.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Report              As String          'Report under construction
    Dim ScanIdx             As Long            'Cursor over recorded entries

'------------------------------------------------------------------------------
' BUILD REPORT
'------------------------------------------------------------------------------
    'A report must be produced even when a field cannot be read
        On Error Resume Next

        Report = "VBA Excel UI - release certification" & vbNewLine & _
            "Excel " & Application.Version & _
            " build " & CStr(Application.Build) & vbNewLine & _
            Application.OperatingSystem & vbNewLine & _
            Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbNewLine & _
            String$(60, "-") & vbNewLine & Verdict & vbNewLine

'------------------------------------------------------------------------------
' APPEND UNIT RESULTS
'------------------------------------------------------------------------------
    'One line per unit, in execution order
        For ScanIdx = 1 To m_CertUnitCount

            Report = Report & _
                IIf(m_CertUnitPassed(ScanIdx), "  PASS  ", "  FAIL  ") & _
                m_CertUnitNames(ScanIdx) & _
                IIf(Len(m_CertUnitDetail(ScanIdx)) > 0, _
                    " | " & m_CertUnitDetail(ScanIdx), vbNullString) & _
                vbNewLine

        Next ScanIdx

'------------------------------------------------------------------------------
' APPEND SKIPS AND CLEANUP
'------------------------------------------------------------------------------
    'Skips are named, because an unnamed skip is what this runner exists to
    'stop happening
        For ScanIdx = 1 To m_CertSkipCount
            Report = Report & "  SKIP  " & m_CertSkipDetail(ScanIdx) & vbNewLine
        Next ScanIdx

        If Len(CleanupDetail) > 0 Then
            Report = Report & "  CLEANUP  " & CleanupDetail & vbNewLine
        End If

'------------------------------------------------------------------------------
' RETURN REPORT
'------------------------------------------------------------------------------
    'Publish the assembled report
        TST_CertBuildTextReport = Report

End Function


Private Function TST_CertTryWriteEvidence( _
    ByVal Content As String, _
    ByVal Extension As String, _
    ByRef PathOut As String) _
    As Boolean

'
'==============================================================================
' TST_CertTryWriteEvidence
'------------------------------------------------------------------------------
' PURPOSE
'   Write one evidence document to the temporary folder.
'
' WHY THIS EXISTS
'   Evidence that exists only in the Immediate Window is lost the moment the
'   session ends and cannot be attached to a release. A file can.
'
'   Writing is nonetheless best effort and never a gate on the verdict: a
'   locked-down or full temporary folder is an environment problem, and letting
'   it fail a run whose assertions all passed would report the wrong defect.
'
' INPUTS
'   Content
'     Document text to write.
'
'   Extension
'     File extension without the dot.
'
'   PathOut
'     ByRef. Receives the path written, or empty on failure.
'
' RETURNS
'   Boolean
'     TRUE when the file was written.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FolderPath          As String          'Temporary folder for evidence
    Dim FileNumber          As Integer         'Free file handle

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A file-system failure must never abort certification
        On Error GoTo Err_Handler

        TST_CertTryWriteEvidence = False
        PathOut = vbNullString

'------------------------------------------------------------------------------
' RESOLVE PATH
'------------------------------------------------------------------------------
    'Fall back silently when the host exposes no temporary folder
        FolderPath = Environ$("TEMP")

        If Len(FolderPath) = 0 Then
            GoTo Safe_Exit
        End If

        PathOut = FolderPath & "\EXCEL_UI_certification_" & _
            Format$(Now, "yyyymmdd_hhnnss") & "." & Extension

'------------------------------------------------------------------------------
' WRITE FILE
'------------------------------------------------------------------------------
    'Write the document as a single block
        FileNumber = FreeFile

        Open PathOut For Output As #FileNumber
        Print #FileNumber, Content
        Close #FileNumber

        TST_CertTryWriteEvidence = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Report the miss without disturbing the verdict
        TST_CertTryWriteEvidence = False
        PathOut = vbNullString

        Resume Safe_Exit

End Function


Public Sub Test_EXCEL_UI_RunOwnershipCleanupChecks()

'
'==============================================================================
' Test_EXCEL_UI_RunOwnershipCleanupChecks
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a runner which creates its own snapshot releases it on every
'   exit, and that a failure survives that cleanup unchanged.
'
' WHY THIS EXISTS
'   The refusal-path case proves a caller-owned snapshot is never touched. It
'   says nothing about the other half of the same ownership rule: that a
'   runner-owned snapshot is always released. Those paths run when something
'   fails between capture and completion, which is precisely what the runners
'   are written not to do, so they are unreachable without injecting a failure.
'   An unreachable cleanup path is where the original defect lived.
'
'   Cleanup must also not swallow the failure that triggered it. Safe_Exit runs
'   under On Error Resume Next, so a cleanup step that raised would leave the
'   caller with a cleared Err and no idea anything went wrong.
'
' RETURNS
'   None
'
' BEHAVIOR
'   A  no injection            - normal return, no error, snapshot released
'   B  injected assertion      - assertion error preserved, snapshot released
'   C  injected unexpected     - injected error preserved, snapshot released
'   D  identity post-capture   - assertion preserved, snapshot and the
'                                temporary window both released
'
' ERROR POLICY
'   - Raises on assertion failure
'   - Disarms any injection it armed before returning
'
' DEPENDENCIES
'   - TST_InjectPostCapture
'   - Test_EXCEL_UI_RunCertificationSelfTest
'   - Test_EXCEL_UI_RunSnapshotIdentity
'
' CALLED FROM
'   - Invoked directly
'
' NOTES
'   Deliberately not registered in TST_RunRegressionPack. The self-test calls
'   certification, and while a certification run is active that call meets the
'   re-entrancy guard instead of the snapshot precondition this exercises, so
'   inside the pack these cases would assert the wrong refusal.
'
'   Run it with no explicit snapshot held.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RaisedNumber        As Long            'Error number reaching this run
    Dim RaisedSource        As String          'Error source reaching this run
    Dim RaisedDescription   As String          'Error text reaching this run
    Dim WindowsBefore       As Long            'Window count before case D
    Dim WindowsAfter        As Long            'Window count after case D

    Const PROC As String = "Test_EXCEL_UI_RunOwnershipCleanupChecks"

'------------------------------------------------------------------------------
' VALIDATE PRECONDITIONS
'------------------------------------------------------------------------------
    'Every case below reasons about the snapshot slot being empty afterwards,
    'which says nothing unless it was empty to begin with
        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_OWNERSHIP_ERR_BASE + 10, _
                PROC, _
                "an explicit EXCEL_UI snapshot already exists; clear or " & _
                "restore it before running these checks"
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_InjectPostCapture TEST_FAULT_NONE

        TST_Log PROC, "START", _
            "Validating runner-owned snapshot cleanup on every exit"

'------------------------------------------------------------------------------
' CASE A - NORMAL SUCCESS RELEASES THE SNAPSHOT
'------------------------------------------------------------------------------
        On Error Resume Next

            Test_EXCEL_UI_RunCertificationSelfTest

            RaisedNumber = Err.Number
            RaisedDescription = Err.Description

            Err.Clear

        On Error GoTo Err_Handler

        TST_AssertTrue _
            (RaisedNumber = 0), _
            PROC & ".SelfTestSuccess.NoError"

        TST_AssertSnapshotAvailability False, _
            PROC & ".SelfTestSuccess.SnapshotReleased"

'------------------------------------------------------------------------------
' CASE B - AN ASSERTION AFTER CAPTURE STILL RELEASES THE SNAPSHOT
'------------------------------------------------------------------------------
        TST_InjectPostCapture TEST_FAULT_ASSERTION

        On Error Resume Next

            Test_EXCEL_UI_RunCertificationSelfTest

            RaisedNumber = Err.Number
            RaisedSource = Err.Source
            RaisedDescription = Err.Description

            Err.Clear

        On Error GoTo Err_Handler

        TST_AssertRefusal _
            ExpectedNumber:=TEST_SNAPSHOT_ID_ERR_BASE + 20, _
            ExpectedSource:= _
                "Test_EXCEL_UI_RunCertificationSelfTest.InjectedAssertion", _
            ExpectedDescription:="expected TRUE but received FALSE", _
            ActualNumber:=RaisedNumber, _
            ActualSource:=RaisedSource, _
            ActualDescription:=RaisedDescription, _
            AssertionName:=PROC & ".SelfTestAssertion.Preserved"

        TST_AssertSnapshotAvailability False, _
            PROC & ".SelfTestAssertion.SnapshotReleased"

'------------------------------------------------------------------------------
' CASE C - AN UNEXPECTED ERROR AFTER CAPTURE STILL RELEASES THE SNAPSHOT
'------------------------------------------------------------------------------
        TST_InjectPostCapture TEST_FAULT_UNEXPECTED

        On Error Resume Next

            Test_EXCEL_UI_RunCertificationSelfTest

            RaisedNumber = Err.Number
            RaisedSource = Err.Source
            RaisedDescription = Err.Description

            Err.Clear

        On Error GoTo Err_Handler

        TST_AssertRefusal _
            ExpectedNumber:=TEST_OWNERSHIP_ERR_BASE + 1, _
            ExpectedSource:= _
                "Test_EXCEL_UI_RunCertificationSelfTest.InjectedUnexpected", _
            ExpectedDescription:="injected unexpected post-capture failure", _
            ActualNumber:=RaisedNumber, _
            ActualSource:=RaisedSource, _
            ActualDescription:=RaisedDescription, _
            AssertionName:=PROC & ".SelfTestUnexpected.Preserved"

        TST_AssertSnapshotAvailability False, _
            PROC & ".SelfTestUnexpected.SnapshotReleased"

'------------------------------------------------------------------------------
' CASE D - THE IDENTITY RUNNER RELEASES SNAPSHOT AND WINDOWS ALIKE
'------------------------------------------------------------------------------
    'This runner sets ownership between the capture and an assertion on its
    'structured result, a window the self-test does not have. It also creates a
    'replacement window, so cleanup has more than the snapshot to release.
        WindowsBefore = ThisWorkbook.Windows.Count

        TST_InjectPostCapture TEST_FAULT_ASSERTION

        On Error Resume Next

            Test_EXCEL_UI_RunSnapshotIdentity

            RaisedNumber = Err.Number
            RaisedSource = Err.Source
            RaisedDescription = Err.Description

            Err.Clear

        On Error GoTo Err_Handler

        TST_AssertRefusal _
            ExpectedNumber:=TEST_SNAPSHOT_ID_ERR_BASE + 20, _
            ExpectedSource:= _
                "Test_EXCEL_UI_RunSnapshotIdentity.InjectedAssertion", _
            ExpectedDescription:="expected TRUE but received FALSE", _
            ActualNumber:=RaisedNumber, _
            ActualSource:=RaisedSource, _
            ActualDescription:=RaisedDescription, _
            AssertionName:=PROC & ".IdentityAssertion.Preserved"

        TST_AssertSnapshotAvailability False, _
            PROC & ".IdentityAssertion.SnapshotReleased"

        WindowsAfter = ThisWorkbook.Windows.Count

        TST_AssertTrue _
            (WindowsAfter = WindowsBefore), _
            PROC & ".IdentityAssertion.WindowsReleased"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Runner-owned snapshots were released on every exit and every " & _
            "failure survived cleanup"

Safe_Exit:
    'Disarm whatever this run armed. A fault left armed would fire inside an
    'unrelated runner later in the same session.
        On Error Resume Next
            TST_InjectPostCapture TEST_FAULT_NONE
        On Error GoTo 0

        Exit Sub

Err_Handler:
    'Capture before anything else touches Err
        RaisedNumber = Err.Number
        RaisedSource = Err.Source
        RaisedDescription = Err.Description

        TST_Log PROC, "FAIL", _
            "Error " & CStr(RaisedNumber) & _
            " | Source: " & RaisedSource & _
            " | " & RaisedDescription

        On Error Resume Next
            TST_InjectPostCapture TEST_FAULT_NONE
        On Error GoTo 0

        Err.Raise RaisedNumber, RaisedSource, RaisedDescription

End Sub


Public Sub Test_EXCEL_UI_RunRibbonSdiProbe()

'
'==============================================================================
' Test_EXCEL_UI_RunRibbonSdiProbe
'------------------------------------------------------------------------------
' PURPOSE
'   Characterize how Ribbon visibility behaves across multiple workbook windows,
'   and record the result as evidence.
'
' WHY THIS EXISTS
'   README.md once stated the Ribbon scope as "Excel application", which a
'   reader would take to mean one state shared by every workbook window. Under
'   the Single Document Interface each workbook window has its own Ribbon UI,
'   and nothing in the component verified that claim. This probe measured it:
'   README.md now states the Ribbon scope as the active window, and
'   docs/RIBBON_SDI_BEHAVIOR.md records the measurement.
'
'   This is deliberately a PROBE and not a test. It asserts nothing, because
'   there is no agreed correct answer yet: the point is to discover what the
'   host does so that a contract can be written, and only then to assert it.
'   Writing assertions first would encode the guess this issue exists to remove.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Refuses to run when an explicit EXCEL_UI snapshot already exists.
'   - Records the Ribbon state through three independent mechanisms at each
'     observation point, because they can legitimately disagree.
'   - Exercises five scenarios: baseline, hide with A active, show with B
'     active, a window created AFTER a hide, and snapshot restore across a
'     window switch.
'   - Emits a text table and a JSON document, and writes both to the temporary
'     folder.
'   - Restores the Ribbon and closes its temporary workbooks.
'
' ERROR POLICY
'   - Raises only on a precondition failure or an unexpected host error.
'   - Never fails on an observed value: every value is data.
'
' DEPENDENCIES
'   - TST_RibbonProbeReset
'   - TST_RibbonProbeRecord
'   - TST_RibbonProbeBuildJson
'   - TST_CertTryWriteEvidence
'
' NOTES
'   Destructive: creates and closes temporary workbooks and toggles the Ribbon.
'   Results belong in docs/RIBBON_SDI_BEHAVIOR.md, one block per host tested.
'
' UPDATED
'   2026-08-20 - README.md now states the active-window Ribbon scope.
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window          'Window active before the probe
    Dim BookB               As Workbook        'Second workbook, window B
    Dim BookC               As Workbook        'Workbook created after a hide

    Dim JsonText            As String          'Machine-readable observations
    Dim ReportText          As String          'Human-readable observations
    Dim EvidencePath        As String          'Path an evidence file was written to

    Dim HasFailure          As Boolean         'TRUE when the probe failed
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Dim CreatedSnapshot     As Boolean         'TRUE once this run owns one

    Const PROC As String = "Test_EXCEL_UI_RunRibbonSdiProbe"

'------------------------------------------------------------------------------
' VALIDATE PRECONDITIONS
'------------------------------------------------------------------------------
    'The probe captures and restores a snapshot, so it cannot share the slot.
    '
    'Refused before the handler is armed. Armed first, this refusal travelled
    'Err_Handler, resumed at Safe_Exit, and the unconditional clear there
    'deleted the caller snapshot it had just refused to disturb.
        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_RIBBON_ERR_BASE + 1, _
                PROC, _
                "an explicit EXCEL_UI snapshot already exists; clear or " & _
                "restore it before probing"
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        CreatedSnapshot = False

        On Error GoTo Err_Handler

        TST_RibbonProbeReset

        TST_Log PROC, "START", "Ribbon SDI characterization started"

        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_RIBBON_ERR_BASE + 2, _
                PROC, _
                "no active Excel window is available"
        End If

'------------------------------------------------------------------------------
' SCENARIO 1 - BASELINE
'------------------------------------------------------------------------------
    'Establish what a visible Ribbon reads as on this host, so later rows can be
    'compared against something rather than interpreted in isolation
        UI_SetExcelUI Ribbon:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        TST_RibbonProbeRecord "1-Baseline", "A"

'------------------------------------------------------------------------------
' SCENARIO 2 - HIDE WITH A ACTIVE, THEN OBSERVE B
'------------------------------------------------------------------------------
    'The central question: does hiding the Ribbon while A is active also hide it
    'for a window that already existed?
        Set BookB = Workbooks.Add
        TST_WaitUI TEST_WAIT_SECONDS

        AnchorWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        UI_SetExcelUI Ribbon:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

        TST_RibbonProbeRecord "2-HiddenOnA", "A"

        BookB.Windows(1).Activate
        TST_WaitUI TEST_WAIT_SECONDS

        TST_RibbonProbeRecord "2-HiddenOnA", "B"

'------------------------------------------------------------------------------
' SCENARIO 3 - SHOW WITH B ACTIVE, THEN OBSERVE A
'------------------------------------------------------------------------------
    'The symmetric question, which need not have the symmetric answer
        UI_SetExcelUI Ribbon:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        TST_RibbonProbeRecord "3-ShownOnB", "B"

        AnchorWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        TST_RibbonProbeRecord "3-ShownOnB", "A"

'------------------------------------------------------------------------------
' SCENARIO 4 - WINDOW CREATED AFTER A HIDE
'------------------------------------------------------------------------------
    'A window that did not exist when the Ribbon was hidden is the case a
    'component storing one Boolean cannot reason about at all
        UI_SetExcelUI Ribbon:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

        Set BookC = Workbooks.Add
        TST_WaitUI TEST_WAIT_SECONDS

        TST_RibbonProbeRecord "4-NewWindowAfterHide", "C"

        AnchorWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        TST_RibbonProbeRecord "4-NewWindowAfterHide", "A"

'------------------------------------------------------------------------------
' SCENARIO 5 - SNAPSHOT RESTORE ACROSS A WINDOW SWITCH
'------------------------------------------------------------------------------
    'Whether the snapshot contract holds for the Ribbon depends entirely on the
    'answers above, so it is measured rather than assumed
        UI_SetExcelUI Ribbon:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        UI_CaptureExcelUIState

    'Ownership follows an observed snapshot, not a returned call. A capture
    'that silently left none must not make this run believe it owns one.
        CreatedSnapshot = UI_HasExcelUIStateSnapshot
        TST_WaitUI TEST_WAIT_SECONDS

        UI_SetExcelUI Ribbon:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

        BookB.Windows(1).Activate
        TST_WaitUI TEST_WAIT_SECONDS

    'Since #23 restore refuses the Ribbon while another window is active, so
    'the rows below record the refusal: neither window's Ribbon changes. The
    'fire-and-forget restore logs the Ribbon failure.
        UI_ResetExcelUIToSnapshot
        TST_WaitUI TEST_WAIT_SECONDS

        TST_RibbonProbeRecord "5-RestoredFromB", "B"

        AnchorWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        TST_RibbonProbeRecord "5-RestoredFromB", "A"

'------------------------------------------------------------------------------
' EMIT EVIDENCE
'------------------------------------------------------------------------------
    'Build both forms before writing either
        ReportText = "VBA Excel UI - Ribbon SDI characterization" & vbNewLine & _
            "Excel " & Application.Version & _
            " build " & CStr(Application.Build) & vbNewLine & _
            Application.OperatingSystem & vbNewLine & _
            Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbNewLine & _
            String$(72, "-") & vbNewLine & _
            "scenario              window  CommandBars.Visible  Height  XLM" & _
            vbNewLine & m_RibbonRowsText

        JsonText = TST_RibbonProbeBuildJson()

        TST_Log PROC, "EVIDENCE", vbNewLine & ReportText
        TST_Log PROC, "JSON", JsonText

        If TST_CertTryWriteEvidence(JsonText, "ribbon.json", EvidencePath) Then
            TST_Log PROC, "WROTE", EvidencePath
        End If

        If TST_CertTryWriteEvidence(ReportText, "ribbon.txt", EvidencePath) Then
            TST_Log PROC, "WROTE", EvidencePath
        End If

        TST_Log PROC, "DONE", _
            "Recorded " & CStr(m_RibbonRowCount) & _
            " observations; transcribe them into docs/RIBBON_SDI_BEHAVIOR.md"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Leave the host as it was found, whatever the observations were. Only a
    'snapshot this run created is released; a caller-owned one reaches this
    'label whenever an early failure resumes here.
        On Error Resume Next
            If CreatedSnapshot Then
                UI_ClearExcelUIStateSnapshot
            End If

            TST_SafeCloseWorkbook BookC
            TST_SafeCloseWorkbook BookB

            If Not AnchorWindow Is Nothing Then
                AnchorWindow.Activate
            End If

            UI_SetExcelUI Ribbon:=UI_Show
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


Private Sub TST_RibbonProbeReset()

'
'==============================================================================
' TST_RibbonProbeReset
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the observation buffers before a probe run.
'
' WHY THIS EXISTS
'   Observations accumulate in module state, so a second run in the same session
'   would otherwise report the first run's rows alongside its own and produce a
'   document describing two different experiments as one.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunRibbonSdiProbe
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR BUFFERS
'------------------------------------------------------------------------------
    'A reset that cannot complete must not abort the probe
        On Error Resume Next

        m_RibbonRowCount = 0
        m_RibbonRowsText = vbNullString
        m_RibbonRowsJson = vbNullString

End Sub


Private Sub TST_RibbonProbeRecord( _
    ByVal ScenarioName As String, _
    ByVal WindowLabel As String)

'
'==============================================================================
' TST_RibbonProbeRecord
'------------------------------------------------------------------------------
' PURPOSE
'   Record one observation of Ribbon state through every mechanism available.
'
' WHY THIS EXISTS
'   The component reads the Ribbon through CommandBars first and falls back to
'   the legacy XLM query. Both are application-scoped calls: neither accepts a
'   window. If the Ribbon really is per-window, the only way that can surface is
'   as a DIFFERENCE between readings taken while different windows are active,
'   or as a disagreement between the mechanisms themselves.
'
'   Height is recorded alongside Visible because it is the more sensitive
'   signal. A Ribbon that is collapsed rather than hidden can report Visible as
'   True while its height collapses to the tab strip, and a component that
'   trusted Visible alone would call that state shown.
'
' INPUTS
'   ScenarioName
'     Scenario the observation belongs to.
'
'   WindowLabel
'     Window that was active when the reading was taken.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Reads CommandBars("Ribbon").Visible, its Height, and the XLM query.
'   - Records "err" for any mechanism that raises, rather than a value that
'     would be indistinguishable from a real reading.
'
' ERROR POLICY
'   - Does not raise. An unreadable mechanism is an observation, not a failure.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunRibbonSdiProbe
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim VisibleText         As String          'CommandBars Visible reading
    Dim HeightText          As String          'CommandBars Height reading
    Dim XlmText             As String          'Legacy XLM query reading
    Dim ProbeValue          As Variant         'Working buffer for each read

'------------------------------------------------------------------------------
' READ EVERY MECHANISM
'------------------------------------------------------------------------------
    'An unreadable mechanism is recorded as such; substituting a default would
    'silently invent data this document exists to gather
        On Error Resume Next

        VisibleText = "err"
        HeightText = "err"
        XlmText = "err"

        ProbeValue = Application.CommandBars("Ribbon").Visible

        If Err.Number = 0 Then
            VisibleText = CStr(CBool(ProbeValue))
        End If

        Err.Clear

        ProbeValue = Application.CommandBars("Ribbon").Height

        If Err.Number = 0 Then
            HeightText = CStr(CLng(ProbeValue))
        End If

        Err.Clear

        ProbeValue = Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")

        If Err.Number = 0 Then
            XlmText = CStr(CBool(ProbeValue))
        End If

        Err.Clear

'------------------------------------------------------------------------------
' APPEND OBSERVATION
'------------------------------------------------------------------------------
    'Text for reading, JSON for comparing between hosts
        m_RibbonRowCount = m_RibbonRowCount + 1

        m_RibbonRowsText = m_RibbonRowsText & _
            Left$(ScenarioName & String$(22, " "), 22) & _
            Left$(WindowLabel & String$(8, " "), 8) & _
            Left$(VisibleText & String$(21, " "), 21) & _
            Left$(HeightText & String$(8, " "), 8) & _
            XlmText & vbNewLine

        If m_RibbonRowCount > 1 Then
            m_RibbonRowsJson = m_RibbonRowsJson & ","
        End If

        m_RibbonRowsJson = m_RibbonRowsJson & _
            "{""scenario"":""" & TST_CertJsonEscape(ScenarioName) & """" & _
            ",""window"":""" & TST_CertJsonEscape(WindowLabel) & """" & _
            ",""commandBarsVisible"":""" & TST_CertJsonEscape(VisibleText) & """" & _
            ",""commandBarsHeight"":""" & TST_CertJsonEscape(HeightText) & """" & _
            ",""xlmGetToolBar"":""" & TST_CertJsonEscape(XlmText) & """}"

End Sub


Private Function TST_RibbonProbeBuildJson() _
    As String

'
'==============================================================================
' TST_RibbonProbeBuildJson
'------------------------------------------------------------------------------
' PURPOSE
'   Wrap the accumulated observations in a document that names the host.
'
' WHY THIS EXISTS
'   Ribbon behavior can differ by Excel build, Office channel and policy. An
'   observation that does not say which host produced it cannot be compared with
'   another, and comparison is the entire purpose of gathering it.
'
' RETURNS
'   String
'     A single-line JSON document.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertJsonEscape
'
' CALLED FROM
'   - Test_EXCEL_UI_RunRibbonSdiProbe
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Bitness             As String          'Office bitness of this build

'------------------------------------------------------------------------------
' BUILD DOCUMENT
'------------------------------------------------------------------------------
    'A document must be produced even when a field cannot be read
        On Error Resume Next

#If VBA7 Then
    #If Win64 Then
        Bitness = "x64"
    #Else
        Bitness = "x86"
    #End If
#Else
        Bitness = "x86"
#End If

        TST_RibbonProbeBuildJson = _
            "{""document"":""ribbon-sdi-characterization""" & _
            ",""schema"":1" & _
            ",""timestampLocal"":""" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & """" & _
            ",""excelVersion"":""" & TST_CertJsonEscape(Application.Version) & """" & _
            ",""excelBuild"":""" & TST_CertJsonEscape(CStr(Application.Build)) & """" & _
            ",""operatingSystem"":""" & _
            TST_CertJsonEscape(Application.OperatingSystem) & """" & _
            ",""bitness"":""" & Bitness & """" & _
            ",""observations"":[" & m_RibbonRowsJson & "]}"

End Function


Public Sub Test_EXCEL_UI_RunAll()

'
'==============================================================================
' Test_EXCEL_UI_RunAll
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
' NOTES
'   This is NOT the release gate. It does not run the snapshot-identity and SDI
'   identity units, it can skip snapshot cases silently when a snapshot
'   already exists, and it produces no machine-readable evidence. Use
'   Test_EXCEL_UI_RunReleaseCertification to certify a release.
'
' UPDATED
'   2026-08-19
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
' Test_EXCEL_UI_RunCore
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
'   2026-07-25
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
'
'------------------------------------------------------------------------------
'
'                          PRIVATE PACK RUNNERS
'
'------------------------------------------------------------------------------
'


Public Sub Test_EXCEL_UI_RunSnapshotIdentity()

'
'==============================================================================
' Test_EXCEL_UI_RunSnapshotIdentity
'------------------------------------------------------------------------------
' PURPOSE
'   Verify identity-safe restoration and structured reporting when a captured
'   Excel Window is closed and replaced after snapshot capture.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Refuses to run when an explicit EXCEL_UI snapshot already exists.
'   - Activates the surviving anchor window before capturing, so the captured
'     title-bar frame belongs to a window that outlives the case.
'   - Captures a surviving anchor window and a temporary second window.
'   - Closes the captured temporary window and creates a replacement.
'   - Restores through UI_ResetExcelUIToSnapshot_WithResult.
'   - Verifies the surviving anchor restores.
'   - Verifies the replacement remains unchanged.
'   - Verifies one ordered WindowIdentity failure is returned.
'
' ERROR POLICY
'   - Raises after best-effort cleanup.
'   - Preserves the original test failure through cleanup.
'
' DEPENDENCIES
'   - UI_CaptureExcelUIState_WithResult
'   - UI_ResetExcelUIToSnapshot_WithResult
'   - TST_AssertResultSuccess
'   - TST_AssertSingleFailurePrefix
'   - TST_AssertSnapshotWindowState
'
' NOTES
'   The snapshot captures every window in Application.Windows regardless of
'   which is active; only the title-bar frame is resolved from the active
'   window. That is why activating the anchor before capture changes the
'   expected failure count without weakening what this case covers.
'
' UPDATED
'   2026-08-19 - Anchor the captured title-bar frame to the surviving window,
'                so the case asserts window identity alone.
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window
    Dim CapturedWindow      As Window
    Dim ReplacementWindow   As Window

    Dim SavedHeadings       As Boolean
    Dim SavedWorkbookTabs   As Boolean
    Dim SavedGridlines      As Boolean

    Dim OK                  As Boolean
    Dim FailureCount        As Long
    Dim FailureList         As Variant

    Dim HasFailure          As Boolean
    Dim FailNumber          As Long
    Dim FailSource          As String
    Dim FailDescription     As String

    Dim CreatedSnapshot     As Boolean         'TRUE once this run owns one

    Const PROC As String = "Test_EXCEL_UI_RunSnapshotIdentity"

'------------------------------------------------------------------------------
' VALIDATE PRECONDITIONS
'------------------------------------------------------------------------------
    'Refused before the handler is armed. Armed first, this refusal travelled
    'Err_Handler, resumed at Safe_Exit, and the unconditional clear there
    'deleted the caller snapshot it had just refused to disturb. This is a
    'certification unit, so that reached the release path.
        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 1, _
                PROC, _
                "an explicit EXCEL_UI snapshot already exists; clear or restore it before running this destructive test"
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        CreatedSnapshot = False

        On Error GoTo Err_Handler

        TST_Log PROC, "START", "Identity-safe structured snapshot test started"

        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 2, _
                PROC, _
                "no active Excel window is available"
        End If

        If Not (AnchorWindow.Parent Is ThisWorkbook) Then
            ThisWorkbook.Activate
            Set AnchorWindow = ActiveWindow
        End If

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 3, _
                PROC, _
                "ThisWorkbook could not provide an active Excel window"
        End If

        SavedHeadings = AnchorWindow.DisplayHeadings
        SavedWorkbookTabs = AnchorWindow.DisplayWorkbookTabs
        SavedGridlines = AnchorWindow.DisplayGridlines

'------------------------------------------------------------------------------
' CREATE AND CONFIGURE CAPTURED WINDOWS
'------------------------------------------------------------------------------
        Set CapturedWindow = ThisWorkbook.NewWindow

        If CapturedWindow Is Nothing Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 4, _
                PROC, _
                "ThisWorkbook.NewWindow did not return a temporary captured window"
        End If

        AnchorWindow.DisplayHeadings = True
        AnchorWindow.DisplayWorkbookTabs = False
        AnchorWindow.DisplayGridlines = True

        CapturedWindow.DisplayHeadings = False
        CapturedWindow.DisplayWorkbookTabs = True
        CapturedWindow.DisplayGridlines = False

'------------------------------------------------------------------------------
' ACTIVATE THE SURVIVING WINDOW BEFORE CAPTURE
'------------------------------------------------------------------------------
    'ThisWorkbook.NewWindow activates the window it creates, so without this the
    'capture would record the TEMPORARY window as the owner of the title-bar
    'frame. Closing it below would then produce a second, correct TitleBar
    'failure alongside the WindowIdentity failure this case exists to assert,
    'and the case would be failing for a reason that is not about window
    'identity at all.
    '
    'Anchoring the frame to the surviving window keeps this case about one
    'thing. The closed-frame behavior it would otherwise trip over has its own
    'dedicated case, TST_Case_TitleBarCapturedFrameClosedIsReported.
        AnchorWindow.Activate

        TST_WaitUI TEST_WAIT_SECONDS

'------------------------------------------------------------------------------
' CAPTURE THROUGH STRUCTURED-RESULT API
'------------------------------------------------------------------------------
        FailureCount = 99
        FailureList = Array("stale capture failure")

        OK = UI_CaptureExcelUIState_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

    'Ownership follows an observed snapshot rather than the returned flag. A
    'capture reporting success that left nothing behind must not make this run
    'believe it owns one, and the assertion below can fail before Safe_Exit.
        CreatedSnapshot = UI_HasExcelUIStateSnapshot

    'Post-capture fault injection. The window between owning a snapshot and the
    'assertion below is genuinely fallible - a capture can report success and
    'leave a partial state - and it is the window the ownership flag exists to
    'cover.
        Select Case TST_TakePostCaptureFault()

            Case TEST_FAULT_ASSERTION
                TST_AssertTrue _
                    ActualValue:=False, _
                    AssertionName:=PROC & ".InjectedAssertion"

            Case TEST_FAULT_UNEXPECTED
                Err.Raise _
                    TEST_OWNERSHIP_ERR_BASE + 1, _
                    PROC & ".InjectedUnexpected", _
                    "injected unexpected post-capture failure"

        End Select

        TST_AssertResultSuccess _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            AssertionName:=PROC & ".CaptureResult"

        TST_AssertTrue _
            ActualValue:=UI_HasExcelUIStateSnapshot, _
            AssertionName:=PROC & ".SnapshotAvailable"

'------------------------------------------------------------------------------
' CLOSE CAPTURED WINDOW AND CREATE REPLACEMENT
'------------------------------------------------------------------------------
        AnchorWindow.DisplayHeadings = False
        AnchorWindow.DisplayWorkbookTabs = True
        AnchorWindow.DisplayGridlines = False

        CapturedWindow.Close
        Set CapturedWindow = Nothing

        Set ReplacementWindow = ThisWorkbook.NewWindow

        If ReplacementWindow Is Nothing Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 5, _
                PROC, _
                "ThisWorkbook.NewWindow did not return a replacement window"
        End If

        ReplacementWindow.DisplayHeadings = True
        ReplacementWindow.DisplayWorkbookTabs = False
        ReplacementWindow.DisplayGridlines = True

'------------------------------------------------------------------------------
' RESTORE AND ASSERT STRUCTURED FAILURE
'------------------------------------------------------------------------------
    'ThisWorkbook.NewWindow activated the replacement. The Ribbon is restored
    'only while the window it was captured from is active (#23), so restoring
    'with the replacement active would add a correct Ribbon failure to the
    'WindowIdentity failure this case exists to assert. Reactivating the anchor
    'keeps the case about window identity; the Ribbon refusal has its own case,
    'TST_Case_RibbonRestoreRefusesOtherActiveWindow.
        AnchorWindow.Activate

        TST_WaitUI TEST_WAIT_SECONDS

        FailureCount = 99
        FailureList = Array("stale restore failure")

        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertSingleFailurePrefix _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedPrefix:="WindowIdentity [", _
            AssertionName:=PROC & ".RestoreResult"

        TST_AssertSnapshotWindowState _
            TargetWindow:=AnchorWindow, _
            ExpectedHeadings:=True, _
            ExpectedWorkbookTabs:=False, _
            ExpectedGridlines:=True, _
            AssertionName:=PROC & ".AnchorRestored"

        TST_AssertSnapshotWindowState _
            TargetWindow:=ReplacementWindow, _
            ExpectedHeadings:=True, _
            ExpectedWorkbookTabs:=False, _
            ExpectedGridlines:=True, _
            AssertionName:=PROC & ".ReplacementUnchanged"

        TST_Log PROC, "PASS", _
            "Identity-safe restore returned the expected ordered failure without touching replacement"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
        On Error Resume Next

    'Only a snapshot this run created is released
        If CreatedSnapshot Then
            UI_ClearExcelUIStateSnapshot
        End If

        TST_SafeCloseWindow ReplacementWindow
        TST_SafeCloseWindow CapturedWindow

        If Not AnchorWindow Is Nothing Then
            AnchorWindow.DisplayHeadings = SavedHeadings
            AnchorWindow.DisplayWorkbookTabs = SavedWorkbookTabs
            AnchorWindow.DisplayGridlines = SavedGridlines
            AnchorWindow.Activate
        End If

        On Error GoTo 0

        If HasFailure Then
            TST_Log PROC, "FAIL", _
                "Error " & CStr(FailNumber) & _
                " | Source: " & FailSource & _
                " | " & FailDescription

            Err.Raise _
                Number:=FailNumber, _
                Source:=FailSource, _
                Description:=FailDescription
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


Private Sub TST_RunRegressionPack( _
    ByVal IncludeTitleBarTests As Boolean, _
    ByVal CallerProc As String)

'
'==============================================================================
' TST_RunRegressionPack
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
'     FALSE => skip the wrapper case and the dedicated title-bar cases
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
' NOTES
'   - This pack is the body of the RegressionPack certification unit, so a case
'     reachable only from the title-bar-only pack is absent from the evidence a
'     release is tagged on. Every frame-state case is registered here as well.
'
' UPDATED
'   2026-11-06 - Put the entry frame's style back through the harness and
'                failed the pack when it could not (#66).
'   2026-11-06 - Registered the WinAPI error-capture case (#27).
'   2026-11-06 - Registered the certification unit-isolation case (#135).
'   2026-08-21 - Registered the frame-refresh-debt and stale-frame-entry cases,
'                which certification did not reach.
'   2026-07-29
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

    Dim HadExplicitSnapshot As Boolean         'TRUE when an explicit EXCEL_UI snapshot already existed before the run
    Dim OldScreenUpdating   As Boolean         'Cached ScreenUpdating state
    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

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

    'Record the active frame's exact style through the harness, so cleanup can
    'put it back without the API under test. Core mode never touches the
    'title bar, so it records nothing and writes nothing.
        If IncludeTitleBarTests Then
            EntryHwnd = Application.hWnd
            EntryFrameCaptured = TST_TryGetWindowStyle( _
                EntryHwnd, EntryFrameStyle, FrameRestoreMsg)

            If Not EntryFrameCaptured Then
                TST_Log CallerProc, "TitleBar", _
                    "entry frame style not recorded | " & FrameRestoreMsg
            End If
        End If

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

    'Run active-window targeting
        TST_Case_TargetScope_ActiveWindow

    'Run active-workbook-window targeting
        TST_Case_TargetScope_ActiveWorkbookWindows

    'Run invalid-target-scope structured failure and continuation
        TST_Case_TargetScope_InvalidValue

    'Run the ScreenUpdating preservation case
        TST_Case_ScreenUpdatingPreserved

    'Verify diagnostics degrade rather than raise when the list cannot grow
        TST_Case_FailureAccumulatorDegradesSafely

    'Verify cleanup is judged against the entry state, not an assumed ideal
        TST_Case_CertificationCleanupUsesBaseline

'------------------------------------------------------------------------------
' RUN OPTIONAL SNAPSHOT CASES
'------------------------------------------------------------------------------
    'Run snapshot-related cases only when no explicit EXCEL_UI snapshot already
    'existed before the run because the harness cannot restore that prior
    'snapshot object safely
        If HadExplicitSnapshot Then

            'Record that snapshot-destructive cases were skipped. Under the
            'certification runner a skip is a failure, not a quiet log line.
                TST_CertRecordSkip CallerProc, _
                    "Snapshot lifecycle cases skipped because an explicit EXCEL_UI snapshot already existed before the run"

        Else

            'Run structured snapshot capture clean-success case
                TST_Case_SnapshotCaptureResultSuccess IncludeTitleBarTests

            'Run structured snapshot restoration clean-success case
                TST_Case_SnapshotResetResultSuccess IncludeTitleBarTests

            'Run structured restoration no-snapshot failure case
                TST_Case_SnapshotResetResultNoSnapshot IncludeTitleBarTests

            'Run per-element application-level capture and restoration case
                TST_Case_SnapshotCapturePartialApplicationRead

            'Run the compatibility-wrapper snapshot lifecycle case
                TST_Case_SnapshotLifecycle IncludeTitleBarTests

            'Run the compatibility-wrapper reset-without-snapshot no-op case
                TST_Case_ResetWithoutSnapshot_NoOp IncludeTitleBarTests

            'Run the destructive-runner snapshot-ownership case
                TST_Case_DestructiveRunnersPreserveCallerSnapshot

            'Run the certification unit-isolation case for a failed capture
                TST_Case_CertUnitIsolatesCaptureFailure

        End If

    'Quiet-scope ownership captures no snapshot, so it runs whether or not the
    'caller already held one
        TST_Case_QuietScopeOwnershipIsVerified

'------------------------------------------------------------------------------
' RUN OPTIONAL TITLE-BAR / WRAPPER CASES
'------------------------------------------------------------------------------
    'Run the convenience-wrapper case only when title-bar testing is enabled
    'because the wrappers also affect TitleBar
        If IncludeTitleBarTests Then
            TST_Case_ConvenienceWrappers True
        Else
            TST_CertRecordSkip CallerProc, _
                "Convenience-wrapper case skipped in core mode because the wrappers also toggle TitleBar"
        End If

    'Run the dedicated title-bar cases when requested. Every frame-state case
    'belongs here, not only in the title-bar-only pack: this pack is what the
    'certification runner executes, so a case reachable only from the other
    'pack is absent from the evidence a release is tagged on.
        If IncludeTitleBarTests Then
            TST_Case_TitleBarRoundTrip
            TST_Case_TitleBarOwnedBitPreservation
            TST_Case_TitleBarShowRecoversWithoutBaseline
            TST_Case_TitleBarShowRejectsCaptionlessBaseline
            TST_Case_TitleBarFrameRefreshDebtRetried
            TST_Case_TitleBarStaleFrameEntryNotReused
            TST_Case_TitleBarSameStyleHandleReuse
            TST_Case_ActiveFramePairRefusesMismatch
            TST_Case_TitleBarApiErrorCapture
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
    'Log successful completion before restoration
        TST_Log CallerProc, "PASS", "All requested regression cases passed"

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

    'Log the failure immediately
        TST_Log CallerProc, "FAIL", _
            CStr(FailNumber) & ": " & FailDescription & _
            IIf(Len(FailSource) > 0, " | Source: " & FailSource, vbNullString)

        Resume Safe_Exit

End Sub


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


Private Sub TST_Case_TitleBarShowRejectsCaptionlessBaseline()
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


Private Sub TST_Case_TitleBarSameStyleHandleReuse()
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
' TST_Case_ShowAllBaseline
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
'   2026-07-25
'==============================================================================
'
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

Private Sub TST_Case_SelectiveHide(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_SelectiveHide
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
'   2026-07-25
'==============================================================================
'
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

Private Sub TST_Case_SelectiveShow(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_SelectiveShow
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
'   2026-07-25
'==============================================================================
'
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

Private Sub TST_Case_NoOpLeaveUnchanged(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_NoOpLeaveUnchanged
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
'   2026-07-25
'==============================================================================
'
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

Private Sub TST_Case_ConvenienceWrappers(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_ConvenienceWrappers
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
'   2026-07-25
'==============================================================================
'
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

Private Sub TST_Case_WithResult_AllSuccess(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_WithResult_AllSuccess
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean         'Boolean success flag returned by the API
    Dim FailureCount        As Long            'Number of recorded failures
    Dim FailureList         As Variant         'Optional array of recorded failures

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

Private Sub TST_Case_WithResult_NoOpSuccess(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_WithResult_NoOpSuccess
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean         'Boolean success flag returned by the API
    Dim FailureCount        As Long            'Number of recorded failures
    Dim FailureList         As Variant         'Optional array of recorded failures

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

Private Sub TST_Case_WithResult_SuccessWithoutFailureList(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_WithResult_SuccessWithoutFailureList
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
'   2026-11-06 - Failure messages carry the structured result (#135).
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean         'Boolean success flag returned by the API
    Dim FailureCount        As Long            'Number of recorded failures
    Dim FailureList         As Variant         'Local untouched Variant proving omission path

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

Private Sub TST_Case_WithResult_InvalidVisibility()

'
'==============================================================================
' TST_Case_WithResult_InvalidVisibility
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
'   2026-11-06 - Failure messages carry the structured result (#135).
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean         'Boolean success flag returned by the API
    Dim FailureCount        As Long            'Number of recorded failures
    Dim FailureList         As Variant         'Recorded structured failures

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

Private Sub TST_Case_TargetScope_ActiveWindow()

'
'==============================================================================
' TST_Case_TargetScope_ActiveWindow
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that UI_TargetActiveWindow changes only the active Excel Window while
'   application-level requests still apply normally.
'
' ERROR POLICY
'   - Raises after best-effort cleanup.
'
' UPDATED
'   2026-08-01
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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

    Const PROC As String = "TST_Case_TargetScope_ActiveWindow"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        ThisWorkbook.Activate
        Set AnchorWindow = Application.ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise TEST_TARGET_ERR_BASE + 1, PROC, _
                "ThisWorkbook could not provide an active Excel window"
        End If

        SavedHeadings = AnchorWindow.DisplayHeadings
        SavedWorkbookTabs = AnchorWindow.DisplayWorkbookTabs
        SavedGridlines = AnchorWindow.DisplayGridlines
        SavedStatusBar = Application.DisplayStatusBar

        Set TargetWindow = ThisWorkbook.NewWindow

        If TargetWindow Is Nothing Then
            Err.Raise TEST_TARGET_ERR_BASE + 2, PROC, _
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
            AssertionName:=PROC & ".Result"

'------------------------------------------------------------------------------
' ASSERT SCOPE
'------------------------------------------------------------------------------
        TST_AssertApplicationProperty _
            Expected:=False, _
            PropertyName:="DisplayStatusBar", _
            AssertionName:=PROC & ".ApplicationLevelUnaffectedByScope"

        TST_AssertSnapshotWindowState _
            TargetWindow:=TargetWindow, _
            ExpectedHeadings:=False, _
            ExpectedWorkbookTabs:=False, _
            ExpectedGridlines:=False, _
            AssertionName:=PROC & ".TargetWindow"

        TST_AssertSnapshotWindowState _
            TargetWindow:=AnchorWindow, _
            ExpectedHeadings:=True, _
            ExpectedWorkbookTabs:=True, _
            ExpectedGridlines:=True, _
            AssertionName:=PROC & ".NonTargetWindow"

        TST_Log PROC, "PASS", _
            "Only the active window received window-level changes"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
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
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        Resume Safe_Exit

End Sub


Private Sub TST_Case_TargetScope_ActiveWorkbookWindows()

'
'==============================================================================
' TST_Case_TargetScope_ActiveWorkbookWindows
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that UI_TargetActiveWorkbookWindows changes every Window belonging to
'   the active workbook and leaves another workbook's Window unchanged.
'
' ERROR POLICY
'   - Raises after best-effort cleanup.
'
' UPDATED
'   2026-08-01
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OriginalWindow      As Window
    Dim TargetBook          As Workbook
    Dim OtherBook           As Workbook
    Dim TargetWindowOne     As Window
    Dim TargetWindowTwo     As Window
    Dim OtherWindow         As Window
    Dim OK                  As Boolean
    Dim FailureCount        As Long
    Dim FailureList         As Variant
    Dim HasFailure          As Boolean
    Dim FailNumber          As Long
    Dim FailSource          As String
    Dim FailDescription     As String

    Const PROC As String = "TST_Case_TargetScope_ActiveWorkbookWindows"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        Set OriginalWindow = Application.ActiveWindow

        Set TargetBook = Application.Workbooks.Add
        Set TargetWindowOne = TargetBook.Windows(1)
        Set TargetWindowTwo = TargetBook.NewWindow

        Set OtherBook = Application.Workbooks.Add
        Set OtherWindow = OtherBook.Windows(1)

        If TargetWindowOne Is Nothing Or _
           TargetWindowTwo Is Nothing Or _
           OtherWindow Is Nothing Then

            Err.Raise TEST_TARGET_ERR_BASE + 3, PROC, _
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
            Err.Raise TEST_TARGET_ERR_BASE + 4, PROC, _
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
            AssertionName:=PROC & ".Result"

'------------------------------------------------------------------------------
' ASSERT SCOPE
'------------------------------------------------------------------------------
        TST_AssertSnapshotWindowState _
            TargetWindow:=TargetWindowOne, _
            ExpectedHeadings:=False, _
            ExpectedWorkbookTabs:=False, _
            ExpectedGridlines:=False, _
            AssertionName:=PROC & ".TargetWindowOne"

        TST_AssertSnapshotWindowState _
            TargetWindow:=TargetWindowTwo, _
            ExpectedHeadings:=False, _
            ExpectedWorkbookTabs:=False, _
            ExpectedGridlines:=False, _
            AssertionName:=PROC & ".TargetWindowTwo"

        TST_AssertSnapshotWindowState _
            TargetWindow:=OtherWindow, _
            ExpectedHeadings:=True, _
            ExpectedWorkbookTabs:=True, _
            ExpectedGridlines:=True, _
            AssertionName:=PROC & ".OtherWorkbookWindow"

        TST_Log PROC, "PASS", _
            "Only windows belonging to the active workbook were changed"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
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
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        Resume Safe_Exit

End Sub


Private Sub TST_Case_TargetScope_InvalidValue()

'
'==============================================================================
' TST_Case_TargetScope_InvalidValue
'------------------------------------------------------------------------------
' PURPOSE
'   Verify ordered invalid-scope diagnostics, application-level continuation,
'   and suppression of window-level writes when TargetScope is invalid.
'
' ERROR POLICY
'   - Raises after best-effort cleanup.
'
' UPDATED
'   2026-08-01
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window
    Dim InvalidScope        As UIWindowTargetScope
    Dim SavedHeadings       As Boolean
    Dim SavedStatusBar      As Boolean
    Dim OK                  As Boolean
    Dim FailureCount        As Long
    Dim FailureList         As Variant
    Dim HasFailure          As Boolean
    Dim FailNumber          As Long
    Dim FailSource          As String
    Dim FailDescription     As String

    Const PROC As String = "TST_Case_TargetScope_InvalidValue"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        ThisWorkbook.Activate
        Set AnchorWindow = Application.ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise TEST_TARGET_ERR_BASE + 5, PROC, _
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
            AssertionName:=PROC & ".Result"

'------------------------------------------------------------------------------
' ASSERT CONTINUATION AND WINDOW SUPPRESSION
'------------------------------------------------------------------------------
        TST_AssertApplicationProperty _
            Expected:=False, _
            PropertyName:="DisplayStatusBar", _
            AssertionName:=PROC & ".ApplicationLevelContinued"

        TST_AssertBooleanEquals _
            Expected:=True, _
            Actual:=AnchorWindow.DisplayHeadings, _
            AssertionName:=PROC & ".WindowLevelSkipped"

        TST_Log PROC, "PASS", _
            "Invalid scope was reported while application-level work continued"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
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
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        Resume Safe_Exit

End Sub


Private Sub TST_Case_SnapshotCaptureResultSuccess( _
    ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_SnapshotCaptureResultSuccess
'------------------------------------------------------------------------------
' PURPOSE
'   Verify the clean-success contract and deterministic output clearing of
'   UI_CaptureExcelUIState_WithResult.
'
' WHY
'   Snapshot capture now exposes the same Boolean/count/list contract as the
'   existing structured apply API.
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE to include title-bar baseline setup; FALSE to leave it unchanged.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Seeds stale output values.
'   - Captures a deterministic mixed UI baseline.
'   - Verifies clean success, zero failures, Empty FailureList, and snapshot
'     availability.
'
' ERROR POLICY
'   - Raises on assertion failure.
'
' DEPENDENCIES
'   - UI_CaptureExcelUIState_WithResult
'   - TST_AssertResultSuccess
'   - TST_AssertSnapshotAvailability
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean
    Dim FailureCount        As Long
    Dim FailureList         As Variant

    Const PROC As String = "TST_Case_SnapshotCaptureResultSuccess"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log PROC, "START", _
            "Validating structured snapshot capture clean-success path"

        UI_ClearExcelUIStateSnapshot

        UI_SetExcelUI _
            Ribbon:=UI_Show, _
            StatusBar:=UI_Hide, _
            ScrollBars:=UI_Show, _
            FormulaBar:=UI_Hide, _
            Headings:=UI_Show, _
            WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Show, _
            TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Show)

        TST_WaitUI TEST_WAIT_SECONDS

        FailureCount = 99
        FailureList = Array("stale capture failure")

'------------------------------------------------------------------------------
' CAPTURE AND ASSERT RESULT
'------------------------------------------------------------------------------
        OK = UI_CaptureExcelUIState_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertResultSuccess _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            AssertionName:=PROC & ".Result"

        TST_AssertSnapshotAvailability True, PROC & ".SnapshotAvailable"

        UI_ClearExcelUIStateSnapshot

        TST_Log PROC, "PASS", _
            "Structured snapshot capture returned clean success and cleared stale outputs"

End Sub


Private Sub TST_Case_SnapshotResetResultSuccess( _
    ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_SnapshotResetResultSuccess
'------------------------------------------------------------------------------
' PURPOSE
'   Verify the clean-success contract and deterministic output clearing of
'   UI_ResetExcelUIToSnapshot_WithResult.
'
' WHY
'   Structured restoration must report success without changing the established
'   snapshot lifecycle or host-state preservation contract.
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE to include title-bar mutation/restoration; FALSE to leave it
'     unchanged.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Captures a deterministic mixed baseline.
'   - Mutates every managed UI surface.
'   - Seeds stale output values.
'   - Restores through the structured API.
'   - Verifies clean result buffers, restored state, retained snapshot, and
'     ScreenUpdating preservation.
'
' ERROR POLICY
'   - Raises on assertion failure.
'
' DEPENDENCIES
'   - UI_ResetExcelUIToSnapshot_WithResult
'   - TST_AssertResultSuccess
'   - managed-state assertion helpers
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OK                  As Boolean
    Dim FailureCount        As Long
    Dim FailureList         As Variant
    Dim SavedScreenUpdating As Boolean

    Const PROC As String = "TST_Case_SnapshotResetResultSuccess"

'------------------------------------------------------------------------------
' INITIALIZE AND CAPTURE BASELINE
'------------------------------------------------------------------------------
        TST_Log PROC, "START", _
            "Validating structured snapshot restoration clean-success path"

        UI_ClearExcelUIStateSnapshot

        UI_SetExcelUI _
            Ribbon:=UI_Show, _
            StatusBar:=UI_Hide, _
            ScrollBars:=UI_Show, _
            FormulaBar:=UI_Hide, _
            Headings:=UI_Show, _
            WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Show, _
            TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Show)

        TST_WaitUI TEST_WAIT_SECONDS

        UI_CaptureExcelUIState
        TST_AssertSnapshotAvailability True, PROC & ".SnapshotAvailable"

'------------------------------------------------------------------------------
' MUTATE AND RESTORE
'------------------------------------------------------------------------------
        UI_SetExcelUI _
            Ribbon:=UI_Hide, _
            StatusBar:=UI_Show, _
            ScrollBars:=UI_Hide, _
            FormulaBar:=UI_Show, _
            Headings:=UI_Hide, _
            WorkbookTabs:=UI_Show, _
            Gridlines:=UI_Hide, _
            TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Hide)

        TST_WaitUI TEST_WAIT_SECONDS

        FailureCount = 99
        FailureList = Array("stale restore failure")

        SavedScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = True

        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        If Not Application.ScreenUpdating Then
            Err.Raise TEST_ERR_BASE + 70, _
                      PROC & ".ScreenUpdating", _
                      "structured restoration did not preserve ScreenUpdating=True"
        End If

        Application.ScreenUpdating = SavedScreenUpdating

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertResultSuccess _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            AssertionName:=PROC & ".Result"

'------------------------------------------------------------------------------
' ASSERT RESTORED STATE
'------------------------------------------------------------------------------
        TST_AssertRibbonVisible True, PROC & ".Ribbon"
        TST_AssertApplicationProperty False, "DisplayStatusBar", PROC & ".StatusBar"
        TST_AssertApplicationProperty True, "DisplayScrollBars", PROC & ".ScrollBars"
        TST_AssertApplicationProperty False, "DisplayFormulaBar", PROC & ".FormulaBar"
        TST_AssertAllWindowsProperty True, "DisplayHeadings", PROC & ".Headings"
        TST_AssertAllWindowsProperty False, "DisplayWorkbookTabs", PROC & ".WorkbookTabs"
        TST_AssertAllWindowsProperty True, "DisplayGridlines", PROC & ".Gridlines"

        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, PROC & ".TitleBar"
        End If

        TST_AssertSnapshotAvailability True, PROC & ".SnapshotRetained"

        UI_ClearExcelUIStateSnapshot

        TST_Log PROC, "PASS", _
            "Structured snapshot restoration returned clean success and restored state"

End Sub


Private Sub TST_Case_SnapshotResetResultNoSnapshot( _
    ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_SnapshotResetResultNoSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Verify deterministic structured failure reporting and no-op behavior when
'   restoration is requested without an available snapshot.
'
' WHY
'   The new result API must make the existing no-snapshot diagnostic
'   machine-readable without altering host state.
'
' INPUTS
'   IncludeTitleBarTests
'     TRUE to include title-bar state assertions; FALSE to omit them.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Clears any existing snapshot.
'   - Establishes a deterministic mixed baseline.
'   - Seeds stale output values.
'   - Verifies one ordered NoSnapshot failure.
'   - Verifies all managed state and ScreenUpdating remain unchanged.
'   - With title-bar tests, verifies the active frame's complete style is
'     unchanged by the refused restore.
'
' ERROR POLICY
'   - Raises on assertion failure.
'
' DEPENDENCIES
'   - UI_ResetExcelUIToSnapshot_WithResult
'   - TST_AssertSingleFailurePrefix
'   - TST_FrameStyleOrRaise
'   - TST_AssertFrameUntouched
'   - managed-state assertion helpers
'
' UPDATED
'   2026-11-06 - Compared the active frame's complete style (#45).
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim ActiveHwnd          As LongPtr         'Frame active at restore time
    Dim FrameBefore         As LongPtr         'Its style before restore
#Else
    Dim ActiveHwnd          As Long            'Frame active at restore time
    Dim FrameBefore         As Long            'Its style before restore
#End If

    Dim OK                  As Boolean
    Dim FailureCount        As Long
    Dim FailureList         As Variant
    Dim SavedScreenUpdating As Boolean

    Const PROC As String = "TST_Case_SnapshotResetResultNoSnapshot"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log PROC, "START", _
            "Validating structured no-snapshot restoration failure"

        UI_ClearExcelUIStateSnapshot

        UI_SetExcelUI _
            Ribbon:=UI_Show, _
            StatusBar:=UI_Hide, _
            ScrollBars:=UI_Show, _
            FormulaBar:=UI_Hide, _
            Headings:=UI_Show, _
            WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Show, _
            TitleBar:=TST_TitleBarMode(IncludeTitleBarTests, UI_Show)

        TST_WaitUI TEST_WAIT_SECONDS

        FailureCount = 99
        FailureList = Array("stale no-snapshot failure")

    'Record the active frame exactly as it stands before the refused restore
        If IncludeTitleBarTests Then
            ActiveHwnd = Application.hWnd
            FrameBefore = TST_FrameStyleOrRaise(ActiveHwnd, PROC)
        End If

        SavedScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = True

'------------------------------------------------------------------------------
' RESTORE WITHOUT SNAPSHOT
'------------------------------------------------------------------------------
        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        If Not Application.ScreenUpdating Then
            Err.Raise TEST_ERR_BASE + 71, _
                      PROC & ".ScreenUpdating", _
                      "no-snapshot restoration did not preserve ScreenUpdating=True"
        End If

        Application.ScreenUpdating = SavedScreenUpdating

        TST_AssertSingleFailurePrefix _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedPrefix:="NoSnapshot | ", _
            AssertionName:=PROC & ".Result"

'------------------------------------------------------------------------------
' ASSERT NO-OP STATE
'------------------------------------------------------------------------------
        TST_AssertRibbonVisible True, PROC & ".Ribbon"
        TST_AssertApplicationProperty False, "DisplayStatusBar", PROC & ".StatusBar"
        TST_AssertApplicationProperty True, "DisplayScrollBars", PROC & ".ScrollBars"
        TST_AssertApplicationProperty False, "DisplayFormulaBar", PROC & ".FormulaBar"
        TST_AssertAllWindowsProperty True, "DisplayHeadings", PROC & ".Headings"
        TST_AssertAllWindowsProperty False, "DisplayWorkbookTabs", PROC & ".WorkbookTabs"
        TST_AssertAllWindowsProperty True, "DisplayGridlines", PROC & ".Gridlines"

        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, PROC & ".TitleBar"
            TST_AssertFrameUntouched _
                ActiveHwnd, FrameBefore, PROC & ".FrameUntouched"
        End If

        TST_Log PROC, "PASS", _
            "Structured no-snapshot failure was ordered and host state remained unchanged"

End Sub


Private Sub TST_Case_SnapshotCapturePartialApplicationRead()
'
'==============================================================================
' TST_Case_SnapshotCapturePartialApplicationRead
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a snapshot capture which cannot read one application-level
'   property still produces a usable snapshot, and that restoration leaves the
'   unreadable element alone rather than writing a default value over it.
'
' WHY THIS EXISTS
'   This is the regression guard for the defect where one failed application-
'   level read discarded the entire snapshot.
'
'   The three application-level properties were read directly under an active
'   On Error GoTo, so an ordinary host refusal reached the module error handler,
'   which clears the snapshot outright. A caller lost the Ribbon state, the
'   frame state and every captured window identity because the status bar
'   happened to be unreadable, and UI_CaptureExcelUIState returns nothing, so
'   the loss was silent until restore time.
'
'   The existing snapshot cases all exercise the clean path, so none of them
'   can detect a regression here.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Captures a baseline through the structured-result API.
'   - Confirms the snapshot is available and the pass reported clean success.
'   - Changes every managed application-level element away from the baseline.
'   - Restores and confirms each captured value came back.
'   - Confirms the Known-flag contract by restoring twice: the second pass must
'     report clean success and leave host state unchanged, proving restoration
'     is idempotent and reads no stale buffer.
'
' ERROR POLICY
'   - Raises a TEST_ERR_BASE assertion error on failure, for the pack handler.
'   - Clears the snapshot before exiting on every path.
'
' DEPENDENCIES
'   - UI_CaptureExcelUIState_WithResult
'   - UI_ResetExcelUIToSnapshot_WithResult
'   - UI_ClearExcelUIStateSnapshot
'   - UI_SetExcelUI
'   - TST_AssertResultSuccess
'   - TST_AssertSnapshotAvailability
'   - TST_AssertApplicationProperty
'   - TST_WaitUI
'   - TST_Log
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   A host refusal on Application.DisplayStatusBar cannot be provoked from VBA
'   in a normal desktop session, so this case cannot force the failing read
'   itself. What it does guard is the contract that broke around it: that every
'   application-level element is captured and restored through its own Known
'   flag, independently of the others, and that no element is written from a
'   value that was never captured.
'
'   A regression that reintroduced the shared failure path would surface here as
'   a lost snapshot or an unrestored element, because capture and restore no
'   longer treat the three properties as one indivisible step.
'
' UPDATED
'   2026-11-06
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CaptureOK           As Boolean         'Structured capture result
    Dim RestoreOK           As Boolean         'Structured restoration result
    Dim SecondRestoreOK     As Boolean         'Second restoration result
    Dim FailureCount        As Long            'Structured failure count
    Dim FailureList         As Variant         'Structured failure list

    Dim BaseStatusBar       As Boolean         'Status bar at capture time
    Dim BaseScrollBars      As Boolean         'Scroll bars at capture time
    Dim BaseFormulaBar      As Boolean         'Formula bar at capture time

    Dim SavedErrNumber      As Long            'Captured assertion error number
    Dim SavedErrSource      As String          'Captured assertion error source
    Dim SavedErrDescription As String          'Captured assertion description

    Const PROC As String = "TST_Case_SnapshotCapturePartialApplicationRead"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route assertion and runtime errors to the clearing handler
        On Error GoTo Err_Handler

    'Announce the case
        TST_Log PROC, "START", _
            "Validating per-element application-level capture and restoration"

    'Start from a known baseline so each element can be flipped individually
        UI_SetExcelUI _
            StatusBar:=UI_Show, _
            ScrollBars:=UI_Show, _
            FormulaBar:=UI_Show

        TST_WaitUI TEST_WAIT_SECONDS

    'Record what the capture is expected to preserve
        BaseStatusBar = Application.DisplayStatusBar
        BaseScrollBars = Application.DisplayScrollBars
        BaseFormulaBar = Application.DisplayFormulaBar

'------------------------------------------------------------------------------
' CAPTURE BASELINE
'------------------------------------------------------------------------------
    'Capture through the structured-result API so the pass can be inspected
        CaptureOK = UI_CaptureExcelUIState_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

    'A readable host must produce a clean capture
        TST_AssertResultSuccess _
            CaptureOK, FailureCount, FailureList, _
            "PartialApplicationRead.Capture"

    'The snapshot must exist regardless of any optional element
        TST_AssertSnapshotAvailability True, _
            "PartialApplicationRead.SnapshotAvailable"

'------------------------------------------------------------------------------
' DISTURB APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Move every managed application-level element away from the baseline
        UI_SetExcelUI _
            StatusBar:=UI_Hide, _
            ScrollBars:=UI_Hide, _
            FormulaBar:=UI_Hide

        TST_WaitUI TEST_WAIT_SECONDS

    'Confirm the disturbance actually took effect, so the restoration
    'assertions below cannot pass against unchanged state
        TST_AssertApplicationProperty False, "DisplayStatusBar", _
            "PartialApplicationRead.Disturbed.StatusBar"

        TST_AssertApplicationProperty False, "DisplayScrollBars", _
            "PartialApplicationRead.Disturbed.ScrollBars"

        TST_AssertApplicationProperty False, "DisplayFormulaBar", _
            "PartialApplicationRead.Disturbed.FormulaBar"

'------------------------------------------------------------------------------
' RESTORE AND VERIFY EACH ELEMENT
'------------------------------------------------------------------------------
    'Restore through the structured-result API
        RestoreOK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertResultSuccess _
            RestoreOK, FailureCount, FailureList, _
            "PartialApplicationRead.Restore"

    'Each application-level element must be restored independently
        TST_AssertApplicationProperty BaseStatusBar, "DisplayStatusBar", _
            "PartialApplicationRead.Restored.StatusBar"

        TST_AssertApplicationProperty BaseScrollBars, "DisplayScrollBars", _
            "PartialApplicationRead.Restored.ScrollBars"

        TST_AssertApplicationProperty BaseFormulaBar, "DisplayFormulaBar", _
            "PartialApplicationRead.Restored.FormulaBar"

'------------------------------------------------------------------------------
' VERIFY RESTORATION IS REPEATABLE
'------------------------------------------------------------------------------
    'The snapshot is retained after a restore, so a second pass must succeed
    'and must not disturb the state the first pass established
        SecondRestoreOK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertResultSuccess _
            SecondRestoreOK, FailureCount, FailureList, _
            "PartialApplicationRead.SecondRestore"

        TST_AssertApplicationProperty BaseStatusBar, "DisplayStatusBar", _
            "PartialApplicationRead.Repeat.StatusBar"

        TST_AssertApplicationProperty BaseScrollBars, "DisplayScrollBars", _
            "PartialApplicationRead.Repeat.ScrollBars"

        TST_AssertApplicationProperty BaseFormulaBar, "DisplayFormulaBar", _
            "PartialApplicationRead.Repeat.FormulaBar"

'------------------------------------------------------------------------------
' RELEASE SNAPSHOT
'------------------------------------------------------------------------------
    'Leave no captured baseline behind for later cases
        UI_ClearExcelUIStateSnapshot

        TST_AssertSnapshotAvailability False, _
            "PartialApplicationRead.SnapshotCleared"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Report the pass and exit before the error-handler block
        TST_Log PROC, "PASS", _
            "Application-level elements captured and restored independently"

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Retain the failure first. UI_ClearExcelUIStateSnapshot executes an On Error
    'statement, and any form of On Error resets the Err object, so reading Err
    'after that call would re-raise error zero.
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description

    'Release the snapshot so a failing case never leaves a stale baseline for
    'the cases that follow
        UI_ClearExcelUIStateSnapshot

    'Hand the original failure to the pack handler
        Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription

End Sub


Private Sub TST_Case_SnapshotLifecycle(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_SnapshotLifecycle
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
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
        TST_Log "TST_Case_SnapshotLifecycle", "PASS", "Explicit snapshot and reset lifecycle behaved as expected"

End Sub

Private Sub TST_Case_ResetWithoutSnapshot_NoOp(ByVal IncludeTitleBarTests As Boolean)

'
'==============================================================================
' TST_Case_ResetWithoutSnapshot_NoOp
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
'     TRUE  => include TitleBar, and the active frame's complete style, in
'              the unchanged-state assertion
'     FALSE => skip TitleBar assertions in this case
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on assertion failure
'
' UPDATED
'   2026-11-06 - Compared the active frame's complete style (#45).
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim ActiveHwnd          As LongPtr         'Frame active at reset time
    Dim FrameBefore         As LongPtr         'Its style before reset
#Else
    Dim ActiveHwnd          As Long            'Frame active at reset time
    Dim FrameBefore         As Long            'Its style before reset
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
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
    'Record the active frame exactly as it stands before the refused reset
        If IncludeTitleBarTests Then
            ActiveHwnd = Application.hWnd
            FrameBefore = TST_FrameStyleOrRaise( _
                ActiveHwnd, "TST_Case_ResetWithoutSnapshot_NoOp")
        End If

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

    'Assert TitleBar remained visible, and no bit of the active frame changed,
    'when title-bar testing is enabled
        If IncludeTitleBarTests Then
            TST_AssertTitleBarVisible True, "ResetWithoutSnapshot.TitleBar"
            TST_AssertFrameUntouched _
                ActiveHwnd, FrameBefore, "ResetWithoutSnapshot.FrameUntouched"
        End If

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log "TST_Case_ResetWithoutSnapshot_NoOp", "PASS", _
            "Reset-to-snapshot no-op behavior without snapshot behaved as expected"

End Sub

Private Sub TST_Case_ScreenUpdatingPreserved()

'
'==============================================================================
' TST_Case_ScreenUpdatingPreserved
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedScreenUpdating As Boolean         'Caller-visible ScreenUpdating baseline

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

Private Sub TST_Case_CertificationCleanupUsesBaseline()

'
'==============================================================================
' TST_Case_CertificationCleanupUsesBaseline
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that the certification cleanup verdict compares the host with the
'   state observed on entry, rather than requiring an assumed ideal.
'
' WHY THIS EXISTS
'   Cleanup required Application.ScreenUpdating to be True. A run started from
'   within a quiet-update scope therefore failed certification even though the
'   regression pack had correctly restored the suppressed value it found.
'
'   The damage was to the verdict's credibility rather than to the host: a
'   counter that fires when nothing is wrong is a counter a reader stops
'   believing, which defeats the verdict from the opposite direction to a missed
'   failure.
'
'   The case drives TST_CertEvaluateCleanup directly. Reaching the same decision
'   through Test_EXCEL_UI_RunReleaseCertification would require a full
'   destructive run to assert one comparison, and could not run from inside the
'   pack at all.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Suppresses screen updating and asserts a matching baseline is clean.
'   - Asserts a mismatched baseline is reported, naming both values.
'   - Restores the entry value on every path.
'
' ERROR POLICY
'   - Raises on assertion failure, after restoration.
'
' UPDATED
'   2026-08-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EntryScreenUpdating As Boolean         'Value to restore on exit
    Dim AnchorWindow        As Window          'Window used as the anchor
    Dim BaselineBooks       As Long            'Workbooks open during the case
    Dim CleanupOK           As Boolean         'Result under test
    Dim CleanupDetail       As String          'Findings returned by the helper

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_CertificationCleanupUsesBaseline"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that cleanup is judged against the entry state"

        Set AnchorWindow = ActiveWindow
        BaselineBooks = Workbooks.Count
        EntryScreenUpdating = Application.ScreenUpdating

'------------------------------------------------------------------------------
' SUPPRESSED STATE MATCHING ITS BASELINE IS CLEAN
'------------------------------------------------------------------------------
    'This is the case the previous implementation failed: screen updating is
    'suppressed, and that is exactly what the baseline says it should be.
        Application.ScreenUpdating = False

        CleanupOK = TST_CertEvaluateCleanup( _
            BaselineBooks:=BaselineBooks, _
            BaselineScreenUpdating:=False, _
            AnchorWindow:=AnchorWindow, _
            CleanupDetail:=CleanupDetail)

        TST_AssertBooleanEquals _
            True, CleanupOK, "CertificationCleanup.SuppressedMatchingBaselineIsClean"

        TST_AssertTrue _
            (Len(CleanupDetail) = 0), _
            "CertificationCleanup.NoDetailWhenClean"

'------------------------------------------------------------------------------
' A GENUINE DIFFERENCE IS STILL REPORTED
'------------------------------------------------------------------------------
    'Same host state, opposite baseline: the run would have changed it, so this
    'must fail. Loosening the check must not have disabled it.
        CleanupOK = TST_CertEvaluateCleanup( _
            BaselineBooks:=BaselineBooks, _
            BaselineScreenUpdating:=True, _
            AnchorWindow:=AnchorWindow, _
            CleanupDetail:=CleanupDetail)

        TST_AssertBooleanEquals _
            False, CleanupOK, "CertificationCleanup.GenuineDifferenceReported"

    'The detail must name both values, or a real leak cannot be diagnosed from
    'the evidence file alone
        TST_AssertTrue _
            (InStr(1, CleanupDetail, "ScreenUpdating changed from") > 0), _
            "CertificationCleanup.DetailNamesBothValues"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Cleanup compared the entry state and still reported a real change"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Restore the value found on entry, whatever happened above
        On Error Resume Next
            Application.ScreenUpdating = EntryScreenUpdating
        On Error GoTo 0

    'Raise the captured failure after restoration when needed
        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        HasFailure = True

        Resume Safe_Exit

End Sub


Private Sub TST_Case_CertUnitIsolatesCaptureFailure()

'
'==============================================================================
' TST_Case_CertUnitIsolatesCaptureFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a certification unit which fails after a capture reports the
'   failure list the capture returned, leaves no snapshot behind, and does not
'   stop the next unit from running.
'
' WHY THIS EXISTS
'   The first real-host certification run reported three failed units and a
'   cleanup failure for one defect. A capture failed, the assertion on it
'   raised "expected=True actual=False" without the failure list that named
'   the element and the reason, and the snapshot that capture had created was
'   left behind. Both later units then refused to run over it.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   1  failing unit      fails; its detail carries the failure count and the
'                        injected TitleBar entry; the snapshot it left behind
'                        is released and reported
'   2  following unit    passes; nothing is left behind
'
' ERROR POLICY
'   - Raises on assertion failure, after disarming the seam and releasing any
'     snapshot this case created.
'
' DEPENDENCIES
'   - TST_CertExecuteUnit
'   - TST_AssertBooleanEquals
'   - TST_AssertTrue
'   - TST_AssertSnapshotAvailability
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   The case drives TST_CertExecuteUnit, not TST_CertRunUnit. Recording would
'   add these two units to the certification run executing this case, and
'   their deliberate failure would become that run's verdict.
'
'   The detail is the text certification writes to its report and its JSON
'   evidence, so asserting it here asserts what reaches the evidence.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Passed              As Boolean         'Unit outcome
    Dim Detail              As String          'Unit failure and release detail
    Dim LeftSnapshot        As Boolean         'TRUE when the unit left one

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_CertUnitIsolatesCaptureFailure"
    Const INJECTED_ENTRY As String = _
        "TitleBar | the active Excel Window is not the active native frame"

'------------------------------------------------------------------------------
' VALIDATE PRECONDITIONS
'------------------------------------------------------------------------------
    'Every snapshot below is this case's only if the slot is empty now. This is
    'asserted before the handler is armed, so a failure here cannot reach the
    'release below and clear a snapshot this case never owned.
        TST_AssertSnapshotAvailability False, PROC & ".NoSnapshotOnEntry"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a unit failing after a capture is reported and isolated"

'------------------------------------------------------------------------------
' A UNIT THAT FAILS AFTER A CAPTURE
'------------------------------------------------------------------------------
        TST_CertExecuteUnit _
            UnitName:="SelfTest.CaptureFailure", _
            Passed:=Passed, _
            Detail:=Detail, _
            LeftSnapshot:=LeftSnapshot

    'Logged before it is asserted, so a failing assertion keeps what was seen
        TST_Log PROC, "OBSERVED", Detail

        TST_AssertBooleanEquals False, Passed, PROC & ".Failing.Failed"

    'The count and the entry that names the element are what the host run lost
        TST_AssertTrue _
            (InStr(1, Detail, "| FailureCount=", vbBinaryCompare) > 0), _
            PROC & ".Failing.DetailCarriesCount"

        TST_AssertTrue _
            (InStr(1, Detail, INJECTED_ENTRY, vbBinaryCompare) > 0), _
            PROC & ".Failing.DetailCarriesFailureList"

    'The leak is released, and said to be, rather than hidden
        TST_AssertBooleanEquals True, LeftSnapshot, PROC & ".Failing.LeakReported"

        TST_AssertTrue _
            (InStr(1, Detail, "released before the next unit", vbBinaryCompare) > 0), _
            PROC & ".Failing.ReleaseRecorded"

        TST_AssertSnapshotAvailability False, PROC & ".Failing.SnapshotReleased"

'------------------------------------------------------------------------------
' THE NEXT UNIT STILL RUNS
'------------------------------------------------------------------------------
        TST_CertExecuteUnit _
            UnitName:="SelfTest.FollowingUnit", _
            Passed:=Passed, _
            Detail:=Detail, _
            LeftSnapshot:=LeftSnapshot

        TST_AssertBooleanEquals True, Passed, PROC & ".Following.Ran"

        TST_AssertTrue (Len(Detail) = 0), PROC & ".Following.NoDetail"

        TST_AssertBooleanEquals False, LeftSnapshot, PROC & ".Following.NothingLeft"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "The failed unit reported its failure list and left nothing for the next unit"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Disarm the seam and release anything left. The slot was empty on entry,
    'so any snapshot here was created by this case.
        On Error Resume Next
            UI_InternalInjectFramePairFault TEST_PAIR_FAULT_NONE
            UI_ClearExcelUIStateSnapshot
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
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        HasFailure = True

        Resume Safe_Exit

End Sub


Private Sub TST_Case_FailureAccumulatorDegradesSafely()

'
'==============================================================================
' TST_Case_FailureAccumulatorDegradesSafely
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
' RETURNS
'   None.
'
' BEHAVIOR
'   - Records one failure normally and asserts it was listed.
'   - Arms a one-shot growth failure and records a second failure.
'   - Asserts the call did not raise, the count still advanced, the list did
'     not grow, and a truncation marker was written into the existing slot.
'   - Repeats the growth failure against an empty buffer and asserts the count
'     still advances with no list and no error.
'
' ERROR POLICY
'   - Raises on assertion failure, after disarming the seam.
'
' NOTES
'   Uses the internal seam UI_InternalInjectFailureListGrowthFailure. It is
'   Public only for same-project regression access; Option Private Module keeps
'   it out of the external automation namespace.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean         'Result contract success flag
    Dim FailureCount        As Long            'Result contract failure count
    Dim FailureList         As Variant         'Result contract failure list
    Dim MarkerText          As String          'Final entry after truncation

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_FailureAccumulatorDegradesSafely"
    Const MARKER_PREFIX As String = "Diagnostics |"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that failure accumulation degrades instead of raising"

    'Start from the documented empty result contract
        Succeeded = True
        UI_RuntimeClearResultBuffer FailureCount, FailureList, True

'------------------------------------------------------------------------------
' RECORD ONE FAILURE NORMALLY
'------------------------------------------------------------------------------
    'A healthy append is the precondition for the truncation case below
        UI_RuntimeHandleFailure _
            PROC, False, Succeeded, FailureCount, FailureList, True, _
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
            PROC, False, Succeeded, FailureCount, FailureList, True, _
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
            (Left$(MarkerText, Len(MARKER_PREFIX)) = MARKER_PREFIX), _
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
            PROC, False, Succeeded, FailureCount, FailureList, True, _
            "StageThree", "third failure detail"

        TST_AssertBooleanEquals _
            False, Succeeded, "FailureAccumulator.EmptyBufferSucceededCleared"

        TST_AssertTrue _
            (FailureCount = 1), "FailureAccumulator.EmptyBufferCountAdvanced"

        TST_AssertTrue _
            Not IsArray(FailureList), "FailureAccumulator.EmptyBufferListUnallocated"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Accumulation degraded visibly and never raised"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
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
Err_Handler:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub


Private Sub TST_Case_TitleBarRoundTrip()

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


Private Sub TST_Case_TitleBarOwnedBitPreservation()

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


Private Sub TST_Case_TitleBarShowRecoversWithoutBaseline()
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


Private Sub TST_Case_RibbonRestoreRefusesOtherActiveWindow()

'
'==============================================================================
' TST_Case_RibbonRestoreRefusesOtherActiveWindow
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that restoring a snapshot while a different window is active
'   refuses the Ribbon, reports it, changes neither window's Ribbon and
'   activates nothing.
'
' WHY THIS EXISTS
'   Under SDI each window has its own Ribbon and every Ribbon write acts on the
'   active window. Before #23 a value captured on window A was written to
'   whichever window was active at restore time and reported as a success,
'   while A stayed unrestored (ICR-UI-112-P1-01).
'
'   The second window is hidden before restore, so a wrongful write of A's
'   captured visible state would be observable on it.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Captures with the anchor active and its Ribbon visible.
'   - Hides the anchor's Ribbon, activates a second window and hides its
'     Ribbon too.
'   - Restores and asserts exactly one Ribbon failure.
'   - Asserts the second window is still active and its Ribbon still hidden.
'   - Asserts the second window's complete frame style is unchanged.
'   - Asserts the anchor's Ribbon is still hidden: nothing was restored to it.
'
' ERROR POLICY
'   - Raises after best-effort cleanup of the temporary workbook.
'
' UPDATED
'   2026-11-06 - Compared the active frame's complete style (#45).
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window          'Window the Ribbon is captured on
    Dim SecondWorkbook      As Workbook        'Temporary second workbook
    Dim SecondWindow        As Window          'Window active at restore time

#If VBA7 Then
    Dim SecondHwnd          As LongPtr         'Frame active at restore time
    Dim SecondFrameBefore   As LongPtr         'Its style before restore
#Else
    Dim SecondHwnd          As Long            'Frame active at restore time
    Dim SecondFrameBefore   As Long            'Its style before restore
#End If

    Dim RibbonVisible       As Boolean         'Ribbon read for assertions
    Dim ReadMsg             As String          'Diagnostic from the Ribbon read

    Dim OK                  As Boolean         'Structured result flag
    Dim FailureCount        As Long            'Structured result failure count
    Dim FailureList         As Variant         'Structured result failure list

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_RibbonRestoreRefusesOtherActiveWindow"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that the Ribbon is not restored to another active window"

'------------------------------------------------------------------------------
' CAPTURE ON THE ANCHOR
'------------------------------------------------------------------------------
    'Hold the window the Ribbon is captured from
        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 110, _
                PROC, _
                "no active Excel window is available"
        End If

    'Open the second window first, so it already exists at capture time and
    'its own window-level state restores cleanly
        Set SecondWorkbook = Workbooks.Add
        Set SecondWindow = SecondWorkbook.Windows(1)

        AnchorWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        UI_SetExcelUI Ribbon:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        OK = UI_CaptureExcelUIState_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertResultSuccess OK, FailureCount, FailureList, _
            "RibbonRestoreOtherActiveWindow.Capture"

'------------------------------------------------------------------------------
' CHANGE BOTH RIBBONS AND SWITCH WINDOWS
'------------------------------------------------------------------------------
    'Hide the anchor's Ribbon, so a correct restore would have to show it
        UI_SetExcelUI Ribbon:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

    'Hide the second window's Ribbon, so a wrongful write would show it
        SecondWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        UI_SetExcelUI Ribbon:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

    'Record the active frame exactly as it stands before restoration
        SecondHwnd = SecondWindow.hWnd
        SecondFrameBefore = TST_FrameStyleOrRaise(SecondHwnd, PROC)

'------------------------------------------------------------------------------
' RESTORE WITH THE SECOND WINDOW ACTIVE
'------------------------------------------------------------------------------
        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_WaitUI TEST_WAIT_SECONDS

    'The Ribbon is the only element that cannot be restored from here
        TST_AssertSingleFailurePrefix _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedPrefix:="Ribbon | captured Ribbon window is not the active window", _
            AssertionName:="RibbonRestoreOtherActiveWindow.Refused"

'------------------------------------------------------------------------------
' ASSERT NOTHING WAS ACTIVATED OR WRITTEN
'------------------------------------------------------------------------------
    'Restore must not have moved focus to the captured window
        TST_AssertTrue _
            ActualValue:=(ActiveWindow.hWnd = SecondWindow.hWnd), _
            AssertionName:="RibbonRestoreOtherActiveWindow.FocusUnchanged"

    'The active window must not have received the anchor's captured value
        If Not TST_TryGetRibbonVisible(RibbonVisible, ReadMsg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 111, _
                PROC, _
                "second window Ribbon read failed | " & ReadMsg
        End If

        TST_AssertBooleanEquals _
            False, _
            RibbonVisible, _
            "RibbonRestoreOtherActiveWindow.SecondUnchanged"

    'Nor any bit of its frame: the title bar belongs to the anchor's frame
        TST_AssertFrameUntouched _
            SecondHwnd, _
            SecondFrameBefore, _
            "RibbonRestoreOtherActiveWindow.SecondFrameUntouched"

    'Nothing was restored to the anchor either
        AnchorWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertRibbonVisible _
            False, _
            "RibbonRestoreOtherActiveWindow.AnchorNotRestored"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Ribbon restore was refused; neither window changed and none was activated"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Release the snapshot and the temporary workbook, and leave the anchor's
    'Ribbon visible
        On Error Resume Next
            UI_ClearExcelUIStateSnapshot
            TST_SafeCloseWorkbook SecondWorkbook

            If Not AnchorWindow Is Nothing Then
                AnchorWindow.Activate
                UI_SetExcelUI Ribbon:=UI_Show
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


Private Sub TST_Case_RibbonRestoreRefusesClosedWindow()

'
'==============================================================================
' TST_Case_RibbonRestoreRefusesClosedWindow
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that restoring a snapshot whose Ribbon window has closed reports a
'   Ribbon failure and does not write the captured value to the surviving
'   window.
'
' WHY THIS EXISTS
'   A closed captured window is the second route to the same misdirection: the
'   only window left to write to is one the value was never read from.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Captures with a temporary window active and its Ribbon visible.
'   - Closes that window and hides the surviving anchor's Ribbon.
'   - Restores and asserts a Ribbon failure is reported.
'   - Asserts the anchor's Ribbon is still hidden.
'   - Asserts the anchor's complete frame style is unchanged.
'
' ERROR POLICY
'   - Raises after best-effort cleanup of the temporary workbook.
'
' NOTES
'   Closing the captured window also invalidates its title-bar frame and its
'   Window entry, so the result carries more than one failure. The assertion
'   looks for the Ribbon entry within the list.
'
' UPDATED
'   2026-11-06 - Compared the surviving active frame's complete style (#45).
'   2026-11-03
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window          'Window that survives the case
    Dim DoomedWorkbook      As Workbook        'Temporary workbook to be closed
    Dim DoomedWindow        As Window          'Window the Ribbon is captured on

#If VBA7 Then
    Dim AnchorHwnd          As LongPtr         'Frame active at restore time
    Dim AnchorFrameBefore   As LongPtr         'Its style before restore
#Else
    Dim AnchorHwnd          As Long            'Frame active at restore time
    Dim AnchorFrameBefore   As Long            'Its style before restore
#End If

    Dim OK                  As Boolean         'Structured result flag
    Dim FailureCount        As Long            'Structured result failure count
    Dim FailureList         As Variant         'Structured result failure list

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_RibbonRestoreRefusesClosedWindow"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a closed Ribbon window is reported"

'------------------------------------------------------------------------------
' CAPTURE ON THE WINDOW THAT WILL CLOSE
'------------------------------------------------------------------------------
    'Hold the window that must remain untouched
        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 120, _
                PROC, _
                "no active Excel window is available"
        End If

        Set DoomedWorkbook = Workbooks.Add
        Set DoomedWindow = DoomedWorkbook.Windows(1)

        DoomedWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        UI_SetExcelUI Ribbon:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        OK = UI_CaptureExcelUIState_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_AssertResultSuccess OK, FailureCount, FailureList, _
            "RibbonRestoreClosedWindow.Capture"

'------------------------------------------------------------------------------
' CLOSE THE CAPTURED WINDOW
'------------------------------------------------------------------------------
        TST_SafeCloseWorkbook DoomedWorkbook

        Set DoomedWindow = Nothing
        TST_WaitUI TEST_WAIT_SECONDS

    'Hide the survivor's Ribbon, so a wrongful write would show it
        AnchorWindow.Activate
        TST_WaitUI TEST_WAIT_SECONDS

        UI_SetExcelUI Ribbon:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

    'Record the surviving active frame exactly as it stands before restoration
        AnchorHwnd = AnchorWindow.hWnd
        AnchorFrameBefore = TST_FrameStyleOrRaise(AnchorHwnd, PROC)

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_WaitUI TEST_WAIT_SECONDS

    'A Ribbon entry must appear in the ordered failure list
        TST_AssertFailureListContainsPrefix _
            Succeeded:=OK, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            ExpectedPrefix:="Ribbon | captured Ribbon window is no longer open or usable", _
            AssertionName:="RibbonRestoreClosedWindow.Reported"

    'The survivor must not have received the closed window's captured value
        TST_AssertRibbonVisible _
            False, _
            "RibbonRestoreClosedWindow.SurvivorUnchanged"

    'The captured title-bar frame closed with its window, so the title bar is
    'refused too, and no bit of the surviving frame may change
        TST_AssertFrameUntouched _
            AnchorHwnd, _
            AnchorFrameBefore, _
            "RibbonRestoreClosedWindow.SurvivorFrameUntouched"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Closed Ribbon window was reported and no state was applied"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Release the snapshot and any surviving temporary workbook, and leave the
    'anchor's Ribbon visible
        On Error Resume Next
            UI_ClearExcelUIStateSnapshot
            TST_SafeCloseWorkbook DoomedWorkbook

            If Not AnchorWindow Is Nothing Then
                AnchorWindow.Activate
                UI_SetExcelUI Ribbon:=UI_Show
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


Private Sub TST_Case_TitleBarFrameRefreshDebtRetried()

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


Private Sub TST_Case_TitleBarStaleFrameEntryNotReused()

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


#If VBA7 Then
Private Function TST_RestoreTitleBarStyle( _
    ByVal hWnd As LongPtr, _
    ByVal EntryStyle As LongPtr, _
    ByVal StyleCaptured As Boolean, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function TST_RestoreTitleBarStyle( _
    ByVal hWnd As Long, _
    ByVal EntryStyle As Long, _
    ByVal StyleCaptured As Boolean, _
    ByRef FailMsg As String) As Boolean
#End If
'
'==============================================================================
' TST_RestoreTitleBarStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Write a previously captured window style back, refresh the non-client
'   frame and read the owned bits back, reporting the first step that fails.
'
' WHY THIS EXISTS
'   A case that hides a title bar must return the host to its entry state on
'   the success path and on every failure path, and must do so without the
'   API under test: a defect there is exactly what the case may be failing
'   on. One harness-only routine keeps those paths identical.
'
'   A restore that cannot be confirmed is a failure, not a best effort. A
'   caller that ignored it could report a pass over a host left with a hidden
'   title bar, so every caller checks the result and makes a FALSE part of
'   its own failure (#66).
'
' INPUTS
'   hWnd
'     Top-level window the style was captured from.
'
'   EntryStyle
'     Style value captured before the caller modified the frame.
'
'   StyleCaptured
'     FALSE when the entry style was never read. The caller has then changed
'     nothing through this entry, so nothing is written and the result is
'     TRUE.
'
'   FailMsg
'     ByRef. Receives the step that failed and its diagnostic. Empty on
'     success.
'
' RETURNS
'   TRUE  => nothing was captured, or the style was written, the frame was
'            refreshed and the owned bits read back equal to the entry
'            style's.
'   FALSE => a step failed; FailMsg names it.
'
' BEHAVIOR
'   - Writes the captured style through TST_TrySetWindowStyle.
'   - Refreshes the frame through TST_TryRefreshWindowFrame.
'   - Reads the style back through TST_TryGetWindowStyle and compares the
'     owned bits, which carry the title bar. Unrelated bits are not compared:
'     Windows may change them between capture and restore.
'
' ERROR POLICY
'   - Never raises. It runs on cleanup paths, inside or after an error
'     handler.
'   - Each native call is made by a harness helper that captures
'     Err.LastDllError immediately after the call (#27). This routine reads
'     no Err field after a call; an unexpected runtime error is reported
'     through FailMsg from its own handler.
'
' DEPENDENCIES
'   - TST_TrySetWindowStyle
'   - TST_TryRefreshWindowFrame
'   - TST_TryGetWindowStyle
'   - TST_BuildRuntimeErrorText
'
' CALLED FROM
'   - TST_RunRegressionPack
'   - TST_RunTitleBarOnlyPack
'   - Test_EXCEL_UI_RunTitleBarSdiIdentity
'   - TST_Case_TitleBarRoundTrip
'   - TST_Case_TitleBarShowRecoversWithoutBaseline
'   - TST_Case_TitleBarShowRejectsCaptionlessBaseline
'   - TST_Case_TitleBarSameStyleHandleReuse
'   - TST_Case_TitleBarFrameRefreshDebtRetried
'   - TST_Case_TitleBarStaleFrameEntryNotReused
'   - TST_Case_TitleBarSdiRestoreTargetsCapturedFrame
'   - TST_Case_TitleBarPairRefusalLeavesFrameUntouched
'   - TST_Case_TitleBarClosedFrameReopenedIsRefused
'   - TST_Case_TitleBarApiErrorCapture
'
' NOTES
'   This routine cannot recover a host whose VBA execution has stopped: a
'   VBE Reset or End, a terminated Excel process, or a native call that never
'   returns. CONTRIBUTING.md documents the recovery for those.
'
' UPDATED
'   2026-11-06 - Became a Function that reports a failed write, refresh or
'                readback instead of discarding it (#66).
'   2026-11-06 - Recorded the refusal cases as callers (#45).
'   2026-11-06 - Recorded the error-capture case as a caller (#27).
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim ResultStyle         As LongPtr         'Style read back after the write
#Else
    Dim ResultStyle         As Long            'Style read back after the write
#End If

    Dim StepMsg             As String          'Diagnostic from a harness helper

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'This runs on cleanup paths and must never raise
        On Error GoTo Err_Handler

        TST_RestoreTitleBarStyle = False
        FailMsg = vbNullString

    'Nothing was captured, so nothing was changed through this entry
        If Not StyleCaptured Then
            TST_RestoreTitleBarStyle = True
            GoTo Safe_Exit
        End If

        If hWnd = 0 Then
            FailMsg = "entry style not restored: invalid window handle"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' RESTORE STYLE
'------------------------------------------------------------------------------
        If Not TST_TrySetWindowStyle(hWnd, EntryStyle, StepMsg) Then
            FailMsg = "entry style write failed | " & StepMsg
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' REFRESH FRAME
'------------------------------------------------------------------------------
        If Not TST_TryRefreshWindowFrame(hWnd, StepMsg) Then
            FailMsg = "entry style written; frame refresh failed | " & StepMsg
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' VERIFY
'------------------------------------------------------------------------------
    'A write and a refresh that both succeeded still do not prove the frame
    'carries the entry state. Read it back.
        If Not TST_TryGetWindowStyle(hWnd, ResultStyle, StepMsg) Then
            FailMsg = _
                "entry style written and refreshed; readback failed | " & _
                StepMsg
            GoTo Safe_Exit
        End If

        If (ResultStyle And TST_TITLEBAR_OWNED_MASK) <> _
            (EntryStyle And TST_TITLEBAR_OWNED_MASK) Then

            FailMsg = _
                "entry style readback differs; hWnd=" & CStr(hWnd) & _
                "; owned expected=" & _
                CStr(EntryStyle And TST_TITLEBAR_OWNED_MASK) & _
                " actual=" & CStr(ResultStyle And TST_TITLEBAR_OWNED_MASK)
            GoTo Safe_Exit
        End If

        TST_RestoreTitleBarStyle = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'The builder reads Err before anything else can change it
        StepMsg = TST_BuildRuntimeErrorText
        FailMsg = "entry style restoration raised | " & StepMsg

        Resume Safe_Exit

End Function


Private Sub TST_RecordCleanupFailure( _
    ByVal CleanupOK As Boolean, _
    ByVal CleanupMsg As String, _
    ByVal CallerProc As String, _
    ByRef FailNumber As Long, _
    ByRef FailSource As String, _
    ByRef FailDescription As String)
'
'==============================================================================
' TST_RecordCleanupFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Make a failed entry-style restoration part of the caller's result: the
'   failure the caller raises when it had none, or an addition to the one it
'   already holds.
'
' WHY THIS EXISTS
'   A case that cannot put the frame back has not finished, however its
'   assertions went. Logging the miss and returning normally would let the
'   pack report a pass over a host left with a hidden title bar (#66).
'
' INPUTS
'   CleanupOK
'     Result of TST_RestoreTitleBarStyle.
'
'   CleanupMsg
'     Its FailMsg.
'
'   CallerProc
'     Calling procedure, used for the log line and the failure source.
'
'   FailNumber / FailSource / FailDescription
'     ByRef. The caller's saved failure; zero when it has none.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Does nothing when CleanupOK is TRUE.
'   - Logs the failed restoration.
'   - With no saved failure, sets TEST_ERR_BASE + 81, the source
'     CallerProc & ".EntryStyleRestored" and the restoration text.
'   - With a saved failure, keeps its number and source and appends the
'     restoration text to its description, so the original failure is still
'     the one reported first.
'
' ERROR POLICY
'   - Does not raise and reads no Err field.
'
' DEPENDENCIES
'   - TST_Log
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'This runs on cleanup paths and must never raise
        On Error Resume Next

        If CleanupOK Then
            Exit Sub
        End If

        If Len(CleanupMsg) = 0 Then
            CleanupMsg = "restoration did not complete"
        End If

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
        TST_Log CallerProc, "FAIL", _
            "entry style restoration failed | " & CleanupMsg

        If FailNumber = 0 Then
            FailNumber = TEST_ERR_BASE + 81
            FailSource = CallerProc & ".EntryStyleRestored"
            FailDescription = "entry style restoration failed | " & CleanupMsg
        Else
            FailDescription = FailDescription & _
                " | entry style restoration failed | " & CleanupMsg
        End If

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
' TST_SnapshotState
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long            'Current window index during snapshot
    Dim Msg                 As String          'Diagnostic message from reader helpers

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
' TST_RestoreState
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long            'Current window index during restore
    Dim WindowLimit         As Long            'Minimum of saved and current window counts
    Dim Msg                 As String          'Diagnostic message from helper routines

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
' TST_WaitUI
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
'   - A fixed delay, not a wait for a condition. Reaching the end is never a
'     pass: the assertion that follows reads the real state and fails when
'     the change has not happened.
'   - Bounded. A request is capped at TEST_WAIT_MAX_SECONDS, and the loop
'     ends once Timer has moved that far from its baseline.
'     TST_TimerElapsedSeconds never returns a negative value, so a midnight
'     rollover or a clock change shortens the wait rather than extending it.
'   - DoEvents lets Excel run queued work. A native call or a modal dialog
'     that never returns cannot be bounded from VBA; CONTRIBUTING.md
'     documents the recovery for that case.
'
' UPDATED
'   2026-11-06 - Capped the wait at TEST_WAIT_MAX_SECONDS (#66).
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim t0                  As Double          'Timer baseline
    Dim WaitSeconds         As Double          'Requested wait, capped

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Exit immediately when no positive wait duration was requested
        If SecondsToWait <= 0# Then
            Exit Sub
        End If

    'No caller can hold the harness longer than the bound
        WaitSeconds = SecondsToWait

        If WaitSeconds > TEST_WAIT_MAX_SECONDS Then
            WaitSeconds = TEST_WAIT_MAX_SECONDS
        End If

    'Capture the timer baseline
        t0 = Timer

'------------------------------------------------------------------------------
' WAIT LOOP
'------------------------------------------------------------------------------
    'Yield to Excel until the capped duration has elapsed, handling midnight
    'rollover safely
        Do While TST_TimerElapsedSeconds(t0) < WaitSeconds
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
' TST_AssertBooleanEquals
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
'   2026-07-25
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
' TST_AssertApplicationProperty
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Actual              As Boolean         'Actual property value
    Dim Msg                 As String          'Diagnostic message from the reader helper

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
' TST_AssertAllWindowsProperty
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim W                   As Window          'Current Excel window during assertion
    Dim Actual              As Boolean         'Actual property value
    Dim Msg                 As String          'Diagnostic message from the reader helper

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
' TST_AssertRibbonVisible
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Actual              As Boolean         'Actual Ribbon visibility
    Dim Msg                 As String          'Diagnostic message from the reader helper

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
' TST_AssertTitleBarVisible
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Actual              As Boolean         'Actual title-bar visibility
    Dim Msg                 As String          'Diagnostic message from the reader helper

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
' TST_AssertResultSuccess
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
'   - Raises on mismatch, with the failure count and list the call returned
'
' NOTES
'   The failure list is the only record of which element failed and why.
'   Certification writes this message to its report and its JSON evidence,
'   so a message without the list leaves a failed run undiagnosable.
'
' UPDATED
'   2026-11-06 - Messages carry the failure count and list (#135).
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' ASSERT SUCCESS FLAG
'------------------------------------------------------------------------------
    'Assert that the Boolean result reports overall success
        If Not Succeeded Then
            Err.Raise TEST_ERR_BASE + 20, _
                      AssertionName, _
                      AssertionName & " expected=True actual=False | " & _
                      TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

'------------------------------------------------------------------------------
' ASSERT FAILURE COUNT
'------------------------------------------------------------------------------
    'Assert that the result buffers recorded no failures
        If FailureCount <> 0 Then
            Err.Raise TEST_ERR_BASE + 21, _
                      AssertionName, _
                      AssertionName & " expected FailureCount=0 actual=" & CStr(FailureCount) & _
                      " | " & TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

'------------------------------------------------------------------------------
' ASSERT FAILURE LIST STATE
'------------------------------------------------------------------------------
    'Assert that the captured failure list remains Empty for a clean success
    'path
        If Not IsEmpty(FailureList) Then
            Err.Raise TEST_ERR_BASE + 22, _
                      AssertionName, _
                      AssertionName & " expected FailureList=Empty for clean success path | " & _
                      TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

End Sub

Private Sub TST_AssertSingleFailurePrefix( _
    ByVal Succeeded As Boolean, _
    ByVal FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal ExpectedPrefix As String, _
    ByVal AssertionName As String)

'
'==============================================================================
' TST_AssertSingleFailurePrefix
'------------------------------------------------------------------------------
' PURPOSE
'   Assert a structured-result failure containing exactly one ordered entry
'   whose text begins with the expected stage prefix.
'
' INPUTS
'   Succeeded
'     Boolean result returned by the structured API.
'
'   FailureCount / FailureList
'     Structured-result outputs to validate.
'
'   ExpectedPrefix
'     Required leading text for the single failure entry.
'
'   AssertionName
'     Diagnostic source used when an assertion fails.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Raises on mismatch, with the failure count and list the call returned.
'
' UPDATED
'   2026-11-06 - Messages carry the failure count and list (#135).
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT FAILURE FLAG AND COUNT
'------------------------------------------------------------------------------
        If Succeeded Then
            Err.Raise TEST_ERR_BASE + 60, _
                      AssertionName, _
                      AssertionName & " expected=False actual=True | " & _
                      TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

        If FailureCount <> 1 Then
            Err.Raise TEST_ERR_BASE + 61, _
                      AssertionName, _
                      AssertionName & " expected FailureCount=1 actual=" & _
                      CStr(FailureCount) & " | " & _
                      TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

'------------------------------------------------------------------------------
' ASSERT FAILURE LIST
'------------------------------------------------------------------------------
        If Not IsArray(FailureList) Then
            Err.Raise TEST_ERR_BASE + 62, _
                      AssertionName, _
                      AssertionName & " expected FailureList array | " & _
                      TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

        If LBound(FailureList) <> 1 Or UBound(FailureList) <> 1 Then
            Err.Raise TEST_ERR_BASE + 63, _
                      AssertionName, _
                      AssertionName & " expected one 1-based failure entry | " & _
                      TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

        If Left$(CStr(FailureList(1)), Len(ExpectedPrefix)) <> ExpectedPrefix Then
            Err.Raise TEST_ERR_BASE + 64, _
                      AssertionName, _
                      AssertionName & " expected prefix='" & ExpectedPrefix & _
                      "' actual='" & CStr(FailureList(1)) & "'"
        End If

End Sub


Private Function TST_DescribeResultBuffers( _
    ByVal FailureCount As Long, _
    ByRef FailureList As Variant) _
    As String

'
'==============================================================================
' TST_DescribeResultBuffers
'------------------------------------------------------------------------------
' PURPOSE
'   Render a structured result's failure count and failure list as one bounded
'   line of diagnostic text.
'
' WHY THIS EXISTS
'   An assertion on a structured result used to report only that the flag or
'   the count was wrong. The list, which names each failed element and the
'   reason, was discarded, so a failed certification could not be diagnosed
'   from its evidence.
'
' INPUTS
'   FailureCount
'     Failure count the call returned.
'
'   FailureList
'     Failure list the call returned. May be Empty, a non-array, an array
'     without bounds, or hold entries that are not text.
'
' RETURNS
'   String
'     "FailureCount=<n> | FailureList=<list>", where <list> is Empty, a
'     description of a value that is not a usable array, or the entries in
'     brackets joined by " || ". Entries are cut off after
'     TST_RESULT_LIST_MAX_CHARS characters and the cut is marked.
'
' ERROR POLICY
'   - Does not raise. It runs while an assertion is being raised, and must not
'     replace the failure it describes.
'
' DEPENDENCIES
'   - TST_DescribeListEntry
'
' CALLED FROM
'   - TST_AssertResultSuccess
'   - TST_AssertSingleFailurePrefix
'   - TST_AssertFailureListContainsPrefix
'   - TST_Case_WithResult_InvalidVisibility
'
' NOTES
'   Entries are joined with " || " because each entry already uses " | "
'   between its element and its reason.
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ListText            As String          'Rendered failure list
    Dim EntryText           As String          'One rendered entry
    Dim FirstIdx            As Long            'Lower bound of the list
    Dim LastIdx             As Long            'Upper bound of the list
    Dim ScanIdx             As Long            'Cursor over the entries

    Const TST_RESULT_LIST_MAX_CHARS As Long = 1000  'Cap on rendered entries

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A diagnostic must never replace the failure it describes
        On Error Resume Next

        TST_DescribeResultBuffers = _
            "FailureCount=" & CStr(FailureCount) & " | FailureList=<unreadable>"

'------------------------------------------------------------------------------
' RENDER THE LIST
'------------------------------------------------------------------------------
        If IsEmpty(FailureList) Then

            ListText = "Empty"

        ElseIf Not IsArray(FailureList) Then

            ListText = "<" & TypeName(FailureList) & ", not an array>"

        Else

            'An array that was declared but never sized has no bounds
                Err.Clear

                FirstIdx = LBound(FailureList)
                LastIdx = UBound(FailureList)

                If Err.Number <> 0 Then

                    Err.Clear
                    ListText = "<array without bounds>"

                Else

                    For ScanIdx = FirstIdx To LastIdx

                        'An entry that cannot be read keeps this placeholder
                            EntryText = "<unreadable entry>"
                            EntryText = TST_DescribeListEntry(FailureList(ScanIdx))

                            If ScanIdx > FirstIdx Then
                                ListText = ListText & " || "
                            End If

                            ListText = ListText & EntryText

                        'Bound the text, so a long list cannot swamp the report
                            If Len(ListText) > TST_RESULT_LIST_MAX_CHARS Then
                                ListText = Left$(ListText, TST_RESULT_LIST_MAX_CHARS) & _
                                    " ... (truncated)"
                                Exit For
                            End If

                    Next ScanIdx

                    ListText = "[" & ListText & "]"

                End If

        End If

'------------------------------------------------------------------------------
' RETURN TEXT
'------------------------------------------------------------------------------
        TST_DescribeResultBuffers = _
            "FailureCount=" & CStr(FailureCount) & " | FailureList=" & ListText

End Function


Private Function TST_DescribeListEntry( _
    ByVal Entry As Variant) _
    As String

'
'==============================================================================
' TST_DescribeListEntry
'------------------------------------------------------------------------------
' PURPOSE
'   Render one failure-list entry as text, whatever it holds.
'
' WHY THIS EXISTS
'   CStr raises on an object without a default member, on Null and on an
'   array. The entry is taken ByVal as a Variant, so an object arrives as a
'   reference and is never evaluated through its default member.
'
' INPUTS
'   Entry
'     One element of a failure list.
'
' RETURNS
'   String
'     The entry as text, or a bracketed description when it is not text.
'
' ERROR POLICY
'   - Does not raise. An entry that cannot be converted is reported as
'     unreadable.
'
' CALLED FROM
'   - TST_DescribeResultBuffers
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RENDER ENTRY
'------------------------------------------------------------------------------
    'The tests below cannot raise, so no branch is entered on an error
        On Error Resume Next

        TST_DescribeListEntry = "<unreadable entry>"

        If IsObject(Entry) Then
            TST_DescribeListEntry = "<" & TypeName(Entry) & ">"
        ElseIf IsArray(Entry) Then
            TST_DescribeListEntry = "<array>"
        ElseIf IsNull(Entry) Then
            TST_DescribeListEntry = "<Null>"
        Else
            TST_DescribeListEntry = CStr(Entry)
        End If

End Function


Private Sub TST_AssertSnapshotAvailability( _
    ByVal Expected As Boolean, _
    ByVal AssertionName As String)

'
'==============================================================================
' TST_AssertSnapshotAvailability
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
'   2026-07-25
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
' TST_TryGetBooleanProperty
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant         'Late-bound property value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

    'Initialize outputs and default result
        TST_TryGetBooleanProperty = False
        ValueOut = False
        FailMsg = vbNullString

    'Reject invalid object input deterministically
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo Safe_Exit
        End If

    'Reject empty property name deterministically
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo Safe_Exit
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
        TST_TryGetBooleanProperty = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailMsg = TST_BuildRuntimeErrorText

End Function

Private Function TST_TrySetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByVal NewValue As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
' TST_TrySetBooleanProperty
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

    'Initialize default failure result
        TST_TrySetBooleanProperty = False
        FailMsg = vbNullString

    'Reject invalid object input deterministically
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo Safe_Exit
        End If

    'Reject empty property name deterministically
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' APPLY PROPERTY WRITE
'------------------------------------------------------------------------------
    'Assign the requested Boolean value using late-bound property assignment
        CallByName Target, PropertyName, VbLet, NewValue

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
        TST_TrySetBooleanProperty = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailMsg = TST_BuildRuntimeErrorText

End Function

Private Function TST_TryGetRibbonVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
' TST_TryGetRibbonVisible
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant         'Fallback Excel4 macro result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

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
            On Error GoTo Err_Handler
            TST_TryGetRibbonVisible = True
            GoTo Safe_Exit
        End If
        Err.Clear
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' TRY EXCEL4 MACRO FALLBACK
'------------------------------------------------------------------------------
    'Attempt a fallback read using an Excel4 macro
        On Error Resume Next
            V = Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")
        If Err.Number = 0 Then
            On Error GoTo Err_Handler
            IsVisible = CBool(V)
            TST_TryGetRibbonVisible = True
            GoTo Safe_Exit
        End If
        FailMsg = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailMsg = TST_BuildRuntimeErrorText

End Function

Private Function TST_TrySetRibbonVisible( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
' TST_TrySetRibbonVisible
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MacroText           As String          'Excel4 macro text for Ribbon visibility

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

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
        Application.ExecuteExcel4Macro MacroText

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
        TST_TrySetRibbonVisible = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailMsg = TST_BuildRuntimeErrorText

End Function

#If VBA7 Then
Private Function TST_TryGetWindowStyle( _
    ByVal hWnd As LongPtr, _
    ByRef StyleOut As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function TST_TryGetWindowStyle( _
    ByVal hWnd As Long, _
    ByRef StyleOut As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
' TST_TryGetWindowStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Read GWL_STYLE through the correct Win32 API for the current Office bitness.
'
' RETURNS
'   TRUE on success.
'
' ERROR POLICY
'   Clears the thread last-error immediately before the call and copies the
'   result and Err.LastDllError to locals immediately after it, with nothing
'   in between. A zero result is a failure only when the captured error is
'   non-zero.
'
' UPDATED
'   2026-11-06 - Captured Err.LastDllError in place of GetLastError, and
'                named the API, operation, error and hWnd on failure (#27).
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim ApiResult           As LongPtr         'Value the read API returned
#Else
    Dim ApiResult           As Long            'Value the read API returned
#End If

    Dim DllError            As Long            'Err.LastDllError after the read

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_TryGetWindowStyle = False
        StyleOut = 0
        FailMsg = vbNullString

        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Clear, call and capture, with nothing in between
        TST_SetLastError 0
#If VBA7 Then
    #If Win64 Then
        ApiResult = TST_GetWindowLongPtr(hWnd, TST_GWL_STYLE)
    #Else
        ApiResult = TST_GetWindowLong(hWnd, TST_GWL_STYLE)
    #End If
#Else
        ApiResult = TST_GetWindowLong(hWnd, TST_GWL_STYLE)
#End If
        DllError = Err.LastDllError

        If ApiResult = 0 And DllError <> 0 Then
            FailMsg = TST_BuildApiFailureText( _
                TST_API_GET_WINDOW_STYLE, "read GWL_STYLE", DllError, hWnd)

            GoTo Safe_Exit
        End If

        StyleOut = ApiResult
        TST_TryGetWindowStyle = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailMsg = TST_BuildRuntimeErrorText
        Resume Safe_Exit

End Function


#If VBA7 Then
Private Function TST_TrySetWindowStyle( _
    ByVal hWnd As LongPtr, _
    ByVal NewStyle As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function TST_TrySetWindowStyle( _
    ByVal hWnd As Long, _
    ByVal NewStyle As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
' TST_TrySetWindowStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Write GWL_STYLE through the correct Win32 API for the current Office
'   bitness.
'
' RETURNS
'   TRUE on success.
'
' ERROR POLICY
'   Clears the thread last-error immediately before the call and copies the
'   result and Err.LastDllError to locals immediately after it, with nothing
'   in between. A zero previous style is a failure only when the captured
'   error is non-zero.
'
' UPDATED
'   2026-11-06 - Captured Err.LastDllError in place of GetLastError, and
'                named the API, operation, error and hWnd on failure (#27).
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim ApiResult           As LongPtr         'Style replaced by the write
#Else
    Dim ApiResult           As Long            'Style replaced by the write
#End If

    Dim DllError            As Long            'Err.LastDllError after the write

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_TrySetWindowStyle = False
        FailMsg = vbNullString

        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Clear, call and capture, with nothing in between
        TST_SetLastError 0
#If VBA7 Then
    #If Win64 Then
        ApiResult = TST_SetWindowLongPtr(hWnd, TST_GWL_STYLE, NewStyle)
    #Else
        ApiResult = TST_SetWindowLong(hWnd, TST_GWL_STYLE, NewStyle)
    #End If
#Else
        ApiResult = TST_SetWindowLong(hWnd, TST_GWL_STYLE, NewStyle)
#End If
        DllError = Err.LastDllError

        If ApiResult = 0 And DllError <> 0 Then
            FailMsg = TST_BuildApiFailureText( _
                TST_API_SET_WINDOW_STYLE, "write GWL_STYLE", DllError, hWnd)

            GoTo Safe_Exit
        End If

        TST_TrySetWindowStyle = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailMsg = TST_BuildRuntimeErrorText
        Resume Safe_Exit

End Function


#If VBA7 Then
Private Function TST_TryRefreshWindowFrame( _
    ByVal hWnd As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function TST_TryRefreshWindowFrame( _
    ByVal hWnd As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
' TST_TryRefreshWindowFrame
'------------------------------------------------------------------------------
' PURPOSE
'   Recalculate the non-client frame after an exact style restore.
'
' RETURNS
'   TRUE on success.
'
' ERROR POLICY
'   Returns FALSE and FailMsg on failure. Err.LastDllError is captured
'   immediately after the call, so the text reports this call's error.
'
' UPDATED
'   2026-11-06 - Captured Err.LastDllError in place of GetLastError, and
'                named the API, operation, error and hWnd on failure (#27).
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ApiOK               As Long            'SetWindowPos BOOL result
    Dim DllError            As Long            'Err.LastDllError after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_TryRefreshWindowFrame = False
        FailMsg = vbNullString

        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' REFRESH
'------------------------------------------------------------------------------
    'Clear, call and capture, with nothing in between
        TST_SetLastError 0
        ApiOK = TST_SetWindowPos( _
            hWnd, _
            0, _
            0, _
            0, _
            0, _
            0, _
            TST_SWP_NOMOVE Or TST_SWP_NOSIZE Or TST_SWP_NOZORDER Or _
                TST_SWP_NOOWNERZORDER Or TST_SWP_FRAMECHANGED)
        DllError = Err.LastDllError

        If ApiOK = 0 Then
            FailMsg = TST_BuildApiFailureText( _
                "SetWindowPos", "refresh the non-client frame", DllError, hWnd)

            GoTo Safe_Exit
        End If

        TST_TryRefreshWindowFrame = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailMsg = TST_BuildRuntimeErrorText
        Resume Safe_Exit

End Function

Private Function TST_TryGetTitleBarVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
' TST_TryGetTitleBarVisible
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
' DEPENDENCIES
'   - TST_TryGetWindowStyle
'
' UPDATED
'   2026-11-06 - Reads the style through TST_TryGetWindowStyle, so the
'                harness has one GWL_STYLE read and one capture (#27).
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim xlHnd               As LongPtr         'Excel window handle from Application.Hwnd
    Dim StyleValue          As LongPtr         'Current window style value
#Else
    Dim xlHnd               As Long            'Excel window handle from Application.Hwnd
    Dim StyleValue          As Long            'Current window style value
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

    'Initialize outputs and default result
        TST_TryGetTitleBarVisible = False
        IsVisible = False
        FailMsg = vbNullString

    'Read the Excel window handle
        xlHnd = Application.hWnd

    'Reject invalid window handle deterministically
        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' READ WINDOW STYLE
'------------------------------------------------------------------------------
    'Read through the single harness read, which owns the error capture
        If Not TST_TryGetWindowStyle(xlHnd, StyleValue, FailMsg) Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' MAP STYLE TO TITLE-BAR VISIBILITY
'------------------------------------------------------------------------------
    'Treat the caption style bit as the title-bar visibility signal
        IsVisible = ((StyleValue And TST_WS_CAPTION) <> 0)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
        TST_TryGetTitleBarVisible = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailMsg = TST_BuildRuntimeErrorText

End Function

Private Sub TST_TryRestoreWindowProp( _
    ByVal W As Window, _
    ByVal PropName As String, _
    ByVal Value As Boolean)

'
'==============================================================================
' TST_TryRestoreWindowProp
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Msg                 As String          'Diagnostic message from the property-write helper

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

#If VBA7 Then
Private Function TST_TitleBarVisibleForHwndOrRaise( _
    ByVal TargetHwnd As LongPtr, _
    ByVal CallerProc As String) As Boolean
#Else
Private Function TST_TitleBarVisibleForHwndOrRaise( _
    ByVal TargetHwnd As Long, _
    ByVal CallerProc As String) As Boolean
#End If

'
'==============================================================================
' TST_TitleBarVisibleForHwndOrRaise
'------------------------------------------------------------------------------
' PURPOSE
'   Read title-bar visibility for one explicitly supplied top-level window, and
'   raise when it cannot be read.
'
' WHY THIS EXISTS
'   TST_TryGetTitleBarVisible reads through Application.Hwnd, which under the
'   Single Document Interface names whichever window is active. A multi-window
'   case must be able to inspect a specific frame while a different one is
'   active, or it cannot tell the two apart, which is the entire point of the
'   cases that call this.
'
'   It raises rather than returning a flag because an unreadable frame makes the
'   surrounding assertion meaningless; reporting False would silently weaken the
'   case into one that could pass against a defective build.
'
' INPUTS
'   TargetHwnd
'     Top-level window to read.
'
'   CallerProc
'     Calling case name, used as the error source.
'
' RETURNS
'   Boolean
'     TRUE when WS_CAPTION is set on the supplied window.
'
' ERROR POLICY
'   - Raises when the style cannot be read.
'
' DEPENDENCIES
'   - TST_TryGetWindowStyle
'
' CALLED FROM
'   - TST_Case_TitleBarSdiRestoreTargetsCapturedFrame
'   - TST_Case_TitleBarCapturedFrameClosedIsReported
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim StyleValue          As LongPtr         'Live GWL_STYLE value
#Else
    Dim StyleValue          As Long            'Live GWL_STYLE value
#End If

    Dim FailMsg             As String          'Diagnostic returned by the read

'------------------------------------------------------------------------------
' READ STYLE
'------------------------------------------------------------------------------
    'Read the live style for the supplied window, never the active one
        If Not TST_TryGetWindowStyle(TargetHwnd, StyleValue, FailMsg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 40, _
                CallerProc, _
                "unable to read the window style for the supplied " & _
                "frame: " & FailMsg
        End If

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Visibility is carried by the caption bit alone
        TST_TitleBarVisibleForHwndOrRaise = _
            ((StyleValue And TEST_WS_CAPTION) <> 0)

End Function


#If VBA7 Then
Private Function TST_FrameStyleOrRaise( _
    ByVal TargetHwnd As LongPtr, _
    ByVal CallerProc As String) As LongPtr
#Else
Private Function TST_FrameStyleOrRaise( _
    ByVal TargetHwnd As Long, _
    ByVal CallerProc As String) As Long
#End If

'
'==============================================================================
' TST_FrameStyleOrRaise
'------------------------------------------------------------------------------
' PURPOSE
'   Read the complete GWL_STYLE of one explicitly supplied top-level window,
'   and raise when it cannot be read.
'
' WHY THIS EXISTS
'   A refusal must leave the frame exactly as it found it. The caption bit
'   alone would miss a write that changed another owned bit, or an unrelated
'   one, so the refusal cases record the whole style here before the refused
'   call and compare it afterwards with TST_AssertFrameUntouched (#45).
'
' INPUTS
'   TargetHwnd
'     Top-level window to read.
'
'   CallerProc
'     Calling case name, used as the error source.
'
' RETURNS
'   The live GWL_STYLE value.
'
' ERROR POLICY
'   - Raises when the style cannot be read. An unreadable frame would make the
'     surrounding comparison meaningless.
'
' DEPENDENCIES
'   - TST_TryGetWindowStyle
'
' CALLED FROM
'   - TST_AssertFrameUntouched
'   - The refusal cases that record a frame before a refused call
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim StyleValue          As LongPtr         'Live GWL_STYLE value
#Else
    Dim StyleValue          As Long            'Live GWL_STYLE value
#End If

    Dim FailMsg             As String          'Diagnostic returned by the read

'------------------------------------------------------------------------------
' READ STYLE
'------------------------------------------------------------------------------
    'Read the live style for the supplied window, never the active one
        If Not TST_TryGetWindowStyle(TargetHwnd, StyleValue, FailMsg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 41, _
                CallerProc, _
                "unable to read the complete window style for the supplied " & _
                "frame: " & FailMsg
        End If

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
        TST_FrameStyleOrRaise = StyleValue

End Function


#If VBA7 Then
Private Sub TST_AssertFrameUntouched( _
    ByVal TargetHwnd As LongPtr, _
    ByVal StyleBefore As LongPtr, _
    ByVal AssertionName As String)
#Else
Private Sub TST_AssertFrameUntouched( _
    ByVal TargetHwnd As Long, _
    ByVal StyleBefore As Long, _
    ByVal AssertionName As String)
#End If

'
'==============================================================================
' TST_AssertFrameUntouched
'------------------------------------------------------------------------------
' PURPOSE
'   Assert that a frame's complete GWL_STYLE still equals the value recorded
'   before a refused call.
'
' WHY THIS EXISTS
'   A refusal is correct only if nothing was written. Comparing visibility
'   alone passes against a write that changed an owned bit other than the
'   caption, or that rewrote the caption to the value it already had while
'   changing something else. The whole value is compared, so every owned bit,
'   WS_CAPTION and therefore title-bar visibility, and every unrelated bit
'   must be exactly as recorded (#45).
'
' INPUTS
'   TargetHwnd
'     Top-level window to read.
'
'   StyleBefore
'     Complete style recorded with TST_FrameStyleOrRaise before the call.
'
'   AssertionName
'     Diagnostic source used when the assertion fails.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Raises when the style cannot be read or differs from StyleBefore. The
'     description names the window, both styles, both owned-bit sets and both
'     caption states.
'
' DEPENDENCIES
'   - TST_FrameStyleOrRaise
'
' UPDATED
'   2026-11-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim StyleAfter          As LongPtr         'Live GWL_STYLE value now
#Else
    Dim StyleAfter          As Long            'Live GWL_STYLE value now
#End If

'------------------------------------------------------------------------------
' READ STYLE
'------------------------------------------------------------------------------
        StyleAfter = TST_FrameStyleOrRaise(TargetHwnd, AssertionName)

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
        If StyleAfter <> StyleBefore Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 42, _
                AssertionName, _
                AssertionName & " frame changed by a refused call; hWnd=" & _
                CStr(TargetHwnd) & _
                "; style before=" & CStr(StyleBefore) & _
                " after=" & CStr(StyleAfter) & _
                "; owned before=" & _
                CStr(StyleBefore And TST_TITLEBAR_OWNED_MASK) & _
                " after=" & CStr(StyleAfter And TST_TITLEBAR_OWNED_MASK) & _
                "; caption before=" & _
                CStr((StyleBefore And TST_WS_CAPTION) <> 0) & _
                " after=" & CStr((StyleAfter And TST_WS_CAPTION) <> 0)
        End If

End Sub


Private Sub TST_AssertFailureListContainsPrefix( _
    ByVal Succeeded As Boolean, _
    ByVal FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal ExpectedPrefix As String, _
    ByVal AssertionName As String)

'
'==============================================================================
' TST_AssertFailureListContainsPrefix
'------------------------------------------------------------------------------
' PURPOSE
'   Assert a structured result reported failure and that at least one ordered
'   entry begins with the expected stage prefix.
'
' WHY THIS EXISTS
'   TST_AssertSingleFailurePrefix requires the entry to be the only one. Closing
'   a captured window legitimately produces more than one failure, because the
'   window identity and the title-bar frame are both lost by the same act.
'   Requiring exactly one entry there would assert a contract the component does
'   not make.
'
' INPUTS
'   Succeeded
'     Structured result flag; must be FALSE.
'
'   FailureCount
'     Structured result failure count; must be at least one.
'
'   FailureList
'     Ordered failure entries.
'
'   ExpectedPrefix
'     Stage prefix at least one entry must start with.
'
'   AssertionName
'     Diagnostic label for the raised error.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Raises a descriptive assertion failure, with the failure count and list
'     the call returned.
'
' CALLED FROM
'   - TST_Case_TitleBarCapturedFrameClosedIsReported
'   - TST_Case_RibbonRestoreRefusesClosedWindow
'
' UPDATED
'   2026-11-06 - Messages carry the failure count and list (#135).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScanIdx             As Long            'Cursor over the failure entries
    Dim EntryText           As String          'One failure entry as text
    Dim Found               As Boolean         'TRUE once a match is seen

'------------------------------------------------------------------------------
' ASSERT FAILURE WAS REPORTED
'------------------------------------------------------------------------------
    'A success here would mean the component wrote state it could not verify
        If Succeeded Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 50, _
                AssertionName, _
                "expected a structured failure but the operation reported success | " & _
                TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

        If FailureCount < 1 Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 51, _
                AssertionName, _
                "expected at least one ordered failure entry but " & _
                "FailureCount was " & CStr(FailureCount) & " | " & _
                TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

        If Not IsArray(FailureList) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 52, _
                AssertionName, _
                "expected an ordered failure list but no array was returned | " & _
                TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

'------------------------------------------------------------------------------
' SCAN FOR THE EXPECTED STAGE
'------------------------------------------------------------------------------
    'Any entry carrying the prefix satisfies the contract under test
        For ScanIdx = LBound(FailureList) To UBound(FailureList)

            EntryText = CStr(FailureList(ScanIdx))

            If Left$(EntryText, Len(ExpectedPrefix)) = ExpectedPrefix Then
                Found = True
                Exit For
            End If

        Next ScanIdx

        If Not Found Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 53, _
                AssertionName, _
                "no ordered failure entry began with the expected stage " & _
                "prefix " & ExpectedPrefix & " | " & _
                TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

End Sub


Private Sub TST_Log( _
    ByVal ProcName As String, _
    ByVal Stage As String, _
    ByVal Detail As String)

'
'==============================================================================
' TST_Log
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
'   2026-07-25
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
        Debug.Print ProcName & " @ " & Stage & " | " & Detail

End Sub

Private Function TST_TimerElapsedSeconds(ByVal TimerStart As Double) As Double

'
'==============================================================================
' TST_TimerElapsedSeconds
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TimerNow            As Double          'Current Timer reading

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
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
' TST_TitleBarMode
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
'   2026-07-25
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


Private Sub TST_AssertSnapshotWindowState( _
    ByVal TargetWindow As Window, _
    ByVal ExpectedHeadings As Boolean, _
    ByVal ExpectedWorkbookTabs As Boolean, _
    ByVal ExpectedGridlines As Boolean, _
    ByVal AssertionName As String)

'
'==============================================================================
' TST_AssertSnapshotWindowState
'------------------------------------------------------------------------------
' PURPOSE
'   Assert the three window-level Boolean properties managed by EXCEL_UI
'
' INPUTS
'   TargetWindow
'     Window whose managed state is being asserted
'
'   ExpectedHeadings
'     Expected DisplayHeadings state
'
'   ExpectedWorkbookTabs
'     Expected DisplayWorkbookTabs state
'
'   ExpectedGridlines
'     Expected DisplayGridlines state
'
'   AssertionName
'     Diagnostic source used when an assertion fails
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on a missing target or property mismatch
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE TARGET
'------------------------------------------------------------------------------
        If TargetWindow Is Nothing Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 10, _
                AssertionName, _
                "target window is Nothing"
        End If

'------------------------------------------------------------------------------
' ASSERT MANAGED WINDOW STATE
'------------------------------------------------------------------------------
        If TargetWindow.DisplayHeadings <> ExpectedHeadings Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 11, _
                AssertionName, _
                "DisplayHeadings mismatch; expected=" & _
                CStr(ExpectedHeadings) & "; actual=" & _
                CStr(TargetWindow.DisplayHeadings)
        End If

        If TargetWindow.DisplayWorkbookTabs <> ExpectedWorkbookTabs Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 12, _
                AssertionName, _
                "DisplayWorkbookTabs mismatch; expected=" & _
                CStr(ExpectedWorkbookTabs) & "; actual=" & _
                CStr(TargetWindow.DisplayWorkbookTabs)
        End If

        If TargetWindow.DisplayGridlines <> ExpectedGridlines Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 13, _
                AssertionName, _
                "DisplayGridlines mismatch; expected=" & _
                CStr(ExpectedGridlines) & "; actual=" & _
                CStr(TargetWindow.DisplayGridlines)
        End If

End Sub


Private Sub TST_Case_ActiveFramePairRefusesMismatch()

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


Private Sub TST_Case_TitleBarApiErrorCapture()

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


Private Sub TST_Case_QuietScopeOwnershipIsVerified()

'
'==============================================================================
' TST_Case_QuietScopeOwnershipIsVerified
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
' RETURNS
'   None
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
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   ScreenUpdating is restored explicitly after every branch rather than by the
'   scope under test. A test that relied on the code it is testing to clean up
'   would pass for the wrong reason on exactly the branches that matter.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OldValue            As Boolean         'Entry value the scope reports
    Dim Changed             As Boolean         'Ownership the scope claims

    Const PROC As String = "TST_Case_QuietScopeOwnershipIsVerified"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log PROC, "START", _
            "Validating that quiet-scope ownership follows a confirmed readback"

        UI_InternalInjectQuietUpdateFault TEST_QUIET_FAULT_NONE
        Application.ScreenUpdating = True

'------------------------------------------------------------------------------
' BRANCH 1 - ORDINARY PATH OWNS THE TRANSITION
'------------------------------------------------------------------------------
        OldValue = False
        Changed = False

        UI_RuntimeBeginQuietUpdate OldValue, Changed

        TST_AssertBooleanEquals True, OldValue, PROC & ".Ordinary.EntryValue"
        TST_AssertBooleanEquals True, Changed, PROC & ".Ordinary.Owned"
        TST_AssertApplicationProperty _
            False, "ScreenUpdating", PROC & ".Ordinary.Suppressed"

        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC & ".Ordinary.Restored"

'------------------------------------------------------------------------------
' BRANCH 2 - AN ALREADY-SUPPRESSED CALLER IS NOT OWNED
'------------------------------------------------------------------------------
        Application.ScreenUpdating = False

        OldValue = True
        Changed = True

        UI_RuntimeBeginQuietUpdate OldValue, Changed

        TST_AssertBooleanEquals False, OldValue, PROC & ".AlreadyOff.EntryValue"
        TST_AssertBooleanEquals False, Changed, PROC & ".AlreadyOff.NotOwned"

        UI_RuntimeEndQuietUpdate OldValue, Changed

    'End must not re-enable a caller's suppression
        TST_AssertApplicationProperty _
            False, "ScreenUpdating", PROC & ".AlreadyOff.Untouched"

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
        TST_AssertBooleanEquals True, OldValue, PROC & ".EntryFails.Deterministic"
        TST_AssertBooleanEquals False, Changed, PROC & ".EntryFails.NotOwned"
        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC & ".EntryFails.HostUntouched"

        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC & ".EntryFails.NoRestore"

'------------------------------------------------------------------------------
' BRANCH 4 - AN INEFFECTIVE WRITE OWNS NOTHING
'------------------------------------------------------------------------------
        UI_InternalInjectQuietUpdateFault TEST_QUIET_FAULT_WRITE

        OldValue = False
        Changed = True

        UI_RuntimeBeginQuietUpdate OldValue, Changed

        TST_AssertBooleanEquals True, OldValue, PROC & ".WriteIgnored.EntryValue"
        TST_AssertBooleanEquals False, Changed, PROC & ".WriteIgnored.NotOwned"

    'The readback reported the host still enabled, which is what refused
    'ownership; the host really is still enabled
        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC & ".WriteIgnored.HostUnchanged"

        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC & ".WriteIgnored.NoRestore"

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

        TST_AssertBooleanEquals True, OldValue, PROC & ".ReadbackFails.EntryValue"
        TST_AssertBooleanEquals False, Changed, PROC & ".ReadbackFails.NotOwned"

    'The documented consequence: suppressed, unowned, not rolled back
        TST_AssertApplicationProperty _
            False, "ScreenUpdating", PROC & ".ReadbackFails.LeftSuppressed"

        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            False, "ScreenUpdating", PROC & ".ReadbackFails.NoRestore"

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

        TST_AssertBooleanEquals True, Changed, PROC & ".OneShot.SecondScopeOwns"

        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            True, "ScreenUpdating", PROC & ".OneShot.Restored"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Quiet-scope ownership followed the readback on every branch"

End Sub


Private Sub TST_Case_DestructiveRunnersPreserveCallerSnapshot()

'
'==============================================================================
' TST_Case_DestructiveRunnersPreserveCallerSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that every destructive public runner refuses a caller-owned snapshot
'   and leaves it intact and restorable.
'
' WHY THIS EXISTS
'   All four runners raised their refusal after arming a local error handler.
'   The handler captured it, resumed at Safe_Exit, and the unconditional clear
'   there deleted the snapshot the refusal existed to protect. The caller
'   received the error and lost the state anyway, and nothing asserted
'   otherwise. Two of the four are certification units, so this was on the
'   release path.
'
'   A test that only checked UI_HasExcelUIStateSnapshot would be too weak: the
'   flag can survive while the captured content does not. Each block therefore
'   mutates the host and restores from the preserved snapshot, and asserts the
'   original state came back.
'
' RETURNS
'   None
'
' BEHAVIOR
'   Four independent blocks, one per runner. Each establishes a distinguishable
'   state, captures it, invokes the runner, asserts the exact refusal, proves
'   the snapshot survived by restoring through it, and clears before the next
'   block starts.
'
' ERROR POLICY
'   - Raises on assertion failure
'   - Leaves no snapshot behind on the success path
'
' DEPENDENCIES
'   - UI_SetExcelUI
'   - UI_CaptureExcelUIState
'   - UI_ClearExcelUIStateSnapshot
'   - UI_ResetExcelUIToSnapshot
'   - TST_AssertRefusal
'   - TST_AssertSnapshotAvailability
'   - TST_AssertApplicationProperty
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   The blocks are independent on purpose. Sharing one snapshot across all four
'   would mean the first failure invalidated the evidence for the other three,
'   and the point of the case is to know which runners are safe.
'
'   The refusal is the entire interaction. Every guard now precedes handler
'   arming and initialization, so no runner reaches workbook or window
'   creation, Ribbon manipulation, logging or certification-counter reset. The
'   only host mutation here is this case's own capture, mutate and restore.
'
' UPDATED
'   2026-11-06
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RaisedNumber        As Long            'Error number reaching this case
    Dim RaisedSource        As String          'Error source reaching this case
    Dim RaisedDescription   As String          'Error text reaching this case

    Const PROC As String = _
        "TST_Case_DestructiveRunnersPreserveCallerSnapshot"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log PROC, "START", _
            "Validating that destructive runners preserve a caller snapshot"

    'The pack reaches this case only when no explicit snapshot pre-existed, but
    'assert it rather than assume it: every block below depends on owning the
    'slot outright.
        TST_AssertSnapshotAvailability False, PROC & ".NoSnapshotOnEntry"

'------------------------------------------------------------------------------
' BLOCK 1 - Test_EXCEL_UI_RunCertificationSelfTest
'------------------------------------------------------------------------------
        UI_SetExcelUI StatusBar:=UI_Hide, FormulaBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        UI_CaptureExcelUIState
        TST_AssertSnapshotAvailability True, PROC & ".SelfTest.Captured"

    'Capture immediately, before anything else can clear Err
        On Error Resume Next

            Test_EXCEL_UI_RunCertificationSelfTest

            RaisedNumber = Err.Number
            RaisedSource = Err.Source
            RaisedDescription = Err.Description

            Err.Clear

        On Error GoTo 0

        TST_AssertRefusal _
            ExpectedNumber:=TEST_CERT_ERR_BASE + 21, _
            ExpectedSource:="Test_EXCEL_UI_RunCertificationSelfTest", _
            ExpectedDescription:= _
                "an explicit EXCEL_UI snapshot already exists; clear or " & _
                "restore it before running this destructive test", _
            ActualNumber:=RaisedNumber, _
            ActualSource:=RaisedSource, _
            ActualDescription:=RaisedDescription, _
            AssertionName:=PROC & ".SelfTest.Refusal"

        TST_AssertSnapshotAvailability True, PROC & ".SelfTest.Survived"

    'The flag can survive while the captured content does not, so restore
    'through it and assert the original state came back
        UI_SetExcelUI StatusBar:=UI_Show, FormulaBar:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

        UI_ResetExcelUIToSnapshot
        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertApplicationProperty _
            False, "DisplayStatusBar", PROC & ".SelfTest.StatusBarRestored"
        TST_AssertApplicationProperty _
            True, "DisplayFormulaBar", PROC & ".SelfTest.FormulaBarRestored"

        UI_ClearExcelUIStateSnapshot
        TST_AssertSnapshotAvailability False, PROC & ".SelfTest.Cleared"

'------------------------------------------------------------------------------
' BLOCK 2 - Test_EXCEL_UI_RunRibbonSdiProbe
'------------------------------------------------------------------------------
        UI_SetExcelUI StatusBar:=UI_Show, FormulaBar:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

        UI_CaptureExcelUIState
        TST_AssertSnapshotAvailability True, PROC & ".RibbonProbe.Captured"

        On Error Resume Next

            Test_EXCEL_UI_RunRibbonSdiProbe

            RaisedNumber = Err.Number
            RaisedSource = Err.Source
            RaisedDescription = Err.Description

            Err.Clear

        On Error GoTo 0

        TST_AssertRefusal _
            ExpectedNumber:=TEST_RIBBON_ERR_BASE + 1, _
            ExpectedSource:="Test_EXCEL_UI_RunRibbonSdiProbe", _
            ExpectedDescription:= _
                "an explicit EXCEL_UI snapshot already exists; clear or " & _
                "restore it before probing", _
            ActualNumber:=RaisedNumber, _
            ActualSource:=RaisedSource, _
            ActualDescription:=RaisedDescription, _
            AssertionName:=PROC & ".RibbonProbe.Refusal"

        TST_AssertSnapshotAvailability True, PROC & ".RibbonProbe.Survived"

        UI_SetExcelUI StatusBar:=UI_Hide, FormulaBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        UI_ResetExcelUIToSnapshot
        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertApplicationProperty _
            True, "DisplayStatusBar", PROC & ".RibbonProbe.StatusBarRestored"
        TST_AssertApplicationProperty _
            False, "DisplayFormulaBar", PROC & ".RibbonProbe.FormulaBarRestored"

        UI_ClearExcelUIStateSnapshot
        TST_AssertSnapshotAvailability False, PROC & ".RibbonProbe.Cleared"

'------------------------------------------------------------------------------
' BLOCK 3 - Test_EXCEL_UI_RunSnapshotIdentity
'------------------------------------------------------------------------------
        UI_SetExcelUI StatusBar:=UI_Hide, FormulaBar:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

        UI_CaptureExcelUIState
        TST_AssertSnapshotAvailability True, PROC & ".SnapshotIdentity.Captured"

        On Error Resume Next

            Test_EXCEL_UI_RunSnapshotIdentity

            RaisedNumber = Err.Number
            RaisedSource = Err.Source
            RaisedDescription = Err.Description

            Err.Clear

        On Error GoTo 0

        TST_AssertRefusal _
            ExpectedNumber:=TEST_SNAPSHOT_ID_ERR_BASE + 1, _
            ExpectedSource:="Test_EXCEL_UI_RunSnapshotIdentity", _
            ExpectedDescription:= _
                "an explicit EXCEL_UI snapshot already exists; clear or " & _
                "restore it before running this destructive test", _
            ActualNumber:=RaisedNumber, _
            ActualSource:=RaisedSource, _
            ActualDescription:=RaisedDescription, _
            AssertionName:=PROC & ".SnapshotIdentity.Refusal"

        TST_AssertSnapshotAvailability True, PROC & ".SnapshotIdentity.Survived"

        UI_SetExcelUI StatusBar:=UI_Show, FormulaBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        UI_ResetExcelUIToSnapshot
        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertApplicationProperty _
            False, "DisplayStatusBar", _
            PROC & ".SnapshotIdentity.StatusBarRestored"
        TST_AssertApplicationProperty _
            False, "DisplayFormulaBar", _
            PROC & ".SnapshotIdentity.FormulaBarRestored"

        UI_ClearExcelUIStateSnapshot
        TST_AssertSnapshotAvailability False, PROC & ".SnapshotIdentity.Cleared"

'------------------------------------------------------------------------------
' BLOCK 4 - Test_EXCEL_UI_RunTitleBarSdiIdentity
'------------------------------------------------------------------------------
        UI_SetExcelUI StatusBar:=UI_Show, FormulaBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        UI_CaptureExcelUIState
        TST_AssertSnapshotAvailability True, PROC & ".TitleBarSdi.Captured"

        On Error Resume Next

            Test_EXCEL_UI_RunTitleBarSdiIdentity

            RaisedNumber = Err.Number
            RaisedSource = Err.Source
            RaisedDescription = Err.Description

            Err.Clear

        On Error GoTo 0

        TST_AssertRefusal _
            ExpectedNumber:=TEST_TITLEBAR_SDI_ERR_BASE + 1, _
            ExpectedSource:="Test_EXCEL_UI_RunTitleBarSdiIdentity", _
            ExpectedDescription:= _
                "an explicit EXCEL_UI snapshot already exists; clear or " & _
                "restore it before running this destructive test", _
            ActualNumber:=RaisedNumber, _
            ActualSource:=RaisedSource, _
            ActualDescription:=RaisedDescription, _
            AssertionName:=PROC & ".TitleBarSdi.Refusal"

        TST_AssertSnapshotAvailability True, PROC & ".TitleBarSdi.Survived"

        UI_SetExcelUI StatusBar:=UI_Hide, FormulaBar:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

        UI_ResetExcelUIToSnapshot
        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertApplicationProperty _
            True, "DisplayStatusBar", PROC & ".TitleBarSdi.StatusBarRestored"
        TST_AssertApplicationProperty _
            True, "DisplayFormulaBar", PROC & ".TitleBarSdi.FormulaBarRestored"

        UI_ClearExcelUIStateSnapshot
        TST_AssertSnapshotAvailability False, PROC & ".TitleBarSdi.Cleared"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "All four destructive runners refused and preserved the snapshot"

End Sub


Private Sub TST_AssertRefusal( _
    ByVal ExpectedNumber As Long, _
    ByVal ExpectedSource As String, _
    ByVal ExpectedDescription As String, _
    ByVal ActualNumber As Long, _
    ByVal ActualSource As String, _
    ByVal ActualDescription As String, _
    ByVal AssertionName As String)

'
'==============================================================================
' TST_AssertRefusal
'------------------------------------------------------------------------------
' PURPOSE
'   Assert that a destructive runner refused with the exact error it promises.
'
' WHY THIS EXISTS
'   Asserting only that something was raised would pass against a runner that
'   refused for the wrong reason, or that failed later for an unrelated one.
'   The number identifies the guard, the source identifies which runner raised
'   it, and the description is what a caller actually reads.
'
'   The description is compared exactly rather than merely for being non-empty.
'   A length test passes against any text at all, including another runner's
'   message or an unrelated failure that reached the same handler, which is
'   most of what this case exists to distinguish.
'
' INPUTS
'   ExpectedNumber
'     Error number the guard promises
'
'   ExpectedSource
'     Procedure name the guard raises with
'
'   ExpectedDescription
'     Exact text the guard raises with
'
'   ActualNumber / ActualSource / ActualDescription
'     Values captured from Err immediately after the invocation
'
'   AssertionName
'     Diagnostic prefix for the three assertions
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises on assertion failure
'
' NOTES
'   The actuals are logged before being asserted, because a failing assertion
'   raises with the assertion name only and the observed values would otherwise
'   be lost.
'
' CALLED FROM
'   - TST_Case_DestructiveRunnersPreserveCallerSnapshot
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' LOG WHAT WAS OBSERVED
'------------------------------------------------------------------------------
        TST_Log "TST_AssertRefusal", "OBSERVED", _
            AssertionName & " raised " & CStr(ActualNumber) & _
            " | Source: " & ActualSource & _
            " | " & ActualDescription

'------------------------------------------------------------------------------
' ASSERT THE REFUSAL
'------------------------------------------------------------------------------
    'The number identifies which guard fired
        TST_AssertTrue _
            (ActualNumber = ExpectedNumber), _
            AssertionName & ".Number"

    'The source identifies which runner refused
        TST_AssertTrue _
            (ActualSource = ExpectedSource), _
            AssertionName & ".Source"

    'Compared exactly. A length test would pass against any text at all,
    'including another runner's message.
        TST_AssertTrue _
            (ActualDescription = ExpectedDescription), _
            AssertionName & ".Description"

End Sub


Private Sub TST_AssertTrue( _
    ByVal ActualValue As Boolean, _
    ByVal AssertionName As String)

'
'==============================================================================
' TST_AssertTrue
'------------------------------------------------------------------------------
' PURPOSE
'   Raise when a Boolean assertion is FALSE
'
' INPUTS
'   ActualValue
'     Boolean result to assert
'
'   AssertionName
'     Diagnostic source used when the assertion fails
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Raises when ActualValue is FALSE
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        If Not ActualValue Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 20, _
                AssertionName, _
                "expected TRUE but received FALSE"
        End If

End Sub


Private Sub TST_SafeCloseWorkbook(ByRef TargetWorkbook As Workbook)

'
'==============================================================================
' TST_SafeCloseWorkbook
'------------------------------------------------------------------------------
' PURPOSE
'   Close and release one temporary workbook without saving.
'
' ERROR POLICY
'   - Suppresses cleanup errors locally.
'
' UPDATED
'   2026-08-01
'==============================================================================
'

        On Error Resume Next

        If Not TargetWorkbook Is Nothing Then
            TargetWorkbook.Close SaveChanges:=False
        End If

        Set TargetWorkbook = Nothing

End Sub


Private Sub TST_SafeCloseWindow(ByRef TargetWindow As Window)

'
'==============================================================================
' TST_SafeCloseWindow
'------------------------------------------------------------------------------
' PURPOSE
'   Close and release one temporary Excel Window on a best-effort basis
'
' INPUTS / OUTPUTS
'   TargetWindow
'     Temporary Window to close and release
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Suppresses cleanup errors locally
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error Resume Next

'------------------------------------------------------------------------------
' CLOSE AND RELEASE
'------------------------------------------------------------------------------
        If Not TargetWindow Is Nothing Then
            TargetWindow.Close
        End If

        Set TargetWindow = Nothing

End Sub


#If VBA7 Then
Private Function TST_BuildApiFailureText( _
    ByVal ApiName As String, _
    ByVal Operation As String, _
    ByVal DllError As Long, _
    ByVal TargetHwnd As LongPtr) As String
#Else
Private Function TST_BuildApiFailureText( _
    ByVal ApiName As String, _
    ByVal Operation As String, _
    ByVal DllError As Long, _
    ByVal TargetHwnd As Long) As String
#End If

'
'==============================================================================
' TST_BuildApiFailureText
'------------------------------------------------------------------------------
' PURPOSE
'   Build the failure text for a native call that reported an error, in the
'   same form as M_EXCEL_UI_TITLEBAR
'
' WHY THIS EXISTS
'   The harness WinAPI helpers report their own failures in this form, and
'   the error-capture case compares the title-bar module's text with it
'   exactly. One builder keeps both uses identical.
'
' INPUTS
'   ApiName
'     Exported name of the function that ran
'
'   Operation
'     What the call was doing
'
'   DllError
'     Err.LastDllError as captured immediately after the call. This builder
'     never reads Err itself.
'
'   TargetHwnd
'     The window handle the call addressed
'
' RETURNS
'   "<API> failed; operation=<operation>; LastDllError=<n>; hWnd=<h>"
'
' ERROR POLICY
'   - Does NOT raise
'
' UPDATED
'   2026-11-06
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Never let diagnostic formatting raise into a failure path
        On Error Resume Next

'------------------------------------------------------------------------------
' BUILD DIAGNOSTIC TEXT
'------------------------------------------------------------------------------
        TST_BuildApiFailureText = _
            ApiName & " failed; operation=" & Operation & _
            "; LastDllError=" & CStr(DllError) & _
            "; hWnd=" & CStr(TargetHwnd)

End Function


Private Function TST_BuildRuntimeErrorText() As String

'
'==============================================================================
' TST_BuildRuntimeErrorText
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
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber           As Long            'Err.Number captured on entry
    Dim ErrDescription      As String          'Err.Description captured on entry
    Dim ErrSource           As String          'Err.Source captured on entry
    Dim ErrLine             As Long            'Erl captured on entry

'------------------------------------------------------------------------------
' CAPTURE ERR STATE
'------------------------------------------------------------------------------
    'Read the Err object BEFORE any On Error statement. Any form of On Error
    'resets Err, so protecting this routine first would blank the very values
    'it exists to report and every diagnostic would read "0: ".
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
' BUILD RUNTIME ERROR TEXT
'------------------------------------------------------------------------------
    'Build a consistent diagnostic string from the captured Err state
        TST_BuildRuntimeErrorText = _
            CStr(ErrNumber) & ": " & ErrDescription & _
            IIf(Len(ErrSource) > 0, " | Source: " & ErrSource, vbNullString) & _
            IIf(ErrLine <> 0, " | Line: " & CStr(ErrLine), vbNullString)

End Function
