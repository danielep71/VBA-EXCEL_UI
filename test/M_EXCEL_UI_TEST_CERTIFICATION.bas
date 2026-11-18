Attribute VB_Name = "M_EXCEL_UI_TEST_CERTIFICATION"
Option Explicit
Option Private Module

'==============================================================================
' M_EXCEL_UI_TEST_CERTIFICATION
'------------------------------------------------------------------------------
' PURPOSE
'   Run the release certification and the combined regression pack, and hold
'   the cases that check the certification and the destructive runners
'
' WHY THIS EXISTS
'   Certification runs every mandatory unit in one pass and writes one
'   complete/incomplete and pass/fail verdict with machine-readable evidence.
'   It and the combined pack are the only runners that span every suite, so
'   they live above the suite modules rather than in any one of them
'
' PUBLIC SURFACE
'   - Test_EXCEL_UI_RunAll
'   - Test_EXCEL_UI_RunCore
'   - Test_EXCEL_UI_RunReleaseCertification
'   - Test_EXCEL_UI_RunCertificationSelfTest
'   - Test_EXCEL_UI_RunOwnershipCleanupChecks
'   - TST_RunRunnerWithLog, which runs one public runner of the suites
'     with the log sink set and returns its verdict instead of raising;
'     tools\Check-EXCEL_UI.ps1 calls it
'
' TEST SCOPE
'   Release certification units
'     - RegressionPack: the full regression pack, title-bar cases included
'     - SnapshotIdentity
'     - TitleBarSdiIdentity
'
'   Certification and ownership cases
'     - the cleanup verdict compares the host with a baseline of its entry
'       state that certification owns: the workbook and window counts, every
'       window by retained object, the application and window properties, the
'       title bar and frame registry of each window, the active window and the
'       Ribbon of its owner; a leaked or replaced window, an altered property,
'       a title-bar mismatch, a failing cleanup helper and the residue of a
'       failed quiet-update readback are each reported as a finding
'     - a forced cleanup failure fails the outer certification call, with its
'       verdict and findings in the error it raises
'     - a unit that fails after a capture reports the failure list it
'       asserted on and leaves no snapshot for the next unit
'     - every destructive public runner refuses a caller-owned snapshot
'     - a certification record that cannot store an entry, or whose
'       buffers disagree with its counts, gives FAIL | INCOMPLETE evidence
'       that names the recording failure, and a reset starts clean
'
' STATE MANAGEMENT
'   - The packs snapshot the current managed Excel UI state before testing
'     and attempt a restore after success or failure; a restoration step that
'     fails or does not stick fails the pack
'   - Certification captures its own baseline on entry, held in a local
'     variable and independent of the snapshot subsystem, and judges the exit
'     state against it. Every difference is a finding in the run's record
'   - A pre-existing explicit EXCEL_UI snapshot is left untouched by skipping
'     the snapshot-destructive lifecycle cases
'   - The packs put the entry frame style back through the harness WinAPI
'     helpers, never through the API under test, and fail when they cannot
'   - Certification releases a snapshot a unit left behind before the next
'     unit runs, and reports it as a cleanup failure
'
' DEPENDENCIES
'   - M_EXCEL_UI_TEST_HARNESS
'   - M_EXCEL_UI_TEST_RUNTIME
'   - M_EXCEL_UI_TEST_SNAPSHOT
'   - M_EXCEL_UI_TEST_TITLEBAR
'
' UPDATED
'   2026-11-18 - Certification cleanup compares a baseline it owns instead of
'                four conditions, reports every difference as a finding, and
'                fails the outer call when cleanup fails; added the cleanup
'                regression cases and the forced-failure ownership check (#35).
'   2026-11-16 - Added TST_RunRunnerWithLog for the scripted Excel checks
'                (#197).
'   2026-11-16 - Certification composes its verdict and evidence through
'                the harness, failing closed when a unit or skip could not
'                be recorded; added the certification-record case (#50).
'   2026-11-15 - Split out of the single regression module; the runners and
'                cases moved here unchanged apart from visibility (#90).
'   2026-11-06 - Certification releases a snapshot a unit left behind before
'                the next unit runs, reporting it as a cleanup failure (#135).
'   2026-08-21 - Registered every frame-state case in the pack certification
'                runs, added the stale-frame-entry and certification-cleanup
'                cases, and stopped certification destroying the error it
'                re-raises.
'   2026-08-19 - Added the certification runner.
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.2
'==============================================================================


'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const TST_CERT_APP_COUNT As Long = 4  'Application-level properties the baseline compares
    Private Const TST_CERT_APP_STATUSBAR As Long = 1  'Baseline index: DisplayStatusBar
    Private Const TST_CERT_APP_SCREEN As Long = 4  'Baseline index: ScreenUpdating
    Private Const TST_CERT_WINPROP_COUNT As Long = 3  'Window-level properties the baseline compares
    Private Const TST_CERT_WINPROP_TABS As Long = 2  'Baseline index: DisplayWorkbookTabs
    Private Const TST_CERT_WS_SYSMENU As Long = &H80000  'WS_SYSMENU, an owned frame bit other than the caption

    'Cleanup helper faults (TST_CertInjectHelperFault)
    Private Const TST_CERT_HELPER_FAULT_NONE As Long = 0  'Helper fault: none armed
    Private Const TST_CERT_HELPER_FAULT_READ As Long = 1  'Helper fault: the next property read fails
    Private Const TST_CERT_HELPER_FAULT_STAGE As Long = 2  'Helper fault: the next window comparison raises

    'Quiet-update fault kinds of UI_InternalInjectQuietUpdateFault, mirrored here
    'as the runtime regression module mirrors them
    Private Const TST_CERT_QUIET_FAULT_NONE As Long = 0  'Quiet scope: no fault
    Private Const TST_CERT_QUIET_FAULT_READBACK As Long = 3  'Quiet scope: readback fails

'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
'The certification-owned baseline of the Excel host. It is held in the caller's
'variable, written only by TST_CertCaptureBaseline, and shares no state with the
'production snapshot subsystem (#35). Every window present on entry is retained
'as an object: it is the proof, on exit, that the same window still responds.
'Windows are matched by the handle each object reports, never with Is, because
'Excel returns a new wrapper from every read of Application.ActiveWindow.
Private Type tCertBaseline
    Captured            As Boolean             'TRUE once every capture stage ran
    BookCount           As Long                'Workbooks.Count on entry
    WindowCount         As Long                'Application.Windows.Count on entry
    AppKnown(1 To 4)    As Boolean             'Application property was read
    AppValue(1 To 4)    As Boolean             'Application property value
    Win()               As Object              'Every window present on entry
    WinLabel()          As String              'Caption label, diagnostic only
    WinHwndKnown()      As Boolean             'The window's hWnd was read
    WinPropKnown()      As Boolean             'Window property was read (window, property)
    WinPropValue()      As Boolean             'Window property value (window, property)
    WinTitleKnown()     As Boolean             'The frame style was read
    WinRegFound()       As Boolean             'A frame-registry slot held the hWnd
    WinRegRefresh()     As Boolean             'Registry: a frame refresh was owed
    WinRegHidden()      As Boolean             'Registry: the component hid the frame
    ActiveKnown         As Boolean             'The active window was identified
    ActiveIndex         As Long                'Index of the active window in Win
    RibbonKnown         As Boolean             'The Ribbon was attributed to the active window
    RibbonVisible       As Boolean             'Ribbon visibility of the active window
#If VBA7 Then
    WinHwnd()           As LongPtr             'hWnd each window reported
    WinStyle()          As LongPtr             'Complete GWL_STYLE of each frame
    ActiveHwnd          As LongPtr             'hWnd of the active window
#Else
    WinHwnd()           As Long                'hWnd each window reported
    WinStyle()          As Long                'Complete GWL_STYLE of each frame
    ActiveHwnd          As Long                'hWnd of the active window
#End If
End Type

'------------------------------------------------------------------------------
' PRIVATE MODULE STATE
'------------------------------------------------------------------------------
'One-shot cleanup helper fault, consumed by the first helper it matches. Only
'TST_CertInjectHelperFault arms it, and only the helper-failure case calls that.
Private m_CertHelperFault               As Long


'
'------------------------------------------------------------------------------
'
'                              PUBLIC RUNNERS
'
'------------------------------------------------------------------------------
'


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
'   - Captures a baseline of the host that certification owns, independent of
'     UI_CaptureExcelUIState: every window by object, the workbook and window
'     counts, the managed application and window properties, each window's
'     title bar and frame registry state, the active window, and the Ribbon of
'     its owner. A surface that cannot be read or addressed is named as
'     unverified, not treated as clean.
'   - Runs the full regression pack, the snapshot-identity runner and the SDI
'     title-bar identity runner, each under its own error boundary so one
'     failing unit does not conceal the rest.
'   - Releases a snapshot a unit left behind before the next unit runs, so the
'     next unit is not refused because of it, and records the leak.
'   - Verifies cleanup by comparing the exit state with that baseline, and
'     records every difference as a finding: a snapshot left behind, a changed
'     workbook or window count, a missing, replaced or unexpected window, an
'     altered application or window property, a title bar or frame registry
'     that differs, a different active window, a different Ribbon. Then makes
'     the entry active window active again, and reports a failure to.
'   - Emits a JSON evidence document and a text summary, and writes both to the
'     temporary folder on a best-effort basis.
'   - Reports FAIL | INCOMPLETE, and names the failure in both documents,
'     when a unit or skip could not be recorded.
'   - Raises when the verdict is anything other than a complete pass, with the
'     verdict and every cleanup finding in the error it raises.
'
' ERROR POLICY
'   - Individual unit failures are captured, not propagated.
'   - Raises once at the end when the verdict is not PASS/COMPLETE.
'
' DEPENDENCIES
'   - TST_CertResetCounters
'   - TST_CertCaptureBaseline
'   - TST_CertRunUnit
'   - TST_CertEvaluateCleanup
'   - TST_CertRestoreActiveWindow
'   - TST_CertComposeEvidence
'   - TST_CertTryWriteEvidence
'
' NOTES
'   Destructive: creates and closes temporary workbooks and toggles every
'   managed UI element. Save unsaved work before running it.
'
'   A one-shot run fault, armed only by Test_EXCEL_UI_RunOwnershipCleanupChecks,
'   makes the run execute no unit, invert the status bar and write no evidence
'   file, so that a forced cleanup failure can be seen to fail this call.
'
' UPDATED
'   2026-11-18 - Cleanup compares a baseline certification owns and records
'                every difference as a finding (#35).
'   2026-11-16 - The verdict and both documents come from
'                TST_CertComposeEvidence, which fails closed when the record
'                could not be stored (#50).
'   2026-11-06 - A snapshot a unit left behind fails cleanup, although it is
'                released before the next unit runs (#135).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Baseline            As tCertBaseline   'Host state captured on entry
    Dim FaultKind           As Long            'Run fault taken for this call

    Dim UnitsFailed         As Long            'Units that reported failure
    Dim CleanupOK           As Boolean         'TRUE when no state leaked
    Dim CleanupDetail       As String          'Detail of cleanup outside the record

    Dim Complete            As Boolean         'TRUE when nothing was skipped
    Dim Passed              As Boolean         'TRUE when the verdict is a pass
    Dim Verdict             As String          'Human-readable verdict line

    Dim JsonText            As String          'Machine-readable evidence
    Dim ReportText          As String          'Human-readable evidence
    Dim JsonPath            As String          'Path the JSON was written to
    Dim ReportPath          As String          'Path the report was written to
    Dim RecordingDetail     As String          'Names a recording failure

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

    'Take the run fault here, so that a fault armed for this call is consumed by
    'it, however the call ends
        FaultKind = TST_CertTakeRunFault()

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

        If Application.ActiveWindow Is Nothing Then
            Err.Raise _
                TEST_CERT_ERR_BASE + 2, _
                PROC, _
                "no active Excel window is available"
        End If

'------------------------------------------------------------------------------
' CAPTURE THE ENTRY BASELINE
'------------------------------------------------------------------------------
    'What cleanup is judged against, held here and read by nothing else. A stage
    'that could not capture is a finding in the record, which fails cleanup, so
    'the run goes on and says so rather than stopping on an unproved baseline.
        If Not TST_CertCaptureBaseline(Baseline, m_CertRecord) Then
            TST_Log PROC, "BASELINE", _
                "The entry baseline was not captured completely; cleanup will " & _
                "be reported as FAILED"
        End If

'------------------------------------------------------------------------------
' RUN MANDATORY UNITS
'------------------------------------------------------------------------------
    'Each unit is run under its own boundary so that one failure does not hide
    'the state of everything after it. An armed run fault runs none of them.
        If FaultKind = TEST_CERT_RUN_FAULT_CLEANUP Then
            TST_CertApplyRunFault
        Else
            TST_CertRunUnit "RegressionPack"
            TST_CertRunUnit "SnapshotIdentity"
            TST_CertRunUnit "TitleBarSdiIdentity"
        End If

'------------------------------------------------------------------------------
' VERIFY CLEANUP
'------------------------------------------------------------------------------
    'Cleanup failure is a run failure. A suite that leaves a snapshot, a stray
    'or replaced window, an altered host setting or a different active window
    'behind has not finished, however many assertions passed on the way. Every
    'difference is a finding in the record; a snapshot a unit left behind was
    'recorded as one when that unit ended, because the release before the next
    'unit means the exit state can no longer show it.
        CleanupOK = TST_CertEvaluateCleanup(Baseline, m_CertRecord)

    'The comparison changed nothing. Putting the entry window back is a
    'separate, test-owned step, and a failure of it is a finding of its own.
        TST_CertRestoreActiveWindow Baseline, m_CertRecord

'------------------------------------------------------------------------------
' DETERMINE VERDICT AND COMPOSE EVIDENCE
'------------------------------------------------------------------------------
    'Completeness and correctness are separate questions and are reported as
    'such. A run that skipped a mandatory unit, or could not record one, is
    'not a pass, whatever the assertions that did execute reported.
    '
    'The composer checks the record before reading it and decides the verdict
    'after, so a unit, skip or finding that could not be stored makes the run
    'FAIL | INCOMPLETE and is named in both documents, and any finding makes
    'cleanup FAILED whatever CleanupOK says. Passed and Complete are
    'authoritative; both forms are built before either is written, so a
    'file-system failure cannot cost the Immediate Window record as well.
        TST_CertComposeEvidence _
            Rec:=m_CertRecord, _
            CleanupOK:=CleanupOK, _
            CleanupDetail:=CleanupDetail, _
            UnitsFailed:=UnitsFailed, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            RecordingDetail:=RecordingDetail, _
            JsonText:=JsonText, _
            ReportText:=ReportText

'------------------------------------------------------------------------------
' EMIT EVIDENCE
'------------------------------------------------------------------------------
        TST_Log PROC, "EVIDENCE", vbNewLine & ReportText
        TST_Log PROC, "JSON", JsonText

    'File output is a convenience, never a gate on the verdict. A run under a
    'run fault writes none: its evidence is a deliberate failure, and the
    'collection of a real run must not find it.
        If FaultKind = TEST_CERT_RUN_FAULT_NONE Then

            If TST_CertTryWriteEvidence(JsonText, "json", JsonPath) Then
                TST_Log PROC, "WROTE", JsonPath
            End If

            If TST_CertTryWriteEvidence(ReportText, "txt", ReportPath) Then
                TST_Log PROC, "WROTE", ReportPath
            End If

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
                Verdict & _
                IIf(Len(CleanupDetail) > 0, " | " & CleanupDetail, vbNullString) & _
                IIf(Len(RecordingDetail) > 0, " | " & RecordingDetail, vbNullString)
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
'   - Records the outcome and, when the unit left a snapshot, records a cleanup
'     finding that names the unit.
'
' ERROR POLICY
'   - Does not raise. A failing unit is recorded and the run continues.
'
' DEPENDENCIES
'   - TST_CertExecuteUnit
'   - TST_CertRecordUnit
'   - TST_CertAddFinding
'   - TST_CertAppendDetail
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
'   2026-11-18 - The leak is recorded as a cleanup finding (#35).
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

    'The release hid the leak from the exit state, so it is recorded here as a
    'cleanup finding; the baseline comparison could not find it afterwards
        If LeftSnapshot Then
            m_CertLeakDetail = TST_CertAppendDetail( _
                m_CertLeakDetail, _
                UnitName & " left an EXCEL_UI snapshot behind")

            TST_CertAddFinding m_CertRecord, "unit.snapshotLeft", _
                UnitName & " left an EXCEL_UI snapshot behind"
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
'   units include Private procedures and a name-based call would fail silently
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


'
'------------------------------------------------------------------------------
'
'                       CERTIFICATION CLEANUP BASELINE
'
'------------------------------------------------------------------------------
'


Private Function TST_CertCaptureBaseline( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_CertCaptureBaseline
'------------------------------------------------------------------------------
' PURPOSE
'   Capture the certification-owned baseline of the Excel host, which cleanup
'   judges the exit state against.
'
' WHY THIS EXISTS
'   Cleanup compared four conditions, so a leaked window, a replaced window or
'   an altered status bar, scroll bar, formula bar, heading, tab, gridline or
'   title bar could coexist with cleanup=OK (#35). The baseline is independent
'   of UI_CaptureExcelUIState: testing the component with its own snapshot
'   would repeat its implementation assumptions, and would collide with the
'   no-snapshot precondition.
'
'   It shares no mutable state with the production snapshot subsystem. It
'   lives in the caller's variable, reads the host through the harness helpers,
'   and touches the production modules only through the read-only
'   UI_InternalTryReadFrameState seam.
'
' INPUTS
'   Base
'     ByRef. Receives the baseline. It is cleared first.
'
'   Rec
'     ByRef. The record that receives a finding for a stage that raised and an
'     unverified entry for every surface that cannot be compared.
'
' RETURNS
'   Boolean
'     TRUE  => every stage ran to its end; Base.Captured is TRUE.
'     FALSE => a stage raised; Base.Captured is FALSE, Rec holds a finding for
'              it, and cleanup cannot be proved.
'
' BEHAVIOR
'   - Retains every Window present on entry by object, with its hWnd, caption,
'     Headings, WorkbookTabs and Gridlines, title-bar style and frame-registry
'     state.
'   - Records Workbooks.Count, Application.Windows.Count, the four
'     application-level properties, and the active window.
'   - Reads the Ribbon with the active window as its owner.
'   - Names every surface it could not read or address as unverified, and the
'     two scopes that no read can widen: the Ribbon of any other window, and
'     the frame-registry entries of windows opened during the run.
'
' ERROR POLICY
'   - Does not raise. A failure is a finding, never a clean baseline.
'
' DEPENDENCIES
'   - TST_CertCaptureCounts
'   - TST_CertCaptureAppProps
'   - TST_CertCaptureWindows
'   - TST_CertCaptureActive
'   - TST_CertCaptureRibbon
'   - TST_CertAddUnverified
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'   - the cleanup regression cases
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Blank               As tCertBaseline   'Never written: the cleared value
    Dim Complete            As Boolean         'TRUE while every stage completed
    Dim StageDone           As Boolean         'Outcome of the stage just run

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_CertCaptureBaseline = False

        Base = Blank
        Complete = True

'------------------------------------------------------------------------------
' CAPTURE STAGES
'------------------------------------------------------------------------------
    'Every stage runs, so one report names every surface that failed
        StageDone = TST_CertCaptureCounts(Base, Rec)
        Complete = Complete And StageDone

        StageDone = TST_CertCaptureAppProps(Base, Rec)
        Complete = Complete And StageDone

        StageDone = TST_CertCaptureWindows(Base, Rec)
        Complete = Complete And StageDone

        StageDone = TST_CertCaptureActive(Base, Rec)
        Complete = Complete And StageDone

        StageDone = TST_CertCaptureRibbon(Base, Rec)
        Complete = Complete And StageDone

'------------------------------------------------------------------------------
' NAME WHAT NO READ CAN COVER
'------------------------------------------------------------------------------
    'Stated on every run, so a clean verdict says what it excludes
        TST_CertAddUnverified Rec, "ribbon.scope", _
            "the Ribbon is per window and every read acts on the active " & _
            "window, so only the Ribbon of the entry active window is " & _
            "compared, and only while that window is active again on exit; " & _
            "the Ribbon of any other window is not verified"

        TST_CertAddUnverified Rec, "frameRegistry.scope", _
            "frame refresh debt and component ownership are compared for the " & _
            "windows present on entry; the read-only seam cannot enumerate " & _
            "the registry entries of windows opened and closed during the run"

'------------------------------------------------------------------------------
' PUBLISH
'------------------------------------------------------------------------------
        Base.Captured = Complete

        TST_CertCaptureBaseline = Complete

End Function


Private Function TST_CertCaptureCounts( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_CertCaptureCounts
'------------------------------------------------------------------------------
' PURPOSE
'   Record Workbooks.Count and Application.Windows.Count on entry.
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline to fill and the record to report a failure to.
'
' RETURNS
'   Boolean
'     TRUE when both counts were read; otherwise FALSE with a finding.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - TST_CertCaptureBaseline
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Msg                 As String          'Diagnostic from an error

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertCaptureCounts = False

        Base.BookCount = Application.Workbooks.Count
        Base.WindowCount = Application.Windows.Count

        TST_CertCaptureCounts = True

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "baseline.counts", _
            "the workbook and window counts could not be read on entry | " & Msg

        TST_CertCaptureCounts = False

        Resume Safe_Exit

End Function


Private Function TST_CertCaptureAppProps( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_CertCaptureAppProps
'------------------------------------------------------------------------------
' PURPOSE
'   Record the four application-level properties the component manages on
'   entry.
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline to fill and the record to report to.
'
' RETURNS
'   Boolean
'     TRUE when the stage ran to its end; a property that cannot be read is
'     named as unverified, not as a failure.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertReadProp
'
' CALLED FROM
'   - TST_CertCaptureBaseline
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PropIdx             As Long            'Cursor over the properties
    Dim PropValue           As Boolean         'Value read
    Dim Msg                 As String          'Diagnostic from a read or error

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertCaptureAppProps = False

        For PropIdx = 1 To TST_CERT_APP_COUNT

            If TST_CertReadProp(Application, TST_CertAppPropName(PropIdx), _
                PropValue, Msg) Then

                Base.AppKnown(PropIdx) = True
                Base.AppValue(PropIdx) = PropValue

            Else

                TST_CertAddUnverified Rec, _
                    "app." & TST_CertAppPropName(PropIdx) & ".entry", _
                    "Application." & TST_CertAppPropName(PropIdx) & _
                    " could not be read on entry, so it is not compared | " & Msg

            End If

        Next PropIdx

        TST_CertCaptureAppProps = True

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "baseline.appProperties", _
            "the application properties could not be captured on entry | " & Msg

        TST_CertCaptureAppProps = False

        Resume Safe_Exit

End Function


Private Function TST_CertCaptureWindows( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_CertCaptureWindows
'------------------------------------------------------------------------------
' PURPOSE
'   Retain every Window present on entry and record its managed state.
'
' WHY THIS EXISTS
'   A window count cannot tell a closed window that another replaced from an
'   untouched one. Retaining the Window object proves, on exit, that the same
'   window still responds; its hWnd, read from that same object, names the
'   frame that carries its title bar (#45).
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline to fill and the record to report to.
'
' RETURNS
'   Boolean
'     TRUE when the stage ran to its end; a value that cannot be read is named
'     as unverified, not as a failure.
'
' BEHAVIOR
'   - Records, per window: the object, a caption label, hWnd, Headings,
'     WorkbookTabs, Gridlines, the complete GWL_STYLE, and the frame-registry
'     fields (refresh pending, component hidden).
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertTryReadHwnd
'   - TST_CertReadProp
'   - TST_TryGetWindowStyle
'   - UI_InternalTryReadFrameState
'
' CALLED FROM
'   - TST_CertCaptureBaseline
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WinIdx              As Long            'Cursor over the windows
    Dim PropIdx             As Long            'Cursor over the window properties
    Dim TargetWindow        As Object          'Window being captured
    Dim PropValue           As Boolean         'Property value read
    Dim Label               As String          'Diagnostic label of the window
    Dim Msg                 As String          'Diagnostic from a read or error

    Dim HasBaseline         As Boolean         'Registry field, not compared
    Dim HiddenNow           As Boolean         'Registry: component hid the frame
    Dim PendingNow          As Boolean         'Registry: frame refresh owed

#If VBA7 Then
    Dim StyleValue          As LongPtr         'GWL_STYLE read
    Dim OwnedBits           As LongPtr         'Registry field, not compared
    Dim WrittenBits         As LongPtr         'Registry field, not compared
#Else
    Dim StyleValue          As Long            'GWL_STYLE read
    Dim OwnedBits           As Long            'Registry field, not compared
    Dim WrittenBits         As Long            'Registry field, not compared
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertCaptureWindows = False

    'No window, nothing to retain
        If Base.WindowCount < 1 Then
            TST_CertCaptureWindows = True
            GoTo Safe_Exit
        End If

        ReDim Base.Win(1 To Base.WindowCount)
        ReDim Base.WinLabel(1 To Base.WindowCount)
        ReDim Base.WinHwnd(1 To Base.WindowCount)
        ReDim Base.WinHwndKnown(1 To Base.WindowCount)
        ReDim Base.WinPropKnown(1 To Base.WindowCount, 1 To TST_CERT_WINPROP_COUNT)
        ReDim Base.WinPropValue(1 To Base.WindowCount, 1 To TST_CERT_WINPROP_COUNT)
        ReDim Base.WinStyle(1 To Base.WindowCount)
        ReDim Base.WinTitleKnown(1 To Base.WindowCount)
        ReDim Base.WinRegFound(1 To Base.WindowCount)
        ReDim Base.WinRegRefresh(1 To Base.WindowCount)
        ReDim Base.WinRegHidden(1 To Base.WindowCount)

'------------------------------------------------------------------------------
' CAPTURE EACH WINDOW
'------------------------------------------------------------------------------
        For WinIdx = 1 To Base.WindowCount

            Set TargetWindow = Application.Windows(WinIdx)
            Set Base.Win(WinIdx) = TargetWindow

            Label = TST_CertWindowLabel(TargetWindow, WinIdx)
            Base.WinLabel(WinIdx) = Label

        'The frame handle, read from the same object
            If TST_CertTryReadHwnd(TargetWindow, Base.WinHwnd(WinIdx), Msg) Then
                Base.WinHwndKnown(WinIdx) = True
            Else
                TST_CertAddUnverified Rec, "window.hWnd.entry", _
                    "entry " & Label & " hWnd could not be read, so its title " & _
                    "bar and frame state are not compared | " & Msg
            End If

        'The window-level properties
            For PropIdx = 1 To TST_CERT_WINPROP_COUNT

                If TST_CertReadProp(TargetWindow, _
                    TST_CertWindowPropName(PropIdx), PropValue, Msg) Then

                    Base.WinPropKnown(WinIdx, PropIdx) = True
                    Base.WinPropValue(WinIdx, PropIdx) = PropValue

                Else

                    TST_CertAddUnverified Rec, _
                        "window." & TST_CertWindowPropName(PropIdx) & ".entry", _
                        "entry " & Label & " " & _
                        TST_CertWindowPropName(PropIdx) & _
                        " could not be read, so it is not compared | " & Msg

                End If

            Next PropIdx

        'The frame style and the registry state need the handle
            If Base.WinHwndKnown(WinIdx) Then

                If TST_TryGetWindowStyle(Base.WinHwnd(WinIdx), StyleValue, Msg) Then
                    Base.WinTitleKnown(WinIdx) = True
                    Base.WinStyle(WinIdx) = StyleValue
                Else
                    TST_CertAddUnverified Rec, "titleBar.entry", _
                        "entry " & Label & " frame style could not be read, " & _
                        "so its title bar is not compared | " & Msg
                End If

                Base.WinRegFound(WinIdx) = UI_InternalTryReadFrameState( _
                    Base.WinHwnd(WinIdx), OwnedBits, HasBaseline, _
                    HiddenNow, PendingNow, WrittenBits)

                Base.WinRegHidden(WinIdx) = HiddenNow
                Base.WinRegRefresh(WinIdx) = PendingNow

            End If

        Next WinIdx

        TST_CertCaptureWindows = True

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "baseline.windows", _
            "the windows could not be captured on entry | " & Msg

        TST_CertCaptureWindows = False

        Resume Safe_Exit

End Function


Private Function TST_CertCaptureActive( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_CertCaptureActive
'------------------------------------------------------------------------------
' PURPOSE
'   Record which retained window was active on entry.
'
' WHY THIS EXISTS
'   Windows are compared by handle, not with Is: Excel returns a new wrapper
'   from every read of Application.ActiveWindow, so two reads of the same
'   window are not Is-equal on a real host. The handle read from the object,
'   matched to a retained window, is the identity.
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline to fill and the record to report to.
'
' RETURNS
'   Boolean
'     TRUE when the stage ran to its end; an active window that cannot be
'     identified is named as unverified, not as a failure.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertTryReadHwnd
'
' CALLED FROM
'   - TST_CertCaptureBaseline
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WinIdx              As Long            'Cursor over the retained windows
    Dim ActiveWin           As Object          'The active window
    Dim Msg                 As String          'Diagnostic from a read or error

#If VBA7 Then
    Dim ActiveHwnd          As LongPtr         'Handle of the active window
#Else
    Dim ActiveHwnd          As Long            'Handle of the active window
#End If

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertCaptureActive = False

        Set ActiveWin = Application.ActiveWindow

        If ActiveWin Is Nothing Then

            TST_CertAddUnverified Rec, "activeWindow.entry", _
                "no window was active on entry, so the active window is " & _
                "not compared"

        ElseIf Not TST_CertTryReadHwnd(ActiveWin, ActiveHwnd, Msg) Then

            TST_CertAddUnverified Rec, "activeWindow.entry", _
                "the active window's hWnd could not be read on entry, so " & _
                "the active window is not compared | " & Msg

        Else

            Base.ActiveHwnd = ActiveHwnd

            For WinIdx = 1 To Base.WindowCount

                If Base.WinHwndKnown(WinIdx) Then
                    If Base.WinHwnd(WinIdx) = ActiveHwnd Then
                        Base.ActiveIndex = WinIdx
                        Exit For
                    End If
                End If

            Next WinIdx

            If Base.ActiveIndex = 0 Then
                TST_CertAddUnverified Rec, "activeWindow.entry", _
                    "the active window's hWnd is not that of any window in " & _
                    "Application.Windows, so the active window is not compared"
            Else
                Base.ActiveKnown = True
            End If

        End If

        TST_CertCaptureActive = True

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "baseline.activeWindow", _
            "the active window could not be captured on entry | " & Msg

        TST_CertCaptureActive = False

        Resume Safe_Exit

End Function


Private Function TST_CertCaptureRibbon( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_CertCaptureRibbon
'------------------------------------------------------------------------------
' PURPOSE
'   Record the Ribbon of the entry active window, when that window can be
'   identified as its owner.
'
' WHY THIS EXISTS
'   The Ribbon is per window and every mechanism that reads it acts on the
'   active window and accepts no target. A value is therefore attributable only
'   to the window that was active when it was read, and only if that window did
'   not change during the read.
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline to fill and the record to report to.
'
' RETURNS
'   Boolean
'     TRUE when the stage ran to its end; a Ribbon that cannot be attributed
'     is named as unverified, not as a failure.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_TryGetRibbonVisible
'   - TST_CertTryReadHwnd
'
' CALLED FROM
'   - TST_CertCaptureBaseline
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RibbonValue         As Boolean         'Ribbon visibility read
    Dim ActiveWin           As Object          'The active window after the read
    Dim Msg                 As String          'Diagnostic from a read or error

#If VBA7 Then
    Dim HwndAfter           As LongPtr         'Active handle after the read
#Else
    Dim HwndAfter           As Long            'Active handle after the read
#End If

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertCaptureRibbon = False

        If Not Base.ActiveKnown Then

            TST_CertAddUnverified Rec, "ribbon.entry", _
                "no identified window owned the Ribbon on entry, so it is " & _
                "not compared"

        ElseIf Not TST_TryGetRibbonVisible(RibbonValue, Msg) Then

            TST_CertAddUnverified Rec, "ribbon.entry", _
                "the Ribbon could not be read on entry, so it is not " & _
                "compared | " & Msg

        Else

        'The owner must be the same window after the read as before it
            Set ActiveWin = Application.ActiveWindow

            If ActiveWin Is Nothing Then
                Msg = "no window was active after the read"
            ElseIf Not TST_CertTryReadHwnd(ActiveWin, HwndAfter, Msg) Then
                Msg = "the active hWnd could not be read after the read | " & Msg
            ElseIf HwndAfter <> Base.ActiveHwnd Then
                Msg = "the active window changed while the Ribbon was read"
            Else
                Msg = vbNullString
            End If

            If Len(Msg) = 0 Then
                Base.RibbonKnown = True
                Base.RibbonVisible = RibbonValue
            Else
                TST_CertAddUnverified Rec, "ribbon.entry", _
                    "the Ribbon cannot be attributed to the entry active " & _
                    "window, so it is not compared | " & Msg
            End If

        End If

        TST_CertCaptureRibbon = True

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "baseline.ribbon", _
            "the Ribbon could not be captured on entry | " & Msg

        TST_CertCaptureRibbon = False

        Resume Safe_Exit

End Function


Private Function TST_CertEvaluateCleanup( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_CertEvaluateCleanup
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the exit state of the host with the entry baseline, and record
'   every difference as a finding.
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
'   failure. A comparison that fails to run must not look like one that found
'   nothing either, so a stage that raises is itself a finding (#35).
'
' INPUTS
'   Base
'     ByRef. The baseline captured on entry.
'
'   Rec
'     ByRef. The record that receives every finding, every unverified surface
'     named on exit, and the count of comparisons completed.
'
' RETURNS
'   Boolean
'     True  => this call added no finding and the record stored all it was
'              given.
'     False => at least one difference or failed comparison was recorded.
'
' BEHAVIOR
'   - Compares, in a fixed order: the snapshot slot, the workbook and window
'     counts, the application properties, every entry window (identity,
'     properties, title bar, frame registry) and every window that was not
'     there on entry, the active window, and the Ribbon.
'   - Runs every stage, so one run reports every way cleanup failed.
'   - Reports the snapshot check absolutely, because a snapshot is rejected as a
'     precondition rather than preserved.
'   - Changes nothing in the host. Recovery is a separate, test-owned step.
'
' ERROR POLICY
'   - Does not raise. Each stage has its own handler, which records the error
'     as a finding.
'
' DEPENDENCIES
'   - TST_CertCmpSnapshot
'   - TST_CertCmpCounts
'   - TST_CertCmpAppProps
'   - TST_CertCmpWindows
'   - TST_CertCmpActive
'   - TST_CertCmpRibbon
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'   - the cleanup regression cases
'
' UPDATED
'   2026-11-18 - Compares the retained entry windows, every managed property,
'                the title bar and the frame registry, and records findings
'                instead of joining text (#35).
'   2026-08-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FindingsBefore      As Long            'Findings before this call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        FindingsBefore = Rec.FindingCount

'------------------------------------------------------------------------------
' COMPARE AGAINST ENTRY STATE
'------------------------------------------------------------------------------
    'A baseline that is not complete cannot prove anything, so the comparison
    'is replaced by a finding. The snapshot check needs no baseline.
        TST_CertCmpSnapshot Rec

        If Base.Captured Then

            TST_CertCmpCounts Base, Rec
            TST_CertCmpAppProps Base, Rec
            TST_CertCmpWindows Base, Rec
            TST_CertCmpActive Base, Rec
            TST_CertCmpRibbon Base, Rec

        Else

            TST_CertAddFinding Rec, "baseline.unavailable", _
                "the entry baseline was not captured completely, so the " & _
                "exit state cannot be proved equal to it"

        End If

        TST_CertEvaluateCleanup = _
            (Rec.FindingCount = FindingsBefore) And Not Rec.RecordFailed

End Function


Private Sub TST_CertCmpSnapshot( _
    ByRef Rec As tCertRecord)

'
'==============================================================================
' TST_CertCmpSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Report an EXCEL_UI snapshot left in the snapshot slot.
'
' WHY THIS EXISTS
'   A snapshot is rejected on entry rather than preserved, so any snapshot on
'   exit was left by the run itself.
'
' INPUTS
'   Rec
'     ByRef. The record that receives the finding.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - UI_HasExcelUIStateSnapshot
'
' CALLED FROM
'   - TST_CertEvaluateCleanup
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Msg                 As String          'Diagnostic from an error

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertCountCheck Rec

        If UI_HasExcelUIStateSnapshot Then
            TST_CertAddFinding Rec, "snapshot.leftBehind", _
                "an EXCEL_UI snapshot was left behind"
        End If

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "cleanup.snapshot", _
            "the snapshot comparison raised | " & Msg

        Resume Safe_Exit

End Sub


Private Sub TST_CertCmpCounts( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord)

'
'==============================================================================
' TST_CertCmpCounts
'------------------------------------------------------------------------------
' PURPOSE
'   Compare Workbooks.Count and Application.Windows.Count with their entry
'   values.
'
' WHY THIS EXISTS
'   The counts catch what identity cannot name: a workbook or window the run
'   added or removed. They do not catch a substitution that keeps the count,
'   which is why the windows are compared by identity as well.
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline and the record that receives the findings.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - TST_CertEvaluateCleanup
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CountNow            As Long            'Count read on exit
    Dim Msg                 As String          'Diagnostic from an error

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        CountNow = Application.Workbooks.Count
        TST_CertCountCheck Rec

        If CountNow <> Base.BookCount Then
            TST_CertAddFinding Rec, "workbooks.count", _
                "workbook count changed from " & CStr(Base.BookCount) & _
                " to " & CStr(CountNow)
        End If

        CountNow = Application.Windows.Count
        TST_CertCountCheck Rec

        If CountNow <> Base.WindowCount Then
            TST_CertAddFinding Rec, "windows.count", _
                "window count changed from " & CStr(Base.WindowCount) & _
                " to " & CStr(CountNow)
        End If

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "cleanup.counts", _
            "the count comparison raised | " & Msg

        Resume Safe_Exit

End Sub


Private Sub TST_CertCmpAppProps( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord)

'
'==============================================================================
' TST_CertCmpAppProps
'------------------------------------------------------------------------------
' PURPOSE
'   Compare DisplayStatusBar, DisplayScrollBars, DisplayFormulaBar and
'   ScreenUpdating with their entry values.
'
' WHY THIS EXISTS
'   These are the application-level elements the component manages. A property
'   left altered is a finding that names both values, so a real leak can be
'   diagnosed from the evidence file alone.
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline and the record that receives the findings.
'
' BEHAVIOR
'   - A property read on entry and not on exit is a finding. One not read on
'     entry was named as unverified when the baseline was captured.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertReadProp
'
' CALLED FROM
'   - TST_CertEvaluateCleanup
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PropIdx             As Long            'Cursor over the properties
    Dim PropNow             As Boolean         'Value read on exit
    Dim Msg                 As String          'Diagnostic from a read or error

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        For PropIdx = 1 To TST_CERT_APP_COUNT

            If Base.AppKnown(PropIdx) Then

                If TST_CertReadProp(Application, _
                    TST_CertAppPropName(PropIdx), PropNow, Msg) Then

                    TST_CertCountCheck Rec

                    If PropNow <> Base.AppValue(PropIdx) Then
                        TST_CertAddFinding Rec, _
                            "app." & TST_CertAppPropName(PropIdx), _
                            TST_CertAppPropName(PropIdx) & " changed from " & _
                            CStr(Base.AppValue(PropIdx)) & " to " & CStr(PropNow)
                    End If

                Else

                    TST_CertAddFinding Rec, _
                        "app." & TST_CertAppPropName(PropIdx) & ".read", _
                        TST_CertAppPropName(PropIdx) & _
                        " could not be read on exit | " & Msg

                End If

            End If

        Next PropIdx

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "cleanup.appProperties", _
            "the application-property comparison raised | " & Msg

        Resume Safe_Exit

End Sub


Private Sub TST_CertCmpWindows( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord)

'
'==============================================================================
' TST_CertCmpWindows
'------------------------------------------------------------------------------
' PURPOSE
'   Prove every window present on entry is still the same usable window with
'   the same managed state, and report every window that was not there on
'   entry.
'
' WHY THIS EXISTS
'   A count comparison is satisfied by a window that was closed and another
'   opened. Each retained Window object must still respond and must still
'   report the handle it had, and every window now in Application.Windows must
'   be one of them (#35). The handle alone is not an identity, because Windows
'   reissues the handle of a closed window; a live window is therefore matched
'   only through a retained object that still responds.
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline and the record that receives the findings.
'
' BEHAVIOR
'   Per entry window, in entry order:
'     - the retained object responds (window.missing) and reports its entry
'       handle (window.replaced), and the window is listed (window.unlisted);
'     - DisplayHeadings, DisplayWorkbookTabs and DisplayGridlines equal their
'       entry values (window.<property>);
'     - the title bar equals its entry state: caption visibility
'       (titleBar.visible) and the complete owned-bit set (titleBar.ownedBits);
'     - the frame registry reports the same refresh debt (frame.refreshPending)
'       and component ownership (frame.componentHidden).
'   Then every listed window no entry window accounts for is window.unexpected.
'
' ERROR POLICY
'   - Does not raise. A value that cannot be read on exit is a finding. The
'     stage's own failure is cleanup.windows.
'
' DEPENDENCIES
'   - TST_CertTryReadHwnd
'   - TST_CertReadProp
'   - TST_TryGetWindowStyle
'   - UI_InternalTryReadFrameState
'
' CALLED FROM
'   - TST_CertEvaluateCleanup
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ExitCount           As Long            'Application.Windows.Count on exit
    Dim WinIdx              As Long            'Cursor over the entry windows
    Dim LiveIdx             As Long            'Cursor over the listed windows
    Dim PropIdx             As Long            'Cursor over the window properties
    Dim Label               As String          'Diagnostic label of a window
    Dim Msg                 As String          'Diagnostic from a read or error
    Dim PropNow             As Boolean         'Property value read on exit
    Dim EntryVisible        As Boolean         'Caption visible on entry
    Dim ExitVisible         As Boolean         'Caption visible on exit
    Dim RegFound            As Boolean         'Registry slot exists on exit
    Dim HasBaseline         As Boolean         'Registry field, not compared
    Dim HiddenNow           As Boolean         'Registry: component hid the frame
    Dim PendingNow          As Boolean         'Registry: frame refresh owed
    Dim Listed              As Boolean         'Entry window found in the list

    Dim LiveKnown()         As Boolean         'Listed window's hWnd was read
    Dim LiveClaimed()       As Boolean         'Listed window accounted for

#If VBA7 Then
    Dim LiveHwnd()          As LongPtr         'hWnd of each listed window
    Dim HwndNow             As LongPtr         'hWnd a retained window reports
    Dim StyleNow            As LongPtr         'GWL_STYLE read on exit
    Dim OwnedBits           As LongPtr         'Registry field, not compared
    Dim WrittenBits         As LongPtr         'Registry field, not compared
#Else
    Dim LiveHwnd()          As Long            'hWnd of each listed window
    Dim HwndNow             As Long            'hWnd a retained window reports
    Dim StyleNow            As Long            'GWL_STYLE read on exit
    Dim OwnedBits           As Long            'Registry field, not compared
    Dim WrittenBits         As Long            'Registry field, not compared
#End If

    Const PROC As String = "TST_CertCmpWindows"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

    'Fault seam: a stage that raises must be recorded, not swallowed
        If TST_CertTakeHelperFault(TST_CERT_HELPER_FAULT_STAGE) Then
            Err.Raise TEST_CERT_ERR_BASE + 50, PROC, _
                "injected cleanup stage failure"
        End If

'------------------------------------------------------------------------------
' LIST THE WINDOWS NOW OPEN
'------------------------------------------------------------------------------
        ExitCount = Application.Windows.Count

        If ExitCount > 0 Then

            ReDim LiveHwnd(1 To ExitCount)
            ReDim LiveKnown(1 To ExitCount)
            ReDim LiveClaimed(1 To ExitCount)

            For LiveIdx = 1 To ExitCount
                LiveKnown(LiveIdx) = TST_CertTryReadHwnd( _
                    Application.Windows(LiveIdx), LiveHwnd(LiveIdx), Msg)
            Next LiveIdx

        End If

'------------------------------------------------------------------------------
' COMPARE EACH ENTRY WINDOW
'------------------------------------------------------------------------------
        For WinIdx = 1 To Base.WindowCount

            Label = "entry " & Base.WinLabel(WinIdx)

        'The retained object must still respond
            If Base.Win(WinIdx) Is Nothing Then
                Msg = "the retained Window is Nothing"
            ElseIf TST_CertTryReadHwnd(Base.Win(WinIdx), HwndNow, Msg) Then
                Msg = vbNullString
            End If

            TST_CertCountCheck Rec

            If Len(Msg) > 0 Then

                TST_CertAddFinding Rec, "window.missing", _
                    Label & " is no longer usable | " & Msg

            Else

            'The handle it reports must be the one it had
                If Base.WinHwndKnown(WinIdx) Then
                    If HwndNow <> Base.WinHwnd(WinIdx) Then
                        TST_CertAddFinding Rec, "window.replaced", _
                            Label & " now reports hWnd " & CStr(HwndNow) & _
                            ", not " & CStr(Base.WinHwnd(WinIdx))
                    End If
                End If

            'It must be among the windows listed, which it then accounts for
                Listed = False

                For LiveIdx = 1 To ExitCount
                    If LiveKnown(LiveIdx) Then
                        If LiveHwnd(LiveIdx) = HwndNow Then
                            LiveClaimed(LiveIdx) = True
                            Listed = True
                        End If
                    End If
                Next LiveIdx

                If Not Listed Then
                    TST_CertAddFinding Rec, "window.unlisted", _
                        Label & " (hWnd " & CStr(HwndNow) & _
                        ") is not in Application.Windows"
                End If

            'Window-level properties
                For PropIdx = 1 To TST_CERT_WINPROP_COUNT

                    If Base.WinPropKnown(WinIdx, PropIdx) Then

                        If TST_CertReadProp(Base.Win(WinIdx), _
                            TST_CertWindowPropName(PropIdx), PropNow, Msg) Then

                            TST_CertCountCheck Rec

                            If PropNow <> Base.WinPropValue(WinIdx, PropIdx) Then
                                TST_CertAddFinding Rec, _
                                    "window." & TST_CertWindowPropName(PropIdx), _
                                    Label & " " & TST_CertWindowPropName(PropIdx) & _
                                    " changed from " & _
                                    CStr(Base.WinPropValue(WinIdx, PropIdx)) & _
                                    " to " & CStr(PropNow)
                            End If

                        Else

                            TST_CertAddFinding Rec, _
                                "window." & TST_CertWindowPropName(PropIdx) & ".read", _
                                Label & " " & TST_CertWindowPropName(PropIdx) & _
                                " could not be read on exit | " & Msg

                        End If

                    End If

                Next PropIdx

            'Title bar, read from the frame the window reported on entry
                If Base.WinTitleKnown(WinIdx) Then

                    If TST_TryGetWindowStyle(Base.WinHwnd(WinIdx), StyleNow, Msg) Then

                        TST_CertCountCheck Rec

                    'The complete caption mask, not one bit of it
                        EntryVisible = _
                            ((Base.WinStyle(WinIdx) And TST_WS_CAPTION) = TST_WS_CAPTION)
                        ExitVisible = _
                            ((StyleNow And TST_WS_CAPTION) = TST_WS_CAPTION)

                        If EntryVisible <> ExitVisible Then
                            TST_CertAddFinding Rec, "titleBar.visible", _
                                Label & " title bar visibility changed from " & _
                                IIf(EntryVisible, "visible", "hidden") & _
                                " to " & IIf(ExitVisible, "visible", "hidden")
                        End If

                        If (StyleNow And TST_TITLEBAR_OWNED_MASK) <> _
                            (Base.WinStyle(WinIdx) And TST_TITLEBAR_OWNED_MASK) Then

                            TST_CertAddFinding Rec, "titleBar.ownedBits", _
                                Label & " owned frame bits changed from &H" & _
                                Hex$(Base.WinStyle(WinIdx) And TST_TITLEBAR_OWNED_MASK) & _
                                " to &H" & Hex$(StyleNow And TST_TITLEBAR_OWNED_MASK)
                        End If

                    Else

                        TST_CertAddFinding Rec, "titleBar.read", _
                            Label & " frame style could not be read on exit | " & Msg

                    End If

                End If

            'Frame registry: refresh debt and component ownership
                If Base.WinHwndKnown(WinIdx) Then

                    RegFound = UI_InternalTryReadFrameState( _
                        Base.WinHwnd(WinIdx), OwnedBits, HasBaseline, _
                        HiddenNow, PendingNow, WrittenBits)

                    TST_CertCountCheck Rec

                    If PendingNow <> Base.WinRegRefresh(WinIdx) Then
                        TST_CertAddFinding Rec, "frame.refreshPending", _
                            Label & " frame refresh debt changed from " & _
                            CStr(Base.WinRegRefresh(WinIdx)) & " to " & _
                            CStr(PendingNow)
                    End If

                    If HiddenNow <> Base.WinRegHidden(WinIdx) Then
                        TST_CertAddFinding Rec, "frame.componentHidden", _
                            Label & " component ownership of the hidden frame " & _
                            "changed from " & CStr(Base.WinRegHidden(WinIdx)) & _
                            " to " & CStr(HiddenNow)
                    End If

                End If

            End If

        Next WinIdx

'------------------------------------------------------------------------------
' REPORT EVERY WINDOW NO ENTRY WINDOW ACCOUNTS FOR
'------------------------------------------------------------------------------
        For LiveIdx = 1 To ExitCount

            If Not LiveKnown(LiveIdx) Then

                TST_CertAddFinding Rec, "window.unreadable", _
                    TST_CertWindowLabel(Application.Windows(LiveIdx), LiveIdx) & _
                    " in Application.Windows has an hWnd that could not be read"

            ElseIf Not LiveClaimed(LiveIdx) Then

                TST_CertAddFinding Rec, "window.unexpected", _
                    TST_CertWindowLabel(Application.Windows(LiveIdx), LiveIdx) & _
                    " (hWnd " & CStr(LiveHwnd(LiveIdx)) & _
                    ") was not open on entry"

            End If

        Next LiveIdx

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "cleanup.windows", _
            "the window comparison raised | " & Msg

        Resume Safe_Exit

End Sub


Private Sub TST_CertCmpActive( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord)

'
'==============================================================================
' TST_CertCmpActive
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the active window with the one active on entry.
'
' WHY THIS EXISTS
'   The cases activate other windows. A run that leaves a different window
'   active has changed what the user sees and what a Ribbon read means.
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline and the record that receives the finding.
'
' BEHAVIOR
'   - Compares handles, never Window objects with Is, because Excel returns a
'     new wrapper from every read of Application.ActiveWindow.
'   - Changes nothing. TST_CertRestoreActiveWindow is the recovery step.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertTryReadHwnd
'
' CALLED FROM
'   - TST_CertEvaluateCleanup
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ActiveWin           As Object          'The active window now
    Dim Msg                 As String          'Diagnostic from a read or error

#If VBA7 Then
    Dim HwndNow             As LongPtr         'Handle of the active window now
#Else
    Dim HwndNow             As Long            'Handle of the active window now
#End If

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        If Not Base.ActiveKnown Then
            GoTo Safe_Exit
        End If

        Set ActiveWin = Application.ActiveWindow

        TST_CertCountCheck Rec

        If ActiveWin Is Nothing Then

            TST_CertAddFinding Rec, "activeWindow.none", _
                "no window is active on exit; the entry active window was " & _
                Base.WinLabel(Base.ActiveIndex)

        ElseIf Not TST_CertTryReadHwnd(ActiveWin, HwndNow, Msg) Then

            TST_CertAddFinding Rec, "activeWindow.read", _
                "the active window's hWnd could not be read on exit | " & Msg

        ElseIf HwndNow <> Base.ActiveHwnd Then

            TST_CertAddFinding Rec, "activeWindow.changed", _
                "active window changed from entry " & _
                Base.WinLabel(Base.ActiveIndex) & " (hWnd " & _
                CStr(Base.ActiveHwnd) & ") to hWnd " & CStr(HwndNow)

        End If

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "cleanup.activeWindow", _
            "the active-window comparison raised | " & Msg

        Resume Safe_Exit

End Sub


Private Sub TST_CertCmpRibbon( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord)

'
'==============================================================================
' TST_CertCmpRibbon
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the Ribbon of the entry active window with its entry value, only
'   when that window is the active window and so the owner of the read.
'
' WHY THIS EXISTS
'   The Ribbon is per window and every read acts on the active window. A read
'   taken while another window is active describes that other window, so
'   comparing it with the entry value would be a false failure or, worse, a
'   false pass. The comparison is made only when the owner can be identified,
'   and is otherwise named as unverified, never treated as clean.
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline and the record that receives the finding and the
'     unverified entry.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_TryGetRibbonVisible
'   - TST_CertTryReadHwnd
'
' CALLED FROM
'   - TST_CertEvaluateCleanup
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ActiveWin           As Object          'The active window
    Dim RibbonNow           As Boolean         'Ribbon visibility read on exit
    Dim Msg                 As String          'Diagnostic from a read or error
    Dim Owned               As Boolean         'TRUE while the owner is active

#If VBA7 Then
    Dim HwndBefore          As LongPtr         'Active handle before the read
    Dim HwndAfter           As LongPtr         'Active handle after the read
#Else
    Dim HwndBefore          As Long            'Active handle before the read
    Dim HwndAfter           As Long            'Active handle after the read
#End If

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        If Not Base.RibbonKnown Then
            GoTo Safe_Exit
        End If

    'The owner must be the active window before the read
        Set ActiveWin = Application.ActiveWindow

        If ActiveWin Is Nothing Then
            Msg = "no window is active on exit"
        ElseIf Not TST_CertTryReadHwnd(ActiveWin, HwndBefore, Msg) Then
            Msg = "the active hWnd could not be read on exit | " & Msg
        ElseIf HwndBefore <> Base.ActiveHwnd Then
            Msg = "the entry active window is not the active window on exit"
        Else
            Msg = vbNullString
        End If

        If Len(Msg) > 0 Then

            TST_CertAddUnverified Rec, "ribbon.exit", _
                "the Ribbon was not compared on exit, because the Ribbon acts " & _
                "only on the active window and its owner is not active | " & Msg

            GoTo Safe_Exit

        End If

    'Read it, and require the owner to be the same window after the read
        If Not TST_TryGetRibbonVisible(RibbonNow, Msg) Then

            TST_CertAddFinding Rec, "ribbon.read", _
                "the Ribbon could not be read on exit | " & Msg

            GoTo Safe_Exit

        End If

        Set ActiveWin = Application.ActiveWindow

        Owned = False

        If Not ActiveWin Is Nothing Then
            If TST_CertTryReadHwnd(ActiveWin, HwndAfter, Msg) Then
                Owned = (HwndAfter = Base.ActiveHwnd)
            End If
        End If

        If Not Owned Then

            TST_CertAddUnverified Rec, "ribbon.exit", _
                "the Ribbon was not compared on exit, because the active " & _
                "window changed while it was read"

            GoTo Safe_Exit

        End If

        TST_CertCountCheck Rec

        If RibbonNow <> Base.RibbonVisible Then
            TST_CertAddFinding Rec, "ribbon.visible", _
                "Ribbon visibility of the entry active window changed from " & _
                CStr(Base.RibbonVisible) & " to " & CStr(RibbonNow)
        End If

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "cleanup.ribbon", _
            "the Ribbon comparison raised | " & Msg

        Resume Safe_Exit

End Sub


Private Sub TST_CertRestoreActiveWindow( _
    ByRef Base As tCertBaseline, _
    ByRef Rec As tCertRecord)

'
'==============================================================================
' TST_CertRestoreActiveWindow
'------------------------------------------------------------------------------
' PURPOSE
'   Make the entry active window active again when a run left another one
'   active, and report a restoration that did not take effect.
'
' WHY THIS EXISTS
'   The comparison reports the difference and changes nothing. This is the
'   test-owned recovery that follows it, so that a failed run does not also
'   leave the user on another window; its own failure is a finding of its own.
'
' INPUTS
'   Base / Rec
'     ByRef. The baseline and the record that receives the finding.
'
' BEHAVIOR
'   - Does nothing when the entry active window is active, or was not
'     identified.
'   - Activates the retained entry window, then reads the active handle again.
'
' ERROR POLICY
'   - Does not raise. A restoration that fails is activeWindow.restore.
'
' DEPENDENCIES
'   - TST_CertTryReadHwnd
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'   - the cleanup regression cases
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ActiveWin           As Object          'The active window
    Dim Msg                 As String          'Diagnostic from a read or error

#If VBA7 Then
    Dim HwndNow             As LongPtr         'Active handle now
#Else
    Dim HwndNow             As Long            'Active handle now
#End If

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        If Not Base.Captured Then
            GoTo Safe_Exit
        End If

        If Not Base.ActiveKnown Then
            GoTo Safe_Exit
        End If

    'Nothing to do while the entry window is active
        Set ActiveWin = Application.ActiveWindow

        If Not ActiveWin Is Nothing Then
            If TST_CertTryReadHwnd(ActiveWin, HwndNow, Msg) Then
                If HwndNow = Base.ActiveHwnd Then
                    GoTo Safe_Exit
                End If
            End If
        End If

    'Activate the retained entry window, and read the result back
        Base.Win(Base.ActiveIndex).Activate

        Set ActiveWin = Application.ActiveWindow

        If ActiveWin Is Nothing Then
            Msg = "no window is active after the activation"
        ElseIf Not TST_CertTryReadHwnd(ActiveWin, HwndNow, Msg) Then
            Msg = "the active hWnd could not be read after the activation | " & Msg
        ElseIf HwndNow <> Base.ActiveHwnd Then
            Msg = "hWnd " & CStr(HwndNow) & " is active, not " & _
                CStr(Base.ActiveHwnd)
        Else
            Msg = vbNullString
        End If

        If Len(Msg) > 0 Then
            TST_CertAddFinding Rec, "activeWindow.restore", _
                "activating the entry " & Base.WinLabel(Base.ActiveIndex) & _
                " did not make it active again | " & Msg
        End If

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Rec, "activeWindow.restore", _
            "restoring the entry active window raised | " & Msg

        Resume Safe_Exit

End Sub


Private Function TST_CertReadProp( _
    ByVal Target As Object, _
    ByVal PropName As String, _
    ByRef ValueOut As Boolean, _
    ByRef FailMsg As String) _
    As Boolean

'
'==============================================================================
' TST_CertReadProp
'------------------------------------------------------------------------------
' PURPOSE
'   Read one Boolean property for the cleanup baseline or its comparison.
'
' WHY THIS EXISTS
'   Capture and comparison read every property through this one function, so
'   the fault seam that stands in for a failing cleanup helper reaches both,
'   and a failed read is returned, never absorbed.
'
' INPUTS
'   Target / PropName
'     The object and the Boolean property to read.
'
'   ValueOut / FailMsg
'     ByRef. The value on success; the reason on failure.
'
' RETURNS
'   Boolean
'     TRUE when the property was read.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_TryGetBooleanProperty
'
' CALLED FROM
'   - the cleanup capture and comparison stages
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'An armed fault stands in for a helper that failed
        If TST_CertTakeHelperFault(TST_CERT_HELPER_FAULT_READ) Then
            ValueOut = False
            FailMsg = "injected cleanup helper failure"
            TST_CertReadProp = False
        Else
            TST_CertReadProp = TST_TryGetBooleanProperty( _
                Target, PropName, ValueOut, FailMsg)
        End If

End Function


#If VBA7 Then
Private Function TST_CertTryReadHwnd( _
    ByVal TargetWindow As Object, _
    ByRef HwndOut As LongPtr, _
    ByRef FailMsg As String) _
    As Boolean
#Else
Private Function TST_CertTryReadHwnd( _
    ByVal TargetWindow As Object, _
    ByRef HwndOut As Long, _
    ByRef FailMsg As String) _
    As Boolean
#End If

'
'==============================================================================
' TST_CertTryReadHwnd
'------------------------------------------------------------------------------
' PURPOSE
'   Read the hWnd of a Window object, which is also the probe that the object
'   still responds.
'
' WHY THIS EXISTS
'   A closed window's wrapper raises on a read, and a handle read from the
'   object, not from a collection index, names the frame that object owns
'   (#45). One read therefore proves both.
'
' INPUTS
'   TargetWindow
'     The Window to read. May be Nothing.
'
'   HwndOut / FailMsg
'     ByRef. The handle on success, zero otherwise; the reason on failure.
'
' RETURNS
'   Boolean
'     TRUE when a non-zero handle was read.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - the cleanup capture and comparison stages
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertTryReadHwnd = False
        HwndOut = 0
        FailMsg = vbNullString

        If TargetWindow Is Nothing Then
            FailMsg = "the Window is Nothing"
            GoTo Safe_Exit
        End If

        HwndOut = TargetWindow.hWnd

        If HwndOut = 0 Then
            FailMsg = "the Window reports hWnd 0"
            GoTo Safe_Exit
        End If

        TST_CertTryReadHwnd = True

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
        FailMsg = TST_BuildRuntimeErrorText

        HwndOut = 0

        TST_CertTryReadHwnd = False

        Resume Safe_Exit

End Function


Private Function TST_CertWindowLabel( _
    ByVal TargetWindow As Object, _
    ByVal Index As Long) _
    As String

'
'==============================================================================
' TST_CertWindowLabel
'------------------------------------------------------------------------------
' PURPOSE
'   Name a window for a finding.
'
' WHY THIS EXISTS
'   A finding must say which window it means. The caption is diagnostic only:
'   it never takes part in an identity comparison, and a caption that cannot be
'   read leaves the label with its index alone.
'
' INPUTS
'   TargetWindow / Index
'     The window and its position in Application.Windows.
'
' RETURNS
'   String
'     "window <index> [<caption>]".
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - the cleanup capture and comparison stages
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CaptionText         As String          'Caption, when it can be read

'------------------------------------------------------------------------------
' READ CAPTION
'------------------------------------------------------------------------------
    'Diagnostic text only: nothing is decided from it, so a miss is harmless
        On Error Resume Next

        CaptionText = TargetWindow.Caption

        On Error GoTo 0

        TST_CertWindowLabel = "window " & CStr(Index) & " [" & CaptionText & "]"

End Function


Private Function TST_CertAppPropName( _
    ByVal Index As Long) _
    As String

'
'==============================================================================
' TST_CertAppPropName
'------------------------------------------------------------------------------
' PURPOSE
'   Name the application-level property at a baseline index.
'
' WHY THIS EXISTS
'   The baseline stores the four properties in arrays, and the order is the
'   order findings are reported in.
'
' INPUTS
'   Index
'     1 DisplayStatusBar, 2 DisplayScrollBars, 3 DisplayFormulaBar,
'     4 ScreenUpdating.
'
' RETURNS
'   String
'     The property name; empty for an index outside 1 to 4.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - the cleanup capture and comparison stages
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        Select Case Index

            Case 1
                TST_CertAppPropName = "DisplayStatusBar"

            Case 2
                TST_CertAppPropName = "DisplayScrollBars"

            Case 3
                TST_CertAppPropName = "DisplayFormulaBar"

            Case 4
                TST_CertAppPropName = "ScreenUpdating"

        End Select

End Function


Private Function TST_CertWindowPropName( _
    ByVal Index As Long) _
    As String

'
'==============================================================================
' TST_CertWindowPropName
'------------------------------------------------------------------------------
' PURPOSE
'   Name the window-level property at a baseline index.
'
' INPUTS
'   Index
'     1 DisplayHeadings, 2 DisplayWorkbookTabs, 3 DisplayGridlines.
'
' RETURNS
'   String
'     The property name; empty for an index outside 1 to 3.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - the cleanup capture and comparison stages
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        Select Case Index

            Case 1
                TST_CertWindowPropName = "DisplayHeadings"

            Case 2
                TST_CertWindowPropName = "DisplayWorkbookTabs"

            Case 3
                TST_CertWindowPropName = "DisplayGridlines"

        End Select

End Function


Private Sub TST_CertInjectHelperFault( _
    ByVal FaultKind As Long)

'
'==============================================================================
' TST_CertInjectHelperFault
'------------------------------------------------------------------------------
' PURPOSE
'   Arm a one-shot fault in the cleanup helpers.
'
' WHY THIS EXISTS
'   A cleanup helper that fails must produce a finding, not a clean verdict.
'   Excel cannot be made to refuse a read on demand, so the path is unreachable
'   without a seam.
'
' INPUTS
'   FaultKind
'     TST_CERT_HELPER_FAULT_NONE disarms. _READ fails the next property read
'     the baseline makes. _STAGE makes the next window comparison raise.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Cannot raise.
'
' CALLED FROM
'   - TST_Case_CertCleanupHelperFailureIsFinding
'
' NOTES
'   Private to this module on purpose: the seam shares no state with the
'   production modules, and only the cases below arm it.
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        m_CertHelperFault = FaultKind

End Sub


Private Function TST_CertTakeHelperFault( _
    ByVal FaultKind As Long) _
    As Boolean

'
'==============================================================================
' TST_CertTakeHelperFault
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether the armed helper fault is this kind, and disarm it in the
'   same step when it is.
'
' INPUTS
'   FaultKind
'     TST_CERT_HELPER_FAULT_READ or TST_CERT_HELPER_FAULT_STAGE.
'
' RETURNS
'   Boolean
'     TRUE when the fault fires; it is then disarmed.
'
' ERROR POLICY
'   - Cannot raise.
'
' CALLED FROM
'   - TST_CertReadProp
'   - TST_CertCmpWindows
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        If m_CertHelperFault = FaultKind Then
            m_CertHelperFault = TST_CERT_HELPER_FAULT_NONE
            TST_CertTakeHelperFault = True
        End If

End Function


Private Sub TST_CertApplyRunFault()

'
'==============================================================================
' TST_CertApplyRunFault
'------------------------------------------------------------------------------
' PURPOSE
'   Alter the host in place of the mandatory units, for the run fault that
'   forces a cleanup failure.
'
' WHY THIS EXISTS
'   The outer certification call must fail when cleanup fails, with its
'   diagnostic fields intact (#35). A real cleanup failure cannot be provoked
'   on demand, so this inverts Application.DisplayStatusBar and leaves it for
'   the cleanup comparison to find. The caller that armed the fault restores it.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Raises what the write raises, into the certification handler.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        TST_Log "TST_CertApplyRunFault", "FAULT", _
            "No unit runs; DisplayStatusBar is inverted before cleanup is verified"

        Application.DisplayStatusBar = Not Application.DisplayStatusBar

End Sub


Private Sub TST_CertNoteHostFailure( _
    ByVal CallerProc As String, _
    ByVal Msg As String, _
    ByRef HasFailure As Boolean, _
    ByRef FailNumber As Long, _
    ByRef FailSource As String, _
    ByRef FailDescription As String)

'
'==============================================================================
' TST_CertNoteHostFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Make a host state that a case could not put back part of its failure.
'
' WHY THIS EXISTS
'   The cleanup cases alter the host on purpose and restore it themselves. A
'   restoration that did not take effect has left Excel altered, so it fails
'   the case, after any failure the case already holds.
'
' INPUTS
'   CallerProc / Msg
'     The case, and what could not be put back.
'
'   HasFailure / FailNumber / FailSource / FailDescription
'     ByRef. The case's saved failure.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - the cleanup regression cases
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        If Not HasFailure Then
            HasFailure = True
            FailNumber = TEST_ERR_BASE + 83
            FailSource = CallerProc & ".HostRestored"
            FailDescription = Msg
        Else
            FailDescription = FailDescription & " | " & Msg
        End If

End Sub


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
'   E  forced cleanup failure  - the outer certification call fails with its
'                                number, source, verdict and finding intact,
'                                the finding is recorded, and both evidence
'                                documents list it
'
' ERROR POLICY
'   - Raises on assertion failure
'   - Disarms any injection it armed before returning
'
' DEPENDENCIES
'   - TST_InjectPostCapture
'   - TST_CertInjectRunFault
'   - Test_EXCEL_UI_RunCertificationSelfTest
'   - Test_EXCEL_UI_RunSnapshotIdentity
'   - Test_EXCEL_UI_RunReleaseCertification
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
'   Case E runs the real certification call with no unit and the status bar
'   inverted, writes no evidence file, and restores the status bar itself.
'
' UPDATED
'   2026-11-18 - Added case E, the forced cleanup failure (#35).
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

    Dim EntryStatusBar      As Boolean         'DisplayStatusBar before case E
    Dim StatusBarAtRisk     As Boolean         'TRUE while case E may have inverted it
    Dim ForcedUnitsFailed   As Long            'Case E: composed failed-unit count
    Dim ForcedComplete      As Boolean         'Case E: composed completeness
    Dim ForcedPassed        As Boolean         'Case E: composed verdict
    Dim ForcedVerdict       As String          'Case E: composed RESULT line
    Dim ForcedDetail        As String          'Case E: composed cleanup detail
    Dim ForcedRecordingDetail As String          'Case E: composed recording failure
    Dim ForcedJson          As String          'Case E: composed JSON evidence
    Dim ForcedReport        As String          'Case E: composed text evidence

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
' CASE E - A FORCED CLEANUP FAILURE FAILS THE OUTER CERTIFICATION
'------------------------------------------------------------------------------
    'The run fault makes certification execute no unit and invert the status bar
    'after its entry baseline is taken. Everything after that is the real run:
    'its cleanup comparison, its evidence and its raise. This proves that a
    'cleanup failure reaches the caller of Test_EXCEL_UI_RunReleaseCertification
    'with its fields intact, rather than being logged and swallowed.
        EntryStatusBar = Application.DisplayStatusBar
        StatusBarAtRisk = True

        TST_CertInjectRunFault TEST_CERT_RUN_FAULT_CLEANUP

        On Error Resume Next

            Test_EXCEL_UI_RunReleaseCertification

            RaisedNumber = Err.Number
            RaisedSource = Err.Source
            RaisedDescription = Err.Description

            Err.Clear

        On Error GoTo Err_Handler

    'Put the host back before anything is asserted, so a failed assertion leaves
    'Excel as it was found
        TST_CertInjectRunFault TEST_CERT_RUN_FAULT_NONE

        Application.DisplayStatusBar = EntryStatusBar
        StatusBarAtRisk = False

        TST_AssertTrue _
            (RaisedNumber = TEST_CERT_ERR_BASE + 3), _
            PROC & ".ForcedCleanup.ErrorNumberPreserved"

        TST_AssertTrue _
            (RaisedSource = "Test_EXCEL_UI_RunReleaseCertification"), _
            PROC & ".ForcedCleanup.ErrorSourcePreserved"

    'The diagnostic fields carry the verdict and the finding, not just a number
        TST_AssertTrue _
            (InStr(1, RaisedDescription, "cleanup=FAILED", vbBinaryCompare) > 0), _
            PROC & ".ForcedCleanup.DescriptionCarriesTheVerdict"

        TST_AssertTrue _
            (InStr(1, RaisedDescription, "DisplayStatusBar changed from", _
                vbBinaryCompare) > 0), _
            PROC & ".ForcedCleanup.DescriptionCarriesTheFinding"

        TST_AssertBooleanEquals False, m_CertActive, _
            PROC & ".ForcedCleanup.RunClosed"

    'The findings were recorded, and nothing was lost in recording them
        TST_AssertTrue _
            (m_CertRecord.FindingCount >= 1) And Not m_CertRecord.RecordFailed, _
            PROC & ".ForcedCleanup.FindingRecorded"

        TST_AssertTrue _
            (InStr(1, TST_CertEntryCodes(m_CertRecord, TEST_CERT_ENTRY_FINDING), _
                "|app.DisplayStatusBar|", vbBinaryCompare) > 0), _
            PROC & ".ForcedCleanup.FindingNamesTheProperty"

    'The fault was one-shot: the call consumed it
        TST_AssertTrue _
            (TST_CertTakeRunFault() = TEST_CERT_RUN_FAULT_NONE), _
            PROC & ".ForcedCleanup.FaultConsumed"

    'The evidence the run composed, rebuilt from the record it left behind
        TST_CertComposeEvidence _
            Rec:=m_CertRecord, _
            CleanupOK:=True, _
            CleanupDetail:=ForcedDetail, _
            UnitsFailed:=ForcedUnitsFailed, _
            Complete:=ForcedComplete, _
            Passed:=ForcedPassed, _
            Verdict:=ForcedVerdict, _
            RecordingDetail:=ForcedRecordingDetail, _
            JsonText:=ForcedJson, _
            ReportText:=ForcedReport

        TST_AssertBooleanEquals False, ForcedPassed, _
            PROC & ".ForcedCleanup.Evidence.NotPassed"

        TST_AssertTrue _
            (InStr(1, ForcedVerdict, "RESULT: FAIL | ", vbBinaryCompare) = 1) And _
            (InStr(1, ForcedVerdict, " cleanup=FAILED", vbBinaryCompare) > 0), _
            PROC & ".ForcedCleanup.Evidence.Verdict"

        TST_AssertTrue _
            (InStr(1, ForcedJson, """cleanup"":""FAILED""", vbBinaryCompare) > 0), _
            PROC & ".ForcedCleanup.Evidence.JsonCleanupFailed"

        TST_AssertTrue _
            (InStr(1, ForcedJson, """code"":""app.DisplayStatusBar""", _
                vbBinaryCompare) > 0), _
            PROC & ".ForcedCleanup.Evidence.JsonListsTheFinding"

        TST_AssertTrue _
            (InStr(1, ForcedReport, vbNewLine & "  CLEANUP  app.DisplayStatusBar | ", _
                vbBinaryCompare) > 0), _
            PROC & ".ForcedCleanup.Evidence.ReportListsTheFinding"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Runner-owned snapshots were released on every exit, every " & _
            "failure survived cleanup, and a forced cleanup failure failed " & _
            "the certification call"

Safe_Exit:
    'Disarm whatever this run armed. A fault left armed would fire inside an
    'unrelated runner later in the same session. Case E inverted the status bar,
    'so it is put back here on every path.
        On Error Resume Next
            TST_InjectPostCapture TEST_FAULT_NONE
            TST_CertInjectRunFault TEST_CERT_RUN_FAULT_NONE

            If StatusBarAtRisk Then
                Application.DisplayStatusBar = EntryStatusBar
            End If
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
            TST_CertInjectRunFault TEST_CERT_RUN_FAULT_NONE

            If StatusBarAtRisk Then
                Application.DisplayStatusBar = EntryStatusBar
            End If
        On Error GoTo 0

        Err.Raise RaisedNumber, RaisedSource, RaisedDescription

End Sub


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


Public Function TST_RunRunnerWithLog( _
    ByVal RunnerName As String, _
    ByVal LogPath As String) _
    As String

'
'==============================================================================
' TST_RunRunnerWithLog
'------------------------------------------------------------------------------
' PURPOSE
'   Run one public runner with the log sink set, and return its verdict as
'   one line instead of raising
'
' WHY THIS EXISTS
'   tools\Check-EXCEL_UI.ps1 runs the runners through Application.Run. A
'   script cannot read the Immediate Window, and a runner that raises into
'   Application.Run stops Excel on a run-time error message that someone must
'   dismiss. Through this function every line the runner logs also reaches
'   LogPath, and its outcome comes back as a value
'
' INPUTS
'   RunnerName
'     The public runner to run. Only the runners listed under RUN are
'     accepted
'
'   LogPath
'     The log file to append to. Empty runs the runner with no sink
'
' RETURNS
'   String
'     One line: "RESULT: PASS" when the runner returned, or
'     "RESULT: FAIL | Error <n> | Source: <source> | <description>" when it
'     raised. When a line could not be written to LogPath,
'     " | LOG WRITE FAILED: <reason>" follows
'
' BEHAVIOR
'   - Logs a START line, runs the runner, then prints the verdict line and
'     writes it to the log as its last line, and stops the sink
'   - What the runner itself prints is unchanged; only its error is caught
'
' ERROR POLICY
'   - Does not raise; an unknown runner name gives a FAIL verdict
'   - A log write failure never changes the runner's PASS or FAIL
'
' NOTES
'   - Test_EXCEL_UI_RunReleaseCertification is not accepted:
'     tools\Certify-EXCEL_UI.ps1 runs it and collects its evidence
'
' UPDATED
'   2026-11-16 - Added for the scripted Excel checks (#197).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Verdict             As String          'The verdict line
    Dim SinkFailure         As String          'First log write failure, if any
    Dim RaisedNumber        As Long            'Error the runner raised
    Dim RaisedSource        As String          'Source of that error
    Dim RaisedDescription   As String          'Text of that error

    Const PROC As String = "TST_RunRunnerWithLog"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_LogSinkStart LogPath

        TST_Log PROC, "START", RunnerName & " | " & _
            Format$(Now, "yyyy-mm-dd hh:nn:ss")

        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        Select Case RunnerName

            Case "Test_EXCEL_UI_RunAll"
                Test_EXCEL_UI_RunAll

            Case "Test_EXCEL_UI_RunCore"
                Test_EXCEL_UI_RunCore

            Case "Test_EXCEL_UI_RunTitleBarOnly"
                Test_EXCEL_UI_RunTitleBarOnly

            Case "Test_EXCEL_UI_RunSnapshotIdentity"
                Test_EXCEL_UI_RunSnapshotIdentity

            Case "Test_EXCEL_UI_RunTitleBarSdiIdentity"
                Test_EXCEL_UI_RunTitleBarSdiIdentity

            Case "Test_EXCEL_UI_RunCertificationSelfTest"
                Test_EXCEL_UI_RunCertificationSelfTest

            Case "Test_EXCEL_UI_RunOwnershipCleanupChecks"
                Test_EXCEL_UI_RunOwnershipCleanupChecks

            Case "Test_EXCEL_UI_RunRibbonSdiProbe"
                Test_EXCEL_UI_RunRibbonSdiProbe

            Case Else
                Err.Raise _
                    TEST_CERT_ERR_BASE + 40, _
                    PROC, _
                    "unknown runner: " & RunnerName

        End Select

        Verdict = "RESULT: PASS"

'------------------------------------------------------------------------------
' REPORT VERDICT
'------------------------------------------------------------------------------
Safe_Exit:
    'Every path arrives here with the handler disabled or released
        On Error GoTo 0

    'One line, whatever line breaks the error text held
        Verdict = Replace(Replace(Replace(Verdict, vbCrLf, " / "), vbCr, " / "), _
            vbLf, " / ")

        Debug.Print Verdict
        TST_LogSinkWriteLine Verdict

    'A failed write is reported, never raised
        SinkFailure = TST_LogSinkStop()

        If Len(SinkFailure) > 0 Then
            Verdict = Verdict & " | LOG WRITE FAILED: " & _
                Replace(Replace(SinkFailure, vbCr, " "), vbLf, " ")
        End If

        TST_RunRunnerWithLog = Verdict

        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Capture before anything else touches Err
        RaisedNumber = Err.Number
        RaisedSource = Err.Source
        RaisedDescription = Err.Description

        Verdict = "RESULT: FAIL | Error " & CStr(RaisedNumber) & _
            " | Source: " & RaisedSource & _
            " | " & RaisedDescription

        Resume Safe_Exit

End Function


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
'   - Attempts to restore the original UI state at the end, and fails the pack
'     when a restoration step failed or did not read back as written
'
' ERROR POLICY
'   - Raises after restoration on assertion failure, unexpected error or an
'     incomplete restoration
'
' DEPENDENCIES
'   - TST_SnapshotState
'   - TST_RestoreState
'   - TST_RecordRestoreFailure
'   - regression case routines
'   - TST_Log
'
' NOTES
'   - This pack is the body of the RegressionPack certification unit, so a case
'     reachable only from the title-bar-only pack is absent from the evidence a
'     release is tagged on. Every frame-state case is registered here as well.
'
' UPDATED
'   2026-11-18 - Registered the cleanup cases and failed the pack when the UI
'                state restoration did not complete (#35).
'   2026-11-16 - Registered the certification-record case (#50).
'   2026-11-14 - Registered the first-allocation result-contract case (#38).
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

    Dim RestoreOK           As Boolean         'TRUE once every UI state step completed
    Dim RestoreFindings     As tCertRecord     'Steps the restoration could not complete

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

    'Verify a failed first list allocation leaves a count and an Empty list
        TST_Case_FirstFailureAllocationKeepsContract

    'Verify cleanup is judged against the entry state, not an assumed ideal
        TST_Case_CertificationCleanupUsesBaseline

    'Verify an untouched host matches its own baseline, and names what it
    'cannot compare
        TST_Case_CertBaselineMatchesUntouchedHost

    'Verify a leaked window, and a window replaced at an unchanged count, are
    'reported
        TST_Case_CertCleanupDetectsLeakedAndReplacedWindows

    'Verify an altered application or window property is reported
        TST_Case_CertCleanupDetectsPropertyDrift

    'Verify a title-bar, owned-bit or frame-registry mismatch is reported
        TST_Case_CertCleanupDetectsTitleBarMismatch

    'Verify a failing cleanup helper or stage is a finding, not a clean verdict
        TST_Case_CertCleanupHelperFailureIsFinding

    'Verify the residue of a failed quiet-update readback is reported before
    'it is recovered
        TST_Case_CertCleanupReportsQuietUpdateResidue

    'Verify a certification record that cannot be stored fails closed
        TST_Case_CertRecordFailsClosed

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
            TST_Case_TitleBarSameHwndGenerationReuse
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
    'Attempt to restore the original pre-test UI state. Both results start
    'FALSE, so a step that raised past this boundary counts as failed.
        RestoreOK = False
        FrameRestored = False

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

        'Put the entry frame's exact style back through the harness. The
        'restore above goes through the API under test, and a defect there
        'must not be able to leave the host with a hidden title bar.
            FrameRestored = TST_RestoreTitleBarStyle( _
                EntryHwnd, EntryFrameStyle, EntryFrameCaptured, FrameRestoreMsg)
        On Error GoTo 0

    'Restore ScreenUpdating before leaving the harness
        Application.ScreenUpdating = OldScreenUpdating

    'A restoration step that did not complete fails the pack, after any case
    'failure
        TST_RecordRestoreFailure _
            RestoreOK, RestoreFindings, CallerProc, _
            FailNumber, FailSource, FailDescription

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


'
'------------------------------------------------------------------------------
'
'                           REGRESSION CASES
'
'------------------------------------------------------------------------------
'


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
'   The case drives TST_CertEvaluateCleanup directly, on a baseline and a record
'   of its own. Reaching the same decision through
'   Test_EXCEL_UI_RunReleaseCertification would require a full destructive run
'   to assert one comparison, and could not run from inside the pack at all.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Suppresses screen updating and asserts a matching baseline is clean.
'   - Asserts a mismatched baseline is reported as a finding naming both
'     values.
'   - Restores the entry value on every path.
'
' ERROR POLICY
'   - Raises on assertion failure, after restoration.
'
' UPDATED
'   2026-11-18 - Drives the retained-window baseline and records findings (#35).
'   2026-08-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EntryScreenUpdating As Boolean         'Value to restore on exit
    Dim Base                As tCertBaseline   'Baseline under test
    Dim Rec                 As tCertRecord     'Record under test
    Dim CleanupOK           As Boolean         'Result under test

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

        EntryScreenUpdating = Application.ScreenUpdating

'------------------------------------------------------------------------------
' SUPPRESSED STATE MATCHING ITS BASELINE IS CLEAN
'------------------------------------------------------------------------------
    'This is the case the previous implementation failed: screen updating is
    'suppressed, and that is exactly what the baseline says it should be.
        Application.ScreenUpdating = False

        TST_AssertBooleanEquals _
            True, TST_CertCaptureBaseline(Base, Rec), _
            "CertificationCleanup.BaselineCaptured"

        CleanupOK = TST_CertEvaluateCleanup(Base, Rec)

        TST_AssertBooleanEquals _
            True, CleanupOK, "CertificationCleanup.SuppressedMatchingBaselineIsClean"

        TST_AssertTrue _
            (Rec.FindingCount = 0), _
            "CertificationCleanup.NoFindingWhenClean"

'------------------------------------------------------------------------------
' A GENUINE DIFFERENCE IS STILL REPORTED
'------------------------------------------------------------------------------
    'Same host state, opposite baseline: the run would have changed it, so this
    'must fail. Loosening the check must not have disabled it.
        TST_AssertTrue _
            Base.AppKnown(TST_CERT_APP_SCREEN), _
            "CertificationCleanup.ScreenUpdatingWasRead"

        Base.AppValue(TST_CERT_APP_SCREEN) = True

        TST_AssertBooleanEquals True, TST_CertResetRecord(Rec), _
            "CertificationCleanup.RecordReset"

        CleanupOK = TST_CertEvaluateCleanup(Base, Rec)

        TST_AssertBooleanEquals _
            False, CleanupOK, "CertificationCleanup.GenuineDifferenceReported"

        TST_AssertTrue _
            (InStr(1, TST_CertEntryCodes(Rec, TEST_CERT_ENTRY_FINDING), _
                "|app.ScreenUpdating|", vbBinaryCompare) > 0), _
            "CertificationCleanup.FindingNamesTheProperty"

    'The detail must name both values, or a real leak cannot be diagnosed from
    'the evidence file alone
        TST_AssertTrue _
            (InStr(1, TST_CertEntrySummary(Rec, TEST_CERT_ENTRY_FINDING), _
                "ScreenUpdating changed from True to False", vbBinaryCompare) > 0), _
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

        If Application.ScreenUpdating <> EntryScreenUpdating Then
            TST_CertNoteHostFailure PROC, _
                "ScreenUpdating could not be restored to " & _
                CStr(EntryScreenUpdating), _
                HasFailure, FailNumber, FailSource, FailDescription
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
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        HasFailure = True

        Resume Safe_Exit

End Sub


Private Sub TST_Case_CertBaselineMatchesUntouchedHost()

'
'==============================================================================
' TST_Case_CertBaselineMatchesUntouchedHost
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a baseline captured from the host and compared at once with the
'   host finds nothing, compared something, and names what it cannot compare.
'
' WHY THIS EXISTS
'   A comparison that always reported a difference would make every other
'   cleanup case pass for the wrong reason, and one that compared nothing would
'   make cleanup=OK mean nothing. This case holds both ends: a clean result, a
'   count of comparisons made, and the unverified surfaces stated in both
'   evidence documents.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   1  capture     every window is retained, the counts equal the host's, and
'                  the four application properties and the active window are
'                  known
'   2  compare     no finding, and several comparisons completed
'   3  evidence    the ribbon and frame-registry scopes are listed as
'                  unverified, in the JSON and in the text report, and do not
'                  fail cleanup
'
' ERROR POLICY
'   - Raises on assertion failure. The case changes nothing in the host.
'
' DEPENDENCIES
'   - TST_CertCaptureBaseline
'   - TST_CertEvaluateCleanup
'   - TST_CertComposeEvidence
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   The baseline and the record are local. The record of the certification run
'   executing this case is not touched.
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Base                As tCertBaseline   'Baseline under test
    Dim Rec                 As tCertRecord     'Record under test
    Dim PropIdx             As Long            'Cursor over the application properties
    Dim KnownCount          As Long            'Application properties read

    Dim UnitsFailed         As Long            'Composed failed-unit count
    Dim Complete            As Boolean         'Composed completeness
    Dim Passed              As Boolean         'Composed verdict
    Dim Verdict             As String          'Composed RESULT line
    Dim CleanupDetail       As String          'Composed cleanup detail
    Dim RecordingDetail     As String          'Composed recording failure
    Dim JsonText            As String          'Composed JSON evidence
    Dim ReportText          As String          'Composed text evidence

    Const PROC As String = "TST_Case_CertBaselineMatchesUntouchedHost"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log PROC, "START", _
            "Validating that an untouched host matches its own baseline"

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertCaptureBaseline(Base, Rec), _
            PROC & ".Captured"

        TST_AssertTrue (Base.BookCount = Application.Workbooks.Count), _
            PROC & ".WorkbookCount"

        TST_AssertTrue (Base.WindowCount = Application.Windows.Count), _
            PROC & ".WindowCount"

    'Every window is retained by object, so the count is not all that is kept
        TST_AssertTrue (Base.WindowCount >= 1), PROC & ".AWindowIsOpen"

        TST_AssertTrue (Not Base.Win(1) Is Nothing), PROC & ".WindowRetained"

        For PropIdx = 1 To TST_CERT_APP_COUNT
            If Base.AppKnown(PropIdx) Then
                KnownCount = KnownCount + 1
            End If
        Next PropIdx

        TST_AssertTrue (KnownCount = TST_CERT_APP_COUNT), _
            PROC & ".EveryApplicationPropertyRead"

        TST_AssertBooleanEquals True, Base.ActiveKnown, PROC & ".ActiveWindowKnown"

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertEvaluateCleanup(Base, Rec), _
            PROC & ".Clean"

        TST_AssertTrue (Rec.FindingCount = 0), PROC & ".NoFinding"

    'cleanup=OK is only as strong as what was compared
        TST_AssertTrue (Rec.CheckCount >= 6), PROC & ".ComparisonsWereMade"

'------------------------------------------------------------------------------
' WHAT CANNOT BE COMPARED IS NAMED
'------------------------------------------------------------------------------
        TST_AssertTrue _
            (InStr(1, TST_CertEntryCodes(Rec, TEST_CERT_ENTRY_UNVERIFIED), _
                "|ribbon.scope|", vbBinaryCompare) > 0), _
            PROC & ".RibbonScopeNamed"

        TST_AssertTrue _
            (InStr(1, TST_CertEntryCodes(Rec, TEST_CERT_ENTRY_UNVERIFIED), _
                "|frameRegistry.scope|", vbBinaryCompare) > 0), _
            PROC & ".FrameRegistryScopeNamed"

    'An unverified surface does not fail cleanup, and is in both documents
        TST_CertComposeEvidence _
            Rec:=Rec, _
            CleanupOK:=True, _
            CleanupDetail:=CleanupDetail, _
            UnitsFailed:=UnitsFailed, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            RecordingDetail:=RecordingDetail, _
            JsonText:=JsonText, _
            ReportText:=ReportText

        TST_AssertBooleanEquals True, Passed, PROC & ".Evidence.Passed"

        TST_AssertTrue _
            (InStr(1, Verdict, " cleanup=OK", vbBinaryCompare) > 0), _
            PROC & ".Evidence.CleanupOK"

        TST_AssertTrue _
            (InStr(1, JsonText, """cleanupFindings"":[]", vbBinaryCompare) > 0), _
            PROC & ".Evidence.JsonNoFindings"

        TST_AssertTrue _
            (InStr(1, JsonText, """code"":""ribbon.scope""", vbBinaryCompare) > 0), _
            PROC & ".Evidence.JsonListsUnverified"

        TST_AssertTrue _
            (InStr(1, JsonText, """cleanupChecked"":0", vbBinaryCompare) = 0), _
            PROC & ".Evidence.JsonCountsComparisons"

        TST_AssertTrue _
            (InStr(1, ReportText, vbNewLine & "  UNVERIFIED  ribbon.scope | ", _
                vbBinaryCompare) > 0), _
            PROC & ".Evidence.ReportListsUnverified"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "An untouched host matched its baseline, and the evidence named what it could not compare"

End Sub


Private Sub TST_Case_CertCleanupDetectsLeakedAndReplacedWindows()

'
'==============================================================================
' TST_Case_CertCleanupDetectsLeakedAndReplacedWindows
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that cleanup reports a window the run leaked, and a window that
'   another replaced while the count stayed the same.
'
' WHY THIS EXISTS
'   A leaked window created by ThisWorkbook.NewWindow leaves Workbooks.Count
'   unchanged, and a closed window replaced by another leaves
'   Application.Windows.Count unchanged. Neither was visible to a comparison of
'   counts and a single anchor window (#35).
'
' RETURNS
'   None.
'
' BEHAVIOR
'   A  leaked window     a baseline, a new window, a comparison: the window
'                        count, the unexpected window and the changed active
'                        window are reported, and no entry window is missing
'   B  replaced window   a baseline that holds a temporary window, that window
'                        closed and another opened: the closed one is missing
'                        and the new one unexpected, with no count finding
'   C  recovery          the temporary windows are closed, the entry window is
'                        active again, and the host matches its baseline
'
' ERROR POLICY
'   - Raises on assertion failure, after closing every window it opened and
'     making the entry window active again.
'
' DEPENDENCIES
'   - TST_CertCaptureBaseline
'   - TST_CertEvaluateCleanup
'   - TST_CertRestoreActiveWindow
'   - TST_SafeCloseWindow
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   Every baseline and record is local. Windows are matched through retained
'   objects, so the case holds if Windows reissues the handle of the window it
'   closed to the one it opens next.
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EntryBase           As tCertBaseline   'Baseline of the host as found
    Dim ReplaceBase         As tCertBaseline   'Baseline that holds a temporary window
    Dim EntryRec            As tCertRecord     'Record of the entry capture
    Dim LeakRec             As tCertRecord     'Findings of scenario A
    Dim ReplaceRec          As tCertRecord     'Findings of scenario B
    Dim RecoverRec          As tCertRecord     'Findings of the recovery
    Dim CleanRec            As tCertRecord     'Findings after recovery
    Dim ScratchRec          As tCertRecord     'Record of the baseline in scenario B
    Dim LeakedWindow        As Window          'Window scenario A leaks
    Dim FirstTemp           As Window          'Window scenario B puts in the baseline
    Dim SecondTemp          As Window          'Window scenario B replaces it with
    Dim LeakCodes           As String          'Finding identifiers of scenario A
    Dim ReplaceCodes        As String          'Finding identifiers of scenario B
    Dim EntryCaptured       As Boolean         'TRUE once the entry baseline exists

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_CertCleanupDetectsLeakedAndReplacedWindows"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a leaked or replaced window is reported"

        EntryCaptured = TST_CertCaptureBaseline(EntryBase, EntryRec)

        TST_AssertBooleanEquals True, EntryCaptured, PROC & ".EntryCaptured"

        TST_AssertBooleanEquals True, EntryBase.ActiveKnown, _
            PROC & ".EntryActiveWindowKnown"

'------------------------------------------------------------------------------
' A - A LEAKED WINDOW
'------------------------------------------------------------------------------
        Set LeakedWindow = ThisWorkbook.NewWindow

        If LeakedWindow Is Nothing Then
            Err.Raise TEST_CERT_ERR_BASE + 60, PROC, _
                "ThisWorkbook.NewWindow did not return a temporary window"
        End If

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(EntryBase, LeakRec), _
            PROC & ".Leak.NotClean"

        LeakCodes = TST_CertEntryCodes(LeakRec, TEST_CERT_ENTRY_FINDING)

        TST_AssertTrue (InStr(1, LeakCodes, "|windows.count|", vbBinaryCompare) > 0), _
            PROC & ".Leak.CountChanged"

        TST_AssertTrue (InStr(1, LeakCodes, "|window.unexpected|", vbBinaryCompare) > 0), _
            PROC & ".Leak.WindowUnexpected"

        TST_AssertTrue (InStr(1, LeakCodes, "|activeWindow.changed|", vbBinaryCompare) > 0), _
            PROC & ".Leak.ActiveWindowChanged"

        TST_AssertTrue (InStr(1, LeakCodes, "|window.missing|", vbBinaryCompare) = 0), _
            PROC & ".Leak.NoEntryWindowMissing"

    'Close it, then make the entry window active again, as the runner does
        TST_SafeCloseWindow LeakedWindow
        TST_WaitUI TEST_WAIT_SECONDS

        TST_CertRestoreActiveWindow EntryBase, RecoverRec

        TST_AssertTrue (RecoverRec.FindingCount = 0), _
            PROC & ".Leak.ActiveWindowRestored"

'------------------------------------------------------------------------------
' B - A WINDOW REPLACED WHILE THE COUNT STAYS THE SAME
'------------------------------------------------------------------------------
        Set FirstTemp = ThisWorkbook.NewWindow

        If FirstTemp Is Nothing Then
            Err.Raise TEST_CERT_ERR_BASE + 61, PROC, _
                "ThisWorkbook.NewWindow did not return the first temporary window"
        End If

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertBooleanEquals True, TST_CertCaptureBaseline(ReplaceBase, ScratchRec), _
            PROC & ".Replace.Captured"

        TST_SafeCloseWindow FirstTemp
        TST_WaitUI TEST_WAIT_SECONDS

        Set SecondTemp = ThisWorkbook.NewWindow

        If SecondTemp Is Nothing Then
            Err.Raise TEST_CERT_ERR_BASE + 62, PROC, _
                "ThisWorkbook.NewWindow did not return the second temporary window"
        End If

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(ReplaceBase, ReplaceRec), _
            PROC & ".Replace.NotClean"

        ReplaceCodes = TST_CertEntryCodes(ReplaceRec, TEST_CERT_ENTRY_FINDING)

    'The count matches, which is exactly why it cannot be the only comparison
        TST_AssertTrue (InStr(1, ReplaceCodes, "|windows.count|", vbBinaryCompare) = 0), _
            PROC & ".Replace.CountUnchanged"

        TST_AssertTrue (InStr(1, ReplaceCodes, "|window.missing|", vbBinaryCompare) > 0), _
            PROC & ".Replace.EntryWindowMissing"

        TST_AssertTrue (InStr(1, ReplaceCodes, "|window.unexpected|", vbBinaryCompare) > 0), _
            PROC & ".Replace.WindowUnexpected"

    'Findings are in the order the comparison runs: an entry window is judged
    'before a window nothing accounts for
        TST_AssertTrue _
            (InStr(1, ReplaceCodes, "|window.missing|", vbBinaryCompare) < _
             InStr(1, ReplaceCodes, "|window.unexpected|", vbBinaryCompare)), _
            PROC & ".Replace.DeterministicOrder"

'------------------------------------------------------------------------------
' C - THE HOST IS BACK AT ITS BASELINE
'------------------------------------------------------------------------------
        TST_SafeCloseWindow SecondTemp
        TST_WaitUI TEST_WAIT_SECONDS

        TST_CertRestoreActiveWindow EntryBase, RecoverRec

        TST_AssertBooleanEquals True, TST_CertEvaluateCleanup(EntryBase, CleanRec), _
            PROC & ".Recovered.Clean"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "A leaked window and a replaced window were both reported, and the host was put back"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Close whatever is still open, whatever happened above
        On Error Resume Next
            TST_SafeCloseWindow LeakedWindow
            TST_SafeCloseWindow FirstTemp
            TST_SafeCloseWindow SecondTemp
        On Error GoTo 0

    'The entry window must be active again; a window that could not be closed
    'or activated leaves Excel altered, which fails this case
        If EntryCaptured Then

            TST_CertRestoreActiveWindow EntryBase, RecoverRec

            If RecoverRec.FindingCount > 0 Then
                TST_CertNoteHostFailure PROC, _
                    "the entry window could not be made active again | " & _
                    TST_CertEntrySummary(RecoverRec, TEST_CERT_ENTRY_FINDING), _
                    HasFailure, FailNumber, FailSource, FailDescription
            End If

            If Application.Windows.Count <> EntryBase.WindowCount Then
                TST_CertNoteHostFailure PROC, _
                    "a temporary window is still open; Application.Windows.Count is " & _
                    CStr(Application.Windows.Count) & ", not " & _
                    CStr(EntryBase.WindowCount), _
                    HasFailure, FailNumber, FailSource, FailDescription
            End If

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
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        HasFailure = True

        Resume Safe_Exit

End Sub


Private Sub TST_Case_CertCleanupDetectsPropertyDrift()

'
'==============================================================================
' TST_Case_CertCleanupDetectsPropertyDrift
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that cleanup reports an application-level property and a
'   window-level property left altered, naming both values, in a fixed order.
'
' WHY THIS EXISTS
'   Status Bar, Scroll Bars, Formula Bar, Headings, Workbook Tabs and
'   Gridlines could be left altered with cleanup=OK, because only the workbook
'   count and ScreenUpdating were compared (#35).
'
' RETURNS
'   None.
'
' BEHAVIOR
'   A  application    Application.DisplayStatusBar inverted: reported with
'                     both values
'   B  window         DisplayWorkbookTabs of the entry active window inverted:
'                     reported with both values
'   C  both           reported together, the application property first
'   D  restored       the host matches its baseline again
'
' ERROR POLICY
'   - Raises on assertion failure, after restoring both properties and
'     checking that they read back as found.
'
' DEPENDENCIES
'   - TST_CertCaptureBaseline
'   - TST_CertEvaluateCleanup
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   The case writes the properties itself, never through the API under test,
'   and puts them back itself.
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Base                As tCertBaseline   'Baseline of the host as found
    Dim Rec                 As tCertRecord     'Record of the capture
    Dim AppRec              As tCertRecord     'Findings of scenario A
    Dim WinRec              As tCertRecord     'Findings of scenario B
    Dim BothRec             As tCertRecord     'Findings of scenario C
    Dim CleanRec            As tCertRecord     'Findings after restoration
    Dim TabsWindow          As Object          'Entry active window, retained
    Dim EntryStatusBar      As Boolean         'DisplayStatusBar as found
    Dim EntryTabs           As Boolean         'DisplayWorkbookTabs as found
    Dim StatusBarAltered    As Boolean         'TRUE while the status bar is inverted
    Dim TabsAltered         As Boolean         'TRUE while the tabs are inverted
    Dim BothCodes           As String          'Finding identifiers of scenario C
    Dim ActiveIdx           As Long            'Index of the entry active window
    Dim Captured            As Boolean         'TRUE once the baseline exists
    Dim ValuesRead          As Boolean         'TRUE once the values found are known

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_CertCleanupDetectsPropertyDrift"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that an altered application or window property is reported"

        Captured = TST_CertCaptureBaseline(Base, Rec)

        TST_AssertBooleanEquals True, Captured, PROC & ".Captured"

        TST_AssertBooleanEquals True, Base.AppKnown(TST_CERT_APP_STATUSBAR), _
            PROC & ".StatusBarWasRead"

        TST_AssertBooleanEquals True, Base.ActiveKnown, PROC & ".ActiveWindowKnown"

        ActiveIdx = Base.ActiveIndex

        TST_AssertBooleanEquals True, Base.WinPropKnown(ActiveIdx, TST_CERT_WINPROP_TABS), _
            PROC & ".WorkbookTabsWasRead"

        Set TabsWindow = Base.Win(ActiveIdx)

        EntryStatusBar = Base.AppValue(TST_CERT_APP_STATUSBAR)
        EntryTabs = Base.WinPropValue(ActiveIdx, TST_CERT_WINPROP_TABS)

        ValuesRead = True

'------------------------------------------------------------------------------
' A - AN APPLICATION PROPERTY
'------------------------------------------------------------------------------
        Application.DisplayStatusBar = Not EntryStatusBar
        StatusBarAltered = True

        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(Base, AppRec), _
            PROC & ".App.NotClean"

        TST_AssertTrue _
            (InStr(1, TST_CertEntryCodes(AppRec, TEST_CERT_ENTRY_FINDING), _
                "|app.DisplayStatusBar|", vbBinaryCompare) > 0), _
            PROC & ".App.Named"

        TST_AssertTrue _
            (InStr(1, TST_CertEntrySummary(AppRec, TEST_CERT_ENTRY_FINDING), _
                "DisplayStatusBar changed from " & CStr(EntryStatusBar) & _
                " to " & CStr(Not EntryStatusBar), vbBinaryCompare) > 0), _
            PROC & ".App.NamesBothValues"

        Application.DisplayStatusBar = EntryStatusBar
        StatusBarAltered = False

'------------------------------------------------------------------------------
' B - A WINDOW PROPERTY
'------------------------------------------------------------------------------
        TabsWindow.DisplayWorkbookTabs = Not EntryTabs
        TabsAltered = True

        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(Base, WinRec), _
            PROC & ".Window.NotClean"

        TST_AssertTrue _
            (InStr(1, TST_CertEntryCodes(WinRec, TEST_CERT_ENTRY_FINDING), _
                "|window.DisplayWorkbookTabs|", vbBinaryCompare) > 0), _
            PROC & ".Window.Named"

        TST_AssertTrue _
            (InStr(1, TST_CertEntrySummary(WinRec, TEST_CERT_ENTRY_FINDING), _
                "DisplayWorkbookTabs changed from " & CStr(EntryTabs) & _
                " to " & CStr(Not EntryTabs), vbBinaryCompare) > 0), _
            PROC & ".Window.NamesBothValues"

        TabsWindow.DisplayWorkbookTabs = EntryTabs
        TabsAltered = False

'------------------------------------------------------------------------------
' C - BOTH, IN A FIXED ORDER
'------------------------------------------------------------------------------
        Application.DisplayStatusBar = Not EntryStatusBar
        StatusBarAltered = True
        TabsWindow.DisplayWorkbookTabs = Not EntryTabs
        TabsAltered = True

        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(Base, BothRec), _
            PROC & ".Both.NotClean"

        BothCodes = TST_CertEntryCodes(BothRec, TEST_CERT_ENTRY_FINDING)

        TST_AssertTrue _
            (InStr(1, BothCodes, "|app.DisplayStatusBar|", vbBinaryCompare) > 0), _
            PROC & ".Both.AppReported"

        TST_AssertTrue _
            (InStr(1, BothCodes, "|window.DisplayWorkbookTabs|", vbBinaryCompare) > 0), _
            PROC & ".Both.WindowReported"

        TST_AssertTrue _
            (InStr(1, BothCodes, "|app.DisplayStatusBar|", vbBinaryCompare) < _
             InStr(1, BothCodes, "|window.DisplayWorkbookTabs|", vbBinaryCompare)), _
            PROC & ".Both.DeterministicOrder"

        Application.DisplayStatusBar = EntryStatusBar
        StatusBarAltered = False
        TabsWindow.DisplayWorkbookTabs = EntryTabs
        TabsAltered = False

'------------------------------------------------------------------------------
' D - RESTORED
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertEvaluateCleanup(Base, CleanRec), _
            PROC & ".Restored.Clean"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "An altered application property and window property were both reported with their values"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Put both properties back, whatever happened above
        On Error Resume Next
            If StatusBarAltered Then
                Application.DisplayStatusBar = EntryStatusBar
            End If

            If TabsAltered Then
                TabsWindow.DisplayWorkbookTabs = EntryTabs
            End If
        On Error GoTo 0

    'A property that did not read back as found leaves Excel altered
        If ValuesRead Then

            If Application.DisplayStatusBar <> EntryStatusBar Then
                TST_CertNoteHostFailure PROC, _
                    "DisplayStatusBar could not be restored to " & _
                    CStr(EntryStatusBar), _
                    HasFailure, FailNumber, FailSource, FailDescription
            End If

            If TabsWindow.DisplayWorkbookTabs <> EntryTabs Then
                TST_CertNoteHostFailure PROC, _
                    "DisplayWorkbookTabs could not be restored to " & _
                    CStr(EntryTabs), _
                    HasFailure, FailNumber, FailSource, FailDescription
            End If

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
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        HasFailure = True

        Resume Safe_Exit

End Sub


Private Sub TST_Case_CertCleanupDetectsTitleBarMismatch()

'
'==============================================================================
' TST_Case_CertCleanupDetectsTitleBarMismatch
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that cleanup reports a title bar, an owned frame bit and a frame
'   registry state that differ from the baseline.
'
' WHY THIS EXISTS
'   A title bar left hidden, or carrying the wrong owned frame bits, or a
'   pending frame refresh, could coexist with cleanup=OK (#35).
'
' RETURNS
'   None.
'
' BEHAVIOR
'   A  caption       the baseline records the caption bits inverted: both
'                    the visibility and the owned bits are reported
'   B  system menu   only WS_SYSMENU differs: the owned bits are reported and
'                    the visibility is not
'   C  registry      the baseline records the opposite refresh debt and
'                    component ownership: both are reported
'   D  untouched     with every record put back, nothing is reported
'
' ERROR POLICY
'   - Raises on assertion failure.
'
' DEPENDENCIES
'   - TST_CertCaptureBaseline
'   - TST_CertEvaluateCleanup
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   The fault is injected into the local baseline record, never into a frame:
'   the live title bar is only read, so no path of this case can leave Excel
'   with a hidden title bar. The write and the refresh of a real frame are
'   covered by the title-bar cases.
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Base                As tCertBaseline   'Baseline whose record is altered
    Dim Rec                 As tCertRecord     'Record of the capture
    Dim CaptionRec          As tCertRecord     'Findings of scenario A
    Dim MenuRec             As tCertRecord     'Findings of scenario B
    Dim RegistryRec         As tCertRecord     'Findings of scenario C
    Dim CleanRec            As tCertRecord     'Findings of scenario D
    Dim ActiveIdx           As Long            'Index of the entry active window
    Dim Codes               As String          'Finding identifiers of a scenario

    Const PROC As String = "TST_Case_CertCleanupDetectsTitleBarMismatch"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        TST_Log PROC, "START", _
            "Validating that a title-bar or frame-state mismatch is reported"

        TST_AssertBooleanEquals True, TST_CertCaptureBaseline(Base, Rec), _
            PROC & ".Captured"

        TST_AssertBooleanEquals True, Base.ActiveKnown, PROC & ".ActiveWindowKnown"

        ActiveIdx = Base.ActiveIndex

    'The active window's frame is the one the host agrees on, so it is readable
        TST_AssertBooleanEquals True, Base.WinTitleKnown(ActiveIdx), _
            PROC & ".TitleBarWasRead"

        TST_AssertBooleanEquals True, Base.WinHwndKnown(ActiveIdx), _
            PROC & ".FrameHandleWasRead"

'------------------------------------------------------------------------------
' A - THE CAPTION BITS DIFFER
'------------------------------------------------------------------------------
        Base.WinStyle(ActiveIdx) = Base.WinStyle(ActiveIdx) Xor TST_WS_CAPTION

        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(Base, CaptionRec), _
            PROC & ".Caption.NotClean"

        Codes = TST_CertEntryCodes(CaptionRec, TEST_CERT_ENTRY_FINDING)

        TST_AssertTrue (InStr(1, Codes, "|titleBar.visible|", vbBinaryCompare) > 0), _
            PROC & ".Caption.VisibilityReported"

        TST_AssertTrue (InStr(1, Codes, "|titleBar.ownedBits|", vbBinaryCompare) > 0), _
            PROC & ".Caption.OwnedBitsReported"

        TST_AssertTrue _
            (InStr(1, TST_CertEntrySummary(CaptionRec, TEST_CERT_ENTRY_FINDING), _
                "title bar visibility changed from", vbBinaryCompare) > 0), _
            PROC & ".Caption.NamesBothStates"

        Base.WinStyle(ActiveIdx) = Base.WinStyle(ActiveIdx) Xor TST_WS_CAPTION

'------------------------------------------------------------------------------
' B - ONLY AN OWNED FRAME BIT OTHER THAN THE CAPTION DIFFERS
'------------------------------------------------------------------------------
        Base.WinStyle(ActiveIdx) = Base.WinStyle(ActiveIdx) Xor TST_CERT_WS_SYSMENU

        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(Base, MenuRec), _
            PROC & ".SysMenu.NotClean"

        Codes = TST_CertEntryCodes(MenuRec, TEST_CERT_ENTRY_FINDING)

        TST_AssertTrue (InStr(1, Codes, "|titleBar.ownedBits|", vbBinaryCompare) > 0), _
            PROC & ".SysMenu.OwnedBitsReported"

        TST_AssertTrue (InStr(1, Codes, "|titleBar.visible|", vbBinaryCompare) = 0), _
            PROC & ".SysMenu.VisibilityUnchanged"

        Base.WinStyle(ActiveIdx) = Base.WinStyle(ActiveIdx) Xor TST_CERT_WS_SYSMENU

'------------------------------------------------------------------------------
' C - THE FRAME REGISTRY DIFFERS
'------------------------------------------------------------------------------
        Base.WinRegRefresh(ActiveIdx) = Not Base.WinRegRefresh(ActiveIdx)
        Base.WinRegHidden(ActiveIdx) = Not Base.WinRegHidden(ActiveIdx)

        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(Base, RegistryRec), _
            PROC & ".Registry.NotClean"

        Codes = TST_CertEntryCodes(RegistryRec, TEST_CERT_ENTRY_FINDING)

        TST_AssertTrue (InStr(1, Codes, "|frame.refreshPending|", vbBinaryCompare) > 0), _
            PROC & ".Registry.RefreshDebtReported"

        TST_AssertTrue (InStr(1, Codes, "|frame.componentHidden|", vbBinaryCompare) > 0), _
            PROC & ".Registry.OwnershipReported"

        Base.WinRegRefresh(ActiveIdx) = Not Base.WinRegRefresh(ActiveIdx)
        Base.WinRegHidden(ActiveIdx) = Not Base.WinRegHidden(ActiveIdx)

'------------------------------------------------------------------------------
' D - WITH EVERY RECORD PUT BACK, NOTHING IS REPORTED
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertEvaluateCleanup(Base, CleanRec), _
            PROC & ".Untouched.Clean"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "A title-bar mismatch, an owned-bit mismatch and a registry mismatch were each reported"

End Sub


Private Sub TST_Case_CertCleanupHelperFailureIsFinding()

'
'==============================================================================
' TST_Case_CertCleanupHelperFailureIsFinding
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a cleanup helper or stage that fails produces a finding, that
'   the remaining comparisons still run, and that no caller can turn a finding
'   into cleanup=OK.
'
' WHY THIS EXISTS
'   Cleanup ran under On Error Resume Next, so a comparison that raised left
'   the verdict as it was and looked, in the evidence, like one that had found
'   nothing (#35).
'
' RETURNS
'   None.
'
' BEHAVIOR
'   A  failed read    the first property read on exit fails: reported as a
'                     finding that carries the helper's message
'   B  failed stage   the window comparison raises: reported as a finding, and
'                     the stages after it still compare
'   C  no baseline    a baseline that was not captured is reported, not assumed
'                     clean
'   D  verdict        the composer reports cleanup=FAILED for a finding even
'                     when the caller passes cleanup OK, and both documents
'                     list it
'   E  storage        a finding that cannot be stored fails the record closed
'
' ERROR POLICY
'   - Raises on assertion failure, after disarming the helper and record
'     faults.
'
' DEPENDENCIES
'   - TST_CertInjectHelperFault
'   - TST_CertInjectRecordFault
'   - TST_CertEvaluateCleanup
'   - TST_CertAddFinding
'   - TST_CertComposeEvidence
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   Both faults are one-shot and local to this module or to a local record. The
'   certification run executing this case has its own baseline and record, and
'   neither is touched.
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Base                As tCertBaseline   'Baseline of the host as found
    Dim Blank               As tCertBaseline   'A baseline that was never captured
    Dim Rec                 As tCertRecord     'Record of the capture
    Dim ReadRec             As tCertRecord     'Findings of scenario A
    Dim StageRec            As tCertRecord     'Findings of scenario B
    Dim BlankRec            As tCertRecord     'Findings of scenario C
    Dim StoreRec            As tCertRecord     'Record of scenario E
    Dim PropIdx             As Long            'Cursor over the properties
    Dim KnownCount          As Long            'Application properties read
    Dim MinChecks           As Long            'Comparisons a surviving stage run makes
    Dim Stored              As Boolean         'Result of storing a finding

    Dim UnitsFailed         As Long            'Composed failed-unit count
    Dim Complete            As Boolean         'Composed completeness
    Dim Passed              As Boolean         'Composed verdict
    Dim Verdict             As String          'Composed RESULT line
    Dim CleanupDetail       As String          'Composed cleanup detail
    Dim RecordingDetail     As String          'Composed recording failure
    Dim JsonText            As String          'Composed JSON evidence
    Dim ReportText          As String          'Composed text evidence

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_CertCleanupHelperFailureIsFinding"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a failing cleanup helper is a finding, not a clean verdict"

        TST_AssertBooleanEquals True, TST_CertCaptureBaseline(Base, Rec), _
            PROC & ".Captured"

'------------------------------------------------------------------------------
' A - A FAILED READ
'------------------------------------------------------------------------------
        TST_CertInjectHelperFault TST_CERT_HELPER_FAULT_READ

        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(Base, ReadRec), _
            PROC & ".Read.NotClean"

        TST_AssertTrue _
            (InStr(1, TST_CertEntryCodes(ReadRec, TEST_CERT_ENTRY_FINDING), _
                "|app.DisplayStatusBar.read|", vbBinaryCompare) > 0), _
            PROC & ".Read.Reported"

        TST_AssertTrue _
            (InStr(1, TST_CertEntrySummary(ReadRec, TEST_CERT_ENTRY_FINDING), _
                "injected cleanup helper failure", vbBinaryCompare) > 0), _
            PROC & ".Read.CarriesTheHelperMessage"

        TST_AssertBooleanEquals False, TST_CertTakeHelperFault(TST_CERT_HELPER_FAULT_READ), _
            PROC & ".Read.OneShot"

'------------------------------------------------------------------------------
' B - A FAILED STAGE
'------------------------------------------------------------------------------
        TST_CertInjectHelperFault TST_CERT_HELPER_FAULT_STAGE

        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(Base, StageRec), _
            PROC & ".Stage.NotClean"

        TST_AssertTrue _
            (InStr(1, TST_CertEntryCodes(StageRec, TEST_CERT_ENTRY_FINDING), _
                "|cleanup.windows|", vbBinaryCompare) > 0), _
            PROC & ".Stage.Reported"

        TST_AssertTrue _
            (InStr(1, TST_CertEntrySummary(StageRec, TEST_CERT_ENTRY_FINDING), _
                "injected cleanup stage failure", vbBinaryCompare) > 0), _
            PROC & ".Stage.CarriesTheErrorText"

    'The stages after the failed one still compared: the snapshot check, both
    'counts, every known application property, and the active window
        For PropIdx = 1 To TST_CERT_APP_COUNT
            If Base.AppKnown(PropIdx) Then
                KnownCount = KnownCount + 1
            End If
        Next PropIdx

        MinChecks = 3 + KnownCount + IIf(Base.ActiveKnown, 1, 0)

        TST_AssertTrue (StageRec.CheckCount >= MinChecks), _
            PROC & ".Stage.LaterStagesStillRan"

'------------------------------------------------------------------------------
' C - NO BASELINE
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(Blank, BlankRec), _
            PROC & ".NoBaseline.NotClean"

        TST_AssertTrue _
            (InStr(1, TST_CertEntryCodes(BlankRec, TEST_CERT_ENTRY_FINDING), _
                "|baseline.unavailable|", vbBinaryCompare) > 0), _
            PROC & ".NoBaseline.Reported"

'------------------------------------------------------------------------------
' D - THE VERDICT AND BOTH DOCUMENTS
'------------------------------------------------------------------------------
    'Cleanup is passed as OK on purpose: the finding alone must fail it
        TST_CertComposeEvidence _
            Rec:=ReadRec, _
            CleanupOK:=True, _
            CleanupDetail:=CleanupDetail, _
            UnitsFailed:=UnitsFailed, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            RecordingDetail:=RecordingDetail, _
            JsonText:=JsonText, _
            ReportText:=ReportText

        TST_AssertBooleanEquals False, Passed, PROC & ".Verdict.NotPassed"

        TST_AssertTrue _
            (InStr(1, Verdict, "RESULT: FAIL | ", vbBinaryCompare) = 1) And _
            (InStr(1, Verdict, " cleanup=FAILED", vbBinaryCompare) > 0), _
            PROC & ".Verdict.CleanupFailed"

        TST_AssertTrue _
            (InStr(1, JsonText, """cleanup"":""FAILED""", vbBinaryCompare) > 0), _
            PROC & ".Json.CleanupFailed"

        TST_AssertTrue _
            (InStr(1, JsonText, """code"":""app.DisplayStatusBar.read""", vbBinaryCompare) > 0), _
            PROC & ".Json.ListsTheFinding"

        TST_AssertTrue _
            (InStr(1, ReportText, vbNewLine & "  CLEANUP  app.DisplayStatusBar.read | ", _
                vbBinaryCompare) > 0), _
            PROC & ".Report.ListsTheFinding"

        TST_AssertTrue _
            (InStr(1, CleanupDetail, "injected cleanup helper failure", vbBinaryCompare) > 0), _
            PROC & ".Verdict.DetailReturnedToTheCaller"

'------------------------------------------------------------------------------
' E - A FINDING THAT CANNOT BE STORED
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertResetRecord(StoreRec), _
            PROC & ".Store.Reset"

        TST_CertInjectRecordFault TEST_CERT_RECORD_FAULT_FIRST, TEST_CERT_BUFFER_ANY

        Stored = TST_CertAddFinding(StoreRec, "test.unstored", "a finding with no room")

        TST_CertInjectRecordFault TEST_CERT_RECORD_FAULT_NONE, TEST_CERT_BUFFER_ANY

        TST_AssertBooleanEquals False, Stored, PROC & ".Store.NotStored"

        TST_AssertBooleanEquals True, StoreRec.RecordFailed, PROC & ".Store.RecordFlagged"

        CleanupDetail = vbNullString

        TST_CertComposeEvidence _
            Rec:=StoreRec, _
            CleanupOK:=True, _
            CleanupDetail:=CleanupDetail, _
            UnitsFailed:=UnitsFailed, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            RecordingDetail:=RecordingDetail, _
            JsonText:=JsonText, _
            ReportText:=ReportText

        TST_AssertBooleanEquals False, Passed, PROC & ".Store.NotPassed"
        TST_AssertBooleanEquals False, Complete, PROC & ".Store.NotComplete"

        TST_AssertTrue _
            (InStr(1, Verdict, "RESULT: FAIL | INCOMPLETE | ", vbBinaryCompare) = 1), _
            PROC & ".Store.FailIncomplete"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "A failing helper, a failing stage and an unstorable finding each failed cleanup"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'A fault left armed would fail the next comparison or record of the run
    'executing this case, so both are disarmed on every path
        On Error Resume Next
            TST_CertInjectHelperFault TST_CERT_HELPER_FAULT_NONE
            TST_CertInjectRecordFault TEST_CERT_RECORD_FAULT_NONE, TEST_CERT_BUFFER_ANY
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


Private Sub TST_Case_CertCleanupReportsQuietUpdateResidue()

'
'==============================================================================
' TST_Case_CertCleanupReportsQuietUpdateResidue
'------------------------------------------------------------------------------
' PURPOSE
'   Verify the quiet-update handoff of #26: when the suppression write succeeds
'   and its readback fails, ownership is refused, End restores nothing, the
'   host stays suppressed, and the final-baseline comparison reports the
'   mismatch before test-owned recovery.
'
' WHY THIS EXISTS
'   UI_RuntimeBeginQuietUpdate claims a suppression only after a readback
'   confirms it. When the readback fails after a real write, the host is left
'   with ScreenUpdating False and unowned, and no rollback is attempted. That
'   is deliberate, and it is the residue this issue exists to detect (#26, #35).
'
' RETURNS
'   None.
'
' BEHAVIOR
'   1  begin        the write succeeds and the readback fails:
'                   QuietModeChanged stays FALSE and the entry value is
'                   reported TRUE
'   2  host         ScreenUpdating is FALSE
'   3  end          UI_RuntimeEndQuietUpdate performs no restore:
'                   ScreenUpdating is still FALSE
'   4  compare      the comparison, made before any recovery, reports
'                   ScreenUpdating changed from True to False
'   5  recover      the case sets ScreenUpdating back itself, and the
'                   comparison is then clean
'
' ERROR POLICY
'   - Raises on assertion failure, after disarming the seam and setting
'     ScreenUpdating back to the value found on entry.
'
' DEPENDENCIES
'   - UI_InternalInjectQuietUpdateFault
'   - UI_RuntimeBeginQuietUpdate
'   - UI_RuntimeEndQuietUpdate
'   - TST_CertCaptureBaseline
'   - TST_CertEvaluateCleanup
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   The seam is one-shot and is disarmed on every path. ScreenUpdating is set
'   back by the case, never by the scope under test, whose restoration is the
'   thing proved absent.
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EntryScreenUpdating As Boolean         'ScreenUpdating as found
    Dim Base                As tCertBaseline   'Baseline taken before the fault
    Dim Rec                 As tCertRecord     'Record of the capture
    Dim ResidueRec          As tCertRecord     'Findings before recovery
    Dim CleanRec            As tCertRecord     'Findings after recovery
    Dim OldValue            As Boolean         'Entry value the scope reports
    Dim Changed             As Boolean         'Ownership the scope claims
    Dim Codes               As String          'Finding identifiers before recovery

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_CertCleanupReportsQuietUpdateResidue"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a failed quiet-update readback is reported as residue"

        EntryScreenUpdating = Application.ScreenUpdating

        UI_InternalInjectQuietUpdateFault TST_CERT_QUIET_FAULT_NONE

    'The scope can only own a suppression it starts from an enabled host
        Application.ScreenUpdating = True

        TST_AssertBooleanEquals True, TST_CertCaptureBaseline(Base, Rec), _
            PROC & ".Captured"

        TST_AssertTrue _
            Base.AppKnown(TST_CERT_APP_SCREEN) And Base.AppValue(TST_CERT_APP_SCREEN), _
            PROC & ".BaselineRecordsAnEnabledHost"

'------------------------------------------------------------------------------
' THE WRITE SUCCEEDS AND THE READBACK FAILS
'------------------------------------------------------------------------------
        UI_InternalInjectQuietUpdateFault TST_CERT_QUIET_FAULT_READBACK

        OldValue = False
        Changed = True

        UI_RuntimeBeginQuietUpdate OldValue, Changed

        TST_AssertBooleanEquals True, OldValue, PROC & ".Begin.EntryValue"

        TST_AssertBooleanEquals False, Changed, PROC & ".Begin.QuietModeChangedStaysFalse"

        TST_AssertApplicationProperty _
            False, "ScreenUpdating", PROC & ".Begin.HostLeftSuppressed"

'------------------------------------------------------------------------------
' END PERFORMS NO RESTORE
'------------------------------------------------------------------------------
        UI_RuntimeEndQuietUpdate OldValue, Changed

        TST_AssertApplicationProperty _
            False, "ScreenUpdating", PROC & ".End.NoRestore"

'------------------------------------------------------------------------------
' THE COMPARISON REPORTS THE MISMATCH BEFORE ANY RECOVERY
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals False, TST_CertEvaluateCleanup(Base, ResidueRec), _
            PROC & ".Residue.NotClean"

        Codes = TST_CertEntryCodes(ResidueRec, TEST_CERT_ENTRY_FINDING)

        TST_AssertTrue (InStr(1, Codes, "|app.ScreenUpdating|", vbBinaryCompare) > 0), _
            PROC & ".Residue.Reported"

        TST_AssertTrue _
            (InStr(1, TST_CertEntrySummary(ResidueRec, TEST_CERT_ENTRY_FINDING), _
                "ScreenUpdating changed from True to False", vbBinaryCompare) > 0), _
            PROC & ".Residue.NamesBothValues"

'------------------------------------------------------------------------------
' TEST-OWNED RECOVERY
'------------------------------------------------------------------------------
    'The scope did not own the suppression, so the case sets the host back
        Application.ScreenUpdating = True

        TST_AssertBooleanEquals True, TST_CertEvaluateCleanup(Base, CleanRec), _
            PROC & ".Recovered.Clean"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "A failed quiet-update readback left the host suppressed, and the comparison reported it before recovery"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Disarm the seam and put ScreenUpdating back to the value found, whatever
    'happened above
        On Error Resume Next
            UI_InternalInjectQuietUpdateFault TST_CERT_QUIET_FAULT_NONE
            Application.ScreenUpdating = EntryScreenUpdating
        On Error GoTo 0

        If Application.ScreenUpdating <> EntryScreenUpdating Then
            TST_CertNoteHostFailure PROC, _
                "ScreenUpdating could not be restored to " & _
                CStr(EntryScreenUpdating), _
                HasFailure, FailNumber, FailSource, FailDescription
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
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        HasFailure = True

        Resume Safe_Exit

End Sub


Private Sub TST_Case_CertRecordFailsClosed()

'
'==============================================================================
' TST_Case_CertRecordFailsClosed
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that a certification record that cannot store an entry, or whose
'   buffers disagree with its counts, produces FAIL | INCOMPLETE evidence that
'   names the recording failure, and that a reset starts the next run clean.
'
' WHY THIS EXISTS
'   Recording raised the count before growing the parallel arrays, under On
'   Error Resume Next. A failed growth left a count describing an entry one
'   array did not hold, and the evidence loops read past its end with every
'   error suppressed (#50). A gate that cannot record its own result must
'   fail, never assemble a clean verdict from what is left.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   1  fault shapes      a first allocation with no buffer allocated, the
'                        first allocation of one chosen buffer, a later
'                        growth with every buffer allocated, the growth of
'                        one chosen buffer, and the growth of one chosen
'                        buffer for a skip: each fails closed, and a reset
'                        then gives a clean pass
'   2  independence      a growth fault ignores a first allocation, and a
'                        first-allocation fault ignores a growth
'   3  invariants        a count beyond the buffers, kinds that disagree with
'                        the kind counts and an unknown kind are each
'                        reported, and nothing is indexed past a bound
'   4  isolation         the record of the certification run executing this
'                        case is unchanged
'
' ERROR POLICY
'   - Raises on assertion failure, after disarming the record fault.
'
' DEPENDENCIES
'   - TST_CertCheckRecordFault
'   - TST_CertAssertFailedEvidence
'   - TST_CertInjectRecordFault
'   - TST_CertResetRecord
'   - TST_CertTryAppendEntry
'   - TST_CertRecordIsConsistent
'   - TST_CertComposeEvidence
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' NOTES
'   Every record here is a local one. The case drives the helpers that
'   certification uses on m_CertRecord, and composes evidence through the
'   same TST_CertComposeEvidence, so what it asserts is what certification
'   writes, while the run executing it records nothing from it.
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Rec                 As tCertRecord     'Record under test
    Dim Stored              As Boolean         'Result of an append

    Dim UnitsFailed         As Long            'Composed failed-unit count
    Dim Complete            As Boolean         'Composed completeness
    Dim Passed              As Boolean         'Composed verdict
    Dim Verdict             As String          'Composed RESULT line
    Dim RecordingDetail     As String          'Composed recording failure
    Dim JsonText            As String          'Composed JSON evidence
    Dim ReportText          As String          'Composed text evidence

    Dim RunEntries          As Long            'Run record: entries on entry
    Dim RunUnits            As Long            'Run record: units on entry
    Dim RunSkips            As Long            'Run record: skips on entry
    Dim RunFailed           As Boolean         'Run record: fatal flag on entry
    Dim RunFailCount        As Long            'Run record: failures on entry

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "TST_Case_CertRecordFailsClosed"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", _
            "Validating that a certification record that cannot be stored fails closed"

    'The record of the run executing this case, compared again at the end
        RunEntries = m_CertRecord.EntryCount
        RunUnits = m_CertRecord.UnitCount
        RunSkips = m_CertRecord.SkipCount
        RunFailed = m_CertRecord.RecordFailed
        RunFailCount = m_CertRecord.RecordFailCount

'------------------------------------------------------------------------------
' EVERY FAULT SHAPE FAILS CLOSED
'------------------------------------------------------------------------------
    'No buffer allocated yet: the first one fails
        TST_CertCheckRecordFault _
            Scenario:=PROC & ".FirstAllocation", _
            Kind:=TEST_CERT_ENTRY_UNIT, _
            PriorUnits:=0, _
            FaultKind:=TEST_CERT_RECORD_FAULT_FIRST, _
            BufferId:=TEST_CERT_BUFFER_ANY, _
            ExpectedStage:=TEST_CERT_BUFFER_KIND

    'Three buffers allocated, the last one not
        TST_CertCheckRecordFault _
            Scenario:=PROC & ".FirstAllocationOfOneBuffer", _
            Kind:=TEST_CERT_ENTRY_UNIT, _
            PriorUnits:=0, _
            FaultKind:=TEST_CERT_RECORD_FAULT_FIRST, _
            BufferId:=TEST_CERT_BUFFER_DETAIL, _
            ExpectedStage:=TEST_CERT_BUFFER_DETAIL

    'Every buffer allocated: the first growth fails
        TST_CertCheckRecordFault _
            Scenario:=PROC & ".LaterGrowth", _
            Kind:=TEST_CERT_ENTRY_UNIT, _
            PriorUnits:=2, _
            FaultKind:=TEST_CERT_RECORD_FAULT_GROWTH, _
            BufferId:=TEST_CERT_BUFFER_ANY, _
            ExpectedStage:=TEST_CERT_BUFFER_KIND

    'Two parallel buffers grow and the third does not
        TST_CertCheckRecordFault _
            Scenario:=PROC & ".GrowthOfOneBuffer", _
            Kind:=TEST_CERT_ENTRY_UNIT, _
            PriorUnits:=2, _
            FaultKind:=TEST_CERT_RECORD_FAULT_GROWTH, _
            BufferId:=TEST_CERT_BUFFER_PASSED, _
            ExpectedStage:=TEST_CERT_BUFFER_PASSED

    'A skip takes the same path as a unit
        TST_CertCheckRecordFault _
            Scenario:=PROC & ".SkipGrowthOfOneBuffer", _
            Kind:=TEST_CERT_ENTRY_SKIP, _
            PriorUnits:=1, _
            FaultKind:=TEST_CERT_RECORD_FAULT_GROWTH, _
            BufferId:=TEST_CERT_BUFFER_NAME, _
            ExpectedStage:=TEST_CERT_BUFFER_NAME

'------------------------------------------------------------------------------
' THE FAULT SHAPES ARE INDEPENDENT
'------------------------------------------------------------------------------
    'A growth fault passes over a first allocation and fails the next growth
        TST_AssertBooleanEquals True, TST_CertResetRecord(Rec), _
            PROC & ".Independent.GrowthReset"

        TST_CertInjectRecordFault TEST_CERT_RECORD_FAULT_GROWTH, TEST_CERT_BUFFER_ANY

        Stored = TST_CertTryAppendEntry(Rec, TEST_CERT_ENTRY_UNIT, "First", True, vbNullString)

        TST_AssertBooleanEquals True, Stored, _
            PROC & ".Independent.GrowthFaultIgnoresFirstAllocation"

        Stored = TST_CertTryAppendEntry(Rec, TEST_CERT_ENTRY_UNIT, "Second", True, vbNullString)

        TST_CertInjectRecordFault TEST_CERT_RECORD_FAULT_NONE, TEST_CERT_BUFFER_ANY

        TST_AssertBooleanEquals False, Stored, _
            PROC & ".Independent.GrowthFaultFailsGrowth"

        TST_AssertTrue (Rec.RecordFailStage = TEST_CERT_BUFFER_KIND), _
            PROC & ".Independent.GrowthFaultStage"

    'A first-allocation fault passes over a growth and fails the next first
    'allocation. A record reset does not disarm it; only TST_CertResetCounters
    'and the fault taking effect do.
        TST_AssertBooleanEquals True, TST_CertResetRecord(Rec), _
            PROC & ".Independent.FirstReset"

        TST_AssertBooleanEquals True, _
            TST_CertTryAppendEntry(Rec, TEST_CERT_ENTRY_UNIT, "First", True, vbNullString), _
            PROC & ".Independent.FirstStored"

        TST_CertInjectRecordFault TEST_CERT_RECORD_FAULT_FIRST, TEST_CERT_BUFFER_ANY

        Stored = TST_CertTryAppendEntry(Rec, TEST_CERT_ENTRY_UNIT, "Second", True, vbNullString)

        TST_AssertBooleanEquals True, Stored, _
            PROC & ".Independent.FirstFaultIgnoresGrowth"

        TST_AssertBooleanEquals True, TST_CertResetRecord(Rec), _
            PROC & ".Independent.FirstResetAgain"

        Stored = TST_CertTryAppendEntry(Rec, TEST_CERT_ENTRY_UNIT, "Third", True, vbNullString)

        TST_CertInjectRecordFault TEST_CERT_RECORD_FAULT_NONE, TEST_CERT_BUFFER_ANY

        TST_AssertBooleanEquals False, Stored, _
            PROC & ".Independent.FirstFaultFailsFirstAllocation"

        TST_AssertTrue (Rec.RecordFailStage = TEST_CERT_BUFFER_KIND), _
            PROC & ".Independent.FirstFaultStage"

'------------------------------------------------------------------------------
' A COUNT BEYOND THE BUFFERS IS NOT INDEXED
'------------------------------------------------------------------------------
    'Publish a count that no buffer holds, as a partial growth used to
        TST_AssertBooleanEquals True, TST_CertResetRecord(Rec), _
            PROC & ".Overcount.Reset"

        TST_AssertBooleanEquals True, _
            TST_CertTryAppendEntry(Rec, TEST_CERT_ENTRY_UNIT, "Only", True, vbNullString), _
            PROC & ".Overcount.EntryStored"

        Rec.EntryCount = 2
        Rec.UnitCount = 2

        TST_AssertBooleanEquals False, TST_CertRecordIsConsistent(Rec), _
            PROC & ".Overcount.Detected"

        TST_CertComposeEvidence _
            Rec:=Rec, _
            CleanupOK:=True, _
            CleanupDetail:=vbNullString, _
            UnitsFailed:=UnitsFailed, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            RecordingDetail:=RecordingDetail, _
            JsonText:=JsonText, _
            ReportText:=ReportText

        TST_CertAssertFailedEvidence _
            Scenario:=PROC & ".Overcount", _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            JsonText:=JsonText, _
            ReportText:=ReportText

    'Found by the check, not by an out-of-range read: one failure, no error
        TST_AssertTrue _
            (Rec.RecordFailStage = TEST_CERT_RECFAIL_INVARIANT) And _
            (Rec.RecordFailCount = 1) And _
            (Rec.RecordFailError = 0), _
            PROC & ".Overcount.NoOutOfRangeRead"

        TST_AssertTrue _
            (InStr(1, JsonText, """unitResults"":[],""skipDetail"":[]}", vbBinaryCompare) > 0), _
            PROC & ".Overcount.NothingListed"

    'A record in this state is not grown further
        TST_AssertBooleanEquals False, _
            TST_CertTryAppendEntry(Rec, TEST_CERT_ENTRY_UNIT, "More", True, vbNullString), _
            PROC & ".Overcount.NotGrown"

        TST_AssertTrue (Rec.EntryCount = 2), PROC & ".Overcount.CountUnchanged"

'------------------------------------------------------------------------------
' KINDS THAT DISAGREE WITH THEIR COUNTS ARE REPORTED
'------------------------------------------------------------------------------
    'One unit stored, counted as a skip: the totals agree, the kinds do not
        TST_AssertBooleanEquals True, TST_CertResetRecord(Rec), _
            PROC & ".KindMismatch.Reset"

        TST_AssertBooleanEquals True, _
            TST_CertTryAppendEntry(Rec, TEST_CERT_ENTRY_UNIT, "Only", True, vbNullString), _
            PROC & ".KindMismatch.EntryStored"

        Rec.UnitCount = 0
        Rec.SkipCount = 1

        TST_CertComposeEvidence _
            Rec:=Rec, _
            CleanupOK:=True, _
            CleanupDetail:=vbNullString, _
            UnitsFailed:=UnitsFailed, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            RecordingDetail:=RecordingDetail, _
            JsonText:=JsonText, _
            ReportText:=ReportText

        TST_CertAssertFailedEvidence _
            Scenario:=PROC & ".KindMismatch", _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            JsonText:=JsonText, _
            ReportText:=ReportText

        TST_AssertTrue (Rec.RecordFailStage = TEST_CERT_RECFAIL_INVARIANT), _
            PROC & ".KindMismatch.Stage"

'------------------------------------------------------------------------------
' AN UNKNOWN KIND IS REFUSED
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertResetRecord(Rec), _
            PROC & ".UnknownKind.Reset"

        TST_AssertBooleanEquals False, _
            TST_CertTryAppendEntry(Rec, 99, "Unknown", True, vbNullString), _
            PROC & ".UnknownKind.Refused"

        TST_AssertTrue _
            (Rec.RecordFailStage = TEST_CERT_RECFAIL_KIND) And (Rec.EntryCount = 0), _
            PROC & ".UnknownKind.Flagged"

'------------------------------------------------------------------------------
' THE RUN'S OWN RECORD IS UNCHANGED
'------------------------------------------------------------------------------
    'Nothing above may reach the record of the run executing this case
        TST_AssertTrue _
            (m_CertRecord.EntryCount = RunEntries) And _
            (m_CertRecord.UnitCount = RunUnits) And _
            (m_CertRecord.SkipCount = RunSkips) And _
            (m_CertRecord.RecordFailed = RunFailed) And _
            (m_CertRecord.RecordFailCount = RunFailCount), _
            PROC & ".RunRecordUnchanged"

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Every injected recording failure produced FAIL | INCOMPLETE evidence, and a reset started clean"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'A fault left armed would fail the next record of the run executing this
    'case, so it is disarmed on every path
        On Error Resume Next
            TST_CertInjectRecordFault TEST_CERT_RECORD_FAULT_NONE, TEST_CERT_BUFFER_ANY
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


Private Sub TST_CertCheckRecordFault( _
    ByVal Scenario As String, _
    ByVal Kind As Long, _
    ByVal PriorUnits As Long, _
    ByVal FaultKind As Long, _
    ByVal BufferId As Long, _
    ByVal ExpectedStage As Long)

'
'==============================================================================
' TST_CertCheckRecordFault
'------------------------------------------------------------------------------
' PURPOSE
'   Inject one record allocation failure into a fresh local record and assert
'   the whole fail-closed contract for it.
'
' WHY THIS EXISTS
'   Each fault shape must meet the same contract, and writing it out five
'   times would let one copy drift from the others.
'
' INPUTS
'   Scenario
'     Assertion-name prefix.
'
'   Kind
'     TEST_CERT_ENTRY_* value of the entry that fails.
'
'   PriorUnits
'     Passing units stored before the fault, so that a growth has buffers to
'     grow.
'
'   FaultKind / BufferId
'     The record fault to arm, as TST_CertInjectRecordFault takes it.
'
'   ExpectedStage
'     The TEST_CERT_BUFFER_* value the failure must be charged to.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   1  inject        the append fails; nothing is published; the fatal flag,
'                    one failure, the stage and error 7 are recorded; the
'                    buffers still agree with the counts
'   2  evidence      FAIL | INCOMPLETE, no passed=true document, the
'                    recording failure named in both documents, the stored
'                    entries still listed, and no further failure while the
'                    evidence was read
'   3  later entry   the next append is stored next to the earlier ones, and
'                    the flag stays set until a reset
'   4  reset         every field is cleared, and one passing unit then gives
'                    PASS | COMPLETE with recording OK
'
' ERROR POLICY
'   - Raises on assertion failure. The caller disarms the record fault.
'
' DEPENDENCIES
'   - TST_CertInjectRecordFault
'   - TST_CertResetRecord
'   - TST_CertTryAppendEntry
'   - TST_CertRecordIsConsistent
'   - TST_CertComposeEvidence
'   - TST_CertAssertFailedEvidence
'
' CALLED FROM
'   - TST_Case_CertRecordFailsClosed
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Rec                 As tCertRecord     'Fresh record for this scenario
    Dim Stored              As Boolean         'Result of an append
    Dim EntryIdx            As Long            'Cursor over the prior units

    Dim UnitsFailed         As Long            'Composed failed-unit count
    Dim Complete            As Boolean         'Composed completeness
    Dim Passed              As Boolean         'Composed verdict
    Dim Verdict             As String          'Composed RESULT line
    Dim RecordingDetail     As String          'Composed recording failure
    Dim JsonText            As String          'Composed JSON evidence
    Dim ReportText          As String          'Composed text evidence

'------------------------------------------------------------------------------
' START FROM A CLEAN RECORD
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals True, TST_CertResetRecord(Rec), Scenario & ".Reset"

        For EntryIdx = 1 To PriorUnits
            TST_AssertBooleanEquals True, _
                TST_CertTryAppendEntry(Rec, TEST_CERT_ENTRY_UNIT, _
                    "Prior" & CStr(EntryIdx), True, vbNullString), _
                Scenario & ".PriorStored"
        Next EntryIdx

'------------------------------------------------------------------------------
' INJECT
'------------------------------------------------------------------------------
    'Armed for this one append only
        TST_CertInjectRecordFault FaultKind, BufferId

        Stored = TST_CertTryAppendEntry(Rec, Kind, "Injected", _
            (Kind = TEST_CERT_ENTRY_UNIT), "injected entry")

        TST_CertInjectRecordFault TEST_CERT_RECORD_FAULT_NONE, TEST_CERT_BUFFER_ANY

'------------------------------------------------------------------------------
' NOTHING WAS PUBLISHED
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals False, Stored, Scenario & ".NotStored"

        TST_AssertBooleanEquals True, Rec.RecordFailed, Scenario & ".FatalFlagSet"

        TST_AssertTrue _
            (Rec.RecordFailCount = 1) And _
            (Rec.RecordFailStage = ExpectedStage) And _
            (Rec.RecordFailError = 7), _
            Scenario & ".FailureNamed count=" & CStr(Rec.RecordFailCount) & _
            " stage=" & CStr(Rec.RecordFailStage) & _
            " error=" & CStr(Rec.RecordFailError)

        TST_AssertTrue _
            (Rec.EntryCount = PriorUnits) And _
            (Rec.UnitCount = PriorUnits) And _
            (Rec.SkipCount = 0), _
            Scenario & ".CountsUnchanged"

        TST_AssertBooleanEquals True, TST_CertRecordIsConsistent(Rec), _
            Scenario & ".BuffersAgreeWithCounts"

'------------------------------------------------------------------------------
' THE EVIDENCE FAILS CLOSED
'------------------------------------------------------------------------------
    'Cleanup is passed as OK, so only the recording failure can fail the run
        TST_CertComposeEvidence _
            Rec:=Rec, _
            CleanupOK:=True, _
            CleanupDetail:=vbNullString, _
            UnitsFailed:=UnitsFailed, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            RecordingDetail:=RecordingDetail, _
            JsonText:=JsonText, _
            ReportText:=ReportText

        TST_CertAssertFailedEvidence _
            Scenario:=Scenario, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            JsonText:=JsonText, _
            ReportText:=ReportText

    'Reading the record raised nothing: a read past a bound would have been
    'counted as a second failure
        TST_AssertTrue (Rec.RecordFailCount = 1), Scenario & ".EvidenceReadCleanly"

        If PriorUnits > 0 Then
            TST_AssertTrue _
                (InStr(1, JsonText, "{""name"":""Prior1"",""passed"":true", vbBinaryCompare) > 0), _
                Scenario & ".StoredEntriesListed"
        End If

'------------------------------------------------------------------------------
' A LATER ENTRY IS STORED, AND THE FLAG STAYS
'------------------------------------------------------------------------------
    'Buffers grown before the failure are reused; nothing has to be repaired
        Stored = TST_CertTryAppendEntry(Rec, Kind, "After", _
            (Kind = TEST_CERT_ENTRY_UNIT), vbNullString)

        TST_AssertBooleanEquals True, Stored, Scenario & ".LaterEntryStored"

        TST_AssertTrue (Rec.EntryCount = PriorUnits + 1), Scenario & ".LaterEntryPublished"

        TST_AssertBooleanEquals True, TST_CertRecordIsConsistent(Rec), _
            Scenario & ".LaterEntryConsistent"

        TST_AssertBooleanEquals True, Rec.RecordFailed, Scenario & ".FlagSurvivesLaterEntry"

        TST_CertComposeEvidence _
            Rec:=Rec, _
            CleanupOK:=True, _
            CleanupDetail:=vbNullString, _
            UnitsFailed:=UnitsFailed, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            RecordingDetail:=RecordingDetail, _
            JsonText:=JsonText, _
            ReportText:=ReportText

        TST_CertAssertFailedEvidence _
            Scenario:=Scenario & ".AfterLaterEntry", _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            JsonText:=JsonText, _
            ReportText:=ReportText

'------------------------------------------------------------------------------
' A RESET STARTS CLEAN
'------------------------------------------------------------------------------
    'What the next certification run does first
        TST_AssertBooleanEquals True, TST_CertResetRecord(Rec), Scenario & ".ResetAfterFailure"

        TST_AssertTrue _
            (Not Rec.RecordFailed) And _
            (Rec.RecordFailCount = 0) And _
            (Rec.RecordFailStage = TEST_CERT_RECFAIL_NONE) And _
            (Rec.RecordFailError = 0) And _
            (Rec.EntryCount = 0) And _
            (Rec.UnitCount = 0) And _
            (Rec.SkipCount = 0), _
            Scenario & ".ResetClearsEveryField"

        TST_AssertBooleanEquals True, _
            TST_CertTryAppendEntry(Rec, TEST_CERT_ENTRY_UNIT, "Clean", True, vbNullString), _
            Scenario & ".CleanEntryStored"

        TST_CertComposeEvidence _
            Rec:=Rec, _
            CleanupOK:=True, _
            CleanupDetail:=vbNullString, _
            UnitsFailed:=UnitsFailed, _
            Complete:=Complete, _
            Passed:=Passed, _
            Verdict:=Verdict, _
            RecordingDetail:=RecordingDetail, _
            JsonText:=JsonText, _
            ReportText:=ReportText

        TST_AssertBooleanEquals True, Passed, Scenario & ".CleanRun.Passed"
        TST_AssertBooleanEquals True, Complete, Scenario & ".CleanRun.Complete"

        TST_AssertTrue _
            (Verdict = "RESULT: PASS | COMPLETE | units=1 failed=0 skipped=0 cleanup=OK"), _
            Scenario & ".CleanRun.Verdict"

        TST_AssertTrue _
            (InStr(1, JsonText, _
                """recording"":""OK"",""recordingDetail"":"""",""complete"":true" & _
                ",""passed"":true,""unitResults"":[{""name"":""Clean"",""passed"":true" & _
                ",""detail"":""""}],""skipDetail"":[]}", vbBinaryCompare) > 0), _
            Scenario & ".CleanRun.Json"

        TST_AssertTrue _
            (InStr(1, ReportText, "RECORDING", vbBinaryCompare) = 0), _
            Scenario & ".CleanRun.NoRecordingLine"

End Sub


Private Sub TST_CertAssertFailedEvidence( _
    ByVal Scenario As String, _
    ByVal Complete As Boolean, _
    ByVal Passed As Boolean, _
    ByVal Verdict As String, _
    ByVal JsonText As String, _
    ByVal ReportText As String)

'
'==============================================================================
' TST_CertAssertFailedEvidence
'------------------------------------------------------------------------------
' PURPOSE
'   Assert that composed certification evidence reports a recording failure
'   and cannot be read as a pass.
'
' WHY THIS EXISTS
'   Every recording failure must produce the same evidence, whatever caused
'   it, so the assertions are written once.
'
' INPUTS
'   Scenario
'     Assertion-name prefix.
'
'   Complete / Passed / Verdict / JsonText / ReportText
'     Outputs of TST_CertComposeEvidence.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Passed and Complete are FALSE, and the verdict is a complete
'     RESULT: FAIL | INCOMPLETE line.
'   - The JSON reports "recording":"FAILED" with a detail naming the
'     failure, "complete":false and "passed":false, and nowhere a top-level
'     "passed":true.
'   - The text report carries the same RESULT line and a RECORDING line.
'
' ERROR POLICY
'   - Raises on assertion failure.
'
' CALLED FROM
'   - TST_Case_CertRecordFailsClosed
'   - TST_CertCheckRecordFault
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT VERDICT
'------------------------------------------------------------------------------
        TST_AssertBooleanEquals False, Passed, Scenario & ".Evidence.NotPassed"
        TST_AssertBooleanEquals False, Complete, Scenario & ".Evidence.NotComplete"

        TST_AssertTrue _
            (InStr(1, Verdict, "RESULT: FAIL | INCOMPLETE | units=", vbBinaryCompare) = 1), _
            Scenario & ".Evidence.FailIncomplete"

'------------------------------------------------------------------------------
' ASSERT JSON
'------------------------------------------------------------------------------
        TST_AssertTrue _
            (InStr(1, JsonText, _
                """recording"":""FAILED"",""recordingDetail"":""certification recording FAILED | ", _
                vbBinaryCompare) > 0), _
            Scenario & ".Evidence.JsonNamesRecordingFailure"

        TST_AssertTrue _
            (InStr(1, JsonText, """complete"":false,""passed"":false,""unitResults"":", _
                vbBinaryCompare) > 0), _
            Scenario & ".Evidence.JsonNotPassed"

        TST_AssertTrue _
            (InStr(1, JsonText, """passed"":true,""unitResults"":", vbBinaryCompare) = 0), _
            Scenario & ".Evidence.NoCleanPassedDocument"

'------------------------------------------------------------------------------
' ASSERT TEXT REPORT
'------------------------------------------------------------------------------
        TST_AssertTrue _
            (InStr(1, ReportText, vbNewLine & Verdict & vbNewLine, vbBinaryCompare) > 0), _
            Scenario & ".Evidence.ReportCarriesVerdict"

        TST_AssertTrue _
            (InStr(1, ReportText, vbNewLine & "  RECORDING  certification recording FAILED | ", _
                vbBinaryCompare) > 0), _
            Scenario & ".Evidence.ReportNamesRecordingFailure"

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
