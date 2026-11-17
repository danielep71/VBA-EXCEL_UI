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
'     - the cleanup verdict compares the host with its entry state
'     - a unit that fails after a capture reports the failure list it
'       asserted on and leaves no snapshot for the next unit
'     - every destructive public runner refuses a caller-owned snapshot
'     - a certification record that cannot store an entry, or whose
'       buffers disagree with its counts, gives FAIL | INCOMPLETE evidence
'       that names the recording failure, and a reset starts clean
'
' STATE MANAGEMENT
'   - The packs snapshot the current managed Excel UI state before testing
'     and attempt a best-effort restore after success or failure
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
'   - Reports FAIL | INCOMPLETE, and names the failure in both documents,
'     when a unit or skip could not be recorded.
'   - Raises when the verdict is anything other than a complete pass.
'
' ERROR POLICY
'   - Individual unit failures are captured, not propagated.
'   - Raises once at the end when the verdict is not PASS/COMPLETE.
'
' DEPENDENCIES
'   - TST_CertResetCounters
'   - TST_CertRunUnit
'   - TST_CertComposeEvidence
'   - TST_CertTryWriteEvidence
'
' NOTES
'   Destructive: creates and closes temporary workbooks and toggles every
'   managed UI element. Save unsaved work before running it.
'
' UPDATED
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
' DETERMINE VERDICT AND COMPOSE EVIDENCE
'------------------------------------------------------------------------------
    'Completeness and correctness are separate questions and are reported as
    'such. A run that skipped a mandatory unit, or could not record one, is
    'not a pass, whatever the assertions that did execute reported.
    '
    'The composer checks the record before reading it and decides the verdict
    'after, so a unit or skip that could not be stored makes the run
    'FAIL | INCOMPLETE and is named in both documents. Passed and Complete are
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
