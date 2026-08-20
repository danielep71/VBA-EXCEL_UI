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
'
'   Environment-preservation tests
'     - ScreenUpdating restored to prior state
'
'   Title-bar tests
'     - hide / show round-trip
'     - preservation of unrelated GWL_STYLE bits across hide / show
'
' STATE MANAGEMENT
'   - The harness snapshots the current managed Excel UI state before testing
'   - The harness attempts a best-effort restore after success or failure
'   - Per-window test state is captured and restored by Application.Windows index
'   - A pre-existing explicit EXCEL_UI snapshot is left untouched by skipping
'     the snapshot-destructive lifecycle cases
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
'   2026-08-18 - Reformatted to the project house style. No behavior change.
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.0
'==============================================================================


'------------------------------------------------------------------------------
' TEST CONFIGURATION
'------------------------------------------------------------------------------
    Private Const TEST_WAIT_SECONDS   As Double = 0.15                'Small UI settle delay after each state change
    Private Const TEST_ERR_BASE       As Long = vbObjectError + 4700  'Base custom error number for test assertions
    Private Const TEST_SNAPSHOT_ID_ERR_BASE As Long = vbObjectError + 4810  'Base custom error for snapshot-identity assertions
    Private Const TEST_TARGET_ERR_BASE As Long = vbObjectError + 4900  'Base custom error for target-scope tests
    Private Const TEST_TITLEBAR_SDI_ERR_BASE As Long = vbObjectError + 5000  'Base custom error for title-bar SDI tests
    Private Const TEST_CERT_ERR_BASE  As Long = vbObjectError + 5100  'Base custom error for certification
    Private Const TEST_RIBBON_ERR_BASE As Long = vbObjectError + 5200  'Base custom error for the Ribbon probe
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

        Private Declare PtrSafe Function TST_GetLastError Lib "kernel32" Alias "GetLastError" () As Long
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

        Private Declare Function TST_GetLastError Lib "kernel32" Alias "GetLastError" () As Long
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

    Private Const TST_SWP_NOSIZE            As Long = &H1
    Private Const TST_SWP_NOMOVE            As Long = &H2
    Private Const TST_SWP_NOZORDER          As Long = &H4
    Private Const TST_SWP_FRAMECHANGED      As Long = &H20
    Private Const TST_SWP_NOOWNERZORDER     As Long = &H200


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

'Ribbon characterization observations, accumulated during a probe run. Rows are
'plain text for reading and JSON for comparison between hosts.
Private m_RibbonRowCount                As Long
Private m_RibbonRowsText                As String
Private m_RibbonRowsJson                As String


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
' UPDATED
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

    Const PROC As String = "Test_EXCEL_UI_RunCertificationSelfTest"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
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

        If Not UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_CERT_ERR_BASE + 20, _
                PROC, _
                "unable to establish a snapshot; the precondition under test " & _
                "cannot be triggered"
        End If

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
    'Release the snapshot this case created
        On Error Resume Next
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
'   The existing runners cannot certify a release. Test_EXCEL_UI_RunAll executes
'   no multi-window case, silently skips the snapshot cases when a snapshot
'   already exists, and reports its outcome only as Immediate Window prose. A
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
'   - Verifies cleanup: no snapshot left behind, ScreenUpdating restored, the
'     workbook count back to its starting value, the anchor window still usable.
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
'   Dispatch is an explicit Select Case rather than Application.Run because the
'   units are private to this module and a name-based call would fail silently
'   or bind to the wrong project. A compile-time reference cannot rot.
'
' INPUTS
'   UnitName
'     Identifier of the unit to execute.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Does not raise. A failing unit is recorded and the run continues.
'
' DEPENDENCIES
'   - TST_RunRegressionPack
'   - Test_EXCEL_UI_RunSnapshotIdentity
'   - Test_EXCEL_UI_RunTitleBarSdiIdentity
'   - TST_CertRecordUnit
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log "TST_CertRunUnit", "UNIT", "Running " & UnitName

'------------------------------------------------------------------------------
' DISPATCH UNIT
'------------------------------------------------------------------------------
    'An unknown identifier is a defect in the registry above, not a pass
        Select Case UnitName

            Case "RegressionPack"
                TST_RunRegressionPack _
                    IncludeTitleBarTests:=True, _
                    CallerProc:="Certification.RegressionPack"

            Case "SnapshotIdentity"
                Test_EXCEL_UI_RunSnapshotIdentity

            Case "TitleBarSdiIdentity"
                Test_EXCEL_UI_RunTitleBarSdiIdentity

            Case Else
                Err.Raise _
                    TEST_CERT_ERR_BASE + 10, _
                    "TST_CertRunUnit", _
                    "unknown certification unit: " & UnitName

        End Select

'------------------------------------------------------------------------------
' RECORD PASS
'------------------------------------------------------------------------------
    'The unit completed without raising
        TST_CertRecordUnit UnitName, True, vbNullString

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Record the failure and continue; the verdict is assembled by the caller
        TST_CertRecordUnit UnitName, False, _
            CStr(Err.Number) & ": " & Err.Description

        Resume Safe_Exit

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
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
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
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
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
'   README.md states the Ribbon scope as "Excel application", which a reader will
'   take to mean one state shared by every workbook window. Under the Single
'   Document Interface each workbook window has its own Ribbon UI, and nothing
'   in the component verifies that the documented scope is the scope Excel
'   actually implements. The claim is currently an assumption.
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

    Const PROC As String = "Test_EXCEL_UI_RunRibbonSdiProbe"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_RibbonProbeReset

        TST_Log PROC, "START", "Ribbon SDI characterization started"

'------------------------------------------------------------------------------
' VALIDATE PRECONDITIONS
'------------------------------------------------------------------------------
    'The probe captures and restores a snapshot, so it cannot share the slot
        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_RIBBON_ERR_BASE + 1, _
                PROC, _
                "an explicit EXCEL_UI snapshot already exists; clear or " & _
                "restore it before probing"
        End If

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
        TST_WaitUI TEST_WAIT_SECONDS

        UI_SetExcelUI Ribbon:=UI_Hide
        TST_WaitUI TEST_WAIT_SECONDS

        BookB.Windows(1).Activate
        TST_WaitUI TEST_WAIT_SECONDS

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
    'Leave the host as it was found, whatever the observations were
        On Error Resume Next
            UI_ClearExcelUIStateSnapshot
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
'   This is NOT the release gate. It runs no multi-window case, it can skip
'   snapshot cases silently when a snapshot already exists, and it produces no
'   machine-readable evidence. Use Test_EXCEL_UI_RunReleaseCertification to
'   certify a release.
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

    Const PROC As String = "Test_EXCEL_UI_RunSnapshotIdentity"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", "Identity-safe structured snapshot test started"

        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_SNAPSHOT_ID_ERR_BASE + 1, _
                PROC, _
                "an explicit EXCEL_UI snapshot already exists; clear or restore it before running this destructive test"
        End If

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

        UI_ClearExcelUIStateSnapshot

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
'   - Runs the redirect case and the closed-frame case.
'   - Clears any snapshot and restores the title bar before returning.
'
' ERROR POLICY
'   - Raises after best-effort cleanup.
'   - Preserves the original test failure through cleanup.
'
' DEPENDENCIES
'   - TST_Case_TitleBarSdiRestoreTargetsCapturedFrame
'   - TST_Case_TitleBarCapturedFrameClosedIsReported
'
' NOTES
'   Destructive: creates and closes temporary workbooks. It is deliberately a
'   separate runner rather than part of Test_EXCEL_UI_RunAll, because RunAll is
'   currently non-destructive. Folding every mandatory case into one runner is
'   tracked separately as the release-certification work.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow        As Window          'Window active before the run
    Dim OldScreenUpdating   As Boolean         'Cached ScreenUpdating state

    Dim HasFailure          As Boolean         'TRUE when a test failure occurred
    Dim FailNumber          As Long            'Captured failure number
    Dim FailSource          As String          'Captured failure source
    Dim FailDescription     As String          'Captured failure description

    Const PROC As String = "Test_EXCEL_UI_RunTitleBarSdiIdentity"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_Log PROC, "START", "SDI title-bar identity test started"

    'Refuse to destroy a snapshot the caller is relying on
        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 1, _
                PROC, _
                "an explicit EXCEL_UI snapshot already exists; clear or restore it before running this destructive test"
        End If

    'Hold the window to return to, so the run leaves the host as it found it
        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 2, _
                PROC, _
                "no active Excel window is available"
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

    'Log successful completion before cleanup
        TST_Log PROC, "PASS", _
            "Restoration targeted the captured frame; a closed frame " & _
            "was reported"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Leave no snapshot behind, and leave the frame visible
        On Error Resume Next
            UI_ClearExcelUIStateSnapshot

            If Not AnchorWindow Is Nothing Then
                AnchorWindow.Activate
            End If

            UI_SetExcelUI TitleBar:=UI_Show
            Application.ScreenUpdating = OldScreenUpdating
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
' UPDATED
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

        End If

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

    'Run the dedicated title-bar cases when requested
        If IncludeTitleBarTests Then
            TST_Case_TitleBarRoundTrip
            TST_Case_TitleBarOwnedBitPreservation
            TST_Case_TitleBarShowRecoversWithoutBaseline
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
        On Error GoTo 0

    'Restore ScreenUpdating before leaving the harness
        Application.ScreenUpdating = OldScreenUpdating

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
'
' UPDATED
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

'------------------------------------------------------------------------------
' RUN REGRESSION CASE
'------------------------------------------------------------------------------
    'Run the dedicated title-bar round-trip case
        TST_Case_TitleBarRoundTrip

    'Verify the exact production merge policy with deterministic style values
        TST_Case_TitleBarOwnedBitPreservation

    'Verify that a show restores the frame with no captured baseline
        TST_Case_TitleBarShowRecoversWithoutBaseline

    'Verify that a failed frame refresh is retried rather than short-circuited
        TST_Case_TitleBarFrameRefreshDebtRetried

    'Verify that frame state which cannot be proved is discarded, not applied
        TST_Case_TitleBarStaleFrameEntryNotReused

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
        On Error GoTo 0

    'Restore ScreenUpdating before leaving the harness
        Application.ScreenUpdating = OldScreenUpdating

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
'
' ERROR POLICY
'   - Raises on assertion failure.
'
' DEPENDENCIES
'   - UI_ResetExcelUIToSnapshot_WithResult
'   - TST_AssertSingleFailurePrefix
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
'   - UI_HasExcelUIStateSnapshot
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
'   2026-07-25
'==============================================================================
'
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
' UPDATED
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
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
        TST_Log "TST_Case_TitleBarRoundTrip", "PASS", "Title-bar round-trip behaved as expected"

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
'   - Restores the entry style before re-raising, so one failure cannot leave
'     the host with a hidden title bar.
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
' RESTORE ENTRY STYLE
'------------------------------------------------------------------------------
    'Leave the host exactly as the case found it
        TST_RestoreTitleBarStyle xlHnd, EntryStyle, StyleCaptured

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Report the pass and exit before the error-handler block
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

    'Never leave the host with a hidden title bar because a case failed
        TST_RestoreTitleBarStyle xlHnd, EntryStyle, StyleCaptured

    'Hand the original failure to the pack handler
        Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription

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
'   - Records the second window's frame state, then restores.
'   - Asserts the anchor frame is restored and the second frame is unchanged.
'
' ERROR POLICY
'   - Raises after best-effort cleanup of the temporary workbook.
'
' UPDATED
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
#Else
    Dim AnchorHwnd          As Long            'Top-level frame of the anchor
    Dim SecondHwnd          As Long            'Top-level frame of the second
#End If

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

        UI_SetExcelUI TitleBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        AnchorHwnd = Application.hWnd

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

    'Record the second frame exactly as it stands before restoration
        SecondVisibleBefore = TST_TitleBarVisibleForHwndOrRaise(SecondHwnd, PROC)

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Restore while the WRONG window is active
        OK = UI_ResetExcelUIToSnapshot_WithResult( _
            FailureCount:=FailureCount, _
            FailureList:=FailureList)

        TST_WaitUI TEST_WAIT_SECONDS

        TST_AssertResultSuccess OK, FailureCount, FailureList, _
            "SdiRestoreTargetsCapturedFrame.Restore"

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

'------------------------------------------------------------------------------
' LOG PASS
'------------------------------------------------------------------------------
        TST_Log PROC, "PASS", _
            "Captured frame restored; the active frame was left unchanged"

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Release the snapshot and the temporary workbook before leaving
        On Error Resume Next
            UI_ClearExcelUIStateSnapshot
            TST_SafeCloseWorkbook SecondWorkbook

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
'   - Closes that workbook, destroying the captured frame.
'   - Restores and asserts a TitleBar failure is reported.
'   - Asserts the surviving window's frame is unchanged.
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
' UPDATED
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
#Else
    Dim AnchorHwnd          As Long            'Top-level frame of the anchor
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

    'Record the surviving frame exactly as it stands before restoration
        AnchorVisibleBefore = TST_TitleBarVisibleForHwndOrRaise(AnchorHwnd, PROC)

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
'   - Raises after best-effort restoration of the frame.
'
' NOTES
'   Uses the internal seams UI_InternalInjectFrameRefreshFailure and
'   UI_InternalIsFrameRefreshPending. They are Public only for same-project
'   regression access; Option Private Module keeps them out of the external
'   automation namespace.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim TargetHwnd          As LongPtr         'Frame under test
#Else
    Dim TargetHwnd          As Long            'Frame under test
#End If

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
' PREPARE A KNOWN FRAME STATE
'------------------------------------------------------------------------------
    'Start from a visible frame and an empty registry
        UI_SetExcelUI TitleBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        UI_InternalResetTitleBarBaseline

        TargetHwnd = Application.hWnd

        If TargetHwnd = 0 Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 30, _
                PROC, _
                "no Excel window handle is available"
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
    'Disarm the seam and put the frame back, whatever happened above
        On Error Resume Next
            UI_InternalInjectFrameRefreshFailure False
            UI_SetExcelUI TitleBar:=UI_Show
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
'   This is the regression for ICR-UI-111-P2-01. Windows reissues a window
'   handle once the window holding it has closed, and IsWindow answers for
'   whichever window holds the handle now, so a handle match was accepted as
'   proof of identity. A show could then restore a closed window's captured
'   frame onto an unrelated window that had merely inherited its handle.
'
'   A reissued handle cannot be forced on demand, so this case reproduces what
'   the registry actually sees: an entry claiming the frame is hidden while the
'   window it names carries owned bits the component never wrote. That is the
'   same evidence a reissued handle presents, and it is the evidence the fix
'   acts on.
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
'   - Raises after best-effort restoration of the frame.
'
' NOTES
'   Windows may normalise or reject individual GWL_STYLE bits, so the frame
'   written here is read back and the case asserts it is distinguishable from
'   the stale baseline before drawing any conclusion from it.
'
' UPDATED
'   2026-08-21
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim TargetHwnd          As LongPtr         'Frame under test
    Dim EntryStyle          As LongPtr         'Style to restore on the way out
    Dim ForeignStyle        As LongPtr         'Style written behind the module
    Dim ResultStyle         As LongPtr         'Style observed after the show
    Dim BaselineOwned       As LongPtr         'Owned bits a stale show writes
    Dim ForeignOwned        As LongPtr         'Owned bits actually accepted
#Else
    Dim TargetHwnd          As Long            'Frame under test
    Dim EntryStyle          As Long            'Style to restore on the way out
    Dim ForeignStyle        As Long            'Style written behind the module
    Dim ResultStyle         As Long            'Style observed after the show
    Dim BaselineOwned       As Long            'Owned bits a stale show writes
    Dim ForeignOwned        As Long            'Owned bits actually accepted
#End If

    Dim StyleCaptured       As Boolean         'TRUE once EntryStyle was read
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
' PREPARE A KNOWN FRAME STATE
'------------------------------------------------------------------------------
    'Start from a visible frame and an empty registry
        UI_SetExcelUI TitleBar:=UI_Show
        TST_WaitUI TEST_WAIT_SECONDS

        UI_InternalResetTitleBarBaseline

        TargetHwnd = Application.hWnd

        If TargetHwnd = 0 Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 60, _
                PROC, _
                "no Excel window handle is available"
        End If

    'Keep the entry style so the frame can be put back exactly
        If Not TST_TryGetWindowStyle(TargetHwnd, EntryStyle, Msg) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 61, _
                PROC, _
                "unable to read the entry window style | " & Msg
        End If

        StyleCaptured = True
        BaselineOwned = EntryStyle And TST_TITLEBAR_OWNED_MASK

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
            (EntryStyle And Not TST_TITLEBAR_OWNED_MASK) Or _
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
    'Put the frame back and drop the registry, whatever happened above
        On Error Resume Next
            TST_RestoreTitleBarStyle TargetHwnd, EntryStyle, StyleCaptured
            UI_InternalResetTitleBarBaseline
            UI_SetExcelUI TitleBar:=UI_Show
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
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume Safe_Exit

End Sub


#If VBA7 Then
Private Sub TST_RestoreTitleBarStyle( _
    ByVal hWnd As LongPtr, _
    ByVal EntryStyle As LongPtr, _
    ByVal StyleCaptured As Boolean)
#Else
Private Sub TST_RestoreTitleBarStyle( _
    ByVal hWnd As Long, _
    ByVal EntryStyle As Long, _
    ByVal StyleCaptured As Boolean)
#End If
'
'==============================================================================
' TST_RestoreTitleBarStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Restore a previously captured window style and refresh the non-client frame.
'
' WHY THIS EXISTS
'   The recovery case must return the host to its entry state on both the
'   success path and the failure path. Isolating the restore keeps those two
'   paths using identical logic, so a failing assertion can never leave the
'   user staring at a hidden title bar.
'
' INPUTS
'   hWnd
'     Excel main-window handle.
'
'   EntryStyle
'     Style value captured before the case modified the frame.
'
'   StyleCaptured
'     FALSE when the entry style was never read, in which case nothing is
'     written.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Writes the captured style and refreshes the frame.
'   - Does nothing when no entry style is available.
'
' ERROR POLICY
'   - Suppresses errors locally. This runs inside an active error handler and
'     must never raise.
'
' DEPENDENCIES
'   - TST_TrySetWindowStyle
'   - TST_TryRefreshWindowFrame
'
' CALLED FROM
'   - TST_Case_TitleBarShowRecoversWithoutBaseline
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IgnoredFailMsg      As String          'Discarded helper diagnostic

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'This may run inside an active error handler and must never raise
        On Error Resume Next

'------------------------------------------------------------------------------
' RESTORE STYLE
'------------------------------------------------------------------------------
    'Restore only when an entry style was actually captured
        If StyleCaptured And hWnd <> 0 Then
            TST_TrySetWindowStyle hWnd, EntryStyle, IgnoredFailMsg
            TST_TryRefreshWindowFrame hWnd, IgnoredFailMsg
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
'   - Handles Timer rollover at midnight
'
' UPDATED
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim t0                  As Double          'Timer baseline

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
'   - Raises on mismatch
'
' UPDATED
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
'   - Raises on mismatch.
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT FAILURE FLAG AND COUNT
'------------------------------------------------------------------------------
        If Succeeded Then
            Err.Raise TEST_ERR_BASE + 60, _
                      AssertionName, _
                      AssertionName & " expected=False actual=True"
        End If

        If FailureCount <> 1 Then
            Err.Raise TEST_ERR_BASE + 61, _
                      AssertionName, _
                      AssertionName & " expected FailureCount=1 actual=" & _
                      CStr(FailureCount)
        End If

'------------------------------------------------------------------------------
' ASSERT FAILURE LIST
'------------------------------------------------------------------------------
        If Not IsArray(FailureList) Then
            Err.Raise TEST_ERR_BASE + 62, _
                      AssertionName, _
                      AssertionName & " expected FailureList array"
        End If

        If LBound(FailureList) <> 1 Or UBound(FailureList) <> 1 Then
            Err.Raise TEST_ERR_BASE + 63, _
                      AssertionName, _
                      AssertionName & " expected one 1-based failure entry"
        End If

        If Left$(CStr(FailureList(1)), Len(ExpectedPrefix)) <> ExpectedPrefix Then
            Err.Raise TEST_ERR_BASE + 64, _
                      AssertionName, _
                      AssertionName & " expected prefix='" & ExpectedPrefix & _
                      "' actual='" & CStr(FailureList(1)) & "'"
        End If

End Sub


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
'   Uses GetLastError to distinguish a valid zero return from failure.
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastErr             As Long

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
        TST_SetLastError 0

#If VBA7 Then
    #If Win64 Then
        StyleOut = TST_GetWindowLongPtr(hWnd, TST_GWL_STYLE)
    #Else
        StyleOut = TST_GetWindowLong(hWnd, TST_GWL_STYLE)
    #End If
#Else
        StyleOut = TST_GetWindowLong(hWnd, TST_GWL_STYLE)
#End If

        LastErr = TST_GetLastError

        If StyleOut = 0 And LastErr <> 0 Then
            FailMsg = _
                "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)

            GoTo Safe_Exit
        End If

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
'   Uses GetLastError to distinguish a valid zero return from failure.
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim PreviousStyle       As LongPtr
#Else
    Dim PreviousStyle       As Long
#End If

    Dim LastErr             As Long

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
        TST_SetLastError 0

#If VBA7 Then
    #If Win64 Then
        PreviousStyle = _
            TST_SetWindowLongPtr(hWnd, TST_GWL_STYLE, NewStyle)
    #Else
        PreviousStyle = _
            TST_SetWindowLong(hWnd, TST_GWL_STYLE, NewStyle)
    #End If
#Else
        PreviousStyle = _
            TST_SetWindowLong(hWnd, TST_GWL_STYLE, NewStyle)
#End If

        LastErr = TST_GetLastError

        If PreviousStyle = 0 And LastErr <> 0 Then
            FailMsg = _
                "SetWindowLong/SetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)

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
'   Returns FALSE and FailMsg on failure.
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ApiOK               As Long
    Dim LastErr             As Long

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

        LastErr = TST_GetLastError

        If ApiOK = 0 Then
            FailMsg = _
                "SetWindowPos failed; GetLastError=" & CStr(LastErr)

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
' UPDATED
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
    Dim LastErr             As Long            'Last Win32 error after API call

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
'   - Raises a descriptive assertion failure.
'
' CALLED FROM
'   - TST_Case_TitleBarCapturedFrameClosedIsReported
'
' UPDATED
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
                "expected a structured failure but the operation reported success"
        End If

        If FailureCount < 1 Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 51, _
                AssertionName, _
                "expected at least one ordered failure entry but " & _
                "FailureCount was " & CStr(FailureCount)
        End If

        If Not IsArray(FailureList) Then
            Err.Raise _
                TEST_TITLEBAR_SDI_ERR_BASE + 52, _
                AssertionName, _
                "expected an ordered failure list but no array was returned"
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
                "prefix " & ExpectedPrefix
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
