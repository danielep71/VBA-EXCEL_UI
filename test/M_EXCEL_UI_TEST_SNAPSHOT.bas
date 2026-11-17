Attribute VB_Name = "M_EXCEL_UI_TEST_SNAPSHOT"
Option Explicit
Option Private Module

'==============================================================================
' M_EXCEL_UI_TEST_SNAPSHOT
'------------------------------------------------------------------------------
' PURPOSE
'   Hold the snapshot regression cases of EXCEL_UI, the snapshot-identity
'   runner and the Ribbon SDI probe
'
' WHY THIS EXISTS
'   The explicit snapshot / reset lifecycle must restore only state that still
'   belongs to the window it was captured from. These cases cover the
'   structured results of capture and restore, partial capture, the lifecycle
'   itself and Ribbon restoration across windows
'
' PUBLIC SURFACE
'   - Test_EXCEL_UI_RunSnapshotIdentity
'   - Test_EXCEL_UI_RunRibbonSdiProbe
'
'   The cases are Public so that the regression pack in
'   M_EXCEL_UI_TEST_CERTIFICATION and the SDI identity runner in
'   M_EXCEL_UI_TEST_TITLEBAR can run them
'
' TEST SCOPE
'   Structured-result tests
'     - snapshot capture clean-success result path
'     - snapshot restoration clean-success result path
'     - snapshot restoration no-snapshot failure path
'
'   Snapshot / restore tests
'     - capture that cannot read one application-level property
'     - explicit snapshot lifecycle
'     - reset without snapshot leaves managed UI unchanged and logs a diagnostic
'     - lifecycle cases are skipped when an explicit EXCEL_UI snapshot already
'       existed before the run because the harness cannot reconstruct that prior
'       module-level snapshot object
'
'   Snapshot-identity runner
'     - identity-safe restoration of surviving captured windows
'     - closed captured-window ordered failure path
'     - replacement-window non-interference
'
'   Ribbon restoration
'     - restore refuses the Ribbon while another window is active
'     - restore reports a Ribbon whose window has closed
'
'   Ribbon SDI probe
'     - records how Ribbon visibility behaves across workbook windows; it
'       asserts nothing
'
' STATE MANAGEMENT
'   - The snapshot-identity runner creates and closes temporary windows, and
'     the Ribbon probe temporary workbooks, only after confirming that no
'     explicit EXCEL_UI snapshot already exists
'   - The probe's observation rows are module-level state reset at the start
'     of each probe run
'
' DEPENDENCIES
'   - M_EXCEL_UI_TEST_HARNESS
'
' UPDATED
'   2026-11-15 - Split out of the single regression module; the runners,
'                cases and probe state moved here unchanged apart from
'                visibility (#90).
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.2
'==============================================================================


'==============================================================================
' RIBBON PROBE STATE
'==============================================================================

'Ribbon characterization observations, accumulated during a probe run. Rows are
'plain text for reading and JSON for comparison between hosts.
Private m_RibbonRowCount                As Long
Private m_RibbonRowsText                As String
Private m_RibbonRowsJson                As String


'
'------------------------------------------------------------------------------
'
'                              PUBLIC RUNNERS
'
'------------------------------------------------------------------------------
'


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


'
'------------------------------------------------------------------------------
'
'                           REGRESSION CASES
'
'------------------------------------------------------------------------------
'


Public Sub TST_Case_SnapshotCaptureResultSuccess( _
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


Public Sub TST_Case_SnapshotResetResultSuccess( _
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


Public Sub TST_Case_SnapshotResetResultNoSnapshot( _
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


Public Sub TST_Case_SnapshotCapturePartialApplicationRead()
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


Public Sub TST_Case_SnapshotLifecycle(ByVal IncludeTitleBarTests As Boolean)

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


Public Sub TST_Case_ResetWithoutSnapshot_NoOp(ByVal IncludeTitleBarTests As Boolean)

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


Public Sub TST_Case_RibbonRestoreRefusesOtherActiveWindow()

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


Public Sub TST_Case_RibbonRestoreRefusesClosedWindow()

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
