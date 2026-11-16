Attribute VB_Name = "M_EXCEL_UI_TEST_HARNESS"
Option Explicit
Option Private Module

'==============================================================================
' M_EXCEL_UI_TEST_HARNESS
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the shared harness of the EXCEL_UI regression suites: assertions,
'   state capture and cleanup, Excel and WinAPI read/write helpers, logging,
'   and the release-certification records and evidence files
'
' WHY THIS EXISTS
'   The runtime, snapshot, title-bar and certification suites each live in
'   their own module, so that a change to one suite is reviewed and
'   re-imported on its own. What they share lives here once: the helpers, the
'   configuration constants and error-number bases, and the module-level
'   state that more than one suite reads or writes
'
' PUBLIC SURFACE
'   None of its own. Members that a suite module calls are Public so that the
'   other modules of the same VBA project can reach them; Option Private
'   Module keeps them out of other projects
'
' CONTENTS
'   - Fault injection: the one-shot post-capture fault that
'     Test_EXCEL_UI_RunOwnershipCleanupChecks arms and the runner it checks
'     takes
'   - Certification records and evidence: unit outcomes, skips, snapshots a
'     unit left behind, and the JSON and text evidence files
'   - Cleanup helpers: entry title-bar style restoration, cleanup-failure
'     reporting, and best-effort window and workbook closing
'   - Snapshot / restore helpers: the harness's own capture and restore of the
'     managed Excel UI state, and the capped settle delay
'   - Assertion helpers, including the structured-result assertions that
'     report the failure count and list they rejected
'   - State read/write helpers, including the GWL_STYLE read and write and the
'     frame refresh, which capture Err.LastDllError immediately after each call
'   - Diagnostics and timing: logging, frame-style probes and error text
'
' STATE MANAGEMENT
'   - Certification records accrue only while a certification run is active
'   - The post-capture fault is one-shot: taking it disarms it
'   - Per-window test state is captured and restored by Application.Windows
'     index
'   - Entry frame styles are put back through the harness WinAPI helpers,
'     never through the API under test, and a failed restoration is reported
'
' LIMITATIONS
'   - Ribbon visibility is read using best-effort logic
'   - Window-level capture and restore use Application.Windows index order; a
'     changed collection can prevent exact restoration to the original windows
'   - The title-bar helpers are the most OS / Excel-version-sensitive area
'
' COMPATIBILITY
'   - Windows only for the title-bar helpers
'   - Assumes the EXCEL_UI modules are present in the same VBA project
'
' UPDATED
'   2026-11-15 - Split out of the single regression module: the shared
'                configuration, helpers and certification records moved here
'                unchanged apart from visibility (#90).
'   2026-11-06 - TST_RestoreTitleBarStyle reports and verifies the restore.
'                TST_WaitUI is capped (#66).
'   2026-11-06 - The title-bar WinAPI helpers capture Err.LastDllError
'                immediately after each call and no longer declare
'                GetLastError (#27).
'   2026-11-06 - Structured-result assertions report the failure count and
'                list they rejected (#135).
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
    Public Const TEST_WAIT_SECONDS   As Double = 0.15                'Small UI settle delay after each state change
    Private Const TEST_WAIT_MAX_SECONDS As Double = 5#                'Upper bound on any one settle delay
    Public Const TEST_ERR_BASE       As Long = vbObjectError + 4700  'Base custom error number for test assertions
    Public Const TEST_SNAPSHOT_ID_ERR_BASE As Long = vbObjectError + 4810  'Base custom error for snapshot-identity assertions
    Public Const TEST_TARGET_ERR_BASE As Long = vbObjectError + 4900  'Base custom error for target-scope tests
    Public Const TEST_TITLEBAR_SDI_ERR_BASE As Long = vbObjectError + 5000  'Base custom error for title-bar SDI tests
    Public Const TEST_CERT_ERR_BASE  As Long = vbObjectError + 5100  'Base custom error for certification
    Public Const TEST_RIBBON_ERR_BASE As Long = vbObjectError + 5200  'Base custom error for the Ribbon probe
    Public Const TEST_OWNERSHIP_ERR_BASE As Long = vbObjectError + 5300  'Base custom error for the snapshot-ownership case

    Public Const TEST_PAIR_FAULT_NONE As Long = 0  'Frame pair: no fault
    Public Const TEST_PAIR_FAULT_MISMATCH As Long = 1  'Frame pair: sampled handle disagrees

    Public Const TEST_FAULT_NONE As Long = 0  'No post-capture fault armed
    Public Const TEST_FAULT_ASSERTION As Long = 1  'Fail an assertion after capture
    Public Const TEST_FAULT_UNEXPECTED As Long = 2  'Raise an unexpected error after capture
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
    Public Const TST_WS_CAPTION            As Long = &HC00000
    Public Const TST_TITLEBAR_OWNED_MASK   As Long = &HCF0000

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
            Public Const TST_API_GET_WINDOW_STYLE As String = "GetWindowLongPtrA"
            Public Const TST_API_SET_WINDOW_STYLE As String = "SetWindowLongPtrA"
        #Else
            Public Const TST_API_GET_WINDOW_STYLE As String = "GetWindowLongA"
            Public Const TST_API_SET_WINDOW_STYLE As String = "SetWindowLongA"
        #End If
    #Else
        Public Const TST_API_GET_WINDOW_STYLE As String = "GetWindowLongA"
        Public Const TST_API_SET_WINDOW_STYLE As String = "SetWindowLongA"
    #End If


'==============================================================================
' RELEASE-CERTIFICATION STATE
'==============================================================================

'Accrued only while a certification run is active. The regression packs are
'shared with the legacy runners, so recording must be inert outside a run rather
'than leaking counts from one invocation into the next.
Public m_CertActive                    As Boolean

Public m_CertUnitCount                 As Long
Private m_CertUnitNames()               As String
Public m_CertUnitPassed()              As Boolean
Private m_CertUnitDetail()              As String

Public m_CertSkipCount                 As Long
Private m_CertSkipDetail()              As String

'Snapshots that units left behind, named by unit. Each one is released before
'the next unit runs, so that one failure cannot fail every unit after it. The
'exit state therefore no longer shows the leak, and the cleanup verdict reads
'it from here instead.
Public m_CertLeakDetail                As String

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


'
'------------------------------------------------------------------------------
'
'                             FAULT INJECTION
'
'------------------------------------------------------------------------------
'


Public Sub TST_InjectPostCapture(ByVal FaultKind As Long)

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


Public Function TST_TakePostCaptureFault() As Long

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


'
'------------------------------------------------------------------------------
'
'                    CERTIFICATION RECORDS AND EVIDENCE
'
'------------------------------------------------------------------------------
'


Public Sub TST_CertResetCounters()

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


Public Sub TST_CertRecordSkip( _
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


Public Sub TST_CertRecordUnit( _
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


Public Function TST_CertAppendDetail( _
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


Public Function TST_CertBuildJsonEvidence( _
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


Public Function TST_CertJsonEscape( _
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


Public Function TST_CertBuildTextReport( _
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


Public Function TST_CertTryWriteEvidence( _
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


'
'------------------------------------------------------------------------------
'
'                             CLEANUP HELPERS
'
'------------------------------------------------------------------------------
'


#If VBA7 Then
Public Function TST_RestoreTitleBarStyle( _
    ByVal hWnd As LongPtr, _
    ByVal EntryStyle As LongPtr, _
    ByVal StyleCaptured As Boolean, _
    ByRef FailMsg As String) As Boolean
#Else
Public Function TST_RestoreTitleBarStyle( _
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
'   - TST_Case_TitleBarSameHwndGenerationReuse
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


Public Sub TST_RecordCleanupFailure( _
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


Public Sub TST_SafeCloseWorkbook(ByRef TargetWorkbook As Workbook)

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


Public Sub TST_SafeCloseWindow(ByRef TargetWindow As Window)

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


'
'------------------------------------------------------------------------------
'
'                      SNAPSHOT / RESTORE HELPERS
'
'------------------------------------------------------------------------------
'


Public Sub TST_SnapshotState( _
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


Public Sub TST_RestoreState( _
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


Public Sub TST_WaitUI(ByVal SecondsToWait As Double)

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


Public Sub TST_AssertBooleanEquals( _
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


Public Sub TST_AssertApplicationProperty( _
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


Public Sub TST_AssertAllWindowsProperty( _
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


Public Sub TST_AssertRibbonVisible( _
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


Public Sub TST_AssertTitleBarVisible( _
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


Public Sub TST_AssertResultSuccess( _
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


Public Sub TST_AssertSingleFailurePrefix( _
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


Public Sub TST_AssertFailureResultExact( _
    ByVal Succeeded As Boolean, _
    ByVal FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal ExpectedCount As Long, _
    ByVal ExpectedEntry As String, _
    ByVal AssertionName As String)

'
'==============================================================================
' TST_AssertFailureResultExact
'------------------------------------------------------------------------------
' PURPOSE
'   Assert a failed structured result exactly: the flag, the count, the shape
'   of the list and the text of its only entry.
'
' WHY THIS EXISTS
'   The degraded-diagnostics contract is defined by those outputs together. A
'   prefix or IsArray check would accept a list holding stale, partial or
'   foreign data.
'
' INPUTS
'   Succeeded / FailureCount / FailureList
'     Structured-result outputs to validate.
'
'   ExpectedCount
'     Exact FailureCount required.
'
'   ExpectedEntry
'     Exact text of the single list entry, or vbNullString when the list must
'     be Empty.
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
' DEPENDENCIES
'   - TST_DescribeResultBuffers
'
' UPDATED
'   2026-11-14
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSERT FAILURE FLAG AND COUNT
'------------------------------------------------------------------------------
        If Succeeded Then
            Err.Raise TEST_ERR_BASE + 90, _
                      AssertionName, _
                      AssertionName & " expected=False actual=True | " & _
                      TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

        If FailureCount <> ExpectedCount Then
            Err.Raise TEST_ERR_BASE + 91, _
                      AssertionName, _
                      AssertionName & " expected FailureCount=" & _
                      CStr(ExpectedCount) & " actual=" & CStr(FailureCount) & _
                      " | " & TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

'------------------------------------------------------------------------------
' ASSERT AN EMPTY LIST
'------------------------------------------------------------------------------
        If Len(ExpectedEntry) = 0 Then

            If VarType(FailureList) <> vbEmpty Then
                Err.Raise TEST_ERR_BASE + 92, _
                          AssertionName, _
                          AssertionName & " expected FailureList=Empty | " & _
                          TST_DescribeResultBuffers(FailureCount, FailureList)
            End If

'------------------------------------------------------------------------------
' ASSERT A SINGLE EXACT ENTRY
'------------------------------------------------------------------------------
        Else

            If VarType(FailureList) <> (vbArray Or vbString) Then
                Err.Raise TEST_ERR_BASE + 93, _
                          AssertionName, _
                          AssertionName & " expected a String array | " & _
                          TST_DescribeResultBuffers(FailureCount, FailureList)
            End If

            If LBound(FailureList) <> 1 Or UBound(FailureList) <> 1 Then
                Err.Raise TEST_ERR_BASE + 94, _
                          AssertionName, _
                          AssertionName & " expected one 1-based failure entry | " & _
                          TST_DescribeResultBuffers(FailureCount, FailureList)
            End If

            If StrComp(FailureList(1), ExpectedEntry, vbBinaryCompare) <> 0 Then
                Err.Raise TEST_ERR_BASE + 95, _
                          AssertionName, _
                          AssertionName & " expected entry='" & ExpectedEntry & _
                          "' actual='" & CStr(FailureList(1)) & "'"
            End If

        End If

End Sub


Public Function TST_DescribeResultBuffers( _
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
'   - TST_AssertFailureResultExact
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


Public Sub TST_AssertSnapshotAvailability( _
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


Public Sub TST_AssertRefusal( _
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


Public Sub TST_AssertTrue( _
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


Public Function TST_TryGetRibbonVisible( _
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
Public Function TST_TryGetWindowStyle( _
    ByVal hWnd As LongPtr, _
    ByRef StyleOut As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Public Function TST_TryGetWindowStyle( _
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
Public Function TST_TrySetWindowStyle( _
    ByVal hWnd As LongPtr, _
    ByVal NewStyle As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Public Function TST_TrySetWindowStyle( _
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
Public Function TST_TryRefreshWindowFrame( _
    ByVal hWnd As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Public Function TST_TryRefreshWindowFrame( _
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
Public Function TST_TitleBarVisibleForHwndOrRaise( _
    ByVal TargetHwnd As LongPtr, _
    ByVal CallerProc As String) As Boolean
#Else
Public Function TST_TitleBarVisibleForHwndOrRaise( _
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
Public Function TST_FrameStyleOrRaise( _
    ByVal TargetHwnd As LongPtr, _
    ByVal CallerProc As String) As LongPtr
#Else
Public Function TST_FrameStyleOrRaise( _
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
Public Sub TST_AssertFrameUntouched( _
    ByVal TargetHwnd As LongPtr, _
    ByVal StyleBefore As LongPtr, _
    ByVal AssertionName As String)
#Else
Public Sub TST_AssertFrameUntouched( _
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


Public Sub TST_AssertFailureListContainsPrefix( _
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


Public Sub TST_Log( _
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


Public Function TST_TitleBarMode( _
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


Public Sub TST_AssertSnapshotWindowState( _
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


#If VBA7 Then
Public Function TST_BuildApiFailureText( _
    ByVal ApiName As String, _
    ByVal Operation As String, _
    ByVal DllError As Long, _
    ByVal TargetHwnd As LongPtr) As String
#Else
Public Function TST_BuildApiFailureText( _
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
