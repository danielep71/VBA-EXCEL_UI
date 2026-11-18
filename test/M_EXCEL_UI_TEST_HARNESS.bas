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
'     takes, and the one-shot certification record fault
'   - Certification records and evidence: one record of unit outcomes,
'     skips, cleanup findings and unverifiable surfaces that publishes an
'     entry only once it is complete, snapshots a unit left behind, the
'     verdict, and the JSON and text evidence files
'   - Cleanup helpers: entry title-bar style restoration, cleanup-failure
'     reporting, and best-effort window and workbook closing
'   - Snapshot / restore helpers: the harness's own capture and restore of the
'     managed Excel UI state, which reports every step it could not complete
'     as a structured finding, and the capped settle delay
'   - Assertion helpers, including the structured-result assertions that
'     report the failure count and list they rejected
'   - State read/write helpers, including the GWL_STYLE read and write and the
'     frame refresh, which capture Err.LastDllError immediately after each call
'   - Diagnostics and timing: logging, the opt-in log sink that mirrors
'     every logged line to a UTF-8 file, frame-style probes and error text
'
' STATE MANAGEMENT
'   - Certification records accrue only while a certification run is active
'   - A certification entry is counted only after every buffer holds it. An
'     entry that cannot be stored sets the record's scalar RecordFailed
'     flag, which forces FAIL | INCOMPLETE and is named in both evidence
'     documents; TST_CertResetCounters clears it before the next run
'   - A cleanup finding is an entry of the same record, so a difference that
'     cannot be stored fails the run, and any finding makes cleanup FAILED
'     whatever the caller passes to the composer
'   - The record fault is one-shot, and certification disarms it on start
'   - The run fault is one-shot: the next certification call takes it
'   - The post-capture fault is one-shot: taking it disarms it
'   - Per-window test state is captured and restored by Application.Windows
'     index. The certification cleanup baseline does not use this capture: it
'     keeps its own state, keyed by the retained Window objects
'   - The log sink is off unless TST_LogSinkStart names a file; with no
'     sink TST_Log prints exactly what it printed before the sink existed
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
'   2026-11-18 - Certification records cleanup findings and unverifiable
'                surfaces through the same record, composes cleanup=FAILED from
'                any finding, and reports them in both documents. TST_RestoreState
'                returns structured findings. Added a one-shot run fault
'                that forces a cleanup failure (#35).
'   2026-11-16 - TST_Log mirrors each line to an opt-in UTF-8 log sink,
'                and a failed write is kept for the caller's verdict
'                instead of being raised (#197).
'   2026-11-16 - Certification records units and skips through one
'                commit-after-allocation path, checks the record before
'                reading it, and fails closed when it cannot be stored (#50).
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

    'Kinds of certification record entry. A case kind is added here when
    'certification records cases (#42); TST_CertTryAppendEntry stores every kind.
    'A finding is a difference between the exit state and the entry baseline and
    'fails cleanup. An unverified entry names a surface cleanup could not compare
    'and does not fail it.
    Public Const TEST_CERT_ENTRY_UNIT As Long = 1  'Entry: one unit outcome
    Public Const TEST_CERT_ENTRY_SKIP As Long = 2  'Entry: one skipped case
    Public Const TEST_CERT_ENTRY_FINDING As Long = 3  'Entry: one cleanup finding
    Public Const TEST_CERT_ENTRY_UNVERIFIED As Long = 4  'Entry: one surface cleanup could not compare

    'The parallel buffers of a certification record, in the order they grow.
    'TST_CertRecordIsConsistent walks them from the first to the last.
    Public Const TEST_CERT_BUFFER_ANY As Long = 0  'Record fault: the first buffer reached
    Public Const TEST_CERT_BUFFER_KIND As Long = 1  'Buffer: entry kinds
    Public Const TEST_CERT_BUFFER_NAME As Long = 2  'Buffer: entry names
    Public Const TEST_CERT_BUFFER_PASSED As Long = 3  'Buffer: unit outcomes
    Public Const TEST_CERT_BUFFER_DETAIL As Long = 4  'Buffer: entry detail

    'Where recording first failed. Stages 1 to 4 are the TEST_CERT_BUFFER_*
    'value of the buffer that could not grow.
    Public Const TEST_CERT_RECFAIL_NONE As Long = 0  'Recording has not failed
    Public Const TEST_CERT_RECFAIL_WRITE As Long = 5  'Writing a reserved entry failed
    Public Const TEST_CERT_RECFAIL_INVARIANT As Long = 6  'Buffers disagree with the counts
    Public Const TEST_CERT_RECFAIL_EVIDENCE As Long = 7  'Reading entries for evidence failed
    Public Const TEST_CERT_RECFAIL_KIND As Long = 8  'An entry of no known kind was offered
    Public Const TEST_CERT_RECFAIL_RESET As Long = 9  'The record could not be cleared

    Public Const TEST_CERT_RECORD_FAULT_NONE As Long = 0  'No record fault armed
    Public Const TEST_CERT_RECORD_FAULT_FIRST As Long = 1  'Fail a buffer's first allocation
    Public Const TEST_CERT_RECORD_FAULT_GROWTH As Long = 2  'Fail an allocated buffer's growth

    Public Const TEST_CERT_RUN_FAULT_NONE As Long = 0  'No run fault armed
    Public Const TEST_CERT_RUN_FAULT_CLEANUP As Long = 1  'Run no unit; alter the host before cleanup is verified
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

'One certification record: every stored entry and the counts that publish
'them. Units and skips share the same parallel buffers and the same append
'path, TST_CertTryAppendEntry, so the rule below is enforced in one place.
'
'EntryCount is the logical length of every buffer. A buffer may be larger,
'because an append that grew some buffers and then failed leaves them grown,
'but an entry exists only once EntryCount includes it, and EntryCount grows
'only after every buffer holds the complete entry. The invariants are
'
'    UnitCount + SkipCount + FindingCount + UnverifiedCount = EntryCount >= 0
'    every buffer is unallocated or 1-based, with UBound >= EntryCount
'
'Cleanup findings and unverified surfaces (#35) are entries of the same record,
'stored by the same path, so
'
'    UnitCount + SkipCount + FindingCount + UnverifiedCount = EntryCount
'
'A finding has its identifier in EntryName and its detail in EntryDetail, and
'any finding makes the cleanup verdict FAILED. CheckCount is the number of
'comparisons cleanup completed.
'
'RecordFailed is the authoritative result when an entry cannot be stored. It
'is a scalar, set before anything else is allocated, and it forces the
'verdict to FAIL | INCOMPLETE whatever the buffers hold. RecordFailCount,
'RecordFailStage and RecordFailError are the diagnostic that survives when
'no text can be composed.
Public Type tCertRecord
    EntryCount          As Long                'Published entries
    UnitCount           As Long                'Published unit entries
    SkipCount           As Long                'Published skip entries
    FindingCount        As Long                'Published cleanup findings
    UnverifiedCount     As Long                'Published unverified surfaces
    CheckCount          As Long                'Comparisons cleanup completed
    EntryKind()         As Long                'TEST_CERT_ENTRY_* per entry
    EntryName()         As String              'Unit name, or the skipping procedure
    EntryPassed()       As Boolean             'Unit outcome; FALSE for a skip
    EntryDetail()       As String              'Unit failure text, or the skip reason
    RecordFailed        As Boolean             'TRUE once an entry was not stored
    RecordFailCount     As Long                'Recording failures since the reset
    RecordFailStage     As Long                'TEST_CERT_RECFAIL_* of the first one
    RecordFailError     As Long                'Runtime error number of the first one
End Type

'Accrued only while a certification run is active. The regression packs are
'shared with the legacy runners, so recording must be inert outside a run rather
'than leaking counts from one invocation into the next.
Public m_CertActive                    As Boolean

'The record of the certification run in progress. Only the TST_Cert* record
'helpers write it; a regression case drives the same helpers on a record of
'its own, so that its injected failures never reach this one.
Public m_CertRecord                    As tCertRecord

'Snapshots that units left behind, named by unit. Each one is released before
'the next unit runs, so that one failure cannot fail every unit after it. The
'exit state therefore no longer shows the leak, so the runner also records each
'one as a cleanup finding. This text is kept for the unit-isolation case.
Public m_CertLeakDetail                As String

'One-shot record fault, consumed by the next buffer allocation it matches.
'Only TST_CertInjectRecordFault arms it, and only the record regression case
'calls that. TST_CertResetCounters disarms it, so a fault left armed cannot
'reach the next certification run.
Private m_CertRecordFault               As Long
Private m_CertRecordFaultBuffer         As Long

'One-shot run fault, consumed by the next call of
'Test_EXCEL_UI_RunReleaseCertification whether or not it fires. Only
'TST_CertInjectRunFault arms it, and only the ownership-cleanup runner calls
'that. It exists so that a cleanup failure can be forced and the outer call
'seen to fail with it (#35).
Private m_CertRunFault                  As Long

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

'Opt-in log sink. While a path is set, TST_Log appends every line it prints
'to that file as UTF-8, so that a script, which cannot read the Immediate
'window, can read the run. Only TST_LogSinkStart sets it. Empty means no
'sink, and TST_Log then prints exactly what it printed before. The first
'write failure is kept and ends writing until the next start; TST_LogSinkStop
'returns it for the caller's verdict, and nothing is raised.
Private m_LogSinkPath                   As String
Private m_LogSinkFailure                As String


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


Public Sub TST_CertInjectRecordFault( _
    ByVal FaultKind As Long, _
    ByVal BufferId As Long)

'
'==============================================================================
' TST_CertInjectRecordFault
'------------------------------------------------------------------------------
' PURPOSE
'   Arm a one-shot allocation failure for the next certification record
'   buffer allocation that the fault matches.
'
' WHY THIS EXISTS
'   A certification record must fail closed when a buffer cannot be allocated
'   (#50). Excel cannot be made to run out of memory on demand, so the path is
'   unreachable without injecting the failure, and the three shapes it must
'   survive differ: no buffer allocated yet, every buffer already allocated,
'   and some parallel buffers grown while another was not.
'
' INPUTS
'   FaultKind
'     TEST_CERT_RECORD_FAULT_NONE disarms.
'     TEST_CERT_RECORD_FAULT_FIRST fails the next first allocation of a
'     buffer that holds nothing yet.
'     TEST_CERT_RECORD_FAULT_GROWTH fails the next growth of a buffer that is
'     already allocated.
'
'   BufferId
'     TEST_CERT_BUFFER_ANY fails the first matching allocation of any buffer.
'     A TEST_CERT_BUFFER_* value fails only that buffer, after the buffers
'     before it have grown.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Cannot raise.
'
' CALLED FROM
'   - TST_CertCheckRecordFault
'   - TST_Case_CertRecordFailsClosed
'
' NOTES
'   The fault is consumed only by an allocation it matches: a FIRST fault is
'   not consumed by a growth, and a GROWTH fault is not consumed by a first
'   allocation. TST_CertResetCounters disarms it.
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

        m_CertRecordFault = FaultKind
        m_CertRecordFaultBuffer = BufferId

End Sub


Private Function TST_CertTakeRecordFault( _
    ByVal BufferId As Long, _
    ByVal Allocated As Boolean) _
    As Boolean

'
'==============================================================================
' TST_CertTakeRecordFault
'------------------------------------------------------------------------------
' PURPOSE
'   Report whether the armed record fault matches this allocation, and disarm
'   it in the same step when it does.
'
' WHY THIS EXISTS
'   Reading and clearing separately would leave the fault armed on a path
'   that read it and then failed, and the only reader fails on purpose
'   immediately after reading it.
'
' INPUTS
'   BufferId
'     TEST_CERT_BUFFER_* value of the buffer about to be allocated.
'
'   Allocated
'     TRUE when the buffer already holds entries and is about to grow.
'
' RETURNS
'   Boolean
'     TRUE when the allocation must fail; the fault is then disarmed.
'
' ERROR POLICY
'   - Cannot raise.
'
' CALLED FROM
'   - TST_CertReserveBuffer
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Fire                As Boolean         'TRUE when the fault matches

'------------------------------------------------------------------------------
' MATCH FAULT
'------------------------------------------------------------------------------
    'A fault matches by allocation shape first, then by buffer
        Fire = False

        Select Case m_CertRecordFault

            Case TEST_CERT_RECORD_FAULT_FIRST
                Fire = Not Allocated

            Case TEST_CERT_RECORD_FAULT_GROWTH
                Fire = Allocated

        End Select

        If Fire Then
            If m_CertRecordFaultBuffer <> TEST_CERT_BUFFER_ANY Then
                Fire = (m_CertRecordFaultBuffer = BufferId)
            End If
        End If

'------------------------------------------------------------------------------
' DISARM AND RETURN
'------------------------------------------------------------------------------
    'One-shot: a fault that fired cannot fire again
        If Fire Then
            m_CertRecordFault = TEST_CERT_RECORD_FAULT_NONE
            m_CertRecordFaultBuffer = TEST_CERT_BUFFER_ANY
        End If

        TST_CertTakeRecordFault = Fire

End Function


Public Sub TST_CertInjectRunFault( _
    ByVal FaultKind As Long)

'
'==============================================================================
' TST_CertInjectRunFault
'------------------------------------------------------------------------------
' PURPOSE
'   Arm a one-shot fault for the next call of
'   Test_EXCEL_UI_RunReleaseCertification.
'
' WHY THIS EXISTS
'   A cleanup failure must fail the outer certification call with its
'   diagnostic fields intact (#35). Certification leaves the host as it found
'   it, so that outcome is unreachable without forcing it, and a full
'   destructive run would be a poor way to check one comparison.
'
' INPUTS
'   FaultKind
'     TEST_CERT_RUN_FAULT_NONE disarms.
'     TEST_CERT_RUN_FAULT_CLEANUP makes the run execute no unit, invert
'     Application.DisplayStatusBar after the entry baseline is captured, and
'     write no evidence file. Cleanup verification, the evidence and the raise
'     are the real ones.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Cannot raise.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunOwnershipCleanupChecks
'
' NOTES
'   The fault is consumed by the next certification call whether or not it
'   fires, so a fault left armed cannot reach a later run. The caller restores
'   the status bar itself; the run under fault does not.
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        m_CertRunFault = FaultKind

End Sub


Public Function TST_CertTakeRunFault() As Long

'
'==============================================================================
' TST_CertTakeRunFault
'------------------------------------------------------------------------------
' PURPOSE
'   Return the armed run fault and disarm it in the same step.
'
' WHY THIS EXISTS
'   Reading and clearing separately would leave the fault armed on a path that
'   read it and then failed, and the only reader acts on it at once.
'
' RETURNS
'   The armed TEST_CERT_RUN_FAULT_* value, or TEST_CERT_RUN_FAULT_NONE.
'
' ERROR POLICY
'   - Cannot raise.
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        TST_CertTakeRunFault = m_CertRunFault
        m_CertRunFault = TEST_CERT_RUN_FAULT_NONE

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
'   Clear the certification record, the leak record and the record fault
'   before a run.
'
' WHY THIS EXISTS
'   The regression packs are shared with the legacy runners, so skip recording
'   is module state rather than a parameter. Without an explicit reset a second
'   certification run in the same session would inherit the first run's counts
'   and report a verdict about work it did not do.
'
'   A run that follows a recording failure must also start clean: the fatal
'   flag of the earlier run is cleared, and a record fault left armed is
'   disarmed, so neither can decide the next verdict (#50).
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Disarms the record fault.
'   - Clears the leak record.
'   - Clears m_CertRecord through TST_CertResetRecord, which verifies that
'     every buffer is released and flags the record when one is not.
'
' ERROR POLICY
'   - Does not raise. A record that cannot be cleared is flagged as a
'     recording failure, which forces FAIL | INCOMPLETE.
'
' DEPENDENCIES
'   - TST_CertResetRecord
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'
' UPDATED
'   2026-11-16 - Clears the record through TST_CertResetRecord and disarms
'                the record fault (#50).
'   2026-11-06 - Clears the record of snapshots units left behind (#135).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR COUNTERS
'------------------------------------------------------------------------------
    'Suppress any error; a reset that cannot complete must not abort the run
        On Error Resume Next

        m_CertRecordFault = TEST_CERT_RECORD_FAULT_NONE
        m_CertRecordFaultBuffer = TEST_CERT_BUFFER_ANY

        m_CertLeakDetail = vbNullString

        If Not TST_CertResetRecord(m_CertRecord) Then
            TST_Log "TST_CertResetCounters", "RECORD FAIL", _
                "the certification record could not be cleared; the verdict " & _
                "will be FAIL | INCOMPLETE"
        End If

End Sub


Public Function TST_CertResetRecord( _
    ByRef Rec As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_CertResetRecord
'------------------------------------------------------------------------------
' PURPOSE
'   Empty one certification record and verify that every buffer is released.
'
' WHY THIS EXISTS
'   A record is reused from one certification run to the next. Clearing it
'   with nothing checked would let a buffer that survived the reset, or a
'   fatal flag from the earlier run, decide the next verdict (#50).
'
' INPUTS
'   Rec
'     ByRef. The record to empty.
'
' RETURNS
'   Boolean
'     TRUE  => every count and failure field is zero and every buffer is
'              unallocated.
'     FALSE => a buffer survived; the record is flagged with
'              TEST_CERT_RECFAIL_RESET.
'
' BEHAVIOR
'   - Unpublishes every entry first, by zeroing the counts.
'   - Releases every buffer.
'   - Clears the failure fields, including a fatal flag left by an earlier
'     run.
'   - Reads every buffer back, and flags the record when one still holds
'     storage.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertBufferCapacity
'   - TST_CertMarkRecordFailure
'
' CALLED FROM
'   - TST_CertResetCounters
'   - TST_CertCheckRecordFault
'   - TST_Case_CertRecordFailsClosed
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BufferId            As Long            'Cursor over the buffers
    Dim Cleared             As Boolean         'TRUE while every buffer reads empty

'------------------------------------------------------------------------------
' CLEAR RECORD
'------------------------------------------------------------------------------
    'Unpublish before releasing, so no count ever describes a released buffer.
    'Every step is attempted; the readback below is what decides the result.
        On Error Resume Next

        Rec.EntryCount = 0
        Rec.UnitCount = 0
        Rec.SkipCount = 0
        Rec.FindingCount = 0
        Rec.UnverifiedCount = 0
        Rec.CheckCount = 0

        Erase Rec.EntryKind
        Erase Rec.EntryName
        Erase Rec.EntryPassed
        Erase Rec.EntryDetail

        Rec.RecordFailed = False
        Rec.RecordFailCount = 0
        Rec.RecordFailStage = TEST_CERT_RECFAIL_NONE
        Rec.RecordFailError = 0

        On Error GoTo 0

'------------------------------------------------------------------------------
' VERIFY RECORD
'------------------------------------------------------------------------------
    'A buffer that still holds storage is a failed reset, not a clean one
        Cleared = True

        For BufferId = TEST_CERT_BUFFER_KIND To TEST_CERT_BUFFER_DETAIL
            If TST_CertBufferCapacity(Rec, BufferId) <> 0 Then
                Cleared = False
            End If
        Next BufferId

        If Not Cleared Then
            TST_CertMarkRecordFailure Rec, TEST_CERT_RECFAIL_RESET, 0
        End If

        TST_CertResetRecord = Cleared

End Function


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
'   - Does not raise. A skip that cannot be stored sets
'     m_CertRecord.RecordFailed, which forces FAIL | INCOMPLETE; that flag,
'     not the skip list, is authoritative when the list is short.
'
' DEPENDENCIES
'   - TST_CertRecordEntry
'
' CALLED FROM
'   - TST_RunRegressionPack
'
' UPDATED
'   2026-11-16 - Stores the skip through the shared record path, which
'                publishes it only once it is complete (#50).
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

            If Not TST_CertRecordEntry(TEST_CERT_ENTRY_SKIP, CallerProc, False, Reason) Then
                TST_Log "TST_CertRecordSkip", "RECORD FAIL", _
                    "a skip by " & CallerProc & " could not be stored; the " & _
                    "verdict will be FAIL | INCOMPLETE"
            End If

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
'   same record, so there is exactly one place where an outcome enters the run
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
'   - Does not raise. An outcome that cannot be stored sets
'     m_CertRecord.RecordFailed, which forces FAIL | INCOMPLETE; that flag,
'     not the unit list, is authoritative when the list is short.
'
' DEPENDENCIES
'   - TST_CertRecordEntry
'
' CALLED FROM
'   - TST_CertRunUnit
'
' UPDATED
'   2026-11-16 - Stores the outcome through the shared record path, which
'                publishes it only once it is complete (#50).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' APPEND OUTCOME
'------------------------------------------------------------------------------
    'Recording must never be the thing that breaks a run
        On Error Resume Next

        If Not TST_CertRecordEntry(TEST_CERT_ENTRY_UNIT, UnitName, Passed, Detail) Then
            TST_Log "TST_CertRecordUnit", "RECORD FAIL", _
                "the outcome of " & UnitName & " could not be stored; the " & _
                "verdict will be FAIL | INCOMPLETE"
        End If

        TST_Log "TST_CertRecordUnit", IIf(Passed, "UNIT PASS", "UNIT FAIL"), _
            UnitName & IIf(Len(Detail) > 0, " | " & Detail, vbNullString)

End Sub


Private Function TST_CertRecordEntry( _
    ByVal Kind As Long, _
    ByVal EntryName As String, _
    ByVal Passed As Boolean, _
    ByVal Detail As String) _
    As Boolean

'
'==============================================================================
' TST_CertRecordEntry
'------------------------------------------------------------------------------
' PURPOSE
'   Store one entry in the record of the certification run in progress, and
'   make sure that an entry that was not stored is counted as a recording
'   failure.
'
' WHY THIS EXISTS
'   TST_CertTryAppendEntry flags the record from its own handler. A failure
'   before that handler is armed, such as copying an argument, would return
'   FALSE with nothing flagged. Comparing the failure count before and after
'   closes that gap, so no missing entry can go unflagged.
'
'   Unit and skip recording use this one path; case recording joins it when
'   certification records cases (#42).
'
' INPUTS
'   Kind / EntryName / Passed / Detail
'     As TST_CertTryAppendEntry.
'
' RETURNS
'   Boolean
'     TRUE when the entry was stored and published.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertTryAppendEntry
'   - TST_CertMarkRecordFailure
'
' CALLED FROM
'   - TST_CertRecordUnit
'   - TST_CertRecordSkip
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FailuresBefore      As Long            'Failure count before the append
    Dim Stored              As Boolean         'TRUE once the entry is published

'------------------------------------------------------------------------------
' APPEND ENTRY
'------------------------------------------------------------------------------
    'An append that raised leaves Stored FALSE
        On Error Resume Next

        FailuresBefore = m_CertRecord.RecordFailCount
        Stored = False

        Stored = TST_CertTryAppendEntry(m_CertRecord, Kind, EntryName, Passed, Detail)

        On Error GoTo 0

'------------------------------------------------------------------------------
' FLAG AN UNCOUNTED MISS
'------------------------------------------------------------------------------
    'A miss the append could not flag itself is still a recording failure
        If Not Stored Then
            If m_CertRecord.RecordFailCount = FailuresBefore Then
                TST_CertMarkRecordFailure m_CertRecord, TEST_CERT_RECFAIL_WRITE, 0
            End If
        End If

        TST_CertRecordEntry = Stored

End Function


Public Function TST_CertAddFinding( _
    ByRef Rec As tCertRecord, _
    ByVal Code As String, _
    ByVal Detail As String) _
    As Boolean

'
'==============================================================================
' TST_CertAddFinding
'------------------------------------------------------------------------------
' PURPOSE
'   Store one cleanup finding, a difference between the exit state and the
'   entry baseline, in a certification record.
'
' WHY THIS EXISTS
'   Cleanup used to join its differences into one string under On Error
'   Resume Next, so a comparison that failed to run left no trace. A finding is
'   an entry of the record, stored by the same path as a unit (#50): one that
'   cannot be stored flags the record, which forces FAIL | INCOMPLETE.
'
' INPUTS
'   Rec
'     ByRef. The record to add to.
'
'   Code
'     Stable machine-readable identifier, such as app.DisplayStatusBar.
'
'   Detail
'     What differs, naming both values. Line breaks are replaced by spaces, so
'     the finding stays one line in the text report.
'
' RETURNS
'   Boolean
'     TRUE  => the finding is stored and counted.
'     FALSE => it was not stored; Rec.RecordFailed is TRUE.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertAddCleanupEntry
'
' CALLED FROM
'   - M_EXCEL_UI_TEST_CERTIFICATION
'   - TST_RestoreState
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        TST_CertAddFinding = TST_CertAddCleanupEntry( _
            Rec, TEST_CERT_ENTRY_FINDING, Code, Detail)

End Function


Public Function TST_CertAddUnverified( _
    ByRef Rec As tCertRecord, _
    ByVal Code As String, _
    ByVal Detail As String) _
    As Boolean

'
'==============================================================================
' TST_CertAddUnverified
'------------------------------------------------------------------------------
' PURPOSE
'   Store the name of one surface that cleanup could not compare, in a
'   certification record.
'
' WHY THIS EXISTS
'   A surface that cannot be read or addressed must be named in the evidence,
'   not treated as clean. It does not fail cleanup, because cleanup=OK claims
'   only what it could verify, and the entry says what that excludes.
'
' INPUTS
'   Rec / Code / Detail
'     As TST_CertAddFinding.
'
' RETURNS
'   Boolean
'     TRUE when the entry is stored; FALSE when it is not, in which case
'     Rec.RecordFailed is TRUE.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertAddCleanupEntry
'
' CALLED FROM
'   - M_EXCEL_UI_TEST_CERTIFICATION
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        TST_CertAddUnverified = TST_CertAddCleanupEntry( _
            Rec, TEST_CERT_ENTRY_UNVERIFIED, Code, Detail)

End Function


Private Function TST_CertAddCleanupEntry( _
    ByRef Rec As tCertRecord, _
    ByVal Kind As Long, _
    ByVal Code As String, _
    ByVal Detail As String) _
    As Boolean

'
'==============================================================================
' TST_CertAddCleanupEntry
'------------------------------------------------------------------------------
' PURPOSE
'   Store one finding or one unverified surface, and make sure that an entry
'   that was not stored is counted as a recording failure.
'
' WHY THIS EXISTS
'   TST_CertTryAppendEntry flags the record from its own handler, but a failure
'   before that handler is armed would return FALSE with nothing flagged.
'   Comparing the failure count before and after closes that gap, as
'   TST_CertRecordEntry does for units and skips.
'
' INPUTS
'   Rec / Kind / Code / Detail
'     The record, TEST_CERT_ENTRY_FINDING or TEST_CERT_ENTRY_UNVERIFIED, the
'     identifier, and the detail.
'
' RETURNS
'   Boolean
'     TRUE when the entry was stored and published.
'
' ERROR POLICY
'   - Does not raise. An error flags the record.
'
' DEPENDENCIES
'   - TST_CertTryAppendEntry
'   - TST_CertMarkRecordFailure
'
' CALLED FROM
'   - TST_CertAddFinding
'   - TST_CertAddUnverified
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FailuresBefore      As Long            'Failure count before the append
    Dim CleanDetail         As String          'Detail as a single line
    Dim Stored              As Boolean         'TRUE once the entry is published
    Dim FailNumber          As Long            'Error number of a failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertAddCleanupEntry = False
        Stored = False

        FailuresBefore = Rec.RecordFailCount

'------------------------------------------------------------------------------
' APPEND ENTRY
'------------------------------------------------------------------------------
    'One line, so the text report can list the entry as one line
        CleanDetail = Replace(Replace(Detail, vbCr, " "), vbLf, " ")

        Stored = TST_CertTryAppendEntry(Rec, Kind, Code, False, CleanDetail)

    'A miss the append could not flag itself is still a recording failure
        If Not Stored Then
            If Rec.RecordFailCount = FailuresBefore Then
                TST_CertMarkRecordFailure Rec, TEST_CERT_RECFAIL_WRITE, 0
            End If
        End If

        TST_CertAddCleanupEntry = Stored

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
    'Flag the record before anything else can allocate
        FailNumber = Err.Number

        TST_CertMarkRecordFailure Rec, TEST_CERT_RECFAIL_WRITE, FailNumber

        TST_CertAddCleanupEntry = False

        Resume Safe_Exit

End Function


Public Sub TST_CertCountCheck( _
    ByRef Rec As tCertRecord)

'
'==============================================================================
' TST_CertCountCheck
'------------------------------------------------------------------------------
' PURPOSE
'   Count one comparison that cleanup completed.
'
' WHY THIS EXISTS
'   cleanup=OK is only as strong as what was compared. The count is reported as
'   cleanupChecked, and a clean verdict that compared nothing is rejected by the
'   evidence validator.
'
' INPUTS
'   Rec
'     ByRef. The record whose count is raised.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   - Does not raise. The count saturates.
'
' CALLED FROM
'   - M_EXCEL_UI_TEST_CERTIFICATION
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

        If Rec.CheckCount < 2147483647 Then
            Rec.CheckCount = Rec.CheckCount + 1
        End If

End Sub


Public Function TST_CertEntryCodes( _
    ByRef Rec As tCertRecord, _
    ByVal Kind As Long) _
    As String

'
'==============================================================================
' TST_CertEntryCodes
'------------------------------------------------------------------------------
' PURPOSE
'   List the identifiers of the findings or unverified surfaces of a record, in
'   recording order, for a regression case to assert on.
'
' WHY THIS EXISTS
'   A case that checks one finding must not depend on how many others the host
'   produced, and a case that checks the order needs the order. Delimiting each
'   identifier with a pipe lets InStr match whole identifiers.
'
' INPUTS
'   Rec
'     ByRef. The record to read. It is not changed.
'
'   Kind
'     TEST_CERT_ENTRY_FINDING or TEST_CERT_ENTRY_UNVERIFIED.
'
' RETURNS
'   String
'     "|code1|code2|" in recording order; "|" when there is none. When the
'     record cannot be read, "|UNREADABLE|".
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertRecordIsConsistent
'
' CALLED FROM
'   - M_EXCEL_UI_TEST_CERTIFICATION
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScanIdx             As Long            'Cursor over the entries
    Dim Codes               As String          'Identifiers listed so far

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertEntryCodes = "|UNREADABLE|"

        If Not TST_CertRecordIsConsistent(Rec) Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' LIST IDENTIFIERS
'------------------------------------------------------------------------------
        Codes = "|"

        For ScanIdx = 1 To Rec.EntryCount

            If Rec.EntryKind(ScanIdx) = Kind Then
                Codes = Codes & Rec.EntryName(ScanIdx) & "|"
            End If

        Next ScanIdx

        TST_CertEntryCodes = Codes

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
    'An unreadable record is reported as one, never as an empty list
        TST_CertEntryCodes = "|UNREADABLE|"

        Resume Safe_Exit

End Function


Public Function TST_CertEntrySummary( _
    ByRef Rec As tCertRecord, _
    ByVal Kind As Long) _
    As String

'
'==============================================================================
' TST_CertEntrySummary
'------------------------------------------------------------------------------
' PURPOSE
'   Join the findings or unverified surfaces of a record into one line of
'   "code: detail" items, in recording order.
'
' WHY THIS EXISTS
'   The error certification raises, and the failure a regression pack raises
'   when its state restoration did not complete, must carry what was found.
'
' INPUTS
'   Rec
'     ByRef. The record to read. It is not changed.
'
'   Kind
'     TEST_CERT_ENTRY_FINDING or TEST_CERT_ENTRY_UNVERIFIED.
'
' RETURNS
'   String
'     The items joined by " | "; empty when there is none. When the record
'     cannot be read, "the record could not be read".
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertRecordIsConsistent
'   - TST_CertAppendDetail
'
' CALLED FROM
'   - TST_RecordRestoreFailure
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScanIdx             As Long            'Cursor over the entries
    Dim Summary             As String          'Items joined so far

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertEntrySummary = "the record could not be read"

        If Not TST_CertRecordIsConsistent(Rec) Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' JOIN ITEMS
'------------------------------------------------------------------------------
        For ScanIdx = 1 To Rec.EntryCount

            If Rec.EntryKind(ScanIdx) = Kind Then
                Summary = TST_CertAppendDetail(Summary, _
                    Rec.EntryName(ScanIdx) & ": " & Rec.EntryDetail(ScanIdx))
            End If

        Next ScanIdx

        TST_CertEntrySummary = Summary

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
    'A summary that cannot be read must not look like an empty one
        TST_CertEntrySummary = "the record could not be read"

        Resume Safe_Exit

End Function


Public Function TST_CertTryAppendEntry( _
    ByRef Rec As tCertRecord, _
    ByVal Kind As Long, _
    ByVal EntryName As String, _
    ByVal Passed As Boolean, _
    ByVal Detail As String) _
    As Boolean

'
'==============================================================================
' TST_CertTryAppendEntry
'------------------------------------------------------------------------------
' PURPOSE
'   Store one entry in a certification record and publish it only once it is
'   complete.
'
' WHY THIS EXISTS
'   Recording used to raise the count first and then grow each parallel array
'   under On Error Resume Next. A growth that failed left a count describing
'   an entry that one array did not hold, and the verdict and evidence loops
'   then read past the end of that array with every error suppressed (#50).
'
'   This is the one append path for every kind of entry: units, skips,
'   cleanup findings and unverified surfaces today, and cases when
'   certification records them (#42). A new kind adds a TEST_CERT_ENTRY_* value
'   and its count, not a new set of arrays.
'
' INPUTS
'   Rec
'     ByRef. The record to append to.
'
'   Kind
'     TEST_CERT_ENTRY_UNIT, TEST_CERT_ENTRY_SKIP, TEST_CERT_ENTRY_FINDING or
'     TEST_CERT_ENTRY_UNVERIFIED.
'
'   EntryName
'     Unit name, the procedure that skipped, or the identifier of a finding or
'     unverified surface.
'
'   Passed
'     Unit outcome. FALSE for every other kind.
'
'   Detail
'     Unit failure text, the skip reason, or the detail of a finding or
'     unverified surface.
'
' RETURNS
'   Boolean
'     TRUE  => the entry is stored and the counts include it.
'     FALSE => nothing was published; Rec.RecordFailed is TRUE.
'
' BEHAVIOR
'   1  validate      reject an unknown kind, and a record whose buffers
'                    already disagree with its counts
'   2  reserve       grow every buffer to the next index, in buffer order; a
'                    buffer already large enough is left alone
'   3  write         write every field of the entry
'   4  publish       raise the kind count, then EntryCount, last
'
'   A failure at any step leaves the counts as they were. Buffers grown before
'   the failure stay grown; they are capacity, not entries, and the next
'   append reuses them.
'
' ERROR POLICY
'   - Does not raise.
'   - On failure the handler copies Err.Number and then sets the scalar
'     failure fields of Rec before anything else is allocated.
'   - Rec.RecordFailed is authoritative when the entry cannot be stored: it
'     forces FAIL | INCOMPLETE through TST_CertComposeEvidence, and
'     RecordFailCount, RecordFailStage and RecordFailError name the failure
'     without needing any text.
'
' DEPENDENCIES
'   - TST_CertRecordIsConsistent
'   - TST_CertReserveBuffer
'   - TST_CertMarkRecordFailure
'
' CALLED FROM
'   - TST_CertRecordEntry
'   - TST_CertAddCleanupEntry
'   - TST_CertCheckRecordFault
'   - TST_Case_CertRecordFailsClosed
'
' UPDATED
'   2026-11-18 - Also stores cleanup findings and unverified surfaces (#35).
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NewIndex            As Long            'Index the entry is written to
    Dim Stage               As Long            'Step a failure is charged to
    Dim FailNumber          As Long            'Error number of a failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertTryAppendEntry = False

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'An entry no evidence section reads would be stored and never reported
        Stage = TEST_CERT_RECFAIL_KIND

        If Kind <> TEST_CERT_ENTRY_UNIT And _
            Kind <> TEST_CERT_ENTRY_SKIP And _
            Kind <> TEST_CERT_ENTRY_FINDING And _
            Kind <> TEST_CERT_ENTRY_UNVERIFIED Then

            TST_CertMarkRecordFailure Rec, Stage, 0
            GoTo Safe_Exit

        End If

    'Growing a record whose buffers already disagree with its counts would
    'publish an entry next to one that is missing
        Stage = TEST_CERT_RECFAIL_INVARIANT

        If Not TST_CertRecordIsConsistent(Rec) Then
            TST_CertMarkRecordFailure Rec, Stage, 0
            GoTo Safe_Exit
        End If

        NewIndex = Rec.EntryCount + 1

'------------------------------------------------------------------------------
' RESERVE EVERY BUFFER
'------------------------------------------------------------------------------
    'Stage names the buffer being grown, so a failure is charged to it
        For Stage = TEST_CERT_BUFFER_KIND To TEST_CERT_BUFFER_DETAIL
            TST_CertReserveBuffer Rec, Stage, NewIndex
        Next Stage

'------------------------------------------------------------------------------
' WRITE EVERY FIELD
'------------------------------------------------------------------------------
    'The slot is beyond EntryCount, so a partial write is never read
        Stage = TEST_CERT_RECFAIL_WRITE

        Rec.EntryKind(NewIndex) = Kind
        Rec.EntryName(NewIndex) = EntryName
        Rec.EntryPassed(NewIndex) = Passed
        Rec.EntryDetail(NewIndex) = Detail

'------------------------------------------------------------------------------
' PUBLISH
'------------------------------------------------------------------------------
    'Scalar assignments only; EntryCount last, once the entry is complete
        Select Case Kind

            Case TEST_CERT_ENTRY_UNIT
                Rec.UnitCount = Rec.UnitCount + 1

            Case TEST_CERT_ENTRY_SKIP
                Rec.SkipCount = Rec.SkipCount + 1

            Case TEST_CERT_ENTRY_FINDING
                Rec.FindingCount = Rec.FindingCount + 1

            Case Else
                Rec.UnverifiedCount = Rec.UnverifiedCount + 1

        End Select

        Rec.EntryCount = NewIndex

        TST_CertTryAppendEntry = True

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
    'Flag the record before anything else can allocate
        FailNumber = Err.Number

        TST_CertMarkRecordFailure Rec, Stage, FailNumber

        TST_CertTryAppendEntry = False

        Resume Safe_Exit

End Function


Private Sub TST_CertReserveBuffer( _
    ByRef Rec As tCertRecord, _
    ByVal BufferId As Long, _
    ByVal Needed As Long)

'
'==============================================================================
' TST_CertReserveBuffer
'------------------------------------------------------------------------------
' PURPOSE
'   Make one record buffer hold at least Needed 1-based elements.
'
' WHY THIS EXISTS
'   Each buffer is grown on its own, so that the caller can charge a failure
'   to the buffer that caused it, and so that the record fault can fail one
'   chosen buffer while the others grow.
'
' INPUTS
'   Rec
'     ByRef. The record that owns the buffer.
'
'   BufferId
'     TEST_CERT_BUFFER_* value of the buffer.
'
'   Needed
'     Upper bound the buffer must reach.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Leaves a buffer that is already large enough alone.
'   - Otherwise consumes a matching record fault, and grows the buffer with
'     ReDim Preserve.
'
' ERROR POLICY
'   - Raises error 7 when the record fault fires, error 9 when the buffer is
'     not 1-based, and whatever ReDim Preserve raises. The caller's handler
'     flags the record.
'
' DEPENDENCIES
'   - TST_CertBufferCapacity
'   - TST_CertTakeRecordFault
'
' CALLED FROM
'   - TST_CertTryAppendEntry
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Capacity            As Long            'Elements the buffer holds now

    Const PROC As String = "TST_CertReserveBuffer"

'------------------------------------------------------------------------------
' READ CAPACITY
'------------------------------------------------------------------------------
    'A buffer that is not 1-based cannot be indexed by entry number
        Capacity = TST_CertBufferCapacity(Rec, BufferId)

        If Capacity < 0 Then
            Err.Raise 9, PROC, _
                "certification record buffer " & CStr(BufferId) & " is not 1-based"
        End If

'------------------------------------------------------------------------------
' GROW BUFFER
'------------------------------------------------------------------------------
    'Only an allocation that is really needed can consume the fault
        If Capacity < Needed Then

            If TST_CertTakeRecordFault(BufferId, (Capacity > 0)) Then
                Err.Raise 7, PROC, _
                    "injected certification record allocation failure"
            End If

            Select Case BufferId

                Case TEST_CERT_BUFFER_KIND
                    ReDim Preserve Rec.EntryKind(1 To Needed)

                Case TEST_CERT_BUFFER_NAME
                    ReDim Preserve Rec.EntryName(1 To Needed)

                Case TEST_CERT_BUFFER_PASSED
                    ReDim Preserve Rec.EntryPassed(1 To Needed)

                Case TEST_CERT_BUFFER_DETAIL
                    ReDim Preserve Rec.EntryDetail(1 To Needed)

                Case Else
                    Err.Raise 5, PROC, _
                        "unknown certification record buffer " & CStr(BufferId)

            End Select

        End If

End Sub


Private Function TST_CertBufferCapacity( _
    ByRef Rec As tCertRecord, _
    ByVal BufferId As Long) _
    As Long

'
'==============================================================================
' TST_CertBufferCapacity
'------------------------------------------------------------------------------
' PURPOSE
'   Report how many 1-based elements one record buffer holds.
'
' WHY THIS EXISTS
'   An unallocated dynamic array raises on LBound and UBound. Reading the
'   bounds in one place, under its own handler, lets every caller compare a
'   capacity with a count instead of indexing and hoping.
'
' INPUTS
'   Rec
'     ByRef. The record that owns the buffer.
'
'   BufferId
'     TEST_CERT_BUFFER_* value of the buffer.
'
' RETURNS
'   Long
'     The upper bound of a 1-based buffer; 0 for an unallocated buffer; -1
'     for a buffer that is not 1-based, an unknown buffer, or a read that
'     failed for any other reason.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - TST_CertReserveBuffer
'   - TST_CertRecordIsConsistent
'   - TST_CertResetRecord
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LowBound            As Long            'LBound of the buffer
    Dim HighBound           As Long            'UBound of the buffer
    Dim FailNumber          As Long            'Error number of a failed read

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertBufferCapacity = -1

'------------------------------------------------------------------------------
' READ BOUNDS
'------------------------------------------------------------------------------
    'An unallocated buffer raises error 9 here
        Select Case BufferId

            Case TEST_CERT_BUFFER_KIND
                LowBound = LBound(Rec.EntryKind)
                HighBound = UBound(Rec.EntryKind)

            Case TEST_CERT_BUFFER_NAME
                LowBound = LBound(Rec.EntryName)
                HighBound = UBound(Rec.EntryName)

            Case TEST_CERT_BUFFER_PASSED
                LowBound = LBound(Rec.EntryPassed)
                HighBound = UBound(Rec.EntryPassed)

            Case TEST_CERT_BUFFER_DETAIL
                LowBound = LBound(Rec.EntryDetail)
                HighBound = UBound(Rec.EntryDetail)

            Case Else
                GoTo Safe_Exit

        End Select

'------------------------------------------------------------------------------
' RETURN CAPACITY
'------------------------------------------------------------------------------
    'Only a 1-based buffer can be indexed by entry number
        If LowBound = 1 Then
            TST_CertBufferCapacity = HighBound
        End If

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
    'Subscript out of range is an unallocated buffer; anything else is unknown
        FailNumber = Err.Number

        If FailNumber = 9 Then
            TST_CertBufferCapacity = 0
        Else
            TST_CertBufferCapacity = -1
        End If

        Resume Safe_Exit

End Function


Private Sub TST_CertMarkRecordFailure( _
    ByRef Rec As tCertRecord, _
    ByVal Stage As Long, _
    ByVal ErrorNumber As Long)

'
'==============================================================================
' TST_CertMarkRecordFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Flag a certification record as incomplete.
'
' WHY THIS EXISTS
'   The flag is what forces FAIL | INCOMPLETE when an entry cannot be stored,
'   so setting it must not depend on the allocation that just failed. This
'   writes scalar fields only: no string, array or object is created.
'
' INPUTS
'   Rec
'     ByRef. The record to flag.
'
'   Stage
'     TEST_CERT_RECFAIL_* value, or the TEST_CERT_BUFFER_* value of the
'     buffer that could not grow.
'
'   ErrorNumber
'     Runtime error number, or 0 when the failure raised none.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Sets RecordFailed and counts the failure.
'   - Keeps the stage and error number of the first failure, which caused
'     any that follow.
'
' ERROR POLICY
'   - Cannot raise. It runs inside error handlers.
'
' CALLED FROM
'   - TST_CertTryAppendEntry
'   - TST_CertRecordEntry
'   - TST_CertResetRecord
'   - TST_CertBuildEntryLists
'   - TST_CertComposeEvidence
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' FLAG RECORD
'------------------------------------------------------------------------------
    'The flag first: it is the part the verdict depends on
        Rec.RecordFailed = True

        If Rec.RecordFailCount < 2147483647 Then
            Rec.RecordFailCount = Rec.RecordFailCount + 1
        End If

        If Rec.RecordFailStage = TEST_CERT_RECFAIL_NONE Then
            Rec.RecordFailStage = Stage
            Rec.RecordFailError = ErrorNumber
        End If

End Sub


Public Function TST_CertRecordIsConsistent( _
    ByRef Rec As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_CertRecordIsConsistent
'------------------------------------------------------------------------------
' PURPOSE
'   Check the invariants of a certification record before anything indexes
'   its buffers.
'
' WHY THIS EXISTS
'   The evidence loops index every buffer up to EntryCount. A count beyond a
'   buffer's bound raised out of range, and the On Error Resume Next around
'   those loops turned that into missing or wrong evidence (#50). Checking
'   first makes a mismatch a recording failure instead.
'
' INPUTS
'   Rec
'     ByRef. The record to check. It is not changed.
'
' RETURNS
'   Boolean
'     TRUE when the counts are non-negative, the four kind counts add up to
'     EntryCount, and every buffer is either unallocated with EntryCount = 0
'     or 1-based with an upper bound of at least EntryCount.
'
' ERROR POLICY
'   - Does not raise. A check that cannot complete returns FALSE.
'
' DEPENDENCIES
'   - TST_CertBufferCapacity
'
' CALLED FROM
'   - TST_CertTryAppendEntry
'   - TST_CertBuildEntryLists
'   - TST_CertCheckRecordFault
'   - TST_Case_CertRecordFailsClosed
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim BufferId            As Long            'Cursor over the buffers
    Dim Consistent          As Boolean         'TRUE while every check passes

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_CertRecordIsConsistent = False

'------------------------------------------------------------------------------
' CHECK COUNTS
'------------------------------------------------------------------------------
    'The kind counts partition the published entries
        Consistent = (Rec.EntryCount >= 0) And _
            (Rec.UnitCount >= 0) And _
            (Rec.SkipCount >= 0) And _
            (Rec.FindingCount >= 0) And _
            (Rec.UnverifiedCount >= 0)

        If Consistent Then
            Consistent = (Rec.UnitCount + Rec.SkipCount + _
                Rec.FindingCount + Rec.UnverifiedCount = Rec.EntryCount)
        End If

'------------------------------------------------------------------------------
' CHECK BUFFERS
'------------------------------------------------------------------------------
    'Every published entry must exist in every buffer
        If Consistent Then

            For BufferId = TEST_CERT_BUFFER_KIND To TEST_CERT_BUFFER_DETAIL

                If TST_CertBufferCapacity(Rec, BufferId) < Rec.EntryCount Then
                    Consistent = False
                End If

            Next BufferId

        End If

        TST_CertRecordIsConsistent = Consistent

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
    'A check that cannot complete is a failed check
        TST_CertRecordIsConsistent = False

        Resume Safe_Exit

End Function


Public Sub TST_CertComposeEvidence( _
    ByRef Rec As tCertRecord, _
    ByVal CleanupOK As Boolean, _
    ByRef CleanupDetail As String, _
    ByRef UnitsFailed As Long, _
    ByRef Complete As Boolean, _
    ByRef Passed As Boolean, _
    ByRef Verdict As String, _
    ByRef RecordingDetail As String, _
    ByRef JsonText As String, _
    ByRef ReportText As String)

'
'==============================================================================
' TST_CertComposeEvidence
'------------------------------------------------------------------------------
' PURPOSE
'   Derive the verdict of a certification record and compose its JSON and
'   text evidence.
'
' WHY THIS EXISTS
'   The verdict used to be counted by the runner and the evidence built by
'   two loops that each trusted the counts. One composer, which checks the
'   record before reading it and decides the verdict after, is what keeps a
'   record that could not be stored from producing a passing document (#50).
'
'   The runner and the record regression case both call it, so the case
'   asserts exactly what certification emits.
'
' INPUTS
'   Rec
'     ByRef. The record to report. A failure found while reading it flags it.
'
'   CleanupOK / CleanupDetail
'     Cleanup verdict computed by the caller, outside the record. Cleanup is
'     also FAILED whenever the record holds a finding, whatever CleanupOK says.
'     CleanupDetail is ByRef: on return it holds the caller's detail followed
'     by the detail of every finding, joined by " | ".
'
'   UnitsFailed / Complete / Passed / Verdict / RecordingDetail /
'   JsonText / ReportText
'     ByRef outputs. Set to a failing verdict and empty documents first, so
'     that any early exit leaves FAIL | INCOMPLETE.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   1  read entries  check the invariants, then list units and skips; a
'                    mismatch or a failed read flags the record and lists
'                    nothing
'   2  verdict       Complete requires no skip and no recording failure;
'                    Passed requires Complete, no failed unit, no cleanup
'                    finding and cleanup OK
'   3  documents     JSON with "cleanupFindings", "cleanupUnverified",
'                    "cleanupChecked", "recording" and "recordingDetail", and a
'                    text report with a CLEANUP line per finding, an UNVERIFIED
'                    line per unverified surface, and a RECORDING line when
'                    recording failed
'
' ERROR POLICY
'   - Does not raise. An error flags the record with
'     TEST_CERT_RECFAIL_EVIDENCE and leaves the failing verdict.
'   - Passed and Complete are authoritative. When a document cannot be
'     composed in full, they and Rec.RecordFailed still describe the run, and
'     neither document can say passed while Rec.RecordFailed is TRUE.
'
' DEPENDENCIES
'   - TST_CertBuildEntryLists
'   - TST_CertVerdictLine
'   - TST_CertRecordingDetail
'   - TST_CertAppendDetail
'   - TST_CertBuildJsonEvidence
'   - TST_CertBuildTextReport
'   - TST_CertMarkRecordFailure
'
' CALLED FROM
'   - Test_EXCEL_UI_RunReleaseCertification
'   - TST_CertCheckRecordFault
'   - TST_Case_CertRecordFailsClosed
'
' UPDATED
'   2026-11-18 - Cleanup is FAILED whenever the record holds a finding, and
'                both documents list the findings and unverified surfaces
'                (#35).
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim UnitJson            As String          'unitResults members
    Dim SkipJson            As String          'skipDetail members
    Dim UnitLines           As String          'Unit lines of the report
    Dim SkipLines           As String          'Skip lines of the report
    Dim FindingJson         As String          'cleanupFindings members
    Dim UnverifiedJson      As String          'cleanupUnverified members
    Dim FindingLines        As String          'CLEANUP lines of the report
    Dim UnverifiedLines     As String          'UNVERIFIED lines of the report
    Dim FindingSummary      As String          'Finding details joined
    Dim CallerDetail        As String          'Detail the caller passed in
    Dim CleanupPassed       As Boolean         'Cleanup verdict after the findings
    Dim FailNumber          As Long            'Error number of a failure

'------------------------------------------------------------------------------
' FAIL CLOSED
'------------------------------------------------------------------------------
    'Every output starts as a failure; only a complete pass below changes it
        UnitsFailed = 0
        Complete = False
        Passed = False
        Verdict = "RESULT: FAIL | INCOMPLETE"
        RecordingDetail = vbNullString
        JsonText = vbNullString
        ReportText = vbNullString

        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' READ ENTRIES
'------------------------------------------------------------------------------
    'Checks the invariants before indexing; a failure flags the record
        TST_CertBuildEntryLists _
            Rec:=Rec, _
            UnitsFailed:=UnitsFailed, _
            UnitJson:=UnitJson, _
            SkipJson:=SkipJson, _
            UnitLines:=UnitLines, _
            SkipLines:=SkipLines, _
            FindingJson:=FindingJson, _
            UnverifiedJson:=UnverifiedJson, _
            FindingLines:=FindingLines, _
            UnverifiedLines:=UnverifiedLines, _
            FindingSummary:=FindingSummary)

'------------------------------------------------------------------------------
' DETERMINE VERDICT
'------------------------------------------------------------------------------
    'Decided after the entries were read, so a failure found there counts.
    'Completeness and correctness are separate questions: a run that skipped
    'a mandatory unit, or could not record one, is not complete.
    '
    'Cleanup is judged on the record as well as on the caller's verdict: a
    'finding fails it whatever the caller passed, so no caller can store a
    'difference and still report cleanup OK.
        CallerDetail = CleanupDetail
        CleanupPassed = CleanupOK And (Rec.FindingCount = 0)
        CleanupDetail = TST_CertAppendDetail(CleanupDetail, FindingSummary)

        Complete = (Rec.SkipCount = 0) And Not Rec.RecordFailed
        Passed = (UnitsFailed = 0) And CleanupPassed And Complete

        Verdict = TST_CertVerdictLine( _
            Passed:=Passed, _
            Complete:=Complete, _
            Units:=Rec.UnitCount, _
            UnitsFailed:=UnitsFailed, _
            Skipped:=Rec.SkipCount, _
            CleanupOK:=CleanupPassed)

        If Rec.RecordFailed Then
            RecordingDetail = TST_CertRecordingDetail(Rec)
        End If

'------------------------------------------------------------------------------
' BUILD DOCUMENTS
'------------------------------------------------------------------------------
    'Both documents carry the same verdict and the same entries
        JsonText = TST_CertBuildJsonEvidence( _
            Rec:=Rec, _
            UnitsFailed:=UnitsFailed, _
            CleanupOK:=CleanupPassed, _
            CleanupDetail:=CleanupDetail, _
            Complete:=Complete, _
            Passed:=Passed, _
            RecordingDetail:=RecordingDetail, _
            UnitJson:=UnitJson, _
            SkipJson:=SkipJson, _
            FindingJson:=FindingJson, _
            UnverifiedJson:=UnverifiedJson)

        ReportText = TST_CertBuildTextReport( _
            Rec:=Rec, _
            Verdict:=Verdict, _
            CleanupDetail:=CallerDetail, _
            RecordingDetail:=RecordingDetail, _
            UnitLines:=UnitLines, _
            SkipLines:=SkipLines, _
            FindingLines:=FindingLines, _
            UnverifiedLines:=UnverifiedLines)

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
    'Flag the record first, then put the failing verdict back
        FailNumber = Err.Number

        TST_CertMarkRecordFailure Rec, TEST_CERT_RECFAIL_EVIDENCE, FailNumber

        Complete = False
        Passed = False
        Verdict = "RESULT: FAIL | INCOMPLETE"

        Resume Safe_Exit

End Sub


Private Function TST_CertBuildEntryLists( _
    ByRef Rec As tCertRecord, _
    ByRef UnitsFailed As Long, _
    ByRef UnitJson As String, _
    ByRef SkipJson As String, _
    ByRef UnitLines As String, _
    ByRef SkipLines As String, _
    ByRef FindingJson As String, _
    ByRef UnverifiedJson As String, _
    ByRef FindingLines As String, _
    ByRef UnverifiedLines As String, _
    ByRef FindingSummary As String) _
    As Boolean

'
'==============================================================================
' TST_CertBuildEntryLists
'------------------------------------------------------------------------------
' PURPOSE
'   Read every published entry once and compose the unit, skip, finding and
'   unverified sections of both evidence documents.
'
' WHY THIS EXISTS
'   Reading the record in one pass, after its invariants are checked, means
'   the JSON and the text report list the same entries, and no index beyond a
'   buffer's bound is ever read.
'
' INPUTS
'   Rec
'     ByRef. The record to read. A mismatch or a failed read flags it.
'
'   UnitsFailed / UnitJson / SkipJson / UnitLines / SkipLines
'     ByRef outputs: the failed-unit count, the members of the unitResults
'     and skipDetail arrays, and the unit and skip lines of the report.
'
'   FindingJson / UnverifiedJson / FindingLines / UnverifiedLines
'     ByRef outputs: the members of the cleanupFindings and cleanupUnverified
'     arrays, and the CLEANUP and UNVERIFIED lines of the report. A member is
'     {"seq":n,"code":"...","detail":"..."}, numbered from 1 in recording
'     order, and a line is "code | detail".
'
'   FindingSummary
'     ByRef output: the details of the findings joined by " | ".
'
' RETURNS
'   Boolean
'     TRUE  => every entry was read, and the kinds read match the counts.
'     FALSE => the record is flagged, every output is empty and
'              UnitsFailed is 0.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_CertRecordIsConsistent
'   - TST_CertJsonEscape
'   - TST_CertAppendDetail
'   - TST_CertMarkRecordFailure
'
' CALLED FROM
'   - TST_CertComposeEvidence
'
' UPDATED
'   2026-11-18 - Also lists cleanup findings and unverified surfaces (#35).
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScanIdx             As Long            'Cursor over the entries
    Dim UnitsSeen           As Long            'Unit entries read
    Dim SkipsSeen           As Long            'Skip entries read
    Dim FindingsSeen        As Long            'Finding entries read
    Dim UnverifiedSeen      As Long            'Unverified entries read
    Dim FailedSeen          As Long            'Failed units read
    Dim SkipText            As String          'One skip, as it is reported
    Dim Matched             As Boolean         'TRUE while kinds match counts
    Dim FailNumber          As Long            'Error number of a failed read

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        UnitsFailed = 0
        UnitJson = vbNullString
        SkipJson = vbNullString
        UnitLines = vbNullString
        SkipLines = vbNullString
        FindingJson = vbNullString
        UnverifiedJson = vbNullString
        FindingLines = vbNullString
        UnverifiedLines = vbNullString
        FindingSummary = vbNullString

        TST_CertBuildEntryLists = False

        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' CHECK INVARIANTS
'------------------------------------------------------------------------------
    'Nothing is indexed until every buffer is known to hold every entry
        Matched = TST_CertRecordIsConsistent(Rec)

'------------------------------------------------------------------------------
' READ ENTRIES
'------------------------------------------------------------------------------
    'One pass, in recording order
        If Matched Then

            For ScanIdx = 1 To Rec.EntryCount

                Select Case Rec.EntryKind(ScanIdx)

                    Case TEST_CERT_ENTRY_UNIT
                        UnitsSeen = UnitsSeen + 1

                        If Not Rec.EntryPassed(ScanIdx) Then
                            FailedSeen = FailedSeen + 1
                        End If

                        If UnitsSeen > 1 Then
                            UnitJson = UnitJson & ","
                        End If

                        UnitJson = UnitJson & "{""name"":""" & _
                            TST_CertJsonEscape(Rec.EntryName(ScanIdx)) & """" & _
                            ",""passed"":" & _
                            IIf(Rec.EntryPassed(ScanIdx), "true", "false") & _
                            ",""detail"":""" & _
                            TST_CertJsonEscape(Rec.EntryDetail(ScanIdx)) & """}"

                        UnitLines = UnitLines & _
                            IIf(Rec.EntryPassed(ScanIdx), "  PASS  ", "  FAIL  ") & _
                            Rec.EntryName(ScanIdx) & _
                            IIf(Len(Rec.EntryDetail(ScanIdx)) > 0, _
                                " | " & Rec.EntryDetail(ScanIdx), vbNullString) & _
                            vbNewLine

                    Case TEST_CERT_ENTRY_SKIP
                        SkipsSeen = SkipsSeen + 1

                        SkipText = Rec.EntryName(ScanIdx) & ": " & _
                            Rec.EntryDetail(ScanIdx)

                        If SkipsSeen > 1 Then
                            SkipJson = SkipJson & ","
                        End If

                        SkipJson = SkipJson & """" & TST_CertJsonEscape(SkipText) & """"
                        SkipLines = SkipLines & "  SKIP  " & SkipText & vbNewLine

                    Case TEST_CERT_ENTRY_FINDING
                        FindingsSeen = FindingsSeen + 1

                        If FindingsSeen > 1 Then
                            FindingJson = FindingJson & ","
                        End If

                        FindingJson = FindingJson & "{""seq"":" & _
                            CStr(FindingsSeen) & _
                            ",""code"":""" & _
                            TST_CertJsonEscape(Rec.EntryName(ScanIdx)) & """" & _
                            ",""detail"":""" & _
                            TST_CertJsonEscape(Rec.EntryDetail(ScanIdx)) & """}"

                        FindingLines = FindingLines & "  CLEANUP  " & _
                            Rec.EntryName(ScanIdx) & " | " & _
                            Rec.EntryDetail(ScanIdx) & vbNewLine

                        FindingSummary = TST_CertAppendDetail(FindingSummary, _
                            Rec.EntryDetail(ScanIdx))

                    Case TEST_CERT_ENTRY_UNVERIFIED
                        UnverifiedSeen = UnverifiedSeen + 1

                        If UnverifiedSeen > 1 Then
                            UnverifiedJson = UnverifiedJson & ","
                        End If

                        UnverifiedJson = UnverifiedJson & "{""seq"":" & _
                            CStr(UnverifiedSeen) & _
                            ",""code"":""" & _
                            TST_CertJsonEscape(Rec.EntryName(ScanIdx)) & """" & _
                            ",""detail"":""" & _
                            TST_CertJsonEscape(Rec.EntryDetail(ScanIdx)) & """}"

                        UnverifiedLines = UnverifiedLines & "  UNVERIFIED  " & _
                            Rec.EntryName(ScanIdx) & " | " & _
                            Rec.EntryDetail(ScanIdx) & vbNewLine

                    Case Else
                        Matched = False

                End Select

                If Not Matched Then
                    Exit For
                End If

            Next ScanIdx

        End If

    'Every published entry must have been read as the kind it was counted as
        If Matched Then
            Matched = (UnitsSeen = Rec.UnitCount) And _
                (SkipsSeen = Rec.SkipCount) And _
                (FindingsSeen = Rec.FindingCount) And _
                (UnverifiedSeen = Rec.UnverifiedCount)
        End If

'------------------------------------------------------------------------------
' RETURN LISTS
'------------------------------------------------------------------------------
    'A record that does not match its counts lists nothing
        If Matched Then
            UnitsFailed = FailedSeen
            TST_CertBuildEntryLists = True
        Else
            TST_CertMarkRecordFailure Rec, TEST_CERT_RECFAIL_INVARIANT, 0

            UnitJson = vbNullString
            SkipJson = vbNullString
            UnitLines = vbNullString
            SkipLines = vbNullString
            FindingJson = vbNullString
            UnverifiedJson = vbNullString
            FindingLines = vbNullString
            UnverifiedLines = vbNullString
            FindingSummary = vbNullString
        End If

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
    'Flag the record first, then drop whatever was half read
        FailNumber = Err.Number

        TST_CertMarkRecordFailure Rec, TEST_CERT_RECFAIL_EVIDENCE, FailNumber

        UnitsFailed = 0
        UnitJson = vbNullString
        SkipJson = vbNullString
        UnitLines = vbNullString
        SkipLines = vbNullString
        FindingJson = vbNullString
        UnverifiedJson = vbNullString
        FindingLines = vbNullString
        UnverifiedLines = vbNullString
        FindingSummary = vbNullString

        TST_CertBuildEntryLists = False

        Resume Safe_Exit

End Function


Private Function TST_CertVerdictLine( _
    ByVal Passed As Boolean, _
    ByVal Complete As Boolean, _
    ByVal Units As Long, _
    ByVal UnitsFailed As Long, _
    ByVal Skipped As Long, _
    ByVal CleanupOK As Boolean) _
    As String

'
'==============================================================================
' TST_CertVerdictLine
'------------------------------------------------------------------------------
' PURPOSE
'   Compose the RESULT line of a certification run.
'
' WHY THIS EXISTS
'   The line is quoted in pull requests and checked against the JSON by
'   tools/check_excel_evidence.py, so its format is composed in one place.
'
' INPUTS
'   Passed / Complete / Units / UnitsFailed / Skipped / CleanupOK
'     The verdict components.
'
' RETURNS
'   String
'     RESULT: <PASS|FAIL> | <COMPLETE|INCOMPLETE> | units=<n> failed=<n>
'     skipped=<n> cleanup=<OK|FAILED>
'
' ERROR POLICY
'   - Raises only if the text cannot be allocated; the caller's handler then
'     keeps its failing verdict.
'
' CALLED FROM
'   - TST_CertComposeEvidence
'
' UPDATED
'   2026-11-16 - Moved out of Test_EXCEL_UI_RunReleaseCertification (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' COMPOSE LINE
'------------------------------------------------------------------------------
        TST_CertVerdictLine = "RESULT: " & IIf(Passed, "PASS", "FAIL") & _
            " | " & IIf(Complete, "COMPLETE", "INCOMPLETE") & _
            " | units=" & CStr(Units) & _
            " failed=" & CStr(UnitsFailed) & _
            " skipped=" & CStr(Skipped) & _
            " cleanup=" & IIf(CleanupOK, "OK", "FAILED")

End Function


Private Function TST_CertRecordingDetail( _
    ByRef Rec As tCertRecord) _
    As String

'
'==============================================================================
' TST_CertRecordingDetail
'------------------------------------------------------------------------------
' PURPOSE
'   Name a certification-recording failure from the scalar fields of the
'   record.
'
' WHY THIS EXISTS
'   The entry that could not be stored is, by definition, not in the record,
'   so the evidence names the failure from what survived it: the count, the
'   stage and the error number of the first failure.
'
' INPUTS
'   Rec
'     ByRef. The flagged record. It is not changed.
'
' RETURNS
'   String
'     "certification recording FAILED | failures=<n> | first: <stage>",
'     followed by " | error <n>" when the failure raised one. When the full
'     text cannot be composed, the first three words alone.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - TST_CertComposeEvidence
'
' UPDATED
'   2026-11-16 - Added (#50).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Detail              As String          'Text under construction
    Dim StageText           As String          'What the first failure was doing

'------------------------------------------------------------------------------
' NAME FAILURE
'------------------------------------------------------------------------------
    'The prefix first, so that a later failure still leaves a named failure
        On Error Resume Next

        Detail = "certification recording FAILED"

        Select Case Rec.RecordFailStage

            Case TEST_CERT_BUFFER_KIND
                StageText = "growing the entry-kind buffer"

            Case TEST_CERT_BUFFER_NAME
                StageText = "growing the entry-name buffer"

            Case TEST_CERT_BUFFER_PASSED
                StageText = "growing the unit-outcome buffer"

            Case TEST_CERT_BUFFER_DETAIL
                StageText = "growing the entry-detail buffer"

            Case TEST_CERT_RECFAIL_WRITE
                StageText = "writing the entry"

            Case TEST_CERT_RECFAIL_INVARIANT
                StageText = "the record buffers disagree with its counts"

            Case TEST_CERT_RECFAIL_EVIDENCE
                StageText = "reading the entries for the evidence"

            Case TEST_CERT_RECFAIL_KIND
                StageText = "an entry of unknown kind"

            Case TEST_CERT_RECFAIL_RESET
                StageText = "clearing the record"

            Case Else
                StageText = "stage " & CStr(Rec.RecordFailStage)

        End Select

        Detail = Detail & _
            " | failures=" & CStr(Rec.RecordFailCount) & _
            " | first: " & StageText & _
            IIf(Rec.RecordFailError <> 0, _
                " | error " & CStr(Rec.RecordFailError), vbNullString)

        TST_CertRecordingDetail = Detail

End Function


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
    ByRef Rec As tCertRecord, _
    ByVal UnitsFailed As Long, _
    ByVal CleanupOK As Boolean, _
    ByVal CleanupDetail As String, _
    ByVal Complete As Boolean, _
    ByVal Passed As Boolean, _
    ByVal RecordingDetail As String, _
    ByVal UnitJson As String, _
    ByVal SkipJson As String, _
    ByVal FindingJson As String, _
    ByVal UnverifiedJson As String) _
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
'   Rec
'     ByRef. The record reported; its counts and RecordFailed are read, never
'     its buffers.
'
'   UnitsFailed / CleanupOK / CleanupDetail / Complete / Passed
'     Verdict components computed by TST_CertComposeEvidence.
'
'   RecordingDetail
'     Text naming a recording failure; empty when recording succeeded.
'
'   UnitJson / SkipJson / FindingJson / UnverifiedJson
'     Members of the unitResults, skipDetail, cleanupFindings and
'     cleanupUnverified arrays, composed by TST_CertBuildEntryLists after the
'     record's invariants were checked.
'
' RETURNS
'   String
'     A single-line JSON document. Cleanup is reported as "cleanup",
'     "cleanupDetail", "cleanupFindings" (every difference found),
'     "cleanupUnverified" (every surface named as not compared) and
'     "cleanupChecked" (the comparisons completed).
'
' ERROR POLICY
'   - Does not raise. A field that cannot be read is reported as unknown.
'   - "recording" is derived from Rec.RecordFailed alone, and "complete" and
'     "passed" are false whenever it is set, so no document can claim a pass
'     over a record that could not be stored.
'
' DEPENDENCIES
'   - TST_CertJsonEscape
'
' CALLED FROM
'   - TST_CertComposeEvidence
'
' NOTES
'   Bitness and VBA generation come from conditional compilation, so they
'   describe the build that is executing rather than what the host reports.
'
' UPDATED
'   2026-11-18 - Reports "cleanupFindings", "cleanupUnverified" and
'                "cleanupChecked" (#35).
'   2026-11-16 - Reads the record it is given, and reports "recording" and
'                "recordingDetail" (#50).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Json                As String          'Document under construction
    Dim Bitness             As String          'Office bitness of this build
    Dim VbaGeneration       As String          'VBA generation of this build
    Dim Recorded            As Boolean         'TRUE when every entry was stored

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Evidence must be produced even when a field cannot be read
        On Error Resume Next

        Recorded = Not Rec.RecordFailed

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
            ",""units"":" & CStr(Rec.UnitCount) & _
            ",""unitsFailed"":" & CStr(UnitsFailed) & _
            ",""skipped"":" & CStr(Rec.SkipCount) & _
            ",""cleanup"":""" & IIf(CleanupOK, "OK", "FAILED") & """" & _
            ",""cleanupDetail"":""" & TST_CertJsonEscape(CleanupDetail) & """" & _
            ",""cleanupFindings"":[" & FindingJson & "]" & _
            ",""cleanupUnverified"":[" & UnverifiedJson & "]" & _
            ",""cleanupChecked"":" & CStr(Rec.CheckCount) & _
            ",""recording"":""" & IIf(Recorded, "OK", "FAILED") & """" & _
            ",""recordingDetail"":""" & TST_CertJsonEscape(RecordingDetail) & """" & _
            ",""complete"":" & IIf(Complete And Recorded, "true", "false") & _
            ",""passed"":" & IIf(Passed And Complete And Recorded, "true", "false")

'------------------------------------------------------------------------------
' BUILD UNIT AND SKIP DETAIL
'------------------------------------------------------------------------------
    'One object per unit, in execution order. Skips are listed explicitly; an
    'empty array is a meaningful result.
        Json = Json & ",""unitResults"":[" & UnitJson & "]"
        Json = Json & ",""skipDetail"":[" & SkipJson & "]}"

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
'   - TST_CertBuildEntryLists
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
    ByRef Rec As tCertRecord, _
    ByVal Verdict As String, _
    ByVal CleanupDetail As String, _
    ByVal RecordingDetail As String, _
    ByVal UnitLines As String, _
    ByVal SkipLines As String, _
    ByVal FindingLines As String, _
    ByVal UnverifiedLines As String) _
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
'   Rec
'     ByRef. The record reported; only RecordFailed is read.
'
'   Verdict
'     Pre-composed verdict line.
'
'   CleanupDetail
'     Cleanup detail the caller passed to the composer, outside the record;
'     empty when there was none. The findings of the record are listed by
'     FindingLines.
'
'   RecordingDetail
'     Text naming a recording failure; empty when recording succeeded.
'
'   UnitLines / SkipLines / FindingLines / UnverifiedLines
'     Unit, skip, CLEANUP and UNVERIFIED lines composed by
'     TST_CertBuildEntryLists.
'
' RETURNS
'   String
'     A multi-line report. A recording failure adds a RECORDING line directly
'     under the RESULT line, named even when RecordingDetail is empty. Every
'     cleanup finding is a CLEANUP line, and every surface cleanup could not
'     compare is an UNVERIFIED line after them.
'
' ERROR POLICY
'   - Does not raise.
'
' CALLED FROM
'   - TST_CertComposeEvidence
'
' UPDATED
'   2026-11-18 - Lists every cleanup finding and unverified surface (#35).
'   2026-11-16 - Takes the composed lines and reports a recording failure
'                (#50).
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Report              As String          'Report under construction

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
' APPEND RECORDING FAILURE
'------------------------------------------------------------------------------
    'Directly under the verdict, because it says the lists below are short
        If Rec.RecordFailed Then

            If Len(RecordingDetail) > 0 Then
                Report = Report & "  RECORDING  " & RecordingDetail & vbNewLine
            Else
                Report = Report & "  RECORDING  certification recording FAILED" & vbNewLine
            End If

        End If

'------------------------------------------------------------------------------
' APPEND UNITS, SKIPS AND CLEANUP
'------------------------------------------------------------------------------
    'One line per unit, in execution order. Skips are named, because an
    'unnamed skip is what this runner exists to stop happening.
        Report = Report & UnitLines & SkipLines & FindingLines

        If Len(CleanupDetail) > 0 Then
            Report = Report & "  CLEANUP  " & CleanupDetail & vbNewLine
        End If

    'What cleanup could not compare is named after what it did compare, so
    'that a clean verdict says what it excludes
        Report = Report & UnverifiedLines

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


Public Function TST_RestoreState( _
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
    ByVal TitleBarVisible As Boolean, _
    ByRef Findings As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_RestoreState
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to restore the pre-test Excel UI state after the regression run,
'   and report every step that did not complete.
'
' WHY THIS EXISTS
'   Regression tests should clean up after themselves as much as possible. A
'   restoration that discards its failures reports a pass over a host it did
'   not put back (#35), so each step is verified and a step that fails or does
'   not stick is a structured finding.
'
' INPUTS
'   [Captured snapshot values]
'     Pre-test UI state captured by TST_SnapshotState
'
'   Findings
'     ByRef. A record that receives one finding per step that failed, with
'     the identifier restore.<element>. The caller owns the record and
'     decides what a finding costs.
'
' RETURNS
'   Boolean
'     TRUE  => every step completed and every written value read back as
'              written.
'     FALSE => at least one step failed, or the restoration itself raised;
'              Findings says which.
'
' BEHAVIOR
'   - Attempts every step, whatever an earlier one reported.
'   - Restores the title bar through the structured API, then the Ribbon, then
'     the application-level properties, then the window-level properties.
'   - Reads each Boolean property back after writing it.
'   - Reports a window count that differs from the captured one, because
'     window-level state is restored by index and may then reach another
'     window.
'
' ERROR POLICY
'   - Does not raise. A step that raises, or the restoration as a whole, is
'     reported as a finding, never discarded.
'
' DEPENDENCIES
'   - TST_RestoreTitleBarElement
'   - TST_RestoreBoolean
'   - TST_RestoreWindowBoolean
'   - TST_TrySetRibbonVisible
'   - TST_CertAddFinding
'
' CALLED FROM
'   - TST_RunRegressionPack
'   - TST_RunTitleBarOnlyPack
'
' NOTES
'   This is the pack-level restoration. Certification judges the whole run
'   against its own entry baseline afterwards, whatever this reports.
'
' UPDATED
'   2026-11-18 - Became a Function that reports every step it could not
'                complete instead of suppressing it (#35).
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long            'Current window index during restore
    Dim WindowLimit         As Long            'Minimum of saved and current window counts
    Dim Msg                 As String          'Diagnostic message from helper routines
    Dim AllRestored         As Boolean         'TRUE while every step has completed
    Dim StepRestored        As Boolean         'Outcome of the step just attempted

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Every step reports for itself, so a failure in one cannot hide the rest
        On Error GoTo Err_Handler

        TST_RestoreState = False
        AllRestored = True

'------------------------------------------------------------------------------
' RESTORE TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Restore TitleBar first when its original state was captured successfully
        If TitleBarKnown Then

            'Restore TitleBar through the structured API and keep its verdict
                StepRestored = TST_RestoreTitleBarElement(TitleBarVisible, Findings)
                AllRestored = AllRestored And StepRestored

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
                    TST_CertAddFinding Findings, "restore.Ribbon", _
                        "the Ribbon could not be restored | " & Msg
                    AllRestored = False
                End If

        End If

'------------------------------------------------------------------------------
' RESTORE APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Restore StatusBar visibility and read it back
        StepRestored = TST_RestoreBoolean(Application, _
            "DisplayStatusBar", StatusBarVisible, _
            "Application.DisplayStatusBar", Findings)
        AllRestored = AllRestored And StepRestored

    'Restore ScrollBars visibility and read it back
        StepRestored = TST_RestoreBoolean(Application, _
            "DisplayScrollBars", ScrollBarsVisible, _
            "Application.DisplayScrollBars", Findings)
        AllRestored = AllRestored And StepRestored

    'Restore FormulaBar visibility and read it back
        StepRestored = TST_RestoreBoolean(Application, _
            "DisplayFormulaBar", FormulaBarVisible, _
            "Application.DisplayFormulaBar", Findings)
        AllRestored = AllRestored And StepRestored

'------------------------------------------------------------------------------
' RESTORE WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'A different window count means the captured index no longer names the same
    'window, which the restoration cannot repair, so it is reported
        WindowLimit = Application.Windows.Count

        If WindowCount <> WindowLimit Then
            TST_CertAddFinding Findings, "restore.windowCount", _
                "Application.Windows.Count is " & CStr(WindowLimit) & _
                ", not the " & CStr(WindowCount) & " captured; window-level " & _
                "state is restored by index and may reach another window"
            AllRestored = False
        End If

    'Compute the number of windows that can be restored safely by index
        If WindowCount < WindowLimit Then WindowLimit = WindowCount

    'Restore each saved window state up to the common window count
        For i = 1 To WindowLimit

            'Restore the current window's Headings visibility
                StepRestored = TST_RestoreWindowBoolean(i, _
                    "DisplayHeadings", HeadingsVisible(i), Findings)
                AllRestored = AllRestored And StepRestored

            'Restore the current window's WorkbookTabs visibility
                StepRestored = TST_RestoreWindowBoolean(i, _
                    "DisplayWorkbookTabs", WorkbookTabsVisible(i), Findings)
                AllRestored = AllRestored And StepRestored

            'Restore the current window's Gridlines visibility
                StepRestored = TST_RestoreWindowBoolean(i, _
                    "DisplayGridlines", GridlinesVisible(i), Findings)
                AllRestored = AllRestored And StepRestored

        Next i

'------------------------------------------------------------------------------
' SETTLE UI
'------------------------------------------------------------------------------
    'Allow the UI a short time to settle after restoration
        TST_WaitUI TEST_WAIT_SECONDS

        TST_RestoreState = AllRestored

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
    'The restoration itself raised: say so, and report failure
        Msg = TST_BuildRuntimeErrorText

        TST_CertAddFinding Findings, "restore.raised", _
            "UI state restoration raised | " & Msg

        TST_RestoreState = False

        Resume Safe_Exit

End Function


Private Function TST_RestoreTitleBarElement( _
    ByVal TitleBarVisible As Boolean, _
    ByRef Findings As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_RestoreTitleBarElement
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the captured title-bar visibility through the structured API and
'   report a failure as a finding.
'
' WHY THIS EXISTS
'   The fire-and-forget call discarded the failure list, so a title bar the API
'   could not restore left no trace. The structured call returns it.
'
' INPUTS
'   TitleBarVisible
'     The title-bar visibility captured before the run.
'
'   Findings
'     ByRef. Receives restore.TitleBar when the restoration fails.
'
' RETURNS
'   Boolean
'     TRUE when the structured call reported success.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - UI_SetExcelUI_WithResult
'   - TST_DescribeResultBuffers
'   - TST_CertAddFinding
'
' CALLED FROM
'   - TST_RestoreState
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean         'Structured result of the call
    Dim FailureCount        As Long            'Structured failure count
    Dim FailureList         As Variant         'Structured failure list
    Dim Msg                 As String          'Diagnostic from an error

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_RestoreTitleBarElement = False

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        If TitleBarVisible Then
            Succeeded = UI_SetExcelUI_WithResult( _
                TitleBar:=UI_Show, _
                FailureCount:=FailureCount, _
                FailureList:=FailureList)
        Else
            Succeeded = UI_SetExcelUI_WithResult( _
                TitleBar:=UI_Hide, _
                FailureCount:=FailureCount, _
                FailureList:=FailureList)
        End If

        If Succeeded Then
            TST_RestoreTitleBarElement = True
        Else
            TST_CertAddFinding Findings, "restore.TitleBar", _
                "the title bar could not be restored through the API | " & _
                TST_DescribeResultBuffers(FailureCount, FailureList)
        End If

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

        TST_CertAddFinding Findings, "restore.TitleBar", _
            "the title-bar restoration raised | " & Msg

        TST_RestoreTitleBarElement = False

        Resume Safe_Exit

End Function


Private Function TST_RestoreBoolean( _
    ByVal Target As Object, _
    ByVal PropName As String, _
    ByVal Value As Boolean, _
    ByVal Label As String, _
    ByRef Findings As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_RestoreBoolean
'------------------------------------------------------------------------------
' PURPOSE
'   Write one Boolean property and read it back, reporting a write that failed
'   or did not stick as a finding.
'
' WHY THIS EXISTS
'   A write that raises no error need not achieve the value. Reading it back is
'   what makes the restoration a restoration.
'
' INPUTS
'   Target / PropName / Value
'     The object, its Boolean property and the value to write.
'
'   Label
'     What the property is called in a finding.
'
'   Findings
'     ByRef. Receives restore.<PropName> when the restoration fails.
'
' RETURNS
'   Boolean
'     TRUE when the value was written and read back as written.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_TrySetBooleanProperty
'   - TST_TryGetBooleanProperty
'   - TST_CertAddFinding
'
' CALLED FROM
'   - TST_RestoreState
'   - TST_RestoreWindowBoolean
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ReadBack            As Boolean         'Value read after the write
    Dim Msg                 As String          'Diagnostic from a helper or error

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_RestoreBoolean = False

'------------------------------------------------------------------------------
' WRITE AND READ BACK
'------------------------------------------------------------------------------
        If Not TST_TrySetBooleanProperty(Target, PropName, Value, Msg) Then
            TST_Log "TST_RestoreState", "Restore." & PropName, Msg
            TST_CertAddFinding Findings, "restore." & PropName, _
                Label & " could not be restored | " & Msg

        ElseIf Not TST_TryGetBooleanProperty(Target, PropName, ReadBack, Msg) Then
            TST_CertAddFinding Findings, "restore." & PropName, _
                Label & " was written but could not be read back | " & Msg

        ElseIf ReadBack <> Value Then
            TST_CertAddFinding Findings, "restore." & PropName, _
                Label & " was written as " & CStr(Value) & _
                " but reads " & CStr(ReadBack)

        Else
            TST_RestoreBoolean = True
        End If

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

        TST_CertAddFinding Findings, "restore." & PropName, _
            Label & " restoration raised | " & Msg

        TST_RestoreBoolean = False

        Resume Safe_Exit

End Function


Private Function TST_RestoreWindowBoolean( _
    ByVal WindowIndex As Long, _
    ByVal PropName As String, _
    ByVal Value As Boolean, _
    ByRef Findings As tCertRecord) _
    As Boolean

'
'==============================================================================
' TST_RestoreWindowBoolean
'------------------------------------------------------------------------------
' PURPOSE
'   Restore one Boolean property of the window at a collection index, and
'   report a failure as a finding.
'
' WHY THIS EXISTS
'   Fetching the window inside the helper keeps a window that has gone from
'   raising in the caller, where it would end the loop over the others.
'
' INPUTS
'   WindowIndex
'     Application.Windows index captured by TST_SnapshotState.
'
'   PropName / Value
'     The window's Boolean property and the value to write.
'
'   Findings
'     ByRef. Receives restore.<PropName> when the restoration fails.
'
' RETURNS
'   Boolean
'     TRUE when the value was written and read back as written.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - TST_RestoreBoolean
'   - TST_CertAddFinding
'
' CALLED FROM
'   - TST_RestoreState
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetWindow        As Object          'Window at the captured index
    Dim Caption             As String          'Its caption, for the finding
    Dim Msg                 As String          'Diagnostic from an error

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        TST_RestoreWindowBoolean = False

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
        Set TargetWindow = Application.Windows(WindowIndex)

        Caption = TargetWindow.Caption

        TST_RestoreWindowBoolean = TST_RestoreBoolean( _
            TargetWindow, PropName, Value, _
            "Window " & CStr(WindowIndex) & " [" & Caption & "]." & PropName, _
            Findings)

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

        TST_CertAddFinding Findings, "restore." & PropName, _
            "Window " & CStr(WindowIndex) & "." & PropName & _
            " restoration raised | " & Msg

        TST_RestoreWindowBoolean = False

        Resume Safe_Exit

End Function


Public Sub TST_RecordRestoreFailure( _
    ByVal RestoreOK As Boolean, _
    ByRef Findings As tCertRecord, _
    ByVal CallerProc As String, _
    ByRef FailNumber As Long, _
    ByRef FailSource As String, _
    ByRef FailDescription As String)

'
'==============================================================================
' TST_RecordRestoreFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Make a state restoration that did not complete part of the caller's
'   result: the failure the caller raises when it had none, or an addition to
'   the one it already holds.
'
' WHY THIS EXISTS
'   A pack that cannot put the UI back has not finished, however its
'   assertions went. Logging the miss and returning normally would let the pack
'   report a pass over a host it left altered (#35).
'
' INPUTS
'   RestoreOK
'     Result of TST_RestoreState. Initialize it to FALSE before the call, so a
'     restoration that raised past the caller's handler counts as failed.
'
'   Findings
'     ByRef. The record TST_RestoreState wrote its findings to.
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
'   - Does nothing when RestoreOK is TRUE and the record stored every finding.
'   - With no saved failure, sets TEST_ERR_BASE + 82, the source
'     CallerProc & ".StateRestored" and the findings.
'   - With a saved failure, keeps its number and source and appends the
'     findings to its description.
'
' ERROR POLICY
'   - Does not raise and reads no Err field. The failure is set before any
'     text is composed, so a composition that fails still leaves a failure.
'
' DEPENDENCIES
'   - TST_CertEntrySummary
'   - TST_Log
'
' CALLED FROM
'   - TST_RunRegressionPack
'   - TST_RunTitleBarOnlyPack
'
' UPDATED
'   2026-11-18 - Added (#35).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Summary             As String          'Findings joined as one line

'------------------------------------------------------------------------------
' DECIDE
'------------------------------------------------------------------------------
        If RestoreOK And Not Findings.RecordFailed Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
    'The failure first, with text that cannot fail to build
        If FailNumber = 0 Then
            FailNumber = TEST_ERR_BASE + 82
            FailSource = CallerProc & ".StateRestored"
            FailDescription = "UI state restoration failed"
        Else
            FailDescription = FailDescription & _
                " | UI state restoration failed"
        End If

    'Then the detail; a failure here leaves the failure above in place
        On Error Resume Next

        Summary = TST_CertEntrySummary(Findings, TEST_CERT_ENTRY_FINDING)

        If Len(Summary) = 0 Then
            Summary = "restoration did not complete"
        End If

        If Findings.RecordFailed Then
            Summary = Summary & " | the findings could not all be stored"
        End If

        TST_Log CallerProc, "FAIL", "UI state restoration failed | " & Summary

        FailDescription = FailDescription & " | " & Summary

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


Public Function TST_TryGetBooleanProperty( _
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
' CALLED FROM
'   - TST_RestoreBoolean
'   - M_EXCEL_UI_TEST_CERTIFICATION
'
' UPDATED
'   2026-11-18 - Public, so that certification reads through it (#35).
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
            ((StyleValue And TST_WS_CAPTION) <> 0)

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
'   regression harness, and to the log sink when one is set
'
' WHY THIS EXISTS
'   The harness needs readable progress and failure logging. A script that
'   runs the suites cannot read the Immediate Window, so the same line also
'   goes to the opt-in log sink
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
'   - A failed sink write is kept by TST_LogSinkWriteLine for the caller's
'     verdict; it never reaches the caller of TST_Log
'
' UPDATED
'   2026-11-16 - Mirrors the line to the opt-in log sink (#197).
'   2026-07-25
'==============================================================================
'
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LineText            As String          'The diagnostic line

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected logging failure
        On Error Resume Next

'------------------------------------------------------------------------------
' WRITE DIAGNOSTIC LINE
'------------------------------------------------------------------------------
        LineText = ProcName & " @ " & Stage & " | " & Detail
        Debug.Print LineText

    'Mirror the line to the opt-in log sink; with no sink this writes nothing
        TST_LogSinkWriteLine LineText

End Sub


Public Sub TST_LogSinkStart(ByVal LogPath As String)

'
'==============================================================================
' TST_LogSinkStart
'------------------------------------------------------------------------------
' PURPOSE
'   Name the file that TST_Log mirrors its lines to from now on
'
' WHY THIS EXISTS
'   A script that runs a runner through Application.Run cannot read the
'   Immediate Window. With a sink set, every line the runner logs is also
'   appended to a file the script reads afterwards
'
' INPUTS
'   LogPath
'     Full path of the log file. Lines are appended; the file is created by
'     the first write. An empty path sets no sink
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Clears any write failure kept from an earlier sink
'   - Writes nothing itself
'
' ERROR POLICY
'   - Does not raise
'
' CALLED FROM
'   - TST_RunRunnerWithLog
'
' UPDATED
'   2026-11-16 - Added for the scripted Excel checks (#197).
'==============================================================================
'

'------------------------------------------------------------------------------
' SET SINK
'------------------------------------------------------------------------------
        m_LogSinkPath = LogPath
        m_LogSinkFailure = vbNullString

End Sub


Public Function TST_LogSinkStop() As String

'
'==============================================================================
' TST_LogSinkStop
'------------------------------------------------------------------------------
' PURPOSE
'   Stop mirroring TST_Log lines and report whether every write succeeded
'
' WHY THIS EXISTS
'   A failed write must not fail or interrupt the runner being logged, but it
'   must not be lost either: the caller puts it in its verdict
'
' RETURNS
'   String
'     Empty when every line reached the file; otherwise the first write
'     failure, with the error number, the path and the error text
'
' BEHAVIOR
'   - After this call TST_Log prints only to the Immediate Window again
'
' ERROR POLICY
'   - Does not raise
'
' CALLED FROM
'   - TST_RunRunnerWithLog
'
' UPDATED
'   2026-11-16 - Added for the scripted Excel checks (#197).
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN FAILURE AND CLEAR SINK
'------------------------------------------------------------------------------
        TST_LogSinkStop = m_LogSinkFailure

        m_LogSinkPath = vbNullString
        m_LogSinkFailure = vbNullString

End Function


Public Sub TST_LogSinkWriteLine(ByVal LineText As String)

'
'==============================================================================
' TST_LogSinkWriteLine
'------------------------------------------------------------------------------
' PURPOSE
'   Append one line to the log sink as UTF-8, when a sink is set
'
' WHY THIS EXISTS
'   TST_Log and the logged-runner verdict write through one path, so that a
'   failure is kept in one place
'
' INPUTS
'   LineText
'     The line, without a line break; CR LF is appended
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Does nothing when no sink is set, or after a write to it failed
'   - Opens, appends to and closes the file for every line, so that nothing
'     stays open when a runner stops part-way
'
' ERROR POLICY
'   - Does not raise. The first failure is kept for TST_LogSinkStop and ends
'     writing until the next TST_LogSinkStart, so an unwritable file costs
'     one failed attempt, not one per line
'
' CALLED FROM
'   - TST_Log
'   - TST_RunRunnerWithLog
'
' UPDATED
'   2026-11-16 - Added for the scripted Excel checks (#197).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LineBytes()         As Byte            'UTF-8 bytes of the line and CR LF
    Dim FileNumber          As Integer         'Free file handle
    Dim FileOpen            As Boolean         'TRUE while the file is open
    Dim FailNumber          As Long            'Captured write error number
    Dim FailDescription     As String          'Captured write error text

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'With no sink, or after a failed write, there is nothing to do
        If Len(m_LogSinkPath) = 0 Or Len(m_LogSinkFailure) > 0 Then
            Exit Sub
        End If

        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' APPEND LINE
'------------------------------------------------------------------------------
        LineBytes = TST_LogSinkEncode(LineText)

        FileNumber = FreeFile
        Open m_LogSinkPath For Binary Access Write As #FileNumber
        FileOpen = True

        Put #FileNumber, LOF(FileNumber) + 1, LineBytes

        Close #FileNumber
        FileOpen = False

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Close a file that a failed write left open
        If FileOpen Then
            On Error Resume Next
                Close #FileNumber
            On Error GoTo 0
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Capture before anything else touches Err, and keep the failure
        FailNumber = Err.Number
        FailDescription = Err.Description

        m_LogSinkFailure = "error " & CStr(FailNumber) & " writing " & _
            m_LogSinkPath & ": " & FailDescription

        Resume Safe_Exit

End Sub


Private Function TST_LogSinkEncode(ByVal LineText As String) As Byte()

'
'==============================================================================
' TST_LogSinkEncode
'------------------------------------------------------------------------------
' PURPOSE
'   Return one log line as UTF-8 bytes followed by CR LF
'
' WHY THIS EXISTS
'   VBA file statements write text in the system code page, which cannot
'   carry every character that a workbook name or a localized Excel error
'   text may hold. The log is UTF-8 so that a script reads it the same way
'   on every host
'
' INPUTS
'   LineText
'     The line to encode, without a line break
'
' RETURNS
'   The UTF-8 bytes of LineText, then CR LF. A UTF-16 surrogate that is not
'   part of a pair is written as U+FFFD
'
' ERROR POLICY
'   - Raises only when memory runs out; TST_LogSinkWriteLine keeps that as a
'     write failure
'
' UPDATED
'   2026-11-16 - Added for the scripted Excel checks (#197).
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Bytes()             As Byte            'Encoded line
    Dim TextLength          As Long            'UTF-16 code units in LineText
    Dim CharIdx             As Long            'Current code unit, 1-based
    Dim ByteIdx             As Long            'Next byte to write, 0-based
    Dim CodeUnit            As Long            'Current UTF-16 code unit
    Dim NextUnit            As Long            'Following UTF-16 code unit
    Dim CodePoint           As Long            'Scalar value being encoded

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Three bytes per code unit is the most UTF-8 needs; two more hold CR LF
        TextLength = Len(LineText)
        ReDim Bytes(0 To 3 * TextLength + 1)

        CharIdx = 1
        ByteIdx = 0

'------------------------------------------------------------------------------
' ENCODE
'------------------------------------------------------------------------------
        Do While CharIdx <= TextLength

            CodeUnit = AscW(Mid$(LineText, CharIdx, 1)) And &HFFFF&
            CodePoint = CodeUnit

        'Join a surrogate pair into one scalar; replace a lone surrogate
            If CodeUnit >= &HD800& And CodeUnit <= &HDFFF& Then
                CodePoint = &HFFFD&
                If CodeUnit <= &HDBFF& And CharIdx < TextLength Then
                    NextUnit = AscW(Mid$(LineText, CharIdx + 1, 1)) And &HFFFF&
                    If NextUnit >= &HDC00& And NextUnit <= &HDFFF& Then
                        CodePoint = &H10000 + (CodeUnit - &HD800&) * &H400& + _
                            (NextUnit - &HDC00&)
                        CharIdx = CharIdx + 1
                    End If
                End If
            End If

        'Write one to four bytes
            If CodePoint < &H80& Then
                Bytes(ByteIdx) = CodePoint
                ByteIdx = ByteIdx + 1
            ElseIf CodePoint < &H800& Then
                Bytes(ByteIdx) = &HC0& Or (CodePoint \ &H40&)
                Bytes(ByteIdx + 1) = &H80& Or (CodePoint And &H3F&)
                ByteIdx = ByteIdx + 2
            ElseIf CodePoint < &H10000 Then
                Bytes(ByteIdx) = &HE0& Or (CodePoint \ &H1000&)
                Bytes(ByteIdx + 1) = &H80& Or ((CodePoint \ &H40&) And &H3F&)
                Bytes(ByteIdx + 2) = &H80& Or (CodePoint And &H3F&)
                ByteIdx = ByteIdx + 3
            Else
                Bytes(ByteIdx) = &HF0& Or (CodePoint \ &H40000)
                Bytes(ByteIdx + 1) = &H80& Or ((CodePoint \ &H1000&) And &H3F&)
                Bytes(ByteIdx + 2) = &H80& Or ((CodePoint \ &H40&) And &H3F&)
                Bytes(ByteIdx + 3) = &H80& Or (CodePoint And &H3F&)
                ByteIdx = ByteIdx + 4
            End If

            CharIdx = CharIdx + 1

        Loop

'------------------------------------------------------------------------------
' RETURN LINE
'------------------------------------------------------------------------------
        Bytes(ByteIdx) = 13
        Bytes(ByteIdx + 1) = 10

        ReDim Preserve Bytes(0 To ByteIdx + 1)

        TST_LogSinkEncode = Bytes

End Function


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


Public Function TST_BuildRuntimeErrorText() As String

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
