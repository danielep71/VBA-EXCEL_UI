Attribute VB_Name = "M_DEMO_CLEANUP_TESTS"
Option Explicit

'==============================================================================
' M_DEMO_CLEANUP_TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Bounded Excel harness for the single-pass cleanup of DEMO_Sheet_Reset and
'   DEMO_Sheet_BuildTemplate (#71).
'
' WHY THIS EXISTS
'   The demo-cleanup source fixtures in tools/ check the shape of the cleanup
'   code, not what Excel does with it. This harness runs both procedures in
'   Excel through success, ordinary failure, unprotect failure and cleanup
'   failure, and reads back the protection Excel actually holds afterwards.
'
' PUBLIC SURFACE
'   - Demo_Test_RunCleanupChecks
'
' SCENARIOS
'   S01  Reset success, protection options set A, reprotected
'   S02  Reset success, protection options set B, reprotected
'   S03  Reset success with the default ReProtectAtEnd=False
'   S04  Reset ordinary error: WS_In is not a worksheet
'   S05  Reset ordinary error after its own unprotect
'   S06  Reset ordinary error on a sheet that was not protected
'   S07  Reset unprotect failure: wrong password
'   S08  Reset unprotect failure: missing password
'   S09  Reset reprotect failure without an operation error
'   S10  Reset reprotect failure after an operation error
'   S11  Reset EnableSelection restoration failure
'   S12  Reset operation, Protect and EnableSelection failures together
'   S13  BuildTemplate ordinary error inside its guarded body
'   S14  BuildTemplate propagating a nested Reset failure
'
' EVIDENCE
'   Each scenario checks the error that reaches the caller (number, source
'   and description), the M_DEMO_BUILDER cleanup probe counters (handler
'   entries, cleanup passes, Protect and EnableSelection attempts), whether
'   the reset body ran, and a readback of the fifteen Protect options and
'   EnableSelection.
'
' SAFETY
'   - Works only in a temporary workbook it creates and closes without saving;
'     no existing workbook or sheet is touched
'   - Every scenario is a straight sequence of calls; nothing waits or retries
'   - The cleanup probe and its faults are disarmed after every call and on
'     every exit
'   - Never calls Unprotect without a password, so Excel never prompts
'
' LIMITATIONS
'   - Reprotection and EnableSelection failures, and an operation error after
'     a successful unprotect, are injected through the M_DEMO_BUILDER seams;
'     no worksheet condition produces them on demand
'   - The wrong-password error number and text are Excel's; only that an error
'     arrived without a cleanup diagnostic is checked
'   - UserInterfaceOnly is read back in the same session only
'   - Not a watchdog: if Excel blocks inside a call, the harness blocks too
'
' UPDATED
'   2026-11-06 - Added for #71.
'
' AUTHOR
'   Daniele Penza
'==============================================================================


'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const PROTECT_PASSWORD As String = "ct-71-pass"     'Password the harness protects with
    Private Const WRONG_PASSWORD   As String = "ct-71-wrong"    'Password that must be refused
    Private Const BODY_MARKER      As String = "ct-71 marker"   'A1 text cleared by the reset body

    'Fifteen Protect options in Worksheet.Protect order after Password:
    'DrawingObjects, Contents, Scenarios, UserInterfaceOnly, then the eleven
    'Allow* options. "1" means True. Set B is the complement of set A.
    Private Const OPTIONS_A        As String = "010101010101010"
    Private Const OPTIONS_B        As String = "101010101010101"
    Private Const OPTION_NAMES     As String = "DrawingObjects,Contents,Scenarios,UserInterfaceOnly," & _
        "AllowFormattingCells,AllowFormattingColumns,AllowFormattingRows,AllowInsertingColumns," & _
        "AllowInsertingRows,AllowInsertingHyperlinks,AllowDeletingColumns,AllowDeletingRows," & _
        "AllowSorting,AllowFiltering,AllowUsingPivotTables"
    Private Const OPTION_COUNT     As Long = 15

    'Cleanup probe sites; mirror M_DEMO_BUILDER
    Private Const PROBE_RESET_HANDLER   As Long = 1
    Private Const PROBE_RESET_CLEANUP   As Long = 2
    Private Const PROBE_RESET_PROTECT   As Long = 3
    Private Const PROBE_RESET_SELECTION As Long = 4
    Private Const PROBE_BUILD_HANDLER   As Long = 5
    Private Const PROBE_BUILD_CLEANUP   As Long = 6

    'One-shot cleanup fault bits; mirror M_DEMO_BUILDER
    Private Const FAULT_NONE            As Long = 0
    Private Const FAULT_BODY            As Long = 1
    Private Const FAULT_PROTECT         As Long = 2
    Private Const FAULT_SELECTION       As Long = 4

    'Injected faults; mirror M_DEMO_BUILDER
    Private Const ERR_BODY              As Long = vbObjectError + 2190
    Private Const SRC_BODY              As String = "M_DEMO_BUILDER.DEMO_Sheet_Reset.InjectedOperationFault"
    Private Const DESC_BODY             As String = "Injected operation fault after the protection steps."
    Private Const ERR_PROTECT           As Long = vbObjectError + 2191
    Private Const SRC_PROTECT           As String = "M_DEMO_BUILDER.DEMO_Sheet_Reset.InjectedProtectFault"
    Private Const DESC_PROTECT          As String = "Injected reprotection fault."
    Private Const ERR_SELECTION         As Long = vbObjectError + 2192
    Private Const SRC_SELECTION         As String = "M_DEMO_BUILDER.DEMO_Sheet_Reset.InjectedSelectionFault"
    Private Const DESC_SELECTION        As String = "Injected EnableSelection restoration fault."

    'Real validation errors of M_DEMO_BUILDER
    Private Const SRC_RESET             As String = "M_DEMO_BUILDER.DEMO_Sheet_Reset"
    Private Const ERR_NO_PASSWORD       As Long = vbObjectError + 2100
    Private Const DESC_NO_PASSWORD      As String = "Worksheet is protected and no password was supplied."
    Private Const ERR_NOT_WORKSHEET     As Long = vbObjectError + 2103
    Private Const DESC_NOT_WORKSHEET    As String = "WS_In must be a Worksheet when supplied."
    Private Const SRC_BUILD             As String = "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate"
    Private Const ERR_BLANK_CONTENT     As Long = vbObjectError + 2021
    Private Const DESC_BLANK_CONTENT    As String = "ContentColumns cannot be blank."

    'Cleanup diagnostic markers written by DEMO_Sheet_Reset
    Private Const CLEANUP_MARKER        As String = "Cleanup failure: "
    Private Const SELECTION_MARKER      As String = " | EnableSelection: "

    'Raised by the runner when any check failed
    Private Const ERR_RUN_FAILED        As Long = vbObjectError + 2199

'------------------------------------------------------------------------------
' PRIVATE MODULE STATE
'------------------------------------------------------------------------------
    Private mChecks           As Long    'Checks evaluated in this run
    Private mFailures         As Long    'Checks failed in this run
    Private mScenarios        As Long    'Scenarios started in this run
    Private mScenarioFailures As Long    'Scenarios with at least one failure


Public Sub Demo_Test_RunCleanupChecks()
'
'==============================================================================
' Demo_Test_RunCleanupChecks
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every #71 cleanup scenario in a temporary workbook and prints one
'   line per scenario and a PASS/FAIL summary to the Immediate window.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates one temporary workbook; each scenario adds its own sheet to it
'   - A scenario that fails or raises unexpectedly is reported and the run
'     continues with the next one
'   - Disarms the cleanup probe and closes the temporary workbook without
'     saving on every exit
'
' ERROR POLICY
'   Raises vbObjectError + 2199 after cleanup when any check failed, and
'   re-raises an error that stopped the run itself, so a failed run cannot be
'   mistaken for a passing one
'
' DEPENDENCIES
'   - M_DEMO_BUILDER (DEMO_Sheet_Reset, DEMO_Sheet_BuildTemplate and the
'     DEMO_InternalCleanupProbe* seams)
'
' CALLED FROM
'   - Invoked directly, from the Immediate window or the Macros dialog
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WB                  As Workbook        'Temporary workbook
    Dim Summary             As String          'Summary line
    Dim RaisedNumber        As Long            'Error that stopped the run
    Dim RaisedSource        As String          'Source of that error
    Dim RaisedDescription   As String          'Text of that error

    Const PROC As String = "Demo_Test_RunCleanupChecks"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start from zero counters and a disarmed probe
        mChecks = 0
        mFailures = 0
        mScenarios = 0
        mScenarioFailures = 0

        On Error GoTo Err_Handler

        DEMO_InternalCleanupProbeEnd

        Debug.Print PROC & ": START " & Format$(Now, "yyyy-mm-dd hh:nn:ss")

    'Create the only workbook the scenarios touch
        Set WB = Workbooks.Add(xlWBATWorksheet)

'------------------------------------------------------------------------------
' RUN SCENARIOS
'------------------------------------------------------------------------------
    'Success and readback
        DEMO_CT_ResetSuccess WB, "S01_SuccessA", OPTIONS_A, xlUnlockedCells
        DEMO_CT_ResetSuccess WB, "S02_SuccessB", OPTIONS_B, xlNoSelection
        DEMO_CT_ResetWithoutReprotect WB, "S03_NoReprotect"

    'Ordinary failure
        DEMO_CT_ResetNotWorksheet WB, "S04_NotWorksheet"
        DEMO_CT_ResetOperationFault WB, "S05_OperationFault"
        DEMO_CT_ResetOperationFaultUnprotected WB, "S06_OperationUnprot"

    'Unprotect failure
        DEMO_CT_ResetUnprotectFailure WB, "S07_WrongPassword", WRONG_PASSWORD
        DEMO_CT_ResetUnprotectFailure WB, "S08_MissingPassword", vbNullString

    'Cleanup failure
        DEMO_CT_ResetProtectFault WB, "S09_ProtectFault", False
        DEMO_CT_ResetProtectFault WB, "S10_ProtectAfterOp", True
        DEMO_CT_ResetSelectionFault WB, "S11_SelectionFault"
        DEMO_CT_ResetAllCleanupFaults WB, "S12_AllFaults"

    'BuildTemplate propagation
        DEMO_CT_BuildOrdinaryError WB, "S13_BuildError"
        DEMO_CT_BuildNestedResetError WB, "S14_BuildNested"

Safe_Exit:
'------------------------------------------------------------------------------
' RELEASE AND REPORT
'------------------------------------------------------------------------------
    'Disarm the probe and discard the temporary workbook
        On Error Resume Next
            DEMO_InternalCleanupProbeEnd
            If Not WB Is Nothing Then
                WB.Close SaveChanges:=False
            End If
            Set WB = Nothing
        On Error GoTo 0

    'Print the summary line
        If RaisedNumber = 0 And mFailures = 0 Then
            Summary = "DEMO CLEANUP CHECKS: PASS"
        Else
            Summary = "DEMO CLEANUP CHECKS: FAIL"
        End If
        Summary = Summary & " | scenarios " & CStr(mScenarios) & _
                  ", failed " & CStr(mScenarioFailures) & _
                  " | checks " & CStr(mChecks) & _
                  ", failed " & CStr(mFailures)
        If RaisedNumber <> 0 Then
            Summary = Summary & " | run stopped: " & CStr(RaisedNumber) & _
                      " | " & RaisedSource & " | " & RaisedDescription
        End If
        Debug.Print Summary

    'Report a failed run to the caller; the handler is already disabled
        If RaisedNumber <> 0 Then
            Err.Raise RaisedNumber, RaisedSource, RaisedDescription
        End If
        If mFailures <> 0 Then
            Err.Raise ERR_RUN_FAILED, PROC, Summary
        End If

        Exit Sub

Err_Handler:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the error that stopped the run, then release and report once
        RaisedNumber = Err.Number
        RaisedSource = Err.Source
        RaisedDescription = Err.Description

        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_ResetSuccess( _
    ByVal WB As Workbook, _
    ByVal Scenario As String, _
    ByVal Options As String, _
    ByVal SelectionPolicy As Long)
'
'==============================================================================
' DEMO_CT_ResetSuccess
'------------------------------------------------------------------------------
' PURPOSE
'   S01/S02: a protected sheet reset with the right password and
'   ReProtectAtEnd=True ends protected again with every entry option.
'
' CHECKS
'   - the host kept the requested options before the reset
'   - no error; one cleanup pass; no handler entry; one Protect and one
'     EnableSelection attempt
'   - the reset body ran
'   - the fifteen options and EnableSelection read back as on entry
'   - the supplied password protects the sheet again
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet       'Scenario sheet
    Dim EntryOptions        As String          'Options read back on entry
    Dim EntrySelection      As Long            'EnableSelection read back on entry
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description

'------------------------------------------------------------------------------
' ARRANGE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        Set WS = DEMO_CT_NewSheet(WB, Scenario)
        DEMO_CT_Protect WS, Options, SelectionPolicy

        EntryOptions = DEMO_CT_ReadOptions(WS)
        EntrySelection = WS.EnableSelection

        DEMO_CT_CheckOptions Scenario & ".Setup.Options", Options, EntryOptions
        DEMO_CT_CheckLong Scenario & ".Setup.EnableSelection", SelectionPolicy, EntrySelection

'------------------------------------------------------------------------------
' ACT
'------------------------------------------------------------------------------
        DEMO_CT_CallReset WS, FAULT_NONE, PROTECT_PASSWORD, True, _
                          RaisedNumber, RaisedSource, RaisedDescription

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        DEMO_CT_CheckNoError Scenario & ".NoError", RaisedNumber, RaisedSource, RaisedDescription
        DEMO_CT_CheckProbe Scenario & ".Probe", 0, 1, 1, 1, 0, 0
        DEMO_CT_Check DEMO_CT_BodyRan(WS), Scenario & ".BodyRan", "the reset body did not clear A1"
        DEMO_CT_CheckOptions Scenario & ".Readback.Options", EntryOptions, DEMO_CT_ReadOptions(WS)
        DEMO_CT_CheckLong Scenario & ".Readback.EnableSelection", EntrySelection, WS.EnableSelection
        DEMO_CT_CheckPasswordReapplied Scenario & ".Readback.Password", WS

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_ResetWithoutReprotect( _
    ByVal WB As Workbook, _
    ByVal Scenario As String)
'
'==============================================================================
' DEMO_CT_ResetWithoutReprotect
'------------------------------------------------------------------------------
' PURPOSE
'   S03: with the default ReProtectAtEnd=False, a successful reset leaves the
'   sheet unprotected and never attempts Protect.
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet       'Scenario sheet
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description

'------------------------------------------------------------------------------
' ARRANGE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        Set WS = DEMO_CT_NewSheet(WB, Scenario)
        DEMO_CT_Protect WS, OPTIONS_A, xlUnlockedCells

'------------------------------------------------------------------------------
' ACT
'------------------------------------------------------------------------------
        DEMO_CT_CallReset WS, FAULT_NONE, PROTECT_PASSWORD, False, _
                          RaisedNumber, RaisedSource, RaisedDescription

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        DEMO_CT_CheckNoError Scenario & ".NoError", RaisedNumber, RaisedSource, RaisedDescription
        DEMO_CT_CheckProbe Scenario & ".Probe", 0, 1, 0, 0, 0, 0
        DEMO_CT_Check DEMO_CT_BodyRan(WS), Scenario & ".BodyRan", "the reset body did not clear A1"
        DEMO_CT_Check Not DEMO_CT_IsProtected(WS), Scenario & ".Unprotected", _
                      "the sheet is protected although ReProtectAtEnd was False"

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_ResetNotWorksheet( _
    ByVal WB As Workbook, _
    ByVal Scenario As String)
'
'==============================================================================
' DEMO_CT_ResetNotWorksheet
'------------------------------------------------------------------------------
' PURPOSE
'   S04: a real ordinary error raised inside the guarded body (WS_In is a
'   Workbook) reaches the caller once, unchanged, after one handler entry and
'   one cleanup pass, with no protection attempt.
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description

'------------------------------------------------------------------------------
' ACT
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        DEMO_CT_CallReset WB, FAULT_NONE, PROTECT_PASSWORD, True, _
                          RaisedNumber, RaisedSource, RaisedDescription

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        DEMO_CT_CheckError Scenario & ".Error", ERR_NOT_WORKSHEET, SRC_RESET, DESC_NOT_WORKSHEET, _
                           RaisedNumber, RaisedSource, RaisedDescription
        DEMO_CT_CheckProbe Scenario & ".Probe", 1, 1, 0, 0, 0, 0

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_ResetOperationFault( _
    ByVal WB As Workbook, _
    ByVal Scenario As String)
'
'==============================================================================
' DEMO_CT_ResetOperationFault
'------------------------------------------------------------------------------
' PURPOSE
'   S05: an ordinary error after this call's own unprotect reaches the caller
'   once and unchanged, and cleanup still restores the entry protection.
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet       'Scenario sheet
    Dim EntryOptions        As String          'Options read back on entry
    Dim EntrySelection      As Long            'EnableSelection read back on entry
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description

'------------------------------------------------------------------------------
' ARRANGE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        Set WS = DEMO_CT_NewSheet(WB, Scenario)
        DEMO_CT_Protect WS, OPTIONS_A, xlUnlockedCells

        EntryOptions = DEMO_CT_ReadOptions(WS)
        EntrySelection = WS.EnableSelection

'------------------------------------------------------------------------------
' ACT
'------------------------------------------------------------------------------
        DEMO_CT_CallReset WS, FAULT_BODY, PROTECT_PASSWORD, True, _
                          RaisedNumber, RaisedSource, RaisedDescription

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        DEMO_CT_CheckError Scenario & ".Error", ERR_BODY, SRC_BODY, DESC_BODY, _
                           RaisedNumber, RaisedSource, RaisedDescription
        DEMO_CT_CheckProbe Scenario & ".Probe", 1, 1, 1, 1, 0, 0
        DEMO_CT_CheckOptions Scenario & ".Readback.Options", EntryOptions, DEMO_CT_ReadOptions(WS)
        DEMO_CT_CheckLong Scenario & ".Readback.EnableSelection", EntrySelection, WS.EnableSelection
        DEMO_CT_CheckPasswordReapplied Scenario & ".Readback.Password", WS

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_ResetOperationFaultUnprotected( _
    ByVal WB As Workbook, _
    ByVal Scenario As String)
'
'==============================================================================
' DEMO_CT_ResetOperationFaultUnprotected
'------------------------------------------------------------------------------
' PURPOSE
'   S06: the same ordinary error on a sheet that was not protected reaches the
'   caller once, and cleanup does not protect a sheet this call never
'   unprotected, even with a password and ReProtectAtEnd=True.
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet       'Scenario sheet
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description

'------------------------------------------------------------------------------
' ARRANGE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        Set WS = DEMO_CT_NewSheet(WB, Scenario)

'------------------------------------------------------------------------------
' ACT
'------------------------------------------------------------------------------
        DEMO_CT_CallReset WS, FAULT_BODY, PROTECT_PASSWORD, True, _
                          RaisedNumber, RaisedSource, RaisedDescription

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        DEMO_CT_CheckError Scenario & ".Error", ERR_BODY, SRC_BODY, DESC_BODY, _
                           RaisedNumber, RaisedSource, RaisedDescription
        DEMO_CT_CheckProbe Scenario & ".Probe", 1, 1, 0, 0, 0, 0
        DEMO_CT_Check Not DEMO_CT_IsProtected(WS), Scenario & ".StillUnprotected", _
                      "cleanup protected a sheet this call never unprotected"

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_ResetUnprotectFailure( _
    ByVal WB As Workbook, _
    ByVal Scenario As String, _
    ByVal Password As String)
'
'==============================================================================
' DEMO_CT_ResetUnprotectFailure
'------------------------------------------------------------------------------
' PURPOSE
'   S07/S08: a wrong or missing password fails the unprotect; the error
'   reaches the caller once without a cleanup diagnostic, no Protect is
'   attempted, the body does not run and the entry protection is unchanged.
'
' INPUTS
'   Password
'     WRONG_PASSWORD, or vbNullString for a missing password (the default of
'     ProtectPassword)
'
' CHECKS
'   - missing password: the exact vbObjectError + 2100 validation error
'   - wrong password: Excel's error, whatever its number, without a cleanup
'     diagnostic
'   - the original password still unprotects the sheet afterwards
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet       'Scenario sheet
    Dim EntryOptions        As String          'Options read back on entry
    Dim EntrySelection      As Long            'EnableSelection read back on entry
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description
    Dim UnprotectNumber     As Long            'Error of the final original-password unprotect

'------------------------------------------------------------------------------
' ARRANGE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        Set WS = DEMO_CT_NewSheet(WB, Scenario)
        DEMO_CT_Protect WS, OPTIONS_A, xlUnlockedCells

        EntryOptions = DEMO_CT_ReadOptions(WS)
        EntrySelection = WS.EnableSelection

'------------------------------------------------------------------------------
' ACT
'------------------------------------------------------------------------------
        DEMO_CT_CallReset WS, FAULT_NONE, Password, True, _
                          RaisedNumber, RaisedSource, RaisedDescription

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        If Len(Password) = 0 Then
            DEMO_CT_CheckError Scenario & ".Error", ERR_NO_PASSWORD, SRC_RESET, DESC_NO_PASSWORD, _
                               RaisedNumber, RaisedSource, RaisedDescription
        Else
            DEMO_CT_Check RaisedNumber <> 0, Scenario & ".Error", _
                          "a wrong password reached the caller without an error"
            DEMO_CT_CheckLong Scenario & ".NoCleanupDiagnostic", 0, _
                              DEMO_CT_CountText(RaisedDescription, CLEANUP_MARKER)
        End If

        DEMO_CT_CheckProbe Scenario & ".Probe", 1, 1, 0, 0, 0, 0
        DEMO_CT_Check Not DEMO_CT_BodyRan(WS), Scenario & ".BodyNotRun", _
                      "the reset body ran although the unprotect failed"
        DEMO_CT_CheckOptions Scenario & ".Readback.Options", EntryOptions, DEMO_CT_ReadOptions(WS)
        DEMO_CT_CheckLong Scenario & ".Readback.EnableSelection", EntrySelection, WS.EnableSelection

    'The protection must still be the original one, not a reprotection
        On Error Resume Next
            WS.Unprotect Password:=PROTECT_PASSWORD
            UnprotectNumber = Err.Number
            Err.Clear
        On Error GoTo Err_Handler

        DEMO_CT_Check UnprotectNumber = 0 And Not DEMO_CT_IsProtected(WS), _
                      Scenario & ".OriginalPassword", _
                      "the original password no longer unprotects the sheet"

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_ResetProtectFault( _
    ByVal WB As Workbook, _
    ByVal Scenario As String, _
    ByVal WithOperationFault As Boolean)
'
'==============================================================================
' DEMO_CT_ResetProtectFault
'------------------------------------------------------------------------------
' PURPOSE
'   S09/S10: the cleanup Protect statement fails once.
'
' CHECKS
'   - without an operation error, the Protect failure itself reaches the
'     caller unchanged and without the cleanup prefix
'   - after an operation error, that error keeps its number and source and
'     carries the Protect failure in exactly one cleanup diagnostic
'   - Protect is attempted once, EnableSelection is still restored, the sheet
'     stays unprotected and the cleanup pass runs once
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet       'Scenario sheet
    Dim EntrySelection      As Long            'EnableSelection read back on entry
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description

'------------------------------------------------------------------------------
' ARRANGE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        Set WS = DEMO_CT_NewSheet(WB, Scenario)
        DEMO_CT_Protect WS, OPTIONS_A, xlUnlockedCells

        EntrySelection = WS.EnableSelection

'------------------------------------------------------------------------------
' ACT AND ASSERT
'------------------------------------------------------------------------------
        If WithOperationFault Then
            DEMO_CT_CallReset WS, FAULT_BODY + FAULT_PROTECT, PROTECT_PASSWORD, True, _
                              RaisedNumber, RaisedSource, RaisedDescription

            DEMO_CT_CheckError Scenario & ".Error", ERR_BODY, SRC_BODY, _
                               DESC_BODY & vbCrLf & CLEANUP_MARKER & _
                               DEMO_CT_ErrText(ERR_PROTECT, SRC_PROTECT, DESC_PROTECT), _
                               RaisedNumber, RaisedSource, RaisedDescription
            DEMO_CT_CheckProbe Scenario & ".Probe", 1, 1, 1, 1, 0, 0
        Else
            DEMO_CT_CallReset WS, FAULT_PROTECT, PROTECT_PASSWORD, True, _
                              RaisedNumber, RaisedSource, RaisedDescription

            DEMO_CT_CheckError Scenario & ".Error", ERR_PROTECT, SRC_PROTECT, DESC_PROTECT, _
                               RaisedNumber, RaisedSource, RaisedDescription
            DEMO_CT_CheckProbe Scenario & ".Probe", 0, 1, 1, 1, 0, 0
        End If

        DEMO_CT_CheckLong Scenario & ".SingleCleanupDiagnostic", IIf(WithOperationFault, 1, 0), _
                          DEMO_CT_CountText(RaisedDescription, CLEANUP_MARKER)
        DEMO_CT_Check DEMO_CT_BodyRan(WS), Scenario & ".BodyRan", "the reset body did not clear A1"
        DEMO_CT_Check Not DEMO_CT_IsProtected(WS), Scenario & ".Unprotected", _
                      "the sheet is protected although the only Protect attempt failed"
        DEMO_CT_CheckLong Scenario & ".EnableSelectionRestored", EntrySelection, WS.EnableSelection

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_ResetSelectionFault( _
    ByVal WB As Workbook, _
    ByVal Scenario As String)
'
'==============================================================================
' DEMO_CT_ResetSelectionFault
'------------------------------------------------------------------------------
' PURPOSE
'   S11: Protect succeeds and the EnableSelection restoration fails; that
'   failure reaches the caller directly, once, and the fifteen options are
'   still restored with the supplied password.
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet       'Scenario sheet
    Dim EntryOptions        As String          'Options read back on entry
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description

'------------------------------------------------------------------------------
' ARRANGE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        Set WS = DEMO_CT_NewSheet(WB, Scenario)
        DEMO_CT_Protect WS, OPTIONS_A, xlUnlockedCells

        EntryOptions = DEMO_CT_ReadOptions(WS)

'------------------------------------------------------------------------------
' ACT
'------------------------------------------------------------------------------
        DEMO_CT_CallReset WS, FAULT_SELECTION, PROTECT_PASSWORD, True, _
                          RaisedNumber, RaisedSource, RaisedDescription

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        DEMO_CT_CheckError Scenario & ".Error", ERR_SELECTION, SRC_SELECTION, DESC_SELECTION, _
                           RaisedNumber, RaisedSource, RaisedDescription
        DEMO_CT_CheckProbe Scenario & ".Probe", 0, 1, 1, 1, 0, 0
        DEMO_CT_CheckOptions Scenario & ".Readback.Options", EntryOptions, DEMO_CT_ReadOptions(WS)
        DEMO_CT_CheckPasswordReapplied Scenario & ".Readback.Password", WS

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_ResetAllCleanupFaults( _
    ByVal WB As Workbook, _
    ByVal Scenario As String)
'
'==============================================================================
' DEMO_CT_ResetAllCleanupFaults
'------------------------------------------------------------------------------
' PURPOSE
'   S12: an operation error followed by Protect and EnableSelection failures.
'   The operation error keeps its number and source; one cleanup diagnostic
'   carries the Protect failure and then the EnableSelection failure.
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet       'Scenario sheet
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description

'------------------------------------------------------------------------------
' ARRANGE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        Set WS = DEMO_CT_NewSheet(WB, Scenario)
        DEMO_CT_Protect WS, OPTIONS_A, xlUnlockedCells

'------------------------------------------------------------------------------
' ACT
'------------------------------------------------------------------------------
        DEMO_CT_CallReset WS, FAULT_BODY + FAULT_PROTECT + FAULT_SELECTION, PROTECT_PASSWORD, True, _
                          RaisedNumber, RaisedSource, RaisedDescription

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        DEMO_CT_CheckError Scenario & ".Error", ERR_BODY, SRC_BODY, _
                           DESC_BODY & vbCrLf & CLEANUP_MARKER & _
                           DEMO_CT_ErrText(ERR_PROTECT, SRC_PROTECT, DESC_PROTECT) & _
                           SELECTION_MARKER & _
                           DEMO_CT_ErrText(ERR_SELECTION, SRC_SELECTION, DESC_SELECTION), _
                           RaisedNumber, RaisedSource, RaisedDescription
        DEMO_CT_CheckLong Scenario & ".SingleCleanupDiagnostic", 1, _
                          DEMO_CT_CountText(RaisedDescription, CLEANUP_MARKER)
        DEMO_CT_CheckProbe Scenario & ".Probe", 1, 1, 1, 1, 0, 0
        DEMO_CT_Check Not DEMO_CT_IsProtected(WS), Scenario & ".Unprotected", _
                      "the sheet is protected although the only Protect attempt failed"

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_BuildOrdinaryError( _
    ByVal WB As Workbook, _
    ByVal Scenario As String)
'
'==============================================================================
' DEMO_CT_BuildOrdinaryError
'------------------------------------------------------------------------------
' PURPOSE
'   S13: DEMO_Sheet_BuildTemplate raises a real validation error inside its
'   guarded body, after its own reset. The error reaches the caller once and
'   unchanged after one handler entry and one cleanup pass.
'
' NOTES
'   The error is raised before any window or Application setting is applied,
'   so this scenario changes nothing outside the temporary workbook.
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description

'------------------------------------------------------------------------------
' ACT
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        DEMO_CT_CallBuild WB, Scenario, RaisedNumber, RaisedSource, RaisedDescription

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        DEMO_CT_CheckError Scenario & ".Error", ERR_BLANK_CONTENT, SRC_BUILD, DESC_BLANK_CONTENT, _
                           RaisedNumber, RaisedSource, RaisedDescription
        DEMO_CT_CheckProbe Scenario & ".Probe", 0, 1, 0, 0, 1, 1

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_BuildNestedResetError( _
    ByVal WB As Workbook, _
    ByVal Scenario As String)
'
'==============================================================================
' DEMO_CT_BuildNestedResetError
'------------------------------------------------------------------------------
' PURPOSE
'   S14: BuildTemplate on a password-protected sheet. Its nested reset has no
'   password and raises; both procedures pass the error on once, unchanged,
'   and the sheet keeps its entry protection.
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet       'Scenario sheet
    Dim EntryOptions        As String          'Options read back on entry
    Dim FailuresBefore      As Long            'Failure count at scenario start
    Dim RaisedNumber        As Long            'Error reaching the caller
    Dim RaisedSource        As String          'Its source
    Dim RaisedDescription   As String          'Its description

'------------------------------------------------------------------------------
' ARRANGE
'------------------------------------------------------------------------------
        On Error GoTo Err_Handler

        FailuresBefore = DEMO_CT_BeginScenario()

        Set WS = DEMO_CT_NewSheet(WB, Scenario)
        DEMO_CT_Protect WS, OPTIONS_A, xlUnlockedCells

        EntryOptions = DEMO_CT_ReadOptions(WS)

'------------------------------------------------------------------------------
' ACT
'------------------------------------------------------------------------------
        DEMO_CT_CallBuild WB, Scenario, RaisedNumber, RaisedSource, RaisedDescription

'------------------------------------------------------------------------------
' ASSERT
'------------------------------------------------------------------------------
        DEMO_CT_CheckError Scenario & ".Error", ERR_NO_PASSWORD, SRC_RESET, DESC_NO_PASSWORD, _
                           RaisedNumber, RaisedSource, RaisedDescription
        DEMO_CT_CheckProbe Scenario & ".Probe", 1, 1, 0, 0, 1, 1
        DEMO_CT_Check Not DEMO_CT_BodyRan(WS), Scenario & ".BodyNotRun", _
                      "the nested reset body ran although the sheet stayed protected"
        DEMO_CT_CheckOptions Scenario & ".Readback.Options", EntryOptions, DEMO_CT_ReadOptions(WS)

Safe_Exit:
        On Error GoTo 0
        DEMO_CT_EndScenario Scenario, FailuresBefore
        Exit Sub

Err_Handler:
        DEMO_CT_Fail Scenario & ".Unexpected", DEMO_CT_ErrText(Err.Number, Err.Source, Err.Description)
        Resume Safe_Exit

End Sub


Private Sub DEMO_CT_CallReset( _
    ByVal Target As Variant, _
    ByVal FaultMask As Long, _
    ByVal Password As String, _
    ByVal ReProtect As Boolean, _
    ByRef RaisedNumber As Long, _
    ByRef RaisedSource As String, _
    ByRef RaisedDescription As String)
'
'==============================================================================
' DEMO_CT_CallReset
'------------------------------------------------------------------------------
' PURPOSE
'   Calls DEMO_Sheet_Reset once with the probe armed and returns the error
'   that reached the caller, if any.
'
' BEHAVIOR
'   - Arms the probe with FaultMask, makes one call, captures Err immediately
'     and disarms the probe whatever happened
'   - Every other reset option keeps its default
'
' ERROR POLICY
'   Does not raise
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' CALL
'------------------------------------------------------------------------------
        DEMO_InternalCleanupProbeBegin FaultMask

        On Error Resume Next
            DEMO_Sheet_Reset WS_In:=Target, ProtectPassword:=Password, ReProtectAtEnd:=ReProtect
            RaisedNumber = Err.Number
            RaisedSource = Err.Source
            RaisedDescription = Err.Description
            Err.Clear
        On Error GoTo 0

        DEMO_InternalCleanupProbeEnd

End Sub


Private Sub DEMO_CT_CallBuild( _
    ByVal WB As Workbook, _
    ByVal SheetName As String, _
    ByRef RaisedNumber As Long, _
    ByRef RaisedSource As String, _
    ByRef RaisedDescription As String)
'
'==============================================================================
' DEMO_CT_CallBuild
'------------------------------------------------------------------------------
' PURPOSE
'   Calls DEMO_Sheet_BuildTemplate once for SheetName in WB, with a blank
'   ContentColumns so that it fails validation after its reset, and returns
'   the error that reached the caller.
'
' ERROR POLICY
'   Does not raise
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' CALL
'------------------------------------------------------------------------------
        DEMO_InternalCleanupProbeBegin FAULT_NONE

        On Error Resume Next
            DEMO_Sheet_BuildTemplate WS_Name:=SheetName, TargetWorkbook:=WB, ContentColumns:=vbNullString
            RaisedNumber = Err.Number
            RaisedSource = Err.Source
            RaisedDescription = Err.Description
            Err.Clear
        On Error GoTo 0

        DEMO_InternalCleanupProbeEnd

End Sub


Private Function DEMO_CT_NewSheet( _
    ByVal WB As Workbook, _
    ByVal SheetName As String) _
    As Worksheet
'
'==============================================================================
' DEMO_CT_NewSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Adds an unprotected scenario sheet to the temporary workbook, with the
'   body marker in A1.
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet       'New sheet

'------------------------------------------------------------------------------
' CREATE
'------------------------------------------------------------------------------
        Set WS = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
        WS.Name = SheetName
        WS.Range("A1").Value = BODY_MARKER

        Set DEMO_CT_NewSheet = WS

End Function


Private Sub DEMO_CT_Protect( _
    ByVal WS As Worksheet, _
    ByVal Options As String, _
    ByVal SelectionPolicy As Long)
'
'==============================================================================
' DEMO_CT_Protect
'------------------------------------------------------------------------------
' PURPOSE
'   Protects WS with PROTECT_PASSWORD, the fifteen options of an option
'   string, and the given EnableSelection.
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' PROTECT
'------------------------------------------------------------------------------
        WS.Protect Password:=PROTECT_PASSWORD, _
            DrawingObjects:=DEMO_CT_Option(Options, 1), _
            Contents:=DEMO_CT_Option(Options, 2), _
            Scenarios:=DEMO_CT_Option(Options, 3), _
            UserInterfaceOnly:=DEMO_CT_Option(Options, 4), _
            AllowFormattingCells:=DEMO_CT_Option(Options, 5), _
            AllowFormattingColumns:=DEMO_CT_Option(Options, 6), _
            AllowFormattingRows:=DEMO_CT_Option(Options, 7), _
            AllowInsertingColumns:=DEMO_CT_Option(Options, 8), _
            AllowInsertingRows:=DEMO_CT_Option(Options, 9), _
            AllowInsertingHyperlinks:=DEMO_CT_Option(Options, 10), _
            AllowDeletingColumns:=DEMO_CT_Option(Options, 11), _
            AllowDeletingRows:=DEMO_CT_Option(Options, 12), _
            AllowSorting:=DEMO_CT_Option(Options, 13), _
            AllowFiltering:=DEMO_CT_Option(Options, 14), _
            AllowUsingPivotTables:=DEMO_CT_Option(Options, 15)

        WS.EnableSelection = SelectionPolicy

End Sub


Private Function DEMO_CT_Option( _
    ByVal Options As String, _
    ByVal Index As Long) _
    As Boolean
'
'==============================================================================
' DEMO_CT_Option
'------------------------------------------------------------------------------
' PURPOSE
'   Returns option Index (1 to 15) of an option string.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_Option = (Mid$(Options, Index, 1) = "1")

End Function


Private Function DEMO_CT_ReadOptions( _
    ByVal WS As Worksheet) _
    As String
'
'==============================================================================
' DEMO_CT_ReadOptions
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the fifteen Protect options back from Excel as an option string.
'   UserInterfaceOnly is read through ProtectionMode.
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OptionText          As String          'Option string being built

'------------------------------------------------------------------------------
' READ BACK
'------------------------------------------------------------------------------
        OptionText = DEMO_CT_Bit(WS.ProtectDrawingObjects) & _
               DEMO_CT_Bit(WS.ProtectContents) & _
               DEMO_CT_Bit(WS.ProtectScenarios) & _
               DEMO_CT_Bit(WS.ProtectionMode)

        With WS.Protection
            OptionText = OptionText & _
                   DEMO_CT_Bit(.AllowFormattingCells) & _
                   DEMO_CT_Bit(.AllowFormattingColumns) & _
                   DEMO_CT_Bit(.AllowFormattingRows) & _
                   DEMO_CT_Bit(.AllowInsertingColumns) & _
                   DEMO_CT_Bit(.AllowInsertingRows) & _
                   DEMO_CT_Bit(.AllowInsertingHyperlinks) & _
                   DEMO_CT_Bit(.AllowDeletingColumns) & _
                   DEMO_CT_Bit(.AllowDeletingRows) & _
                   DEMO_CT_Bit(.AllowSorting) & _
                   DEMO_CT_Bit(.AllowFiltering) & _
                   DEMO_CT_Bit(.AllowUsingPivotTables)
        End With

        DEMO_CT_ReadOptions = OptionText

End Function


Private Function DEMO_CT_Bit( _
    ByVal Value As Boolean) _
    As String
'
'==============================================================================
' DEMO_CT_Bit
'------------------------------------------------------------------------------
' PURPOSE
'   Returns "1" for True and "0" for False.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_Bit = IIf(Value, "1", "0")

End Function


Private Function DEMO_CT_IsProtected( _
    ByVal WS As Worksheet) _
    As Boolean
'
'==============================================================================
' DEMO_CT_IsProtected
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when any of the three protection kinds is active, the test
'   DEMO_Sheet_Reset itself uses.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_IsProtected = WS.ProtectContents Or WS.ProtectDrawingObjects Or WS.ProtectScenarios

End Function


Private Function DEMO_CT_BodyRan( _
    ByVal WS As Worksheet) _
    As Boolean
'
'==============================================================================
' DEMO_CT_BodyRan
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when the reset body cleared the A1 marker.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_BodyRan = IsEmpty(WS.Range("A1").Value)

End Function


Private Function DEMO_CT_CountText( _
    ByVal Haystack As String, _
    ByVal Fragment As String) _
    As Long
'
'==============================================================================
' DEMO_CT_CountText
'------------------------------------------------------------------------------
' PURPOSE
'   Counts the occurrences of Fragment in Haystack, without a loop.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_CountText = (Len(Haystack) - Len(Replace(Haystack, Fragment, vbNullString))) \ Len(Fragment)

End Function


Private Function DEMO_CT_ErrText( _
    ByVal Number As Long, _
    ByVal Source As String, _
    ByVal Description As String) _
    As String
'
'==============================================================================
' DEMO_CT_ErrText
'------------------------------------------------------------------------------
' PURPOSE
'   Formats an error as "number | source | description", the form
'   DEMO_Sheet_Reset uses in its cleanup diagnostics.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_ErrText = CStr(Number) & " | " & Source & " | " & Description

End Function


Private Function DEMO_CT_BeginScenario() As Long
'
'==============================================================================
' DEMO_CT_BeginScenario
'------------------------------------------------------------------------------
' PURPOSE
'   Counts a scenario and returns the failure count at its start.
'
' UPDATED
'   2026-11-06
'==============================================================================

        mScenarios = mScenarios + 1
        DEMO_CT_BeginScenario = mFailures

End Function


Private Sub DEMO_CT_EndScenario( _
    ByVal Scenario As String, _
    ByVal FailuresBefore As Long)
'
'==============================================================================
' DEMO_CT_EndScenario
'------------------------------------------------------------------------------
' PURPOSE
'   Disarms the probe and prints the scenario result line.
'
' ERROR POLICY
'   Does not raise, so it is safe at Safe_Exit
'
' UPDATED
'   2026-11-06
'==============================================================================

        On Error Resume Next
            DEMO_InternalCleanupProbeEnd
        On Error GoTo 0

        If mFailures = FailuresBefore Then
            Debug.Print "  PASS " & Scenario
        Else
            mScenarioFailures = mScenarioFailures + 1
            Debug.Print "  FAIL " & Scenario & " (" & CStr(mFailures - FailuresBefore) & " check(s) failed)"
        End If

End Sub


Private Sub DEMO_CT_Check( _
    ByVal Condition As Boolean, _
    ByVal CheckName As String, _
    ByVal Detail As String)
'
'==============================================================================
' DEMO_CT_Check
'------------------------------------------------------------------------------
' PURPOSE
'   Records one check and prints its name and detail when it fails.
'
' UPDATED
'   2026-11-06
'==============================================================================

        mChecks = mChecks + 1

        If Not Condition Then
            DEMO_CT_Fail CheckName, Detail
        End If

End Sub


Private Sub DEMO_CT_Fail( _
    ByVal CheckName As String, _
    ByVal Detail As String)
'
'==============================================================================
' DEMO_CT_Fail
'------------------------------------------------------------------------------
' PURPOSE
'   Records and prints one failure.
'
' UPDATED
'   2026-11-06
'==============================================================================

        mFailures = mFailures + 1
        Debug.Print "    FAIL " & CheckName & ": " & Detail

End Sub


Private Sub DEMO_CT_CheckLong( _
    ByVal CheckName As String, _
    ByVal Expected As Long, _
    ByVal Actual As Long)
'
'==============================================================================
' DEMO_CT_CheckLong
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that two numbers are equal.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_Check Expected = Actual, CheckName, _
                      "expected " & CStr(Expected) & " but read " & CStr(Actual)

End Sub


Private Sub DEMO_CT_CheckText( _
    ByVal CheckName As String, _
    ByVal Expected As String, _
    ByVal Actual As String)
'
'==============================================================================
' DEMO_CT_CheckText
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that two strings are equal, case included.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_Check StrComp(Expected, Actual, vbBinaryCompare) = 0, CheckName, _
                      "expected [" & Expected & "] but read [" & Actual & "]"

End Sub


Private Sub DEMO_CT_CheckNoError( _
    ByVal CheckName As String, _
    ByVal Number As Long, _
    ByVal Source As String, _
    ByVal Description As String)
'
'==============================================================================
' DEMO_CT_CheckNoError
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that no error reached the caller.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_Check Number = 0, CheckName, _
                      "unexpected error " & DEMO_CT_ErrText(Number, Source, Description)

End Sub


Private Sub DEMO_CT_CheckError( _
    ByVal CheckName As String, _
    ByVal ExpectedNumber As Long, _
    ByVal ExpectedSource As String, _
    ByVal ExpectedDescription As String, _
    ByVal Number As Long, _
    ByVal Source As String, _
    ByVal Description As String)
'
'==============================================================================
' DEMO_CT_CheckError
'------------------------------------------------------------------------------
' PURPOSE
'   Checks the number, source and full description of the error that reached
'   the caller. An exact description also proves no cleanup diagnostic was
'   added or repeated beyond the expected text.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_CheckLong CheckName & ".Number", ExpectedNumber, Number
        DEMO_CT_CheckText CheckName & ".Source", ExpectedSource, Source
        DEMO_CT_CheckText CheckName & ".Description", ExpectedDescription, Description

End Sub


Private Sub DEMO_CT_CheckProbe( _
    ByVal CheckName As String, _
    ByVal ResetHandler As Long, _
    ByVal ResetCleanup As Long, _
    ByVal ResetProtect As Long, _
    ByVal ResetSelection As Long, _
    ByVal BuildHandler As Long, _
    ByVal BuildCleanup As Long)
'
'==============================================================================
' DEMO_CT_CheckProbe
'------------------------------------------------------------------------------
' PURPOSE
'   Checks all six cleanup probe counters of the last call at once.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_CheckText CheckName, _
            DEMO_CT_ProbeText(ResetHandler, ResetCleanup, ResetProtect, _
                              ResetSelection, BuildHandler, BuildCleanup), _
            DEMO_CT_ProbeText(DEMO_InternalCleanupProbeCount(PROBE_RESET_HANDLER), _
                              DEMO_InternalCleanupProbeCount(PROBE_RESET_CLEANUP), _
                              DEMO_InternalCleanupProbeCount(PROBE_RESET_PROTECT), _
                              DEMO_InternalCleanupProbeCount(PROBE_RESET_SELECTION), _
                              DEMO_InternalCleanupProbeCount(PROBE_BUILD_HANDLER), _
                              DEMO_InternalCleanupProbeCount(PROBE_BUILD_CLEANUP))

End Sub


Private Function DEMO_CT_ProbeText( _
    ByVal ResetHandler As Long, _
    ByVal ResetCleanup As Long, _
    ByVal ResetProtect As Long, _
    ByVal ResetSelection As Long, _
    ByVal BuildHandler As Long, _
    ByVal BuildCleanup As Long) _
    As String
'
'==============================================================================
' DEMO_CT_ProbeText
'------------------------------------------------------------------------------
' PURPOSE
'   Formats the six probe counters for comparison and failure messages.
'
' UPDATED
'   2026-11-06
'==============================================================================

        DEMO_CT_ProbeText = "Reset handler=" & CStr(ResetHandler) & _
                            " cleanup=" & CStr(ResetCleanup) & _
                            " protect=" & CStr(ResetProtect) & _
                            " selection=" & CStr(ResetSelection) & _
                            "; Build handler=" & CStr(BuildHandler) & _
                            " cleanup=" & CStr(BuildCleanup)

End Function


Private Sub DEMO_CT_CheckOptions( _
    ByVal CheckName As String, _
    ByVal Expected As String, _
    ByVal Actual As String)
'
'==============================================================================
' DEMO_CT_CheckOptions
'------------------------------------------------------------------------------
' PURPOSE
'   Checks two option strings and names every option that differs.
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim OptionNames         As Variant         'The fifteen option names
    Dim Index               As Long            'Option index
    Dim Differences         As String          'Differing options

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        OptionNames = Split(OPTION_NAMES, ",")

        For Index = 1 To OPTION_COUNT
            If Mid$(Expected, Index, 1) <> Mid$(Actual, Index, 1) Then
                Differences = Differences & " " & OptionNames(Index - 1) & _
                              "=" & Mid$(Actual, Index, 1) & _
                              " (expected " & Mid$(Expected, Index, 1) & ")"
            End If
        Next Index

        DEMO_CT_Check Len(Differences) = 0 And Len(Actual) = OPTION_COUNT, CheckName, _
                      "options differ:" & Differences & " [expected " & Expected & _
                      ", read " & Actual & "]"

End Sub


Private Sub DEMO_CT_CheckPasswordReapplied( _
    ByVal CheckName As String, _
    ByVal WS As Worksheet)
'
'==============================================================================
' DEMO_CT_CheckPasswordReapplied
'------------------------------------------------------------------------------
' PURPOSE
'   Checks that WS is protected and refuses WRONG_PASSWORD, so the reapplied
'   protection carries a password. Leaves WS as it found it when the check
'   passes.
'
' UPDATED
'   2026-11-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WrongNumber         As Long            'Error of the wrong-password unprotect

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        On Error Resume Next
            WS.Unprotect Password:=WRONG_PASSWORD
            WrongNumber = Err.Number
            Err.Clear
        On Error GoTo 0

        DEMO_CT_Check WrongNumber <> 0 And DEMO_CT_IsProtected(WS), CheckName, _
                      "the sheet is unprotected or accepted a wrong password"

End Sub
