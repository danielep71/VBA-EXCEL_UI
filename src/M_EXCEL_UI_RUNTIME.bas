Attribute VB_Name = "M_EXCEL_UI_RUNTIME"
Option Explicit
Option Private Module

'==============================================================================
' M_EXCEL_UI_RUNTIME
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the shared fail-soft runtime services used by the public Excel UI
'   facade and by the snapshot engine: structured result buffering, ordered
'   diagnostics, Ribbon access, generic Boolean property access, and the
'   quiet-update scope that preserves Application.ScreenUpdating.
'
' WHY THIS EXISTS
'   M_EXCEL_UI and M_EXCEL_UI_SNAPSHOT both need identical result buffering,
'   diagnostic formatting, Ribbon reads and writes, object-model property reads
'   and writes, and ScreenUpdating preservation. Duplicating those operations in
'   each module would allow the two copies to drift apart, and making either
'   module depend on the other would create a circular module dependency.
'
'   Centralizing them here gives one implementation of the "Stage | Detail"
'   diagnostic contract, one no-op suppression policy, and one quiet-update
'   scope, while leaving this module free of any project-module dependency of
'   its own.
'
' INTERNAL SURFACE
'   Diagnostics and result buffers:
'     - UI_RuntimeHandleFailure
'     - UI_RuntimeClearResultBuffer
'     - UI_RuntimeBuildErrorText
'     - UI_RuntimeBuildWindowLabel
'     - UI_RuntimeLogFailure
'
'   Quiet-update scope:
'     - UI_RuntimeBeginQuietUpdate
'     - UI_RuntimeEndQuietUpdate
'
'   Host access:
'     - UI_RuntimeTryGetRibbonVisible
'     - UI_RuntimeTrySetRibbonVisibleIfNeeded
'     - UI_RuntimeTryGetBooleanProperty
'     - UI_RuntimeTrySetBooleanPropertyIfNeeded
'
' DESIGN PRINCIPLES
'   - Every Try... entry point returns a Boolean and reports its reason through
'     a ByRef FailMsg rather than raising to the caller.
'   - Reads are attempted before writes so that a property already holding the
'     requested value is never written again.
'   - A failed read never blocks the corresponding write; it only disables the
'     no-op short circuit.
'   - Diagnostics are data, not user interface. Nothing here displays a dialog.
'   - The module holds no mutable state, so it has no lifecycle of its own.
'
' DIAGNOSTIC CONTRACT
'   Failures are recorded in insertion order as:
'
'       Stage | Detail
'
'   Stage names the managed element or phase that failed ("Ribbon",
'   "StatusBar", "Headings [Book1 :: Book1]", "Unexpected"). Detail carries the
'   host error text or an explicit validation reason. Callers that requested a
'   failure list receive a 1-based String array of those entries.
'
' ERROR POLICY
'   - Entry points are fail-soft and do not raise to callers.
'   - Immediate Window logging happens only when LogFailures is True.
'   - No MsgBox or other user-interface message is raised.
'
' DEPENDENCIES
'   None. This module deliberately has no project-module dependency so that it
'   can sit underneath both M_EXCEL_UI and M_EXCEL_UI_SNAPSHOT.
'
' PLATFORM / COMPATIBILITY
'   - Windows and macOS safe: this module contains no WinAPI declaration.
'   - Ribbon control uses Application.ExecuteExcel4Macro and is therefore
'     best effort and dependent on legacy macro support in the host.
'
' DIAGNOSTIC DURABILITY
'   The failure accumulator is called FROM error handlers. Anything it can do
'   that raises therefore destroys the very failure it was invoked to record,
'   and turns a documented fail-soft contract into an unexpected error for the
'   whole operation.
'
'   It is consequently structured so that the most important outputs are the
'   ones that cannot fail:
'
'       Succeeded     a Boolean assignment; cannot fail
'       FailureCount  a Long increment; cannot fail
'       FailureList   an allocation; can fail under memory pressure or when the
'                     buffer holds something other than the expected array
'
'   FailureCount is therefore authoritative. FailureList is best effort and can
'   hold fewer entries than the count, but never silently: a list that could not
'   be grown carries a truncation marker written into a slot that already
'   exists, so no allocation is required to report that allocation failed.
'
' NOTES
'   - VBA's And and Or are not short-circuit: both operands are always
'     evaluated. Guards whose right operand can fault must be nested.
'   - UI_RuntimeAddFailure grows the failure list with ReDim Preserve. Growth
'     is attempted, not assumed; see DIAGNOSTIC DURABILITY above.
'
' UPDATED
'   2026-08-19 - Failure accumulation made non-raising, so a diagnostic failure
'                can no longer replace the failure being diagnosed. Fixes
'                ICR-UI-P2-02.
'   2026-08-18 - Reformatted to the project house style. No behavior change.
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.1
'==============================================================================


'==============================================================================
' PRIVATE MODULE STATE
'==============================================================================

'Regression seam. When True the next failure-list growth reports failure without
'allocating, so the diagnostic-degradation path can be exercised without a way
'to exhaust memory on demand.
Private m_InjectFailureListGrowthFailure As Boolean


Public Sub UI_RuntimeHandleFailure( _
    ByVal ProcName As String, _
    ByVal LogFailures As Boolean, _
    ByRef Succeeded As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean, _
    ByVal Stage As String, _
    ByVal Detail As String)
'
'==============================================================================
' UI_RuntimeHandleFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Records one best-effort operation failure and optionally logs it.
'
' WHY THIS EXISTS
'   Both the fire-and-forget path and the structured-result path must record a
'   failure identically; only the logging differs. Routing both through one
'   helper keeps the ordered diagnostic contract from drifting between them.
'
' INPUTS
'   ProcName
'     Public caller name used as the log prefix.
'
'   LogFailures
'     True to emit an Immediate Window line; False for result-only use.
'
'   Succeeded
'     ByRef success flag. Set to False by this call.
'
'   FailureCount
'     ByRef running failure count. Incremented by this call.
'
'   FailureList
'     ByRef failure-list buffer. Appended to only when CaptureFailureList.
'
'   CaptureFailureList
'     True when FailureList should be populated.
'
'   Stage
'     Managed element or phase that failed.
'
'   Detail
'     Host error text or explicit validation reason.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Appends the failure to the Boolean / count / list result contract.
'   - Emits one Immediate Window line when logging is requested.
'
' ERROR POLICY
'   - Does not raise, under any circumstances. This is a hard requirement, not
'     a best effort: the procedure is called from error handlers.
'   - Clears the success flag before anything that can fail is attempted.
'
' DEPENDENCIES
'   - UI_RuntimeAddFailure
'   - UI_RuntimeLogFailure
'
' CALLED FROM
'   - M_EXCEL_UI
'   - M_EXCEL_UI_SNAPSHOT
'
' NOTES
'   The procedure name parameter is retained for compatibility with the
'   established internal apply path; the helper is also used by snapshot
'   operations.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'This procedure runs inside the caller's error handler. Nothing it does may
    'propagate, or the failure being recorded is replaced by the failure to
    'record it.
        On Error GoTo Err_Handler

    'Mark the operation unsuccessful FIRST. It is the one output that cannot
    'fail, and the one a caller must never be denied.
        Succeeded = False

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
    'Append the failure to the standard result contract
        UI_RuntimeAddFailure _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:=Stage, _
            Detail:=Detail

'------------------------------------------------------------------------------
' LOG FAILURE
'------------------------------------------------------------------------------
    'Emit an Immediate Window line only for the fire-and-forget path
        If LogFailures Then
            UI_RuntimeLogFailure ProcName, Stage, Detail
        End If

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
    'A diagnostic path must never raise into an error handler. The success flag
    'was already cleared above, so the caller still learns that the operation
    'failed even when nothing else could be recorded.
        Resume Safe_Exit

End Sub


Public Sub UI_RuntimeClearResultBuffer( _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean)
'
'==============================================================================
' UI_RuntimeClearResultBuffer
'------------------------------------------------------------------------------
' PURPOSE
'   Initializes the structured result buffers to a known empty state.
'
' WHY THIS EXISTS
'   Callers may reuse the same FailureCount and FailureList variables across
'   several operations. Clearing deterministically on entry means a later
'   success can never be read as carrying an earlier run's failures.
'
' INPUTS
'   FailureCount
'     ByRef failure count. Reset to zero.
'
'   FailureList
'     ByRef failure-list buffer. Reset to Empty only when CaptureFailureList.
'
'   CaptureFailureList
'     True when FailureList is in use.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Resets the count unconditionally.
'   - Resets the list only when the caller asked for one, leaving an unused
'     Optional argument untouched.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - M_EXCEL_UI
'   - M_EXCEL_UI_SNAPSHOT
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR RESULT BUFFERS
'------------------------------------------------------------------------------
    'Reset the running failure count
        FailureCount = 0

    'Reset the failure list only when the caller requested one
        If CaptureFailureList Then
            FailureList = Empty
        End If

End Sub


Private Sub UI_RuntimeAddFailure( _
    ByRef Succeeded As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean, _
    ByVal Stage As String, _
    ByVal Detail As String)
'
'==============================================================================
' UI_RuntimeAddFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one failure to the standard Boolean / count / list result contract,
'   without ever raising.
'
' WHY THIS EXISTS
'   The three result outputs must move together, and updating them in one place
'   prevents a caller from seeing a False success flag with a zero count.
'
'   The three are not equally reliable, however, and pretending otherwise was a
'   defect. Clearing a Boolean and incrementing a Long cannot fail. Growing an
'   array can: under memory pressure, or when the buffer holds something other
'   than the expected array. Because this procedure is reached from error
'   handlers, a raise here does not merely lose one diagnostic entry - it
'   replaces the original failure with a failure to record it, and can abort a
'   pass that was designed to continue.
'
'   The order below is therefore deliberate: everything that cannot fail
'   happens before anything that can.
'
' INPUTS
'   Succeeded
'     ByRef success flag. Set to False before any fallible work is attempted.
'
'   FailureCount
'     ByRef running failure count. Incremented unconditionally, because a
'     failure occurred whether or not it could be described.
'
'   FailureList
'     ByRef failure-list buffer. Grown by one entry when captured and when
'     growth succeeds.
'
'   CaptureFailureList
'     True when FailureList should be populated.
'
'   Stage / Detail
'     Components of the ordered "Stage | Detail" diagnostic entry.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Marks the operation unsuccessful.
'   - Increments the failure count.
'   - Composes the entry text defensively, degrading to Stage alone and then to
'     a fixed marker rather than failing.
'   - Attempts to grow the list; on failure marks the list as truncated using a
'     slot that already exists, which needs no allocation.
'
' ERROR POLICY
'   - Does not raise, under any circumstances.
'   - Diagnostic degradation is reported in the list rather than hidden.
'
' DEPENDENCIES
'   - UI_RuntimeTryAppendFailureEntry
'   - UI_RuntimeMarkFailureListTruncated
'
' CALLED FROM
'   - UI_RuntimeHandleFailure
'
' NOTES
'   FailureCount is authoritative and FailureList is best effort. The list can
'   therefore hold fewer entries than the count, but never silently: a
'   truncation marker is written whenever growth failed. The buffers are always
'   cleared together by UI_RuntimeClearResultBuffer.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EntryText           As String          'Composed "Stage | Detail" entry

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' UPDATE STATUS
'------------------------------------------------------------------------------
    'Mark the overall operation unsuccessful. This assignment cannot fail and
    'is the single output a caller must never be denied.
        Succeeded = False

    'Advance the ordered failure position. A failure occurred whether or not it
    'can be described below, so the count is incremented unconditionally.
        FailureCount = FailureCount + 1

'------------------------------------------------------------------------------
' COMPOSE ENTRY TEXT
'------------------------------------------------------------------------------
    'Concatenation allocates and can therefore fail on a pathological Detail.
    'Degrade to the stage name, and then to a fixed marker, rather than losing
    'the entry altogether.
        On Error Resume Next

        EntryText = Stage & " | " & Detail

        If Err.Number <> 0 Then
            Err.Clear
            EntryText = Stage
        End If

        If Err.Number <> 0 Or Len(EntryText) = 0 Then
            Err.Clear
            EntryText = "Unknown | diagnostic text unavailable"
        End If

        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' APPEND DIAGNOSTIC ENTRY
'------------------------------------------------------------------------------
    'Grow and write the failure list only when the caller requested one
        If CaptureFailureList Then

            'A list that cannot grow must still say so. Marking an existing
            'slot needs no allocation, which is what makes the report survive
            'the very condition that caused the growth to fail.
                If Not UI_RuntimeTryAppendFailureEntry( _
                    FailureList:=FailureList, _
                    EntryText:=EntryText) Then

                    UI_RuntimeMarkFailureListTruncated _
                        FailureList:=FailureList, _
                        FailureCount:=FailureCount
                End If

        End If

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
    'The status outputs were already set above, so the caller still learns that
    'the operation failed even when nothing could be recorded about it.
        Resume Safe_Exit

End Sub


Private Function UI_RuntimeTryAppendFailureEntry( _
    ByRef FailureList As Variant, _
    ByVal EntryText As String) _
    As Boolean
'
'==============================================================================
' UI_RuntimeTryAppendFailureEntry
'------------------------------------------------------------------------------
' PURPOSE
'   Grows the failure-list buffer by one element and writes one entry into it.
'
' WHY THIS EXISTS
'   Isolating the allocation is what lets the caller treat it as fallible. It
'   also removes an assumption the previous implementation made: that the
'   buffer either was Empty or already held a String array whose bound agreed
'   with FailureCount. Neither is guaranteed. A buffer holding a Variant array,
'   a scalar, or an array whose bound has drifted would previously raise from
'   inside an error handler.
'
'   The new element index is taken from the array itself rather than from the
'   failure count, so the list stays internally consistent even if the count
'   and the buffer have diverged.
'
' INPUTS
'   FailureList
'     ByRef buffer. Replaced with the grown array on success, untouched on
'     failure.
'
'   EntryText
'     Pre-composed ordered entry to write.
'
' RETURNS
'   Boolean
'     True  => the entry was appended and published.
'     False => the buffer could not be grown; it is left as it was.
'
' BEHAVIOR
'   - Starts a new 1-based array when the buffer is Empty or unusable.
'   - Extends an existing 1-based array by one element otherwise.
'   - Publishes the grown array only after the write has succeeded.
'
' ERROR POLICY
'   - Does not raise. Any failure is reported as False.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - UI_RuntimeAddFailure
'
' NOTES
'   A regression seam can force this to report failure; see
'   UI_InternalInjectFailureListGrowthFailure.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Arr()               As String          'Working failure-list buffer
    Dim NewIndex            As Long            'Index the entry is written to

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the entry has been published
        UI_RuntimeTryAppendFailureEntry = False

'------------------------------------------------------------------------------
' CONSUME REGRESSION SEAM
'------------------------------------------------------------------------------
    'When armed, report failure without allocating. The seam is one-shot, so a
    'test that forgets to disarm it cannot suppress diagnostics indefinitely.
        If m_InjectFailureListGrowthFailure Then
            m_InjectFailureListGrowthFailure = False
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' GROW BUFFER
'------------------------------------------------------------------------------
    'Start a fresh array whenever the buffer holds nothing usable. Treating an
    'unusable buffer as empty is deliberate: losing earlier entries is bad, but
    'raising from inside an error handler is worse.
        If Not IsArray(FailureList) Then
            ReDim Arr(1 To 1)
            NewIndex = 1
        Else
            Arr = FailureList
            NewIndex = UBound(Arr) + 1
            ReDim Preserve Arr(1 To NewIndex)
        End If

'------------------------------------------------------------------------------
' WRITE ENTRY
'------------------------------------------------------------------------------
    'Write the ordered entry, then publish the grown array to the caller
        Arr(NewIndex) = EntryText

        FailureList = Arr

        UI_RuntimeTryAppendFailureEntry = True

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
    'Leave the caller's buffer exactly as it was found
        UI_RuntimeTryAppendFailureEntry = False

End Function


Private Sub UI_RuntimeMarkFailureListTruncated( _
    ByRef FailureList As Variant, _
    ByVal FailureCount As Long)
'
'==============================================================================
' UI_RuntimeMarkFailureListTruncated
'------------------------------------------------------------------------------
' PURPOSE
'   Records, inside the failure list itself, that the list could not be grown
'   and no longer describes every counted failure.
'
' WHY THIS EXISTS
'   A list that silently stops growing is worse than a short one, because a
'   caller comparing FailureCount with the number of entries has no way to tell
'   a truncated list from a bug in the count.
'
'   The marker is written into a slot that ALREADY EXISTS. That is the whole
'   point: the report has to survive the condition that caused the growth to
'   fail, so it must not itself allocate. Overwriting the final entry costs one
'   diagnostic and buys the knowledge that diagnostics were lost, which is the
'   better trade.
'
' INPUTS
'   FailureList
'     ByRef buffer to mark. Left untouched when it holds no usable slot.
'
'   FailureCount
'     Authoritative failure count, reported in the marker text.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Overwrites the last existing element with the truncation marker.
'   - Does nothing when the buffer holds no array to write into.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - UI_RuntimeAddFailure
'
' NOTES
'   The marker text is a literal joined to one CStr call. If even that cannot
'   be composed the procedure leaves the list alone rather than raising, and
'   FailureCount remains the caller's signal that failures occurred.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Arr()               As String          'Working failure-list buffer

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' VALIDATE BUFFER
'------------------------------------------------------------------------------
    'Without an existing slot there is nowhere to write without allocating,
    'which is precisely what must be avoided here
        If Not IsArray(FailureList) Then
            GoTo Safe_Exit
        End If

        Arr = FailureList

        If UBound(Arr) < 1 Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' WRITE MARKER
'------------------------------------------------------------------------------
    'Overwrite the final entry in place; no element is added
        Arr(UBound(Arr)) = _
            "Diagnostics | failure list could not be grown; " & _
            CStr(FailureCount) & " failures were counted and fewer are listed"

        FailureList = Arr

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
    'Leave the caller's buffer exactly as it was found
        Resume Safe_Exit

End Sub


Public Sub UI_InternalInjectFailureListGrowthFailure( _
    ByVal FailNextGrowth As Boolean)
'
'==============================================================================
' UI_InternalInjectFailureListGrowthFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Arms or disarms a one-shot failure of the next failure-list growth.
'
' WHY THIS EXISTS
'   The degradation path exists for conditions that cannot be produced on
'   demand - exhausted memory, or a buffer corrupted by something outside this
'   component. Without a seam the path could be reasoned about but never
'   executed, which is indistinguishable from not having written it.
'
' INPUTS
'   FailNextGrowth
'     True arms the seam; False disarms it.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - The armed state is consumed by the next growth attempt and is not
'     re-armed, so a test cannot leave the component silently unable to record
'     diagnostics.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - M_EXCEL_UI_REGRESSION_TESTS
'
' NOTES
'   - Public only for same-project regression access. Option Private Module
'     prevents exposure to external VBA projects.
'   - Production code must never call it.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' SET SEAM
'------------------------------------------------------------------------------
    'Record the armed state for the next growth attempt to consume
        m_InjectFailureListGrowthFailure = FailNextGrowth

End Sub


Public Sub UI_RuntimeBeginQuietUpdate( _
    ByRef OldScreenUpdating As Boolean, _
    ByRef QuietModeChanged As Boolean)
'
'==============================================================================
' UI_RuntimeBeginQuietUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Enters a best-effort Application.ScreenUpdating suppression scope.
'
' WHY THIS EXISTS
'   Applying several UI elements in sequence causes visible flicker. Suppressing
'   redraw for the duration of the pass removes most of it. Suppression must be
'   recorded rather than assumed, because a caller may already have disabled
'   ScreenUpdating and would not expect this module to re-enable it.
'
' INPUTS
'   OldScreenUpdating
'     ByRef. Receives the ScreenUpdating value observed on entry.
'
'   QuietModeChanged
'     ByRef. Receives True only when this scope actually changed the setting.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Captures the current ScreenUpdating value.
'   - Disables redraw only when it was enabled.
'   - Records whether the change was made, so the matching End can be exact.
'
' ERROR POLICY
'   - Suppresses errors locally.
'   - Leaves QuietModeChanged False when the host refuses the change.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - M_EXCEL_UI
'   - M_EXCEL_UI_SNAPSHOT
'
' NOTES
'   ScreenUpdating suppression cannot fully eliminate Ribbon or non-client
'   frame repaint; those surfaces are redrawn by the host.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Never let a host refusal escape this scope
        On Error Resume Next

'------------------------------------------------------------------------------
' ENTER QUIET SCOPE
'------------------------------------------------------------------------------
    'Capture the value observed on entry
        OldScreenUpdating = Application.ScreenUpdating

    'Assume no change until one is actually made
        QuietModeChanged = False

    'Suppress redraw only when it was enabled, and record that we did so
        If OldScreenUpdating Then
            Application.ScreenUpdating = False
            QuietModeChanged = True
        End If

End Sub


Public Sub UI_RuntimeEndQuietUpdate( _
    ByVal OldScreenUpdating As Boolean, _
    ByVal QuietModeChanged As Boolean)
'
'==============================================================================
' UI_RuntimeEndQuietUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Restores Application.ScreenUpdating when this module changed it.
'
' WHY THIS EXISTS
'   Restoring unconditionally would re-enable redraw for a caller that had
'   deliberately disabled it before calling in. Only the change this module
'   made is undone.
'
' INPUTS
'   OldScreenUpdating
'     Value captured by UI_RuntimeBeginQuietUpdate.
'
'   QuietModeChanged
'     True only when this module actually suppressed redraw.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Restores the captured value only when this module changed it.
'   - Is safe to call on paths where Begin was never reached, because
'     QuietModeChanged is then False.
'
' ERROR POLICY
'   - Suppresses errors locally.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - M_EXCEL_UI
'   - M_EXCEL_UI_SNAPSHOT
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Never let a host refusal escape this scope
        On Error Resume Next

'------------------------------------------------------------------------------
' EXIT QUIET SCOPE
'------------------------------------------------------------------------------
    'Undo only the change this module made
        If QuietModeChanged Then
            Application.ScreenUpdating = OldScreenUpdating
        End If

End Sub


Public Function UI_RuntimeTrySetRibbonVisibleIfNeeded( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_RuntimeTrySetRibbonVisibleIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Sets Ribbon visibility only when the current state differs from the request.
'
' WHY THIS EXISTS
'   Showing or hiding the Ribbon forces a full repaint of the command surface.
'   Skipping the write when the Ribbon already holds the requested state removes
'   that flicker entirely for the common no-op case.
'
' INPUTS
'   IsVisible
'     Requested Ribbon visibility.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => already correct, or successfully updated.
'     False => the write was attempted and failed.
'
' BEHAVIOR
'   - Attempts a read and short-circuits when the state already matches.
'   - Falls through to the write when the state differs OR cannot be read.
'   - Clears any read diagnostic before writing, so a successful write never
'     leaves a stale message behind.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_RuntimeTryGetRibbonVisible
'   - UI_RuntimeTrySetRibbonVisible
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - M_EXCEL_UI
'   - M_EXCEL_UI_SNAPSHOT
'
' NOTES
'   A failed read is not itself a failure of this routine. It only disables the
'   no-op short circuit, and the write is still attempted.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentVisible      As Boolean         'Ribbon state observed on entry

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the operation completes
        UI_RuntimeTrySetRibbonVisibleIfNeeded = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT
'------------------------------------------------------------------------------
    'Skip the write when the Ribbon can be read and already matches
        If UI_RuntimeTryGetRibbonVisible(CurrentVisible, FailMsg) Then
            If CurrentVisible = IsVisible Then
                UI_RuntimeTrySetRibbonVisibleIfNeeded = True
                GoTo Safe_Exit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Discard any read diagnostic; the write result is what matters now
        FailMsg = vbNullString

    'Apply the requested Ribbon state
        UI_RuntimeTrySetRibbonVisibleIfNeeded = _
            UI_RuntimeTrySetRibbonVisible(IsVisible, FailMsg)

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
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_RuntimeBuildErrorText

End Function


Public Function UI_RuntimeTrySetBooleanPropertyIfNeeded( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByVal NewValue As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_RuntimeTrySetBooleanPropertyIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Sets a Boolean object-model property only when its value differs from the
'   requested value.
'
' WHY THIS EXISTS
'   This is the single write path for every managed Application-level and
'   Window-level property. Late binding through CallByName keeps one
'   implementation for all of them instead of one branch per property name, and
'   the read-before-write step gives every managed element the same no-op
'   suppression behavior.
'
' INPUTS
'   Target
'     Application or Window object carrying the property.
'
'   PropertyName
'     Name of the Boolean property, for example "DisplayHeadings".
'
'   NewValue
'     Requested value.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => already correct, or successfully updated.
'     False => the write was attempted and failed.
'
' BEHAVIOR
'   - Attempts a read and short-circuits when the value already matches.
'   - Falls through to the write when the value differs OR cannot be read.
'   - Clears any read diagnostic before writing.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_RuntimeTryGetBooleanProperty
'   - UI_RuntimeTrySetBooleanProperty
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - M_EXCEL_UI
'   - M_EXCEL_UI_SNAPSHOT
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentValue        As Boolean         'Property value observed on entry

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the operation completes
        UI_RuntimeTrySetBooleanPropertyIfNeeded = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT
'------------------------------------------------------------------------------
    'Skip the write when the property can be read and already matches
        If UI_RuntimeTryGetBooleanProperty( _
            Target, PropertyName, CurrentValue, FailMsg) Then

            If CurrentValue = NewValue Then
                UI_RuntimeTrySetBooleanPropertyIfNeeded = True
                GoTo Safe_Exit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Discard any read diagnostic; the write result is what matters now
        FailMsg = vbNullString

    'Apply the requested property value
        UI_RuntimeTrySetBooleanPropertyIfNeeded = _
            UI_RuntimeTrySetBooleanProperty(Target, PropertyName, NewValue, FailMsg)

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
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_RuntimeBuildErrorText

End Function


Public Function UI_RuntimeTryGetRibbonVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_RuntimeTryGetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the current Ribbon visibility, preferring the CommandBars object model
'   and falling back to a legacy Excel 4 macro query.
'
' WHY THIS EXISTS
'   Excel exposes no first-class Ribbon visibility property. CommandBars("Ribbon")
'   works in most desktop builds; where it does not, the legacy Get.ToolBar query
'   usually still answers. Trying both gives no-op suppression the best chance of
'   working without ever failing the surrounding operation.
'
' INPUTS
'   IsVisible
'     ByRef. Receives the observed Ribbon visibility.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the state was read by one of the two mechanisms.
'     False => neither mechanism answered; IsVisible is not meaningful.
'
' BEHAVIOR
'   - Tries the CommandBars read first.
'   - Falls back to Get.ToolBar(7, "Ribbon") when that read raises.
'   - Reports the fallback error text when both mechanisms fail.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - UI_RuntimeTrySetRibbonVisibleIfNeeded
'   - M_EXCEL_UI_SNAPSHOT
'
' NOTES
'   Ribbon reads are explicitly best effort. A False return is not an error in
'   the surrounding operation; it only means no-op suppression is unavailable.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant         'Raw Excel 4 macro result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until one mechanism answers
        UI_RuntimeTryGetRibbonVisible = False

    'Initialize the output and the failure message buffer
        IsVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' TRY COMMANDBARS
'------------------------------------------------------------------------------
    'Probe the object model without letting a missing CommandBar raise
        On Error Resume Next
            IsVisible = Application.CommandBars("Ribbon").Visible

    'Accept the object-model answer when the probe did not raise
        If Err.Number = 0 Then
            On Error GoTo Err_Handler
            UI_RuntimeTryGetRibbonVisible = True
            GoTo Safe_Exit
        End If

    'Discard the probe error and continue to the fallback
        Err.Clear
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' TRY EXCEL 4 FALLBACK
'------------------------------------------------------------------------------
    'Probe the legacy toolbar query without letting a disabled host raise
        On Error Resume Next
            V = Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")

    'Accept the legacy answer when the probe did not raise
        If Err.Number = 0 Then
            On Error GoTo Err_Handler
            IsVisible = CBool(V)
            UI_RuntimeTryGetRibbonVisible = True
            GoTo Safe_Exit
        End If

    'Both mechanisms failed; report the fallback reason
        FailMsg = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        On Error GoTo Err_Handler

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
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_RuntimeBuildErrorText

End Function


Public Function UI_RuntimeTryGetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByRef ValueOut As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_RuntimeTryGetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Reads a Boolean object-model property through CallByName.
'
' WHY THIS EXISTS
'   One late-bound read serves every managed Application-level and Window-level
'   property. It is also the liveness probe used by the snapshot engine to test
'   whether a retained Window object is still usable, because a dead COM wrapper
'   raises here rather than returning a value.
'
' INPUTS
'   Target
'     Application or Window object carrying the property.
'
'   PropertyName
'     Name of the Boolean property.
'
'   ValueOut
'     ByRef. Receives the observed value. False when the read fails.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the read succeeded and ValueOut is meaningful.
'     False => the read failed and ValueOut is not meaningful.
'
' BEHAVIOR
'   - Rejects a Nothing target and an empty property name explicitly, so those
'     two cases produce a readable reason rather than a raw COM error.
'   - Reads through CallByName and coerces the result to Boolean.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - UI_RuntimeTrySetBooleanPropertyIfNeeded
'   - M_EXCEL_UI_SNAPSHOT
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant         'Raw CallByName result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the read completes
        UI_RuntimeTryGetBooleanProperty = False

    'Initialize the output and the failure message buffer
        ValueOut = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the target object
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo Safe_Exit
        End If

    'Validate the property name
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' READ PROPERTY
'------------------------------------------------------------------------------
    'Read the property through late binding
        V = CallByName(Target, PropertyName, VbGet)

    'Coerce and publish the observed value
        ValueOut = CBool(V)
        UI_RuntimeTryGetBooleanProperty = True

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
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_RuntimeBuildErrorText

End Function


Private Function UI_RuntimeTrySetRibbonVisible( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_RuntimeTrySetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Shows or hides the Ribbon using legacy Excel 4 macro execution.
'
' WHY THIS EXISTS
'   Excel exposes no supported object-model call that hides the Ribbon outright.
'   Show.TOOLBAR remains the only mechanism available from VBA without a COM
'   add-in, which is why Ribbon control is documented as best effort.
'
' INPUTS
'   IsVisible
'     Requested Ribbon visibility.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the macro executed.
'     False => the host refused or the macro raised.
'
' BEHAVIOR
'   - Builds the Show.TOOLBAR command text for the requested state.
'   - Executes it through Application.ExecuteExcel4Macro.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - UI_RuntimeTrySetRibbonVisibleIfNeeded
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MacroText           As String          'Excel 4 command to execute

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the macro executes
        UI_RuntimeTrySetRibbonVisible = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Build the command text for the requested state
        If IsVisible Then
            MacroText = "Show.TOOLBAR(""Ribbon"",True)"
        Else
            MacroText = "Show.TOOLBAR(""Ribbon"",False)"
        End If

    'Execute the legacy command
        Application.ExecuteExcel4Macro MacroText

    'Report success
        UI_RuntimeTrySetRibbonVisible = True

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
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_RuntimeBuildErrorText

End Function


Private Function UI_RuntimeTrySetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByVal NewValue As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_RuntimeTrySetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a Boolean object-model property through CallByName.
'
' WHY THIS EXISTS
'   One late-bound write serves every managed Application-level and
'   Window-level property, so a new managed element needs no new write branch.
'
' INPUTS
'   Target
'     Application or Window object carrying the property.
'
'   PropertyName
'     Name of the Boolean property.
'
'   NewValue
'     Value to write.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the write succeeded.
'     False => the host refused the write.
'
' BEHAVIOR
'   - Rejects a Nothing target and an empty property name explicitly.
'   - Writes through CallByName.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - UI_RuntimeTrySetBooleanPropertyIfNeeded
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the write completes
        UI_RuntimeTrySetBooleanProperty = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the target object
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo Safe_Exit
        End If

    'Validate the property name
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Write the property through late binding
        CallByName Target, PropertyName, VbLet, NewValue

    'Report success
        UI_RuntimeTrySetBooleanProperty = True

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
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_RuntimeBuildErrorText

End Function


Public Function UI_RuntimeBuildErrorText() _
    As String
'
'==============================================================================
' UI_RuntimeBuildErrorText
'------------------------------------------------------------------------------
' PURPOSE
'   Builds a consistent diagnostic string from the active Err object.
'
' WHY THIS EXISTS
'   Every error handler in the package reports through the same Detail format,
'   so a failure list is readable without knowing which module produced it.
'
' RETURNS
'   String
'     Best-effort error number, description, source and Erl text.
'
' BEHAVIOR
'   - Always emits number and description.
'   - Appends Source and Line only when they carry information.
'
' ERROR POLICY
'   - Suppresses formatting errors locally.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - This module
'   - M_EXCEL_UI
'   - M_EXCEL_UI_SNAPSHOT
'
' NOTES
'   Erl returns zero unless the procedure carries line numbers, so the Line
'   fragment is omitted for the unnumbered source used throughout this project.
'
' UPDATED
'   2026-08-18
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
' BUILD DIAGNOSTIC TEXT
'------------------------------------------------------------------------------
    'Never let diagnostic formatting raise inside an error handler
        On Error Resume Next

    'Compose number, description, and the optional source and line fragments
        UI_RuntimeBuildErrorText = _
            CStr(ErrNumber) & ": " & ErrDescription & _
            IIf(Len(ErrSource) > 0, _
                " | Source: " & ErrSource, _
                vbNullString) & _
            IIf(ErrLine <> 0, _
                " | Line: " & CStr(ErrLine), _
                vbNullString)

End Function


Public Function UI_RuntimeBuildWindowLabel( _
    ByVal TargetWindow As Object) _
    As String
'
'==============================================================================
' UI_RuntimeBuildWindowLabel
'------------------------------------------------------------------------------
' PURPOSE
'   Builds a stable best-effort diagnostic label for one Excel Window.
'
' WHY THIS EXISTS
'   Failure entries must name the window they refer to, and both the apply path
'   and the snapshot path need that label. Building it here means the two agree
'   on the format, and means neither has to read Window.Caption while composing
'   a failure message: that read can itself raise on a window that is closing
'   or otherwise unusable, turning one property failure into an unexpected
'   error for the whole pass.
'
' INPUTS
'   TargetWindow
'     Window to describe. May be Nothing or unusable.
'
' RETURNS
'   String
'     A diagnostic label. Never used for identity matching.
'
' BEHAVIOR
'   - Prefers "Workbook :: Caption".
'   - Falls back to whichever of the two is readable.
'   - Falls back to a generic label when Excel exposes neither.
'
' ERROR POLICY
'   - Does not raise. A label is always produced.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - M_EXCEL_UI
'   - M_EXCEL_UI_SNAPSHOT
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WorkbookName        As String          'Parent workbook name, if readable
    Dim WindowCaption       As String          'Window caption, if readable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A label must always be produced, so no read may raise
        On Error Resume Next

'------------------------------------------------------------------------------
' READ IDENTIFYING FIELDS
'------------------------------------------------------------------------------
    'Read both descriptive fields on a best-effort basis
        If Not TargetWindow Is Nothing Then
            WorkbookName = TargetWindow.Parent.Name
            WindowCaption = TargetWindow.Caption
        End If

'------------------------------------------------------------------------------
' BUILD LABEL
'------------------------------------------------------------------------------
    'Prefer the fullest description Excel was able to supply
        If Len(WorkbookName) > 0 And Len(WindowCaption) > 0 Then
            UI_RuntimeBuildWindowLabel = WorkbookName & " :: " & WindowCaption
        ElseIf Len(WindowCaption) > 0 Then
            UI_RuntimeBuildWindowLabel = WindowCaption
        ElseIf Len(WorkbookName) > 0 Then
            UI_RuntimeBuildWindowLabel = WorkbookName
        Else
            UI_RuntimeBuildWindowLabel = "Excel window"
        End If

End Function


Public Sub UI_RuntimeLogFailure( _
    ByVal ProcName As String, _
    ByVal Stage As String, _
    ByVal Detail As String)
'
'==============================================================================
' UI_RuntimeLogFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one consistent diagnostic line to the Immediate Window.
'
' WHY THIS EXISTS
'   The fire-and-forget public procedures return nothing, so the Immediate
'   Window is their only diagnostic channel. Keeping the line format in one
'   place makes those lines greppable across the whole package.
'
' INPUTS
'   ProcName
'     Public caller name used as the line prefix.
'
'   Stage
'     Managed element or phase that failed.
'
'   Detail
'     Host error text or explicit validation reason.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Emits "ProcName failed @ Stage | Detail".
'
' ERROR POLICY
'   - Suppresses logging errors locally.
'   - Raises no user-interface message.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - UI_RuntimeHandleFailure
'   - M_EXCEL_UI
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' WRITE DIAGNOSTIC LINE
'------------------------------------------------------------------------------
    'Never let logging raise inside an error handler
        On Error Resume Next

    'Emit the standard diagnostic line
        Debug.Print ProcName & " failed @ " & Stage & " | " & Detail

End Sub
