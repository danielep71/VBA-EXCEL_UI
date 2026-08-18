Attribute VB_Name = "M_EXCEL_UI_SNAPSHOT"
Option Explicit
Option Private Module

'==============================================================================
' M_EXCEL_UI_SNAPSHOT
'------------------------------------------------------------------------------
' PURPOSE
'   Owns the in-memory Excel UI snapshot: its state, the retained Window object
'   identities, and the best-effort capture and restoration passes.
'
' WHY THIS EXISTS
'   Snapshot state has a lifecycle of its own, independent of the public facade.
'   It is created by a capture, consumed by a restore, invalidated by a project
'   reset, and it holds live COM references for as long as it exists.
'
'   Isolating that subsystem keeps M_EXCEL_UI focused on API compatibility and
'   general UI application logic, and keeps all the mutable state that survives
'   between calls in one module where its lifetime can be reasoned about.
'
' INTERNAL SURFACE
'   - UI_SnapshotCaptureCore
'   - UI_SnapshotRestoreCore
'   - UI_SnapshotHasState
'   - UI_SnapshotClear
'
' IDENTITY MODEL
'   Window state is keyed by the captured Window OBJECT, never by its position
'   in Application.Windows and never by caption, workbook name or hWnd.
'
'   Re-enumerating Application.Windows can return a different COM wrapper for
'   the same live window, so comparing a fresh wrapper with the retained one
'   using Is can reject a window that is perfectly alive. The retained
'   reference is therefore treated as authoritative and used directly, subject
'   to a liveness probe.
'
'   The consequences are deliberate and are what "identity-safe" means here:
'
'     reordered windows   restore correctly, because position is never used
'     new windows         are left unchanged, because nothing was captured
'     closed windows      are reported and skipped, never redirected to a
'                         replacement window that happens to sit at the same
'                         index
'
' DESIGN PRINCIPLES
'   - Capture replaces any prior snapshot outright; snapshots do not merge.
'   - Capture and restore are both best effort and continue after an
'     element-level failure.
'   - A partial capture is still a usable snapshot; per-element Known flags
'     record which values are meaningful.
'   - Restore never writes a value that was not successfully captured.
'   - Restore retains the snapshot afterwards, so it can be replayed.
'
' STATE LIFETIME
'   The snapshot lives in module memory only. It is lost when Excel closes, when
'   the host workbook or add-in unloads, when the VBA project is reset, and when
'   code editing resets project state. It is not durable recovery across
'   sessions.
'
'   While it exists it retains live Window references. Callers that open and
'   close many workbooks should clear the snapshot when it is no longer needed.
'
' ERROR POLICY
'   - Internal entry points are fail-soft and do not raise.
'   - Capture and restoration continue after element-level failure.
'   - An unexpected capture failure clears the partial snapshot, so a caller
'     never restores from a half-built baseline.
'   - An unexpected restoration failure leaves the snapshot in place.
'   - No user-interface message is displayed.
'
' DEPENDENCIES
'   - M_EXCEL_UI_RUNTIME  shared host operations, result buffers, diagnostics
'   - M_EXCEL_UI_TITLEBAR title-bar capture and restoration
'
' NOTES
'   - The parallel label array is diagnostic only and never participates in
'     identity matching.
'   - Restoration order is title bar, Ribbon, application-level properties,
'     then window-level properties, which is the reverse of the order in which
'     they visibly settle.
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

'==============================================================================
' PRIVATE MODULE STATE
'==============================================================================

'True once a capture pass has completed, whether or not it was complete.
Private m_HasExcelUIStateSnapshot       As Boolean

'Application-level captured values. The Ribbon carries a Known flag because it
'cannot always be read; the other three are read directly.
Private m_SnapshotRibbonKnown           As Boolean
Private m_SnapshotRibbonVisible         As Boolean
Private m_SnapshotStatusBarVisible      As Boolean
Private m_SnapshotScrollBarsVisible     As Boolean
Private m_SnapshotFormulaBarVisible     As Boolean

'Number of windows captured, and the upper bound of every parallel array below.
Private m_SnapshotWindowCount           As Long

'Object references provide identity-safe in-memory matching. Parallel labels
'are diagnostic only and never participate in matching.
Private m_SnapshotWindows()             As Object
Private m_SnapshotWindowLabels()        As String

'Per-window managed properties, each with a Known flag so a failed read is
'never replayed as a False value on restore.
Private m_SnapshotHeadingsKnown()       As Boolean
Private m_SnapshotHeadingsVisible()     As Boolean

Private m_SnapshotWorkbookTabsKnown()   As Boolean
Private m_SnapshotWorkbookTabsVisible() As Boolean

Private m_SnapshotGridlinesKnown()      As Boolean
Private m_SnapshotGridlinesVisible()    As Boolean

'Main-window frame state.
Private m_SnapshotTitleBarKnown         As Boolean
Private m_SnapshotTitleBarVisible       As Boolean


Public Function UI_SnapshotCaptureCore( _
    ByVal ProcName As String, _
    ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) _
    As Boolean
'
'==============================================================================
' UI_SnapshotCaptureCore
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the shared snapshot-capture pass for both the fire-and-forget
'   compatibility wrapper and the structured-result API.
'
' WHY THIS EXISTS
'   The two public capture entry points differ only in how they report. Sharing
'   one worker keeps capture ordering, Known-flag policy and the diagnostic
'   contract identical between them.
'
' INPUTS
'   ProcName
'     Public caller name used for diagnostics.
'
'   LogFailures
'     True to emit Immediate Window diagnostics; False for result-only use.
'
'   FailureCount
'     ByRef structured-result count buffer.
'
'   FailureList
'     ByRef structured-result list buffer.
'
'   CaptureFailureList
'     True when FailureList should be populated.
'
' RETURNS
'   Boolean
'     True  => no failure was recorded.
'     False => at least one element could not be captured.
'
' BEHAVIOR
'   - Clears any prior snapshot before capturing.
'   - Captures application-level state.
'   - Captures Ribbon and title-bar state on a best-effort basis, recording a
'     Known flag for each.
'   - Captures each window's retained object identity, its diagnostic label and
'     its managed properties.
'   - Records failures in deterministic capture order.
'   - Marks the snapshot available once the pass completes, even when optional
'     elements were unreadable.
'
' ERROR POLICY
'   - Does not raise.
'   - Optional-element failures are recorded and capture continues.
'   - An unexpected failure clears the partial snapshot and records one entry,
'     so no caller can restore from a half-built baseline.
'
' DEPENDENCIES
'   - UI_SnapshotClear
'   - UI_SnapshotBuildWindowIdentityText
'   - UI_RuntimeClearResultBuffer
'   - UI_RuntimeHandleFailure
'   - UI_RuntimeTryGetRibbonVisible
'   - UI_RuntimeTryGetBooleanProperty
'   - UI_RuntimeBuildErrorText
'   - UI_TryGetTitleBarVisible
'
' CALLED FROM
'   - M_EXCEL_UI
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long            'Parallel-array write position
    Dim W                   As Window          'Window being captured
    Dim Msg                 As String          'Per-element failure reason
    Dim UnexpectedDetail    As String          'Unexpected-error diagnostic
    Dim Succeeded           As Boolean         'Pass-level success flag

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Clear the caller's result buffers deterministically
        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList

    'Assume success until an element fails
        Succeeded = True

    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Replace any prior snapshot; snapshots never merge
        UI_SnapshotClear

'------------------------------------------------------------------------------
' CAPTURE: APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Capture the three directly readable application-level properties
        m_SnapshotStatusBarVisible = Application.DisplayStatusBar
        m_SnapshotScrollBarsVisible = Application.DisplayScrollBars
        m_SnapshotFormulaBarVisible = Application.DisplayFormulaBar

'------------------------------------------------------------------------------
' CAPTURE: RIBBON AND TITLE BAR
'------------------------------------------------------------------------------
    'Capture Ribbon visibility, recording whether it could be read at all
        m_SnapshotRibbonKnown = UI_RuntimeTryGetRibbonVisible( _
            IsVisible:=m_SnapshotRibbonVisible, _
            FailMsg:=Msg)

        If Not m_SnapshotRibbonKnown Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Ribbon", Msg
        End If

    'Capture title-bar visibility, recording whether it could be read at all
        m_SnapshotTitleBarKnown = UI_TryGetTitleBarVisible( _
            IsVisible:=m_SnapshotTitleBarVisible, _
            FailMsg:=Msg)

        If Not m_SnapshotTitleBarKnown Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "TitleBar", Msg
        End If

'------------------------------------------------------------------------------
' CAPTURE: WINDOW IDENTITIES AND STATE
'------------------------------------------------------------------------------
    'Size every parallel array to the current window count
        m_SnapshotWindowCount = Application.Windows.Count

        If m_SnapshotWindowCount > 0 Then

            ReDim m_SnapshotWindows(1 To m_SnapshotWindowCount)
            ReDim m_SnapshotWindowLabels(1 To m_SnapshotWindowCount)

            ReDim m_SnapshotHeadingsKnown(1 To m_SnapshotWindowCount)
            ReDim m_SnapshotHeadingsVisible(1 To m_SnapshotWindowCount)

            ReDim m_SnapshotWorkbookTabsKnown(1 To m_SnapshotWindowCount)
            ReDim m_SnapshotWorkbookTabsVisible(1 To m_SnapshotWindowCount)

            ReDim m_SnapshotGridlinesKnown(1 To m_SnapshotWindowCount)
            ReDim m_SnapshotGridlinesVisible(1 To m_SnapshotWindowCount)

            'Start before the first parallel-array slot
                i = 0

            'Capture each window's identity and managed properties in order
                For Each W In Application.Windows

                    i = i + 1

                    'Retain the exact object; this is the restore key
                        Set m_SnapshotWindows(i) = W

                    'Build the diagnostic label once, for use by both passes
                        m_SnapshotWindowLabels(i) = _
                            UI_SnapshotBuildWindowIdentityText(W)

                    'Capture Headings
                        m_SnapshotHeadingsKnown(i) = _
                            UI_RuntimeTryGetBooleanProperty( _
                                Target:=W, _
                                PropertyName:="DisplayHeadings", _
                                ValueOut:=m_SnapshotHeadingsVisible(i), _
                                FailMsg:=Msg)

                        If Not m_SnapshotHeadingsKnown(i) Then
                            UI_RuntimeHandleFailure _
                                ProcName, LogFailures, Succeeded, FailureCount, _
                                FailureList, CaptureFailureList, _
                                "Headings [" & m_SnapshotWindowLabels(i) & "]", Msg
                        End If

                    'Capture Workbook Tabs
                        m_SnapshotWorkbookTabsKnown(i) = _
                            UI_RuntimeTryGetBooleanProperty( _
                                Target:=W, _
                                PropertyName:="DisplayWorkbookTabs", _
                                ValueOut:=m_SnapshotWorkbookTabsVisible(i), _
                                FailMsg:=Msg)

                        If Not m_SnapshotWorkbookTabsKnown(i) Then
                            UI_RuntimeHandleFailure _
                                ProcName, LogFailures, Succeeded, FailureCount, _
                                FailureList, CaptureFailureList, _
                                "WorkbookTabs [" & m_SnapshotWindowLabels(i) & "]", Msg
                        End If

                    'Capture Gridlines
                        m_SnapshotGridlinesKnown(i) = _
                            UI_RuntimeTryGetBooleanProperty( _
                                Target:=W, _
                                PropertyName:="DisplayGridlines", _
                                ValueOut:=m_SnapshotGridlinesVisible(i), _
                                FailMsg:=Msg)

                        If Not m_SnapshotGridlinesKnown(i) Then
                            UI_RuntimeHandleFailure _
                                ProcName, LogFailures, Succeeded, FailureCount, _
                                FailureList, CaptureFailureList, _
                                "Gridlines [" & m_SnapshotWindowLabels(i) & "]", Msg
                        End If

                Next W

        End If

'------------------------------------------------------------------------------
' MARK SNAPSHOT AVAILABLE
'------------------------------------------------------------------------------
    'A partial capture is still a usable snapshot; Known flags carry the detail
        m_HasExcelUIStateSnapshot = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Publish the pass-level result and exit before the error handler
        UI_SnapshotCaptureCore = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Capture the diagnostic before the clear disturbs anything
        UnexpectedDetail = UI_RuntimeBuildErrorText

    'Discard the partial snapshot so it can never be restored from
        UI_SnapshotClear

    'Record the unexpected failure
        UI_RuntimeHandleFailure _
            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, "Unexpected", UnexpectedDetail

        Resume Safe_Exit

End Function


Public Function UI_SnapshotHasState() _
    As Boolean
'
'==============================================================================
' UI_SnapshotHasState
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether an explicit in-memory Excel UI snapshot is available.
'
' RETURNS
'   Boolean
'     True  => a snapshot exists and can be restored from.
'     False => no snapshot has been captured, or it has been cleared.
'
' BEHAVIOR
'   - Reports availability only. It does not report completeness: a snapshot
'     built from a partially readable UI still reports True.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - M_EXCEL_UI
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN AVAILABILITY
'------------------------------------------------------------------------------
    'Report the snapshot availability flag
        UI_SnapshotHasState = m_HasExcelUIStateSnapshot

End Function


Public Function UI_SnapshotRestoreCore( _
    ByVal ProcName As String, _
    ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) _
    As Boolean
'
'==============================================================================
' UI_SnapshotRestoreCore
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the shared snapshot-restoration pass for both the fire-and-forget
'   compatibility wrapper and the structured-result API.
'
' WHY THIS EXISTS
'   The two public restore entry points differ only in how they report. Sharing
'   one worker keeps restoration ordering, identity resolution and the
'   diagnostic contract identical between them.
'
' INPUTS
'   ProcName
'     Public caller name used for diagnostics.
'
'   LogFailures
'     True to emit Immediate Window diagnostics; False for result-only use.
'
'   FailureCount
'     ByRef structured-result count buffer.
'
'   FailureList
'     ByRef structured-result list buffer.
'
'   CaptureFailureList
'     True when FailureList should be populated.
'
' RETURNS
'   Boolean
'     True  => no failure was recorded.
'     False => at least one element could not be restored, or no snapshot
'              existed.
'
' BEHAVIOR
'   - Reports a NoSnapshot failure and exits when nothing was captured.
'   - Restores title bar and Ribbon only when their captured states were
'     readable.
'   - Restores application-level object-model properties.
'   - Resolves each captured Window by retained object identity.
'   - Restores state only to matching, still-usable windows.
'   - Leaves newly opened windows unchanged, because nothing was captured for
'     them.
'   - Records failures in deterministic restoration order.
'   - Preserves ScreenUpdating through a quiet-update scope.
'   - Retains the snapshot afterwards so it can be replayed.
'
' ERROR POLICY
'   - Does not raise.
'   - Continues after element-level or window-level failure.
'
' DEPENDENCIES
'   - UI_SnapshotTryResolveWindow
'   - UI_RuntimeClearResultBuffer
'   - UI_RuntimeHandleFailure
'   - UI_RuntimeBeginQuietUpdate
'   - UI_RuntimeEndQuietUpdate
'   - UI_RuntimeTrySetRibbonVisibleIfNeeded
'   - UI_RuntimeTrySetBooleanPropertyIfNeeded
'   - UI_RuntimeBuildErrorText
'   - UI_TrySetTitleBarVisibleIfNeeded
'
' CALLED FROM
'   - M_EXCEL_UI
'
' NOTES
'   The quiet-update scope is entered only after the NoSnapshot check, and the
'   matching End is safe on that path because QuietModeChanged is still False.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long            'Snapshot slot being restored
    Dim MatchedWindow       As Object          'Resolved captured Window
    Dim Msg                 As String          'Per-element failure reason
    Dim UnexpectedDetail    As String          'Unexpected-error diagnostic
    Dim OldScreenUpdating   As Boolean         'ScreenUpdating value on entry
    Dim QuietModeChanged    As Boolean         'True when this pass suppressed it
    Dim Succeeded           As Boolean         'Pass-level success flag

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Clear the caller's result buffers deterministically
        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList

    'Assume success until an element fails
        Succeeded = True

    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' VALIDATE SNAPSHOT AVAILABILITY
'------------------------------------------------------------------------------
    'There is nothing to restore without a captured baseline
        If Not m_HasExcelUIStateSnapshot Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "NoSnapshot", _
                "no captured Excel UI snapshot is available"

            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' ENTER QUIET SCOPE
'------------------------------------------------------------------------------
    'Suppress redraw for the duration of the restoration pass
        UI_RuntimeBeginQuietUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

'------------------------------------------------------------------------------
' RESTORE: TITLE BAR
'------------------------------------------------------------------------------
    'Restore the frame only when its captured state was readable
        If m_SnapshotTitleBarKnown Then
            If Not UI_TrySetTitleBarVisibleIfNeeded( _
                IsVisible:=m_SnapshotTitleBarVisible, _
                FailMsg:=Msg) Then

                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "TitleBar", Msg
            End If
        End If

'------------------------------------------------------------------------------
' RESTORE: RIBBON
'------------------------------------------------------------------------------
    'Restore the Ribbon only when its captured state was readable
        If m_SnapshotRibbonKnown Then
            If Not UI_RuntimeTrySetRibbonVisibleIfNeeded( _
                IsVisible:=m_SnapshotRibbonVisible, _
                FailMsg:=Msg) Then

                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "Ribbon", Msg
            End If
        End If

'------------------------------------------------------------------------------
' RESTORE: APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Restore the status bar
        If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
            Target:=Application, _
            PropertyName:="DisplayStatusBar", _
            NewValue:=m_SnapshotStatusBarVisible, _
            FailMsg:=Msg) Then

            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "StatusBar", Msg
        End If

    'Restore the scroll bars
        If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
            Target:=Application, _
            PropertyName:="DisplayScrollBars", _
            NewValue:=m_SnapshotScrollBarsVisible, _
            FailMsg:=Msg) Then

            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "ScrollBars", Msg
        End If

    'Restore the formula bar
        If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
            Target:=Application, _
            PropertyName:="DisplayFormulaBar", _
            NewValue:=m_SnapshotFormulaBarVisible, _
            FailMsg:=Msg) Then

            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "FormulaBar", Msg
        End If

'------------------------------------------------------------------------------
' RESTORE: WINDOW-LEVEL STATE BY OBJECT IDENTITY
'------------------------------------------------------------------------------
    'Walk the captured slots; never the live Application.Windows ordering
        For i = 1 To m_SnapshotWindowCount

            'Release any window resolved on the previous iteration
                Set MatchedWindow = Nothing

            'Resolve the retained object, or report why it is unusable
                If UI_SnapshotTryResolveWindow( _
                    SnapshotIndex:=i, _
                    WindowOut:=MatchedWindow, _
                    FailMsg:=Msg) Then

                    'Restore Headings when the captured value is meaningful
                        If m_SnapshotHeadingsKnown(i) Then
                            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                                Target:=MatchedWindow, _
                                PropertyName:="DisplayHeadings", _
                                NewValue:=m_SnapshotHeadingsVisible(i), _
                                FailMsg:=Msg) Then

                                UI_RuntimeHandleFailure _
                                    ProcName, LogFailures, Succeeded, FailureCount, _
                                    FailureList, CaptureFailureList, _
                                    "Headings [" & m_SnapshotWindowLabels(i) & "]", Msg
                            End If
                        End If

                    'Restore Workbook Tabs when the captured value is meaningful
                        If m_SnapshotWorkbookTabsKnown(i) Then
                            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                                Target:=MatchedWindow, _
                                PropertyName:="DisplayWorkbookTabs", _
                                NewValue:=m_SnapshotWorkbookTabsVisible(i), _
                                FailMsg:=Msg) Then

                                UI_RuntimeHandleFailure _
                                    ProcName, LogFailures, Succeeded, FailureCount, _
                                    FailureList, CaptureFailureList, _
                                    "WorkbookTabs [" & m_SnapshotWindowLabels(i) & "]", Msg
                            End If
                        End If

                    'Restore Gridlines when the captured value is meaningful
                        If m_SnapshotGridlinesKnown(i) Then
                            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                                Target:=MatchedWindow, _
                                PropertyName:="DisplayGridlines", _
                                NewValue:=m_SnapshotGridlinesVisible(i), _
                                FailMsg:=Msg) Then

                                UI_RuntimeHandleFailure _
                                    ProcName, LogFailures, Succeeded, FailureCount, _
                                    FailureList, CaptureFailureList, _
                                    "Gridlines [" & m_SnapshotWindowLabels(i) & "]", Msg
                            End If
                        End If

                Else

                    'The captured window is gone; report it rather than
                    'redirecting its state to a replacement window
                        UI_RuntimeHandleFailure _
                            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                            CaptureFailureList, _
                            "WindowIdentity [" & m_SnapshotWindowLabels(i) & "]", Msg

                End If

        Next i

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Leave the quiet scope exactly as it was entered
        UI_RuntimeEndQuietUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

    'Publish the pass-level result and exit before the error handler
        UI_SnapshotRestoreCore = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Record the unexpected failure; the snapshot is deliberately retained
        UnexpectedDetail = UI_RuntimeBuildErrorText

        UI_RuntimeHandleFailure _
            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, "Unexpected", UnexpectedDetail

        Resume Safe_Exit

End Function


Public Sub UI_SnapshotClear()
'
'==============================================================================
' UI_SnapshotClear
'------------------------------------------------------------------------------
' PURPOSE
'   Removes all captured state and releases every retained Window reference.
'
' WHY THIS EXISTS
'   The snapshot holds live COM references. Clearing is therefore not only a
'   logical reset but the point at which those references are released, which
'   matters when many workbooks are opened and closed in one session.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Resets the availability and Known flags.
'   - Resets every captured value.
'   - Erases the parallel arrays, releasing the retained Window objects.
'
' ERROR POLICY
'   - Does not raise during normal operation.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - UI_SnapshotCaptureCore
'   - M_EXCEL_UI
'
' NOTES
'   Erase on a dynamic array that was never sized is legal, so this procedure
'   is safe to call before any capture has occurred.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET FLAGS
'------------------------------------------------------------------------------
    'Mark the snapshot unavailable
        m_HasExcelUIStateSnapshot = False

    'Reset the best-effort Known flags
        m_SnapshotRibbonKnown = False
        m_SnapshotTitleBarKnown = False

'------------------------------------------------------------------------------
' RESET VALUES
'------------------------------------------------------------------------------
    'Reset the application-level and frame values
        m_SnapshotRibbonVisible = False
        m_SnapshotStatusBarVisible = False
        m_SnapshotScrollBarsVisible = False
        m_SnapshotFormulaBarVisible = False
        m_SnapshotTitleBarVisible = False

    'Reset the captured window count
        m_SnapshotWindowCount = 0

'------------------------------------------------------------------------------
' RELEASE WINDOW REFERENCES AND ARRAYS
'------------------------------------------------------------------------------
    'Release the retained Window objects and their diagnostic labels
        Erase m_SnapshotWindows
        Erase m_SnapshotWindowLabels

    'Release the per-window Headings state
        Erase m_SnapshotHeadingsKnown
        Erase m_SnapshotHeadingsVisible

    'Release the per-window Workbook Tabs state
        Erase m_SnapshotWorkbookTabsKnown
        Erase m_SnapshotWorkbookTabsVisible

    'Release the per-window Gridlines state
        Erase m_SnapshotGridlinesKnown
        Erase m_SnapshotGridlinesVisible

End Sub


Private Function UI_SnapshotTryResolveWindow( _
    ByVal SnapshotIndex As Long, _
    ByRef WindowOut As Object, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_SnapshotTryResolveWindow
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one captured Excel Window by validating and returning the retained
'   Window object reference held in the snapshot.
'
' WHY THIS EXISTS
'   Re-enumerating Application.Windows may return a different COM wrapper for
'   the same live Excel window. Comparing that wrapper with the retained one
'   through the Is operator can therefore reject a valid surviving window.
'
'   The retained object reference is the authoritative identity and is used
'   directly, provided it remains usable. That is what makes restoration safe
'   against reordering, and what stops captured state from ever landing on a
'   replacement window that happens to occupy the same collection index.
'
' INPUTS
'   SnapshotIndex
'     1-based index into the internal snapshot arrays.
'
'   WindowOut
'     ByRef. Receives the retained captured Window object, or Nothing.
'
'   FailMsg
'     ByRef diagnostic reason on failure.
'
' RETURNS
'   Boolean
'     True  => the retained captured Window reference remains usable.
'     False => the captured Window was closed or its reference is unavailable.
'
' BEHAVIOR
'   - Validates the snapshot index against the captured range.
'   - Retrieves the exact Window object retained during capture.
'   - Performs a non-mutating property read to confirm the object still works.
'   - Never searches by collection index, caption, workbook name or hWnd.
'   - Never redirects captured state to a newly created replacement window.
'
' ERROR POLICY
'   - Does not raise to callers.
'   - Returns False and populates FailMsg.
'
' DEPENDENCIES
'   - UI_RuntimeTryGetBooleanProperty
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - UI_SnapshotRestoreCore
'
' NOTES
'   The liveness probe reads DisplayHeadings because it is a managed property
'   that exists on every Window and has no side effect. A dead COM wrapper
'   raises on that read, which is exactly the signal required.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CapturedWindow      As Object          'Exact object retained at capture
    Dim ProbeValue          As Boolean         'Non-mutating probe output
    Dim ProbeMsg            As String          'Diagnostic returned by the probe

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the retained reference is proven usable
        UI_SnapshotTryResolveWindow = False

    'Initialize the output and the failure message buffer
        Set WindowOut = Nothing
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE SNAPSHOT INDEX
'------------------------------------------------------------------------------
    'Reject an index outside the captured range
        If SnapshotIndex < 1 Or SnapshotIndex > m_SnapshotWindowCount Then
            FailMsg = "snapshot index is outside the captured window range"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' RETRIEVE CAPTURED WINDOW REFERENCE
'------------------------------------------------------------------------------
    'Take the exact object retained at capture time
        Set CapturedWindow = m_SnapshotWindows(SnapshotIndex)

        If CapturedWindow Is Nothing Then
            FailMsg = "captured window reference is unavailable"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' VALIDATE RETAINED WINDOW REFERENCE
'------------------------------------------------------------------------------
    'Confirm the retained object still works using a non-mutating read of an
    'existing managed property
        If Not UI_RuntimeTryGetBooleanProperty( _
            Target:=CapturedWindow, _
            PropertyName:="DisplayHeadings", _
            ValueOut:=ProbeValue, _
            FailMsg:=ProbeMsg) Then

            FailMsg = _
                "captured window is no longer open or usable; no state was applied"

            'Append the underlying probe reason when one is available
                If Len(ProbeMsg) > 0 Then
                    FailMsg = FailMsg & " | " & ProbeMsg
                End If

            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' RETURN EXACT CAPTURED WINDOW
'------------------------------------------------------------------------------
    'Publish the retained object; never a freshly enumerated wrapper
        Set WindowOut = CapturedWindow
        UI_SnapshotTryResolveWindow = True

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
    'Report the unexpected runtime error and publish no window
        FailMsg = UI_RuntimeBuildErrorText
        Set WindowOut = Nothing

        Resume Safe_Exit

End Function


Private Function UI_SnapshotBuildWindowIdentityText( _
    ByVal TargetWindow As Object) _
    As String
'
'==============================================================================
' UI_SnapshotBuildWindowIdentityText
'------------------------------------------------------------------------------
' PURPOSE
'   Builds a stable best-effort diagnostic label for a captured Excel Window.
'
' WHY THIS EXISTS
'   Failure messages need to name the window they refer to. The label is built
'   once at capture time so that the restore pass can report on a window that
'   has since been closed, when reading its caption would raise.
'
' INPUTS
'   TargetWindow
'     Window whose identifying text should be described.
'
' RETURNS
'   String
'     A diagnostic label. This text is never used for identity matching.
'
' BEHAVIOR
'   - Prefers "Workbook :: Caption".
'   - Falls back to whichever of the two is available.
'   - Falls back to a generic label when Excel exposes neither.
'
' ERROR POLICY
'   - Does not raise.
'   - Suppresses errors locally so an unreadable window still yields a label.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - UI_SnapshotCaptureCore
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
        WorkbookName = TargetWindow.Parent.Name
        WindowCaption = TargetWindow.Caption

'------------------------------------------------------------------------------
' BUILD LABEL
'------------------------------------------------------------------------------
    'Prefer the fullest description Excel was able to supply
        If Len(WorkbookName) > 0 And Len(WindowCaption) > 0 Then
            UI_SnapshotBuildWindowIdentityText = _
                WorkbookName & " :: " & WindowCaption
        ElseIf Len(WindowCaption) > 0 Then
            UI_SnapshotBuildWindowIdentityText = WindowCaption
        ElseIf Len(WorkbookName) > 0 Then
            UI_SnapshotBuildWindowIdentityText = WorkbookName
        Else
            UI_SnapshotBuildWindowIdentityText = "captured Excel window"
        End If

End Function
