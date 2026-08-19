Attribute VB_Name = "M_EXCEL_UI"
Option Explicit
Option Private Module

'==============================================================================
' M_EXCEL_UI
'------------------------------------------------------------------------------
' PURPOSE
'   Centralizes visibility control for the Excel UI elements managed by this
'   package, combining Excel object-model elements with delegated WinAPI
'   title-bar control for the main window represented by Application.Hwnd.
'
' WHY THIS EXISTS
'   Workbook-driven solutions often need a constrained or application-style
'   Excel shell. Written directly, that means UI writes scattered across every
'   module that happens to need one, with no single owner, no consistent
'   failure behavior, and no reliable way back to the original interface.
'
'   This module provides one explicit, defensive, fail-soft API instead: a
'   single owner for UI state changes, tri-state intent rather than ambiguous
'   optional Booleans, diagnostics available as data, and a deterministic
'   show-everything recovery path that is separate from snapshot restoration.
'
' PUBLIC SURFACE
'   Enums:
'     - UIVisibility
'     - UIWindowTargetScope
'
'   Selective application:
'     - UI_SetExcelUI
'     - UI_SetExcelUI_WithResult
'
'   Convenience wrappers:
'     - UI_HideExcelUI
'     - UI_ShowExcelUI
'
'   Snapshot lifecycle:
'     - UI_CaptureExcelUIState
'     - UI_CaptureExcelUIState_WithResult
'     - UI_ResetExcelUIToSnapshot
'     - UI_ResetExcelUIToSnapshot_WithResult
'     - UI_HasExcelUIStateSnapshot
'     - UI_ClearExcelUIStateSnapshot
'
' MANAGED UI SURFACE
'   Application-level, affecting the current Excel instance:
'     Ribbon, Status Bar, Scroll Bars, Formula Bar
'
'   Window-level, affecting the requested target scope:
'     Headings, Workbook Tabs, Gridlines
'
'   Main-window frame:
'     Title Bar
'
' DESIGN PRINCIPLES
'   1. Explicit caller intent. Tri-state values rather than ambiguous optional
'      Booleans, because an omitted Boolean defaults to False and would read as
'      an instruction to hide.
'   2. One owner for UI state changes.
'   3. Continue where safe. One failed element does not prevent unrelated
'      requested changes from being attempted.
'   4. Expose diagnostics when required. Simple callers may log; orchestration
'      code can request structured results.
'   5. Separate show-all from restore. A deterministic visible shell and a
'      captured custom baseline are different operations.
'   6. Preserve host state where possible. Avoid unnecessary writes and restore
'      ScreenUpdating.
'   7. Document platform-sensitive behavior. Ribbon and title-bar paths remain
'      explicitly best effort.
'
' ERROR POLICY
'   - Public entry points are fail-soft and never raise to callers.
'   - Fire-and-forget procedures log failures to the Immediate Window.
'   - UI_SetExcelUI_WithResult and the snapshot WithResult APIs return
'     structured failure information as data.
'   - One failed element does not prevent later requested elements from being
'     attempted.
'
' DEPENDENCIES
'   - M_EXCEL_UI_RUNTIME   shared fail-soft host operations and diagnostics
'   - M_EXCEL_UI_TITLEBAR  WinAPI title-bar control
'   - M_EXCEL_UI_SNAPSHOT  snapshot state and lifecycle
'
' PLATFORM / COMPATIBILITY
'   - Windows only, because title-bar control depends on WinAPI.
'   - Supports 32-bit and 64-bit Office through conditional compilation in
'     M_EXCEL_UI_TITLEBAR.
'
' NOTES
'   - Snapshot state is stored in memory only and is lost after a project reset
'     or when Excel closes.
'   - Window-level snapshot state is keyed by the captured Window object
'     identity, not by Application.Windows collection index, so reordered
'     windows restore correctly, newly opened windows are left unchanged, and
'     closed or recreated windows are skipped rather than having another
'     window's state applied to them.
'   - Title-bar ownership is limited to the caption, system-menu, sizing-frame,
'     minimize-box and maximize-box style bits. Showing merges only those owned
'     bits into the current style, preserving unrelated changes made by Excel or
'     another component.
'   - Selective apply APIs accept an optional trailing TargetScope that affects
'     only Headings, Workbook Tabs and Gridlines.
'   - Hidden Excel UI is not a security boundary.
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
' PUBLIC ENUMS
'==============================================================================

'Tri-state visibility. An omitted argument is equivalent to UI_LeaveUnchanged,
'which is why the neutral member is negative rather than zero: a caller cannot
'accidentally request a hide by omitting an argument.
Public Enum UIVisibility
    UI_LeaveUnchanged = -1                'Do not touch this UI element
    UI_Hide = 0                           'Hide this UI element
    UI_Show = 1                           'Show this UI element
End Enum

'Window-level targeting. The default is zero so that pre-v1.1.0 callers, which
'pass no scope at all, keep the original all-windows behavior.
Public Enum UIWindowTargetScope
    UI_TargetAllExcelWindows = 0          'Apply to every current Excel window
    UI_TargetActiveWindow = 1             'Apply only to Application.ActiveWindow
    UI_TargetActiveWorkbookWindows = 2    'Apply to ActiveWorkbook.Windows
End Enum


Public Sub UI_SetExcelUI( _
    Optional ByVal Ribbon As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal StatusBar As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal ScrollBars As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal FormulaBar As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal Headings As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal WorkbookTabs As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal Gridlines As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal TitleBar As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal TargetScope As UIWindowTargetScope = _
        UI_TargetAllExcelWindows)
'
'==============================================================================
' UI_SetExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the requested visibility state to the managed Excel UI elements.
'
' WHY THIS EXISTS
'   This is the primary selective entry point and the one most workbook code
'   should call. It accepts best-effort completion and reports through the
'   Immediate Window, which is sufficient when the caller has no decision to
'   make about a partial result.
'
' INPUTS
'   Ribbon, StatusBar, ScrollBars, FormulaBar,
'   Headings, WorkbookTabs, Gridlines, TitleBar
'     UI_Show, UI_Hide or UI_LeaveUnchanged. Omitted arguments are equivalent
'     to UI_LeaveUnchanged.
'
'   TargetScope
'     Controls only Headings, Workbook Tabs and Gridlines. The default remains
'     UI_TargetAllExcelWindows for backward compatibility.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Applies application-level settings to the current Excel instance.
'   - Applies window-level settings to the requested target scope.
'   - Applies title-bar visibility to Application.Hwnd.
'   - Continues after an element-level failure.
'
' ERROR POLICY
'   - Does not raise to callers.
'   - Logs failures to the Immediate Window.
'
' DEPENDENCIES
'   - UI_ApplyExcelUIState
'   - UI_RuntimeLogFailure
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - Workbook and add-in code
'   - UI_HideExcelUI
'   - UI_ShowExcelUI
'
' NOTES
'   Use UI_SetExcelUI_WithResult instead when initialization, orchestration or
'   test code must distinguish complete success from partial application.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IgnoredFailureCount As Long            'Discarded result buffer
    Dim IgnoredFailureList  As Variant         'Discarded result buffer

    Const PROC              As String = "UI_SetExcelUI"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Delegate to the shared worker with logging on and result capture off
        UI_ApplyExcelUIState _
            ProcName:=PROC, _
            Ribbon:=Ribbon, _
            StatusBar:=StatusBar, _
            ScrollBars:=ScrollBars, _
            FormulaBar:=FormulaBar, _
            Headings:=Headings, _
            WorkbookTabs:=WorkbookTabs, _
            Gridlines:=Gridlines, _
            TitleBar:=TitleBar, _
            TargetScope:=TargetScope, _
            LogFailures:=True, _
            FailureCount:=IgnoredFailureCount, _
            FailureList:=IgnoredFailureList, _
            CaptureFailureList:=False

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
    'Log the unexpected wrapper failure without raising to the caller
        UI_RuntimeLogFailure PROC, "Unexpected", UI_RuntimeBuildErrorText
        Resume Safe_Exit

End Sub


Public Sub UI_HideExcelUI()
'
'==============================================================================
' UI_HideExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Hides every Excel UI element managed by this package.
'
' WHY THIS EXISTS
'   Applying a fully constrained shell is common enough to deserve a named
'   operation rather than eight repeated arguments at every call site.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Delegates to UI_SetExcelUI with UI_Hide for every managed element.
'   - Uses the default all-windows target scope.
'
' ERROR POLICY
'   - Does not raise to callers.
'   - Logs an unexpected wrapper failure to the Immediate Window.
'
' DEPENDENCIES
'   - UI_SetExcelUI
'   - UI_RuntimeLogFailure
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - Workbook and add-in code
'
' NOTES
'   Keep an accessible UI_ShowExcelUI recovery path available whenever this is
'   used in a constrained-shell solution.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC              As String = "UI_HideExcelUI"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Request the hidden state for every managed element
        UI_SetExcelUI _
            Ribbon:=UI_Hide, _
            StatusBar:=UI_Hide, _
            ScrollBars:=UI_Hide, _
            FormulaBar:=UI_Hide, _
            Headings:=UI_Hide, _
            WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Hide, _
            TitleBar:=UI_Hide

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
    'Log the unexpected wrapper failure without raising to the caller
        UI_RuntimeLogFailure PROC, "Unexpected", UI_RuntimeBuildErrorText
        Resume Safe_Exit

End Sub


Public Sub UI_ShowExcelUI()
'
'==============================================================================
' UI_ShowExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Shows every Excel UI element managed by this package.
'
' WHY THIS EXISTS
'   This is the emergency recovery path. It requires no snapshot, which is what
'   makes it usable when a workflow was interrupted, when the snapshot has been
'   cleared, or when VBA project state was reset and the prior baseline is no
'   longer available.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Delegates to UI_SetExcelUI with UI_Show for every managed element.
'   - Means "show all", not "restore captured state".
'
' ERROR POLICY
'   - Does not raise to callers.
'   - Logs an unexpected wrapper failure to the Immediate Window.
'
' DEPENDENCIES
'   - UI_SetExcelUI
'   - UI_RuntimeLogFailure
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - Workbook and add-in code
'   - Manual recovery from the VBA editor or Quick Access Toolbar
'
' NOTES
'   For development work, keep a simple macro that calls this procedure within
'   reach of the VBA editor.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC              As String = "UI_ShowExcelUI"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Request the visible state for every managed element
        UI_SetExcelUI _
            Ribbon:=UI_Show, _
            StatusBar:=UI_Show, _
            ScrollBars:=UI_Show, _
            FormulaBar:=UI_Show, _
            Headings:=UI_Show, _
            WorkbookTabs:=UI_Show, _
            Gridlines:=UI_Show, _
            TitleBar:=UI_Show

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
    'Log the unexpected wrapper failure without raising to the caller
        UI_RuntimeLogFailure PROC, "Unexpected", UI_RuntimeBuildErrorText
        Resume Safe_Exit

End Sub


Public Function UI_SetExcelUI_WithResult( _
    Optional ByVal Ribbon As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal StatusBar As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal ScrollBars As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal FormulaBar As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal Headings As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal WorkbookTabs As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal Gridlines As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal TitleBar As UIVisibility = UI_LeaveUnchanged, _
    Optional ByRef FailureCount As Long = 0, _
    Optional ByRef FailureList As Variant, _
    Optional ByVal TargetScope As UIWindowTargetScope = _
        UI_TargetAllExcelWindows) _
    As Boolean
'
'==============================================================================
' UI_SetExcelUI_WithResult
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the requested managed UI state and returns structured diagnostics.
'
' WHY THIS EXISTS
'   Initialization, orchestration and automated tests need to distinguish
'   complete success from partial application, and to inspect which elements
'   failed rather than reading the Immediate Window. This entry point returns
'   that information as data.
'
' INPUTS
'   Ribbon, StatusBar, ScrollBars, FormulaBar,
'   Headings, WorkbookTabs, Gridlines, TitleBar
'     UI_Show, UI_Hide or UI_LeaveUnchanged.
'
'   FailureCount
'     Optional ByRef. Receives the number of recorded failures.
'
'   FailureList
'     Optional ByRef. Receives a 1-based String array of ordered
'     "Stage | Detail" entries. Populated only when the argument is supplied.
'
'   TargetScope
'     Controls only Headings, Workbook Tabs and Gridlines. The default remains
'     UI_TargetAllExcelWindows.
'
' RETURNS
'   Boolean
'     True  => no failure was recorded.
'     False => one or more failures were recorded.
'
' BEHAVIOR
'   - Mirrors UI_SetExcelUI while suppressing Immediate Window logging.
'   - Clears the output buffers deterministically on entry.
'   - Accumulates into an internal buffer and publishes it once, so a caller's
'     variable is never left holding a partially built array.
'
' ERROR POLICY
'   - Does not raise for ordinary failures.
'   - Captures unexpected failures in the result rather than propagating them.
'
' DEPENDENCIES
'   - UI_ApplyExcelUIState
'   - UI_RuntimeClearResultBuffer
'   - UI_RuntimeHandleFailure
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - Workbook and add-in initialization and orchestration code
'   - M_EXCEL_UI_REGRESSION_TESTS
'
' NOTES
'   TargetScope is deliberately declared after FailureCount and FailureList so
'   that existing positional callers written against v1.0 are unaffected.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean         'Pass-level success flag
    Dim CaptureFailureList  As Boolean         'True when a list was requested
    Dim InternalFailureList As Variant         'Buffer published once on exit

    Const PROC              As String = "UI_SetExcelUI_WithResult"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A failure list is built only when the caller supplied the argument
        CaptureFailureList = Not IsMissing(FailureList)

    'Clear the result buffers deterministically
        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList

    'Assume success until an element fails
        Succeeded = True

    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Delegate to the shared worker with logging off and result capture on
        Succeeded = UI_ApplyExcelUIState( _
            ProcName:=PROC, _
            Ribbon:=Ribbon, _
            StatusBar:=StatusBar, _
            ScrollBars:=ScrollBars, _
            FormulaBar:=FormulaBar, _
            Headings:=Headings, _
            WorkbookTabs:=WorkbookTabs, _
            Gridlines:=Gridlines, _
            TitleBar:=TitleBar, _
            TargetScope:=TargetScope, _
            LogFailures:=False, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Publish the accumulated list only when one was requested
        If CaptureFailureList Then
            FailureList = InternalFailureList
        End If

    'Publish the pass-level result and exit before the error handler
        UI_SetExcelUI_WithResult = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Record the unexpected failure in the result instead of raising
        UI_RuntimeHandleFailure _
            ProcName:=PROC, _
            LogFailures:=False, _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", _
            Detail:=UI_RuntimeBuildErrorText

        Resume Safe_Exit

End Function


Public Sub UI_CaptureExcelUIState()
'
'==============================================================================
' UI_CaptureExcelUIState
'------------------------------------------------------------------------------
' PURPOSE
'   Captures the current managed Excel UI state for later restoration.
'
' WHY THIS EXISTS
'   A workflow that constrains the shell should be able to hand the user back
'   the interface they actually had, rather than a generic show-everything
'   state. This is the fire-and-forget form of that capture.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Delegates to the shared capture worker.
'   - Replaces any prior snapshot.
'   - Logs ordered best-effort failures to the Immediate Window.
'
' ERROR POLICY
'   - Does not raise to callers.
'   - Logs failures and continues where capture remains meaningful.
'   - Leaves the snapshot unavailable after an unexpected capture failure.
'
' DEPENDENCIES
'   - UI_SnapshotCaptureCore
'   - UI_RuntimeLogFailure
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - Workbook and add-in code
'
' NOTES
'   This form returns nothing, so a caller that must confirm the snapshot was
'   taken should use UI_CaptureExcelUIState_WithResult instead.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IgnoredFailureCount As Long            'Discarded result buffer
    Dim IgnoredFailureList  As Variant         'Discarded result buffer

    Const PROC              As String = "UI_CaptureExcelUIState"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
    'Delegate to the shared worker with logging on and result capture off
        UI_SnapshotCaptureCore _
            ProcName:=PROC, _
            LogFailures:=True, _
            FailureCount:=IgnoredFailureCount, _
            FailureList:=IgnoredFailureList, _
            CaptureFailureList:=False

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
    'Log the unexpected wrapper failure without raising to the caller
        UI_RuntimeLogFailure PROC, "Unexpected", UI_RuntimeBuildErrorText
        Resume Safe_Exit

End Sub


Public Function UI_CaptureExcelUIState_WithResult( _
    Optional ByRef FailureCount As Long = 0, _
    Optional ByRef FailureList As Variant) _
    As Boolean
'
'==============================================================================
' UI_CaptureExcelUIState_WithResult
'------------------------------------------------------------------------------
' PURPOSE
'   Captures the current managed Excel UI state and returns structured
'   diagnostics.
'
' WHY THIS EXISTS
'   Code that is about to constrain the shell needs to know whether the
'   baseline it will later restore was actually recorded, and which elements
'   were unreadable. The fire-and-forget form cannot report either.
'
' INPUTS
'   FailureCount
'     Optional ByRef. Receives the number of recorded capture failures.
'
'   FailureList
'     Optional ByRef. Receives a 1-based String array of ordered
'     "Stage | Detail" entries.
'
' RETURNS
'   Boolean
'     True  => the capture pass recorded no failure.
'     False => at least one element could not be captured.
'
' BEHAVIOR
'   - Clears the output buffers deterministically on entry.
'   - Replaces any prior snapshot.
'   - Preserves best-effort partial-capture semantics.
'   - Marks the snapshot available after the capture pass completes, even when
'     optional elements were unreadable.
'
' ERROR POLICY
'   - Does not raise for ordinary capture failures.
'   - Returns ordered element and window-specific diagnostics.
'   - Leaves the snapshot unavailable after an unexpected capture failure.
'
' DEPENDENCIES
'   - UI_SnapshotCaptureCore
'   - UI_RuntimeClearResultBuffer
'   - UI_RuntimeHandleFailure
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - Workbook and add-in initialization and orchestration code
'   - M_EXCEL_UI_REGRESSION_TESTS
'
' NOTES
'   A False return does not mean no snapshot exists. A partial capture is still
'   available for restoration; the failure list names what was missed.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean         'Pass-level success flag
    Dim CaptureFailureList  As Boolean         'True when a list was requested
    Dim InternalFailureList As Variant         'Buffer published once on exit

    Const PROC              As String = "UI_CaptureExcelUIState_WithResult"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A failure list is built only when the caller supplied the argument
        CaptureFailureList = Not IsMissing(FailureList)

    'Clear the result buffers deterministically
        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList

    'Assume success until an element fails
        Succeeded = True

    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
    'Delegate to the shared worker with logging off and result capture on
        Succeeded = UI_SnapshotCaptureCore( _
            ProcName:=PROC, _
            LogFailures:=False, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Publish the accumulated list only when one was requested
        If CaptureFailureList Then
            FailureList = InternalFailureList
        End If

    'Publish the pass-level result and exit before the error handler
        UI_CaptureExcelUIState_WithResult = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Record the unexpected failure in the result instead of raising
        UI_RuntimeHandleFailure _
            ProcName:=PROC, _
            LogFailures:=False, _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", _
            Detail:=UI_RuntimeBuildErrorText

        Resume Safe_Exit

End Function


Public Function UI_HasExcelUIStateSnapshot() _
    As Boolean
'
'==============================================================================
' UI_HasExcelUIStateSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether an explicit in-memory Excel UI snapshot is available.
'
' WHY THIS EXISTS
'   Callers that offer a "restore" action need to know whether it would do
'   anything, so the action can be disabled rather than failing when pressed.
'
' RETURNS
'   Boolean
'     True  => a snapshot exists and can be restored from.
'     False => no snapshot has been captured, or it has been cleared.
'
' BEHAVIOR
'   - Reports availability only, not completeness.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - UI_SnapshotHasState
'
' CALLED FROM
'   - Workbook and add-in code
'   - M_EXCEL_UI_DEMO
'   - M_EXCEL_UI_REGRESSION_TESTS
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN AVAILABILITY
'------------------------------------------------------------------------------
    'Delegate to the snapshot engine, which owns the flag
        UI_HasExcelUIStateSnapshot = UI_SnapshotHasState

End Function


Public Sub UI_ResetExcelUIToSnapshot()
'
'==============================================================================
' UI_ResetExcelUIToSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Restores the managed Excel UI to the most recently captured snapshot.
'
' WHY THIS EXISTS
'   This returns the user to the interface they actually had, which is a
'   different and usually better outcome than the deterministic show-everything
'   state produced by UI_ShowExcelUI.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Delegates to the shared restoration worker.
'   - Logs ordered best-effort failures to the Immediate Window.
'   - Retains the snapshot afterwards.
'
' ERROR POLICY
'   - Does not raise to callers.
'   - Logs restore failures and continues where possible.
'
' DEPENDENCIES
'   - UI_SnapshotRestoreCore
'   - UI_RuntimeLogFailure
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - Workbook and add-in code
'
' NOTES
'   Requires a snapshot. Prefer UI_ShowExcelUI for emergency recovery, because
'   it works when no snapshot is available.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IgnoredFailureCount As Long            'Discarded result buffer
    Dim IgnoredFailureList  As Variant         'Discarded result buffer

    Const PROC              As String = "UI_ResetExcelUIToSnapshot"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Delegate to the shared worker with logging on and result capture off
        UI_SnapshotRestoreCore _
            ProcName:=PROC, _
            LogFailures:=True, _
            FailureCount:=IgnoredFailureCount, _
            FailureList:=IgnoredFailureList, _
            CaptureFailureList:=False

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
    'Log the unexpected wrapper failure without raising to the caller
        UI_RuntimeLogFailure PROC, "Unexpected", UI_RuntimeBuildErrorText
        Resume Safe_Exit

End Sub


Public Function UI_ResetExcelUIToSnapshot_WithResult( _
    Optional ByRef FailureCount As Long = 0, _
    Optional ByRef FailureList As Variant) _
    As Boolean
'
'==============================================================================
' UI_ResetExcelUIToSnapshot_WithResult
'------------------------------------------------------------------------------
' PURPOSE
'   Restores the managed Excel UI to the current snapshot and returns structured
'   diagnostics.
'
' WHY THIS EXISTS
'   Restoration is the operation most likely to be partially satisfiable, since
'   captured windows may have been closed in the meantime. Callers that must
'   report or persist what could not be restored need that as data.
'
' INPUTS
'   FailureCount
'     Optional ByRef. Receives the number of recorded restoration failures.
'
'   FailureList
'     Optional ByRef. Receives a 1-based String array of ordered
'     "Stage | Detail" entries.
'
' RETURNS
'   Boolean
'     True  => restoration recorded no failure.
'     False => at least one element could not be restored, or no snapshot
'              existed.
'
' BEHAVIOR
'   - Clears the output buffers deterministically on entry.
'   - Restores every available captured element on a best-effort basis.
'   - Leaves newly opened windows unchanged.
'   - Reports closed, recreated or unusable captured windows without applying
'     their state to a replacement window.
'   - Retains the snapshot after the restoration attempt.
'
' ERROR POLICY
'   - Does not raise for ordinary restoration failures.
'   - Returns ordered element and window-specific diagnostics.
'
' DEPENDENCIES
'   - UI_SnapshotRestoreCore
'   - UI_RuntimeClearResultBuffer
'   - UI_RuntimeHandleFailure
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - Workbook and add-in orchestration code
'   - M_EXCEL_UI_REGRESSION_TESTS
'
' NOTES
'   Calling this without a snapshot is not an unexpected error; it returns
'   False with a single NoSnapshot entry.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean         'Pass-level success flag
    Dim CaptureFailureList  As Boolean         'True when a list was requested
    Dim InternalFailureList As Variant         'Buffer published once on exit

    Const PROC              As String = "UI_ResetExcelUIToSnapshot_WithResult"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A failure list is built only when the caller supplied the argument
        CaptureFailureList = Not IsMissing(FailureList)

    'Clear the result buffers deterministically
        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList

    'Assume success until an element fails
        Succeeded = True

    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Delegate to the shared worker with logging off and result capture on
        Succeeded = UI_SnapshotRestoreCore( _
            ProcName:=PROC, _
            LogFailures:=False, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Publish the accumulated list only when one was requested
        If CaptureFailureList Then
            FailureList = InternalFailureList
        End If

    'Publish the pass-level result and exit before the error handler
        UI_ResetExcelUIToSnapshot_WithResult = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Record the unexpected failure in the result instead of raising
        UI_RuntimeHandleFailure _
            ProcName:=PROC, _
            LogFailures:=False, _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", _
            Detail:=UI_RuntimeBuildErrorText

        Resume Safe_Exit

End Function


Public Sub UI_ClearExcelUIStateSnapshot()
'
'==============================================================================
' UI_ClearExcelUIStateSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the current in-memory Excel UI snapshot.
'
' WHY THIS EXISTS
'   The snapshot retains live Window references for as long as it exists.
'   Clearing it when the captured baseline is no longer needed releases those
'   references rather than holding them until the project resets.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Releases retained Window object references and clears all captured values.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - UI_SnapshotClear
'
' CALLED FROM
'   - Workbook and add-in code
'   - M_EXCEL_UI_REGRESSION_TESTS
'
' NOTES
'   Worth calling from Workbook_BeforeClose in solutions that capture a
'   baseline, so no Window reference outlives the workbook that produced it.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Clearing must never raise, whatever state the snapshot is in
        On Error Resume Next

    'Delegate to the snapshot engine, which owns the state
        UI_SnapshotClear

End Sub


Private Function UI_ApplyExcelUIState( _
    ByVal ProcName As String, _
    ByVal Ribbon As UIVisibility, _
    ByVal StatusBar As UIVisibility, _
    ByVal ScrollBars As UIVisibility, _
    ByVal FormulaBar As UIVisibility, _
    ByVal Headings As UIVisibility, _
    ByVal WorkbookTabs As UIVisibility, _
    ByVal Gridlines As UIVisibility, _
    ByVal TitleBar As UIVisibility, _
    ByVal TargetScope As UIWindowTargetScope, _
    ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) _
    As Boolean
'
'==============================================================================
' UI_ApplyExcelUIState
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the requested UI state through the shared fire-and-forget and
'   structured-result worker.
'
' WHY THIS EXISTS
'   Both selective entry points delegate here so that validation stays
'   consistent, operation ordering stays stable, best-effort semantics do not
'   drift between them, and the logging and structured-result paths share one
'   implementation rather than two that must be kept in step.
'
' INPUTS
'   ProcName
'     Public caller name used for diagnostics.
'
'   Ribbon .. TitleBar
'     Requested UIVisibility values, validated individually.
'
'   TargetScope
'     Requested window targeting, validated before use.
'
'   LogFailures
'     True to emit Immediate Window diagnostics.
'
'   FailureCount / FailureList / CaptureFailureList
'     Standard structured-result buffers.
'
' RETURNS
'   Boolean
'     True  => no failure was recorded.
'     False => at least one element failed or was invalid.
'
' BEHAVIOR
'   - Validates every UIVisibility argument and the TargetScope.
'   - Records an invalid value as a failure and skips only that element, so one
'     bad argument does not discard the rest of the request.
'   - Skips UI_LeaveUnchanged values without reading or writing them.
'   - Avoids no-op writes wherever the current state can be read.
'   - Applies window-level state to all Excel windows, the active window, or
'     all windows belonging to the active workbook.
'   - TargetScope never changes the Ribbon, application-level properties, or
'     the Excel main-window title bar.
'   - Preserves ScreenUpdating through a quiet-update scope.
'
' ERROR POLICY
'   - Does not raise.
'   - Records and optionally logs failures in insertion order.
'   - Leaves the quiet-update scope exactly as it was entered, on every path.
'
' DEPENDENCIES
'   - UI_ApplyWindowLevelState
'   - UI_IsValidVisibility
'   - UI_IsValidTargetScope
'   - UI_VisibilityToBoolean
'   - M_EXCEL_UI_RUNTIME
'   - M_EXCEL_UI_TITLEBAR
'
' CALLED FROM
'   - UI_SetExcelUI
'   - UI_SetExcelUI_WithResult
'
' NOTES
'   Every requested value is validated up front rather than at the point of
'   use, so the failure list reports all invalid arguments in one pass instead
'   of stopping at the first.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean         'Pass-level success flag
    Dim W                   As Window          'Window being written
    Dim ActiveTargetWindow  As Window          'Resolved active window
    Dim ActiveTargetBook    As Workbook        'Resolved active workbook
    Dim ShowFlag            As Boolean         'Requested value as a Boolean
    Dim Msg                 As String          'Per-element failure reason

    Dim ValidRibbon         As Boolean         'Ribbon argument is a valid enum
    Dim ValidStatusBar      As Boolean         'StatusBar argument is valid
    Dim ValidScrollBars     As Boolean         'ScrollBars argument is valid
    Dim ValidFormulaBar     As Boolean         'FormulaBar argument is valid
    Dim ValidHeadings       As Boolean         'Headings argument is valid
    Dim ValidWorkbookTabs   As Boolean         'WorkbookTabs argument is valid
    Dim ValidGridlines      As Boolean         'Gridlines argument is valid
    Dim ValidTitleBar       As Boolean         'TitleBar argument is valid
    Dim ValidTargetScope    As Boolean         'TargetScope argument is valid

    Dim DoHeadings          As Boolean         'Headings must be written
    Dim DoWorkbookTabs      As Boolean         'Workbook Tabs must be written
    Dim DoGridlines         As Boolean         'Gridlines must be written
    Dim DoWindowState       As Boolean         'Any window-level work is needed

    Dim ShowHeadings        As Boolean         'Requested Headings value
    Dim ShowWorkbookTabs    As Boolean         'Requested Workbook Tabs value
    Dim ShowGridlines       As Boolean         'Requested Gridlines value

    Dim OldScreenUpdating   As Boolean         'ScreenUpdating value on entry
    Dim QuietModeChanged    As Boolean         'True when this pass suppressed it

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Clear the result buffers deterministically
        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList

    'Assume success until an element fails
        Succeeded = True

    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Suppress redraw for the duration of the apply pass
        UI_RuntimeBeginQuietUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the Ribbon request
        ValidRibbon = UI_IsValidVisibility(Ribbon)
        If Not ValidRibbon Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Ribbon", _
                "invalid UIVisibility value: " & CStr(Ribbon)
        End If

    'Validate the Status Bar request
        ValidStatusBar = UI_IsValidVisibility(StatusBar)
        If Not ValidStatusBar Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "StatusBar", _
                "invalid UIVisibility value: " & CStr(StatusBar)
        End If

    'Validate the Scroll Bars request
        ValidScrollBars = UI_IsValidVisibility(ScrollBars)
        If Not ValidScrollBars Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "ScrollBars", _
                "invalid UIVisibility value: " & CStr(ScrollBars)
        End If

    'Validate the Formula Bar request
        ValidFormulaBar = UI_IsValidVisibility(FormulaBar)
        If Not ValidFormulaBar Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "FormulaBar", _
                "invalid UIVisibility value: " & CStr(FormulaBar)
        End If

    'Validate the Headings request
        ValidHeadings = UI_IsValidVisibility(Headings)
        If Not ValidHeadings Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Headings", _
                "invalid UIVisibility value: " & CStr(Headings)
        End If

    'Validate the Workbook Tabs request
        ValidWorkbookTabs = UI_IsValidVisibility(WorkbookTabs)
        If Not ValidWorkbookTabs Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "WorkbookTabs", _
                "invalid UIVisibility value: " & CStr(WorkbookTabs)
        End If

    'Validate the Gridlines request
        ValidGridlines = UI_IsValidVisibility(Gridlines)
        If Not ValidGridlines Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Gridlines", _
                "invalid UIVisibility value: " & CStr(Gridlines)
        End If

    'Validate the Title Bar request
        ValidTitleBar = UI_IsValidVisibility(TitleBar)
        If Not ValidTitleBar Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "TitleBar", _
                "invalid UIVisibility value: " & CStr(TitleBar)
        End If

    'Validate the window targeting request
        ValidTargetScope = UI_IsValidTargetScope(TargetScope)
        If Not ValidTargetScope Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "TargetScope", _
                "invalid UIWindowTargetScope value: " & CStr(TargetScope)
        End If

'------------------------------------------------------------------------------
' APPLY APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Apply the Ribbon state
        If ValidRibbon And Ribbon <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(Ribbon)

            If Not UI_RuntimeTrySetRibbonVisibleIfNeeded(ShowFlag, Msg) Then
                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "Ribbon", Msg
            End If
        End If

    'Apply the Status Bar state
        If ValidStatusBar And StatusBar <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(StatusBar)

            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                Application, "DisplayStatusBar", ShowFlag, Msg) Then

                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "StatusBar", Msg
            End If
        End If

    'Apply the Scroll Bars state
        If ValidScrollBars And ScrollBars <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(ScrollBars)

            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                Application, "DisplayScrollBars", ShowFlag, Msg) Then

                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "ScrollBars", Msg
            End If
        End If

    'Apply the Formula Bar state
        If ValidFormulaBar And FormulaBar <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(FormulaBar)

            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                Application, "DisplayFormulaBar", ShowFlag, Msg) Then

                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "FormulaBar", Msg
            End If
        End If

'------------------------------------------------------------------------------
' PRECOMPUTE WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Decide once which window properties need writing, rather than re-testing
    'the same conditions inside every window iteration
        DoHeadings = ValidHeadings And (Headings <> UI_LeaveUnchanged)
        DoWorkbookTabs = _
            ValidWorkbookTabs And (WorkbookTabs <> UI_LeaveUnchanged)
        DoGridlines = ValidGridlines And (Gridlines <> UI_LeaveUnchanged)
        DoWindowState = DoHeadings Or DoWorkbookTabs Or DoGridlines

    'Convert only the values that will actually be written
        If DoHeadings Then
            ShowHeadings = UI_VisibilityToBoolean(Headings)
        End If

        If DoWorkbookTabs Then
            ShowWorkbookTabs = UI_VisibilityToBoolean(WorkbookTabs)
        End If

        If DoGridlines Then
            ShowGridlines = UI_VisibilityToBoolean(Gridlines)
        End If

'------------------------------------------------------------------------------
' APPLY WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Resolve the requested target scope and write each window in it
        If DoWindowState And ValidTargetScope Then
            Select Case TargetScope

                Case UI_TargetAllExcelWindows

                    'Every window in the current Excel instance
                        For Each W In Application.Windows
                            UI_ApplyWindowLevelState _
                                ProcName:=ProcName, _
                                TargetWindow:=W, _
                                DoHeadings:=DoHeadings, _
                                ShowHeadings:=ShowHeadings, _
                                DoWorkbookTabs:=DoWorkbookTabs, _
                                ShowWorkbookTabs:=ShowWorkbookTabs, _
                                DoGridlines:=DoGridlines, _
                                ShowGridlines:=ShowGridlines, _
                                LogFailures:=LogFailures, _
                                Succeeded:=Succeeded, _
                                FailureCount:=FailureCount, _
                                FailureList:=FailureList, _
                                CaptureFailureList:=CaptureFailureList
                        Next W

                Case UI_TargetActiveWindow

                    'The active window only, when one exists
                        Set ActiveTargetWindow = Application.ActiveWindow

                        If ActiveTargetWindow Is Nothing Then
                            UI_RuntimeHandleFailure _
                                ProcName, LogFailures, Succeeded, FailureCount, _
                                FailureList, CaptureFailureList, "TargetScope", _
                                "active Excel window is unavailable"
                        Else
                            UI_ApplyWindowLevelState _
                                ProcName:=ProcName, _
                                TargetWindow:=ActiveTargetWindow, _
                                DoHeadings:=DoHeadings, _
                                ShowHeadings:=ShowHeadings, _
                                DoWorkbookTabs:=DoWorkbookTabs, _
                                ShowWorkbookTabs:=ShowWorkbookTabs, _
                                DoGridlines:=DoGridlines, _
                                ShowGridlines:=ShowGridlines, _
                                LogFailures:=LogFailures, _
                                Succeeded:=Succeeded, _
                                FailureCount:=FailureCount, _
                                FailureList:=FailureList, _
                                CaptureFailureList:=CaptureFailureList
                        End If

                Case UI_TargetActiveWorkbookWindows

                    'Every window belonging to the active workbook
                        Set ActiveTargetBook = Application.ActiveWorkbook

                        If ActiveTargetBook Is Nothing Then
                            UI_RuntimeHandleFailure _
                                ProcName, LogFailures, Succeeded, FailureCount, _
                                FailureList, CaptureFailureList, "TargetScope", _
                                "active workbook is unavailable"
                        Else
                            For Each W In ActiveTargetBook.Windows
                                UI_ApplyWindowLevelState _
                                    ProcName:=ProcName, _
                                    TargetWindow:=W, _
                                    DoHeadings:=DoHeadings, _
                                    ShowHeadings:=ShowHeadings, _
                                    DoWorkbookTabs:=DoWorkbookTabs, _
                                    ShowWorkbookTabs:=ShowWorkbookTabs, _
                                    DoGridlines:=DoGridlines, _
                                    ShowGridlines:=ShowGridlines, _
                                    LogFailures:=LogFailures, _
                                    Succeeded:=Succeeded, _
                                    FailureCount:=FailureCount, _
                                    FailureList:=FailureList, _
                                    CaptureFailureList:=CaptureFailureList
                            Next W
                        End If

            End Select
        End If

'------------------------------------------------------------------------------
' APPLY TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Apply the frame state; TargetScope never affects the main window
        If ValidTitleBar And TitleBar <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(TitleBar)

            If Not UI_TrySetTitleBarVisibleIfNeeded(ShowFlag, Msg) Then
                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "TitleBar", Msg
            End If
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Leave the quiet scope exactly as it was entered
        UI_RuntimeEndQuietUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

    'Publish the pass-level result and exit before the error handler
        UI_ApplyExcelUIState = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Record the unexpected failure, then restore ScreenUpdating via Safe_Exit
        UI_RuntimeHandleFailure _
            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, "Unexpected", UI_RuntimeBuildErrorText

        Resume Safe_Exit

End Function


Private Sub UI_ApplyWindowLevelState( _
    ByVal ProcName As String, _
    ByVal TargetWindow As Window, _
    ByVal DoHeadings As Boolean, _
    ByVal ShowHeadings As Boolean, _
    ByVal DoWorkbookTabs As Boolean, _
    ByVal ShowWorkbookTabs As Boolean, _
    ByVal DoGridlines As Boolean, _
    ByVal ShowGridlines As Boolean, _
    ByVal LogFailures As Boolean, _
    ByRef Succeeded As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean)
'
'==============================================================================
' UI_ApplyWindowLevelState
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the requested managed Window properties to one resolved Excel
'   Window.
'
' WHY THIS EXISTS
'   All three target scopes need identical per-window behavior. Factoring the
'   property writes out of the scope resolution keeps the three Select Case
'   branches to their actual difference, which is which windows to visit.
'
'   It also gives each window its own error boundary. One window that becomes
'   unusable mid-pass is recorded and skipped; the windows after it are still
'   attempted.
'
' INPUTS
'   ProcName
'     Public caller name used for diagnostics.
'
'   TargetWindow
'     Window resolved by the caller's target scope.
'
'   DoHeadings / DoWorkbookTabs / DoGridlines
'     True when the corresponding property was requested.
'
'   ShowHeadings / ShowWorkbookTabs / ShowGridlines
'     Requested values, meaningful only when the matching Do flag is True.
'
'   LogFailures
'     True to emit Immediate Window diagnostics.
'
'   Succeeded / FailureCount / FailureList / CaptureFailureList
'     Standard structured-result buffers, updated in place.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Writes only the requested properties.
'   - Records a property-level failure and continues with the later properties
'     on the same window.
'   - Names the window in each failure entry so a multi-window pass stays
'     readable.
'
' ERROR POLICY
'   - Records property-level failures and continues with later properties.
'   - Handles unexpected errors locally, records one entry against this window,
'     and returns normally so the caller's enumeration continues.
'   - Does not raise to the caller.
'
' DEPENDENCIES
'   - UI_RuntimeTrySetBooleanPropertyIfNeeded
'   - UI_RuntimeHandleFailure
'   - UI_RuntimeBuildWindowLabel
'   - UI_RuntimeBuildErrorText
'
' CALLED FROM
'   - UI_ApplyExcelUIState
'
' NOTES
'   Handling errors here rather than letting them reach the caller is what
'   keeps the documented best-effort contract true across a multi-window pass.
'   The caller's handler ends in Resume Safe_Exit, so an error escaping this
'   procedure would abandon every window still to be visited.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Msg                 As String          'Per-property failure reason
    Dim WindowLabel         As String          'Diagnostic label for this window

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain unexpected errors here so the caller's enumeration survives them
        On Error GoTo Err_Handler

    'Build the label once. Reading Window.Caption while composing each failure
    'message would put a raising read inside the failure path itself.
        WindowLabel = UI_RuntimeBuildWindowLabel(TargetWindow)

'------------------------------------------------------------------------------
' APPLY HEADINGS
'------------------------------------------------------------------------------
    'Write the row and column headings state
        If DoHeadings Then
            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                TargetWindow, "DisplayHeadings", ShowHeadings, Msg) Then

                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, _
                    "Headings [" & WindowLabel & "]", Msg
            End If
        End If

'------------------------------------------------------------------------------
' APPLY WORKBOOK TABS
'------------------------------------------------------------------------------
    'Write the sheet-tab strip state
        If DoWorkbookTabs Then
            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                TargetWindow, "DisplayWorkbookTabs", ShowWorkbookTabs, Msg) Then

                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, _
                    "WorkbookTabs [" & WindowLabel & "]", Msg
            End If
        End If

'------------------------------------------------------------------------------
' APPLY GRIDLINES
'------------------------------------------------------------------------------
    'Write the gridline state
        If DoGridlines Then
            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                TargetWindow, "DisplayGridlines", ShowGridlines, Msg) Then

                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, _
                    "Gridlines [" & WindowLabel & "]", Msg
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
    'Record one entry against this window and return normally, so the caller
    'continues with the windows that follow
        UI_RuntimeHandleFailure _
            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, _
            "Window [" & WindowLabel & "]", UI_RuntimeBuildErrorText

        Resume Safe_Exit

End Sub


Private Function UI_IsValidTargetScope( _
    ByVal TargetScope As UIWindowTargetScope) _
    As Boolean
'
'==============================================================================
' UI_IsValidTargetScope
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a UIWindowTargetScope value defensively.
'
' WHY THIS EXISTS
'   VBA does not constrain an enum-typed parameter to its declared members, so
'   any Long can arrive here at runtime. Validating explicitly turns that into
'   a reported failure rather than a silently ignored Select Case.
'
' INPUTS
'   TargetScope
'     Value to validate.
'
' RETURNS
'   Boolean
'     True only for the three documented targeting scopes.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - UI_ApplyExcelUIState
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Accept only the three documented members
        UI_IsValidTargetScope = _
            (TargetScope = UI_TargetAllExcelWindows) Or _
            (TargetScope = UI_TargetActiveWindow) Or _
            (TargetScope = UI_TargetActiveWorkbookWindows)

End Function


Private Function UI_IsValidVisibility( _
    ByVal Visibility As UIVisibility) _
    As Boolean
'
'==============================================================================
' UI_IsValidVisibility
'------------------------------------------------------------------------------
' PURPOSE
'   Validates a UIVisibility value defensively.
'
' WHY THIS EXISTS
'   Invalid numeric values can reach a VBA enum-typed parameter at runtime. The
'   shared worker validates every requested value before converting it to a
'   Boolean target, so an out-of-range argument is reported rather than
'   silently treated as a hide.
'
' INPUTS
'   Visibility
'     Value to validate.
'
' RETURNS
'   Boolean
'     True only for UI_LeaveUnchanged, UI_Hide or UI_Show.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - UI_ApplyExcelUIState
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Accept only the three documented members
        UI_IsValidVisibility = _
            (Visibility = UI_LeaveUnchanged) Or _
            (Visibility = UI_Hide) Or _
            (Visibility = UI_Show)

End Function


Private Function UI_VisibilityToBoolean( _
    ByVal Visibility As UIVisibility) _
    As Boolean
'
'==============================================================================
' UI_VisibilityToBoolean
'------------------------------------------------------------------------------
' PURPOSE
'   Converts UI_Show or UI_Hide to a Boolean visible state.
'
' INPUTS
'   Visibility
'     Value to convert. Callers validate the enum first.
'
' RETURNS
'   Boolean
'     True for UI_Show; False otherwise.
'
' ERROR POLICY
'   - Does not raise. Callers validate the enum and exclude UI_LeaveUnchanged
'     before calling.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - UI_ApplyExcelUIState
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Only UI_Show means visible
        UI_VisibilityToBoolean = (Visibility = UI_Show)

End Function




