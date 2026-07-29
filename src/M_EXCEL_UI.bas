Attribute VB_Name = "M_EXCEL_UI"
'==============================================================================
'                           MODULE: M_EXCEL_UI
'------------------------------------------------------------------------------
' PURPOSE
'   Centralize visibility control for the Excel UI elements managed by this
'   module, combining:
'     - Excel object-model UI elements
'     - Delegated WinAPI title-bar control for the Excel main window represented
'       by Application.Hwnd
'
' WHY
'   Workbook-driven solutions often need a constrained or application-style
'   Excel shell. This module provides one explicit, defensive, fail-soft API
'   instead of scattering UI writes throughout the project.
'
' PUBLIC SURFACE
'   - UIVisibility
'   - UI_SetExcelUI
'   - UI_SetExcelUI_WithResult
'   - UI_HideExcelUI
'   - UI_ShowExcelUI
'   - UI_CaptureExcelUIState
'   - UI_CaptureExcelUIState_WithResult
'   - UI_ResetExcelUIToSnapshot
'   - UI_ResetExcelUIToSnapshot_WithResult
'   - UI_HasExcelUIStateSnapshot
'   - UI_ClearExcelUIStateSnapshot
'
' BEHAVIOR
'   - Application-level UI:
'       * Ribbon
'       * Status Bar
'       * Scroll Bars
'       * Formula Bar
'   - Window-level UI:
'       * Headings
'       * Workbook Tabs
'       * Gridlines
'   - Main-window frame:
'       * Title Bar
'
' ERROR POLICY
'   - Public entry points are fail-soft.
'   - Fire-and-forget procedures log failures to the Immediate Window.
'   - UI_SetExcelUI_WithResult and the snapshot WithResult APIs return
'     structured failure information.
'   - One failed element does not prevent later requested elements from being
'     attempted.
'
' PLATFORM / COMPATIBILITY
'   - Windows only.
'   - Supports 32-bit and 64-bit Office / VBA through conditional compilation.
'
' NOTES
'   - Snapshot state is stored in memory only and is lost after project reset
'     or when Excel closes.
'   - Window-level snapshot state is keyed by the captured Excel Window object
'     identity, not by Application.Windows collection index.
'   - Reordered windows therefore restore correctly.
'   - Newly opened windows are left unchanged because no state was captured for
'     them.
'   - Closed or recreated captured windows are skipped rather than allowing
'     their state to be applied to a different window.
'   - Title-bar ownership is limited to the caption, system-menu, sizing-frame,
'     minimize-box, and maximize-box style bits.
'   - M_EXCEL_UI_TITLEBAR owns title-bar WinAPI declarations and mutable style
'     state; this module remains the public facade.
'   - M_EXCEL_UI_SNAPSHOT owns snapshot state, retained Window identities, and
'     capture / restoration orchestration.
'   - M_EXCEL_UI_RUNTIME owns shared host operations and result diagnostics.
'   - Showing the title bar merges only those owned bits into the current style,
'     preserving unrelated changes made by Excel or another component.
'   - Snapshot capture and restoration expose optional structured-result APIs
'     while retaining the original fire-and-forget compatibility wrappers.
'
' UPDATED
'   2026-07-29
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.0
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' DECLARE: PUBLIC ENUMS
'------------------------------------------------------------------------------
Public Enum UIVisibility
    UI_LeaveUnchanged = -1     'Do not touch this UI element
    UI_Hide = 0                'Hide this UI element
    UI_Show = 1                'Show this UI element
End Enum


'------------------------------------------------------------------------------
' INTERNAL MODULE DEPENDENCIES
'------------------------------------------------------------------------------
'   - M_EXCEL_UI_RUNTIME for shared fail-soft host operations and diagnostics.
'   - M_EXCEL_UI_TITLEBAR for WinAPI title-bar control.
'   - M_EXCEL_UI_SNAPSHOT for snapshot state and lifecycle.

Public Sub UI_SetExcelUI( _
    Optional ByVal Ribbon As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal StatusBar As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal ScrollBars As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal FormulaBar As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal Headings As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal WorkbookTabs As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal Gridlines As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal TitleBar As UIVisibility = UI_LeaveUnchanged)

'
'==============================================================================
'                               UI_SetExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the requested visibility state to the managed Excel UI elements.
'
' INPUTS
'   Each optional argument accepts UI_Show, UI_Hide, or UI_LeaveUnchanged.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Applies application-level settings to the current Excel instance.
'   - Applies window-level settings to every current Excel window.
'   - Applies title-bar visibility to Application.Hwnd.
'   - Continues after element-level failure.
'
' ERROR POLICY
'   - Does not raise to callers.
'   - Logs failures to the Immediate Window.
'
' DEPENDENCIES
'   - UI_ApplyExcelUIState
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IgnoredFailureCount As Long
    Dim IgnoredFailureList  As Variant

    Const PROC As String = "UI_SetExcelUI"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
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
            LogFailures:=True, _
            FailureCount:=IgnoredFailureCount, _
            FailureList:=IgnoredFailureList, _
            CaptureFailureList:=False

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UI_RuntimeLogFailure PROC, "Unexpected", UI_RuntimeBuildErrorText
        Resume SafeExit

End Sub


Public Sub UI_HideExcelUI()

'
'==============================================================================
'                               UI_HideExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Hide all Excel UI elements managed by this module.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Delegates to UI_SetExcelUI with UI_Hide for every managed element.
'
' ERROR POLICY
'   - Does not raise to callers.
'   - Logs an unexpected wrapper failure to the Immediate Window.
'
' DEPENDENCIES
'   - UI_SetExcelUI
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC As String = "UI_HideExcelUI"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
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
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UI_RuntimeLogFailure PROC, "Unexpected", UI_RuntimeBuildErrorText
        Resume SafeExit

End Sub


Public Sub UI_ShowExcelUI()

'
'==============================================================================
'                               UI_ShowExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Show all Excel UI elements managed by this module.
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
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC As String = "UI_ShowExcelUI"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
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
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UI_RuntimeLogFailure PROC, "Unexpected", UI_RuntimeBuildErrorText
        Resume SafeExit

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
    Optional ByRef FailureList As Variant) As Boolean

'
'==============================================================================
'                         UI_SetExcelUI_WithResult
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the requested managed UI state and return structured diagnostics.
'
' INPUTS
'   Visibility arguments
'     UI_Show, UI_Hide, or UI_LeaveUnchanged.
'
'   FailureCount (optional, ByRef)
'     Receives the number of recorded failures.
'
'   FailureList (optional, ByRef)
'     Receives a 1-based String array containing "Stage | Detail" entries.
'
' RETURNS
'   TRUE when no failure was recorded; otherwise FALSE.
'
' BEHAVIOR
'   Mirrors UI_SetExcelUI while suppressing Immediate Window logging.
'
' ERROR POLICY
'   - Does not raise for ordinary failures.
'   - Captures unexpected failures in the result.
'
' DEPENDENCIES
'   - UI_ApplyExcelUIState
'   - UI_RuntimeHandleFailure
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean
    Dim CaptureFailureList  As Boolean
    Dim InternalFailureList As Variant

    Const PROC As String = "UI_SetExcelUI_WithResult"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        CaptureFailureList = Not IsMissing(FailureList)

        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
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
            LogFailures:=False, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        If CaptureFailureList Then
            FailureList = InternalFailureList
        End If

        UI_SetExcelUI_WithResult = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UI_RuntimeHandleFailure _
            ProcName:=PROC, _
            LogFailures:=False, _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", _
            Detail:=UI_RuntimeBuildErrorText

        Resume SafeExit

End Function


Public Sub UI_CaptureExcelUIState()

'
'==============================================================================
'                           UI_CaptureExcelUIState
'------------------------------------------------------------------------------
' PURPOSE
'   Capture the current managed Excel UI state for later restoration.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Delegates to the shared capture worker and logs ordered best-effort failures
'   to the Immediate Window.
'
' ERROR POLICY
'   - Does not raise to callers.
'   - Logs failures and continues where capture remains meaningful.
'   - Leaves the snapshot unavailable after an unexpected capture failure.
'
' DEPENDENCIES
'   - UI_SnapshotCaptureCore
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IgnoredFailureCount As Long
    Dim IgnoredFailureList  As Variant

    Const PROC As String = "UI_CaptureExcelUIState"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
        UI_SnapshotCaptureCore _
            ProcName:=PROC, _
            LogFailures:=True, _
            FailureCount:=IgnoredFailureCount, _
            FailureList:=IgnoredFailureList, _
            CaptureFailureList:=False

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UI_RuntimeLogFailure PROC, "Unexpected", UI_RuntimeBuildErrorText
        Resume SafeExit

End Sub


Public Function UI_CaptureExcelUIState_WithResult( _
    Optional ByRef FailureCount As Long = 0, _
    Optional ByRef FailureList As Variant) As Boolean

'
'==============================================================================
'                    UI_CaptureExcelUIState_WithResult
'------------------------------------------------------------------------------
' PURPOSE
'   Capture the current managed Excel UI state and return structured diagnostics.
'
' INPUTS / OUTPUTS
'   FailureCount (optional, ByRef)
'     Receives the number of recorded capture failures.
'
'   FailureList (optional, ByRef)
'     Receives a 1-based String array containing ordered "Stage | Detail"
'     entries.
'
' RETURNS
'   TRUE when the capture pass recorded no failure; otherwise FALSE.
'
' BEHAVIOR
'   - Clears output buffers deterministically on entry.
'   - Replaces any prior snapshot.
'   - Preserves best-effort partial-capture semantics.
'   - Marks the snapshot available after the capture pass completes, even when
'     optional elements were unreadable.
'
' ERROR POLICY
'   - Does not raise for ordinary capture failures.
'   - Returns ordered element/window-specific diagnostics.
'   - Leaves the snapshot unavailable after an unexpected capture failure.
'
' DEPENDENCIES
'   - UI_SnapshotCaptureCore
'   - UI_RuntimeHandleFailure
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean
    Dim CaptureFailureList  As Boolean
    Dim InternalFailureList As Variant

    Const PROC As String = "UI_CaptureExcelUIState_WithResult"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        CaptureFailureList = Not IsMissing(FailureList)

        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
        Succeeded = UI_SnapshotCaptureCore( _
            ProcName:=PROC, _
            LogFailures:=False, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        If CaptureFailureList Then
            FailureList = InternalFailureList
        End If

        UI_CaptureExcelUIState_WithResult = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UI_RuntimeHandleFailure _
            ProcName:=PROC, _
            LogFailures:=False, _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", _
            Detail:=UI_RuntimeBuildErrorText

        Resume SafeExit

End Function


Public Function UI_HasExcelUIStateSnapshot() As Boolean

'
'==============================================================================
'                        UI_HasExcelUIStateSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Return whether an explicit in-memory Excel UI snapshot is available.
'
' RETURNS
'   TRUE when a snapshot is available; otherwise FALSE.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        UI_HasExcelUIStateSnapshot = UI_SnapshotHasState

End Function


Public Sub UI_ResetExcelUIToSnapshot()

'
'==============================================================================
'                        UI_ResetExcelUIToSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the managed Excel UI to the most recently captured snapshot.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Delegates to the shared restoration worker and logs ordered best-effort
'   failures to the Immediate Window.
'
' ERROR POLICY
'   - Does not raise to callers.
'   - Logs restore failures and continues where possible.
'
' DEPENDENCIES
'   - UI_SnapshotRestoreCore
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IgnoredFailureCount As Long
    Dim IgnoredFailureList  As Variant

    Const PROC As String = "UI_ResetExcelUIToSnapshot"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
        UI_SnapshotRestoreCore _
            ProcName:=PROC, _
            LogFailures:=True, _
            FailureCount:=IgnoredFailureCount, _
            FailureList:=IgnoredFailureList, _
            CaptureFailureList:=False

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UI_RuntimeLogFailure PROC, "Unexpected", UI_RuntimeBuildErrorText
        Resume SafeExit

End Sub


Public Function UI_ResetExcelUIToSnapshot_WithResult( _
    Optional ByRef FailureCount As Long = 0, _
    Optional ByRef FailureList As Variant) As Boolean

'
'==============================================================================
'                 UI_ResetExcelUIToSnapshot_WithResult
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the managed Excel UI to the current snapshot and return structured
'   diagnostics.
'
' INPUTS / OUTPUTS
'   FailureCount (optional, ByRef)
'     Receives the number of recorded restoration failures.
'
'   FailureList (optional, ByRef)
'     Receives a 1-based String array containing ordered "Stage | Detail"
'     entries.
'
' RETURNS
'   TRUE when restoration recorded no failure; otherwise FALSE.
'
' BEHAVIOR
'   - Clears output buffers deterministically on entry.
'   - Restores every available captured element on a best-effort basis.
'   - Leaves newly opened windows unchanged.
'   - Reports closed, recreated, or unusable captured windows without applying
'     their state to a replacement window.
'   - Retains the snapshot after the restoration attempt.
'
' ERROR POLICY
'   - Does not raise for ordinary restoration failures.
'   - Returns ordered element/window-specific diagnostics.
'
' DEPENDENCIES
'   - UI_SnapshotRestoreCore
'   - UI_RuntimeHandleFailure
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean
    Dim CaptureFailureList  As Boolean
    Dim InternalFailureList As Variant

    Const PROC As String = "UI_ResetExcelUIToSnapshot_WithResult"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        CaptureFailureList = Not IsMissing(FailureList)

        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
        Succeeded = UI_SnapshotRestoreCore( _
            ProcName:=PROC, _
            LogFailures:=False, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        If CaptureFailureList Then
            FailureList = InternalFailureList
        End If

        UI_ResetExcelUIToSnapshot_WithResult = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UI_RuntimeHandleFailure _
            ProcName:=PROC, _
            LogFailures:=False, _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", _
            Detail:=UI_RuntimeBuildErrorText

        Resume SafeExit

End Function


Public Sub UI_ClearExcelUIStateSnapshot()

'
'==============================================================================
'                      UI_ClearExcelUIStateSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Remove the current in-memory Excel UI snapshot.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Releases retained Window object references and clears all captured values.
'
' ERROR POLICY
'   Does not raise.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        On Error Resume Next

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
    ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) As Boolean

'
'==============================================================================
'                           UI_ApplyExcelUIState
'------------------------------------------------------------------------------
' PURPOSE
'   Apply requested UI state through the shared fire-and-forget / result worker.
'
' RETURNS
'   TRUE when no failure was recorded; otherwise FALSE.
'
' BEHAVIOR
'   - Validates every UIVisibility argument.
'   - Skips UI_LeaveUnchanged values.
'   - Avoids no-op writes where current state can be read.
'   - Applies window-level state to every current Excel window.
'   - Preserves ScreenUpdating.
'
' ERROR POLICY
'   - Does not raise.
'   - Records and optionally logs failures in insertion order.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean
    Dim W                   As Window
    Dim ShowFlag            As Boolean
    Dim Msg                 As String

    Dim ValidRibbon         As Boolean
    Dim ValidStatusBar      As Boolean
    Dim ValidScrollBars     As Boolean
    Dim ValidFormulaBar     As Boolean
    Dim ValidHeadings       As Boolean
    Dim ValidWorkbookTabs   As Boolean
    Dim ValidGridlines      As Boolean
    Dim ValidTitleBar       As Boolean

    Dim DoHeadings          As Boolean
    Dim DoWorkbookTabs      As Boolean
    Dim DoGridlines         As Boolean

    Dim ShowHeadings        As Boolean
    Dim ShowWorkbookTabs    As Boolean
    Dim ShowGridlines       As Boolean

    Dim OldScreenUpdating   As Boolean
    Dim QuietModeChanged    As Boolean

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

        UI_RuntimeBeginQuietUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
        ValidRibbon = UI_IsValidVisibility(Ribbon)
        If Not ValidRibbon Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Ribbon", _
                "invalid UIVisibility value: " & CStr(Ribbon)
        End If

        ValidStatusBar = UI_IsValidVisibility(StatusBar)
        If Not ValidStatusBar Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "StatusBar", _
                "invalid UIVisibility value: " & CStr(StatusBar)
        End If

        ValidScrollBars = UI_IsValidVisibility(ScrollBars)
        If Not ValidScrollBars Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "ScrollBars", _
                "invalid UIVisibility value: " & CStr(ScrollBars)
        End If

        ValidFormulaBar = UI_IsValidVisibility(FormulaBar)
        If Not ValidFormulaBar Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "FormulaBar", _
                "invalid UIVisibility value: " & CStr(FormulaBar)
        End If

        ValidHeadings = UI_IsValidVisibility(Headings)
        If Not ValidHeadings Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Headings", _
                "invalid UIVisibility value: " & CStr(Headings)
        End If

        ValidWorkbookTabs = UI_IsValidVisibility(WorkbookTabs)
        If Not ValidWorkbookTabs Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "WorkbookTabs", _
                "invalid UIVisibility value: " & CStr(WorkbookTabs)
        End If

        ValidGridlines = UI_IsValidVisibility(Gridlines)
        If Not ValidGridlines Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Gridlines", _
                "invalid UIVisibility value: " & CStr(Gridlines)
        End If

        ValidTitleBar = UI_IsValidVisibility(TitleBar)
        If Not ValidTitleBar Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "TitleBar", _
                "invalid UIVisibility value: " & CStr(TitleBar)
        End If

'------------------------------------------------------------------------------
' APPLY APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
        If ValidRibbon And Ribbon <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(Ribbon)

            If Not UI_RuntimeTrySetRibbonVisibleIfNeeded(ShowFlag, Msg) Then
                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "Ribbon", Msg
            End If
        End If

        If ValidStatusBar And StatusBar <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(StatusBar)

            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                Application, "DisplayStatusBar", ShowFlag, Msg) Then

                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "StatusBar", Msg
            End If
        End If

        If ValidScrollBars And ScrollBars <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(ScrollBars)

            If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                Application, "DisplayScrollBars", ShowFlag, Msg) Then

                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "ScrollBars", Msg
            End If
        End If

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
        DoHeadings = ValidHeadings And (Headings <> UI_LeaveUnchanged)
        DoWorkbookTabs = _
            ValidWorkbookTabs And (WorkbookTabs <> UI_LeaveUnchanged)
        DoGridlines = ValidGridlines And (Gridlines <> UI_LeaveUnchanged)

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
        If DoHeadings Or DoWorkbookTabs Or DoGridlines Then
            For Each W In Application.Windows

                If DoHeadings Then
                    If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                        W, "DisplayHeadings", ShowHeadings, Msg) Then

                        UI_RuntimeHandleFailure _
                            ProcName, LogFailures, Succeeded, FailureCount, _
                            FailureList, CaptureFailureList, _
                            "Headings [" & W.Caption & "]", Msg
                    End If
                End If

                If DoWorkbookTabs Then
                    If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                        W, "DisplayWorkbookTabs", ShowWorkbookTabs, Msg) Then

                        UI_RuntimeHandleFailure _
                            ProcName, LogFailures, Succeeded, FailureCount, _
                            FailureList, CaptureFailureList, _
                            "WorkbookTabs [" & W.Caption & "]", Msg
                    End If
                End If

                If DoGridlines Then
                    If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
                        W, "DisplayGridlines", ShowGridlines, Msg) Then

                        UI_RuntimeHandleFailure _
                            ProcName, LogFailures, Succeeded, FailureCount, _
                            FailureList, CaptureFailureList, _
                            "Gridlines [" & W.Caption & "]", Msg
                    End If
                End If

            Next W
        End If

'------------------------------------------------------------------------------
' APPLY TITLE-BAR STATE
'------------------------------------------------------------------------------
        If ValidTitleBar And TitleBar <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(TitleBar)

            If Not UI_TrySetTitleBarVisibleIfNeeded(ShowFlag, Msg) Then
                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "TitleBar", Msg
            End If
        End If

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        UI_RuntimeEndQuietUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

        UI_ApplyExcelUIState = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UI_RuntimeHandleFailure _
            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, "Unexpected", UI_RuntimeBuildErrorText

        Resume SafeExit

End Function


Private Function UI_IsValidVisibility( _
    ByVal Visibility As UIVisibility) As Boolean

'
'==============================================================================
'                           UI_IsValidVisibility
'------------------------------------------------------------------------------
' PURPOSE
'   Validate a UIVisibility value defensively.
'
' RETURNS
'   TRUE only for UI_LeaveUnchanged, UI_Hide, or UI_Show.
'
' ERROR POLICY
'   Does not raise.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        UI_IsValidVisibility = _
            (Visibility = UI_LeaveUnchanged) Or _
            (Visibility = UI_Hide) Or _
            (Visibility = UI_Show)

End Function
Private Function UI_VisibilityToBoolean( _
    ByVal Visibility As UIVisibility) As Boolean

'
'==============================================================================
'                           UI_VisibilityToBoolean
'------------------------------------------------------------------------------
' PURPOSE
'   Convert UI_Show/UI_Hide to a Boolean visible state.
'
' RETURNS
'   TRUE for UI_Show; FALSE otherwise.
'
' ERROR POLICY
'   Does not raise. Callers validate the enum first.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        UI_VisibilityToBoolean = (Visibility = UI_Show)

End Function
