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
' DECLARE: PRIVATE MODULE STATE
'------------------------------------------------------------------------------
    Private m_HasExcelUIStateSnapshot    As Boolean

    Private m_SnapshotRibbonKnown        As Boolean
    Private m_SnapshotRibbonVisible      As Boolean
    Private m_SnapshotStatusBarVisible   As Boolean
    Private m_SnapshotScrollBarsVisible  As Boolean
    Private m_SnapshotFormulaBarVisible  As Boolean

    Private m_SnapshotWindowCount        As Long

    'Object references provide identity-safe in-memory matching. Parallel labels
    'are diagnostic only and never participate in matching.
    Private m_SnapshotWindows()             As Object
    Private m_SnapshotWindowLabels()        As String

    Private m_SnapshotHeadingsKnown()       As Boolean
    Private m_SnapshotHeadingsVisible()     As Boolean

    Private m_SnapshotWorkbookTabsKnown()   As Boolean
    Private m_SnapshotWorkbookTabsVisible() As Boolean

    Private m_SnapshotGridlinesKnown()      As Boolean
    Private m_SnapshotGridlinesVisible()    As Boolean

    Private m_SnapshotTitleBarKnown      As Boolean
    Private m_SnapshotTitleBarVisible    As Boolean


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
        UI_LogFailure PROC, "Unexpected", UI_BuildRuntimeErrorText
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
        UI_LogFailure PROC, "Unexpected", UI_BuildRuntimeErrorText
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
        UI_LogFailure PROC, "Unexpected", UI_BuildRuntimeErrorText
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
'   - UI_HandleApplyFailure
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

        UI_ClearResultBuffer _
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
        UI_HandleApplyFailure _
            ProcName:=PROC, _
            LogFailures:=False, _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", _
            Detail:=UI_BuildRuntimeErrorText

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
'   - UI_CaptureExcelUIState_Core
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
        UI_CaptureExcelUIState_Core _
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
        UI_LogFailure PROC, "Unexpected", UI_BuildRuntimeErrorText
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
'   - UI_CaptureExcelUIState_Core
'   - UI_HandleApplyFailure
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

        UI_ClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

'------------------------------------------------------------------------------
' CAPTURE
'------------------------------------------------------------------------------
        Succeeded = UI_CaptureExcelUIState_Core( _
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
        UI_HandleApplyFailure _
            ProcName:=PROC, _
            LogFailures:=False, _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", _
            Detail:=UI_BuildRuntimeErrorText

        Resume SafeExit

End Function


Private Function UI_CaptureExcelUIState_Core( _
    ByVal ProcName As String, _
    ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) As Boolean

'
'==============================================================================
'                    UI_CaptureExcelUIState_Core
'------------------------------------------------------------------------------
' PURPOSE
'   Execute the shared snapshot-capture pass for the compatibility wrapper and
'   structured-result API.
'
' INPUTS
'   ProcName
'     Public caller name used for diagnostics.
'
'   LogFailures
'     TRUE to emit Immediate Window diagnostics; FALSE for result-only use.
'
'   FailureCount / FailureList
'     Standard structured-result buffers.
'
'   CaptureFailureList
'     TRUE when FailureList should be populated.
'
' RETURNS
'   TRUE when no failure was recorded; otherwise FALSE.
'
' BEHAVIOR
'   - Clears any prior snapshot.
'   - Captures application-level state.
'   - Captures Ribbon and title-bar state on a best-effort basis.
'   - Captures each window's retained object identity and managed properties.
'   - Records failures in deterministic capture order.
'   - Marks the snapshot available after the capture pass completes.
'
' ERROR POLICY
'   - Does not raise.
'   - Optional-element failures are recorded and capture continues.
'   - Unexpected failures clear the partial snapshot and record one failure.
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                As Long
    Dim W                As Window
    Dim Msg              As String
    Dim UnexpectedDetail As String
    Dim Succeeded        As Boolean

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        UI_ClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

        UI_ClearExcelUIStateSnapshot

'------------------------------------------------------------------------------
' CAPTURE: APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
        m_SnapshotStatusBarVisible = Application.DisplayStatusBar
        m_SnapshotScrollBarsVisible = Application.DisplayScrollBars
        m_SnapshotFormulaBarVisible = Application.DisplayFormulaBar

'------------------------------------------------------------------------------
' CAPTURE: RIBBON / TITLE BAR
'------------------------------------------------------------------------------
        m_SnapshotRibbonKnown = UI_TryGetRibbonVisible( _
            IsVisible:=m_SnapshotRibbonVisible, _
            FailMsg:=Msg)

        If Not m_SnapshotRibbonKnown Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Ribbon", Msg
        End If

        m_SnapshotTitleBarKnown = UI_TryGetTitleBarVisible( _
            IsVisible:=m_SnapshotTitleBarVisible, _
            FailMsg:=Msg)

        If Not m_SnapshotTitleBarKnown Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "TitleBar", Msg
        End If

'------------------------------------------------------------------------------
' CAPTURE: WINDOW IDENTITIES AND STATE
'------------------------------------------------------------------------------
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

            i = 0

            For Each W In Application.Windows
                i = i + 1

                Set m_SnapshotWindows(i) = W
                m_SnapshotWindowLabels(i) = UI_BuildWindowIdentityText(W)

                m_SnapshotHeadingsKnown(i) = UI_TryGetBooleanProperty( _
                    Target:=W, _
                    PropertyName:="DisplayHeadings", _
                    ValueOut:=m_SnapshotHeadingsVisible(i), _
                    FailMsg:=Msg)

                If Not m_SnapshotHeadingsKnown(i) Then
                    UI_HandleApplyFailure _
                        ProcName, LogFailures, Succeeded, FailureCount, _
                        FailureList, CaptureFailureList, _
                        "Headings [" & m_SnapshotWindowLabels(i) & "]", Msg
                End If

                m_SnapshotWorkbookTabsKnown(i) = UI_TryGetBooleanProperty( _
                    Target:=W, _
                    PropertyName:="DisplayWorkbookTabs", _
                    ValueOut:=m_SnapshotWorkbookTabsVisible(i), _
                    FailMsg:=Msg)

                If Not m_SnapshotWorkbookTabsKnown(i) Then
                    UI_HandleApplyFailure _
                        ProcName, LogFailures, Succeeded, FailureCount, _
                        FailureList, CaptureFailureList, _
                        "WorkbookTabs [" & m_SnapshotWindowLabels(i) & "]", Msg
                End If

                m_SnapshotGridlinesKnown(i) = UI_TryGetBooleanProperty( _
                    Target:=W, _
                    PropertyName:="DisplayGridlines", _
                    ValueOut:=m_SnapshotGridlinesVisible(i), _
                    FailMsg:=Msg)

                If Not m_SnapshotGridlinesKnown(i) Then
                    UI_HandleApplyFailure _
                        ProcName, LogFailures, Succeeded, FailureCount, _
                        FailureList, CaptureFailureList, _
                        "Gridlines [" & m_SnapshotWindowLabels(i) & "]", Msg
                End If
            Next W
        End If

'------------------------------------------------------------------------------
' MARK SNAPSHOT AVAILABLE
'------------------------------------------------------------------------------
        m_HasExcelUIStateSnapshot = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        UI_CaptureExcelUIState_Core = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UnexpectedDetail = UI_BuildRuntimeErrorText
        UI_ClearExcelUIStateSnapshot

        UI_HandleApplyFailure _
            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, "Unexpected", UnexpectedDetail

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

        UI_HasExcelUIStateSnapshot = m_HasExcelUIStateSnapshot

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
'   - UI_ResetExcelUIToSnapshot_Core
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
        UI_ResetExcelUIToSnapshot_Core _
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
        UI_LogFailure PROC, "Unexpected", UI_BuildRuntimeErrorText
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
'   - UI_ResetExcelUIToSnapshot_Core
'   - UI_HandleApplyFailure
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

        UI_ClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
        Succeeded = UI_ResetExcelUIToSnapshot_Core( _
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
        UI_HandleApplyFailure _
            ProcName:=PROC, _
            LogFailures:=False, _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", _
            Detail:=UI_BuildRuntimeErrorText

        Resume SafeExit

End Function


Private Function UI_ResetExcelUIToSnapshot_Core( _
    ByVal ProcName As String, _
    ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) As Boolean

'
'==============================================================================
'                 UI_ResetExcelUIToSnapshot_Core
'------------------------------------------------------------------------------
' PURPOSE
'   Execute the shared snapshot-restoration pass for the compatibility wrapper
'   and structured-result API.
'
' INPUTS
'   ProcName
'     Public caller name used for diagnostics.
'
'   LogFailures
'     TRUE to emit Immediate Window diagnostics; FALSE for result-only use.
'
'   FailureCount / FailureList
'     Standard structured-result buffers.
'
'   CaptureFailureList
'     TRUE when FailureList should be populated.
'
' RETURNS
'   TRUE when no failure was recorded; otherwise FALSE.
'
' BEHAVIOR
'   - Restores title bar and Ribbon when their captured states were readable.
'   - Restores application-level object-model properties.
'   - Resolves each captured Window by retained object identity.
'   - Restores state only to matching still-usable Windows.
'   - Records failures in deterministic restoration order.
'   - Preserves ScreenUpdating through a quiet-update scope.
'
' ERROR POLICY
'   - Does not raise.
'   - Continues after element/window-level failure.
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long
    Dim MatchedWindow       As Object
    Dim Msg                 As String
    Dim UnexpectedDetail    As String
    Dim OldScreenUpdating   As Boolean
    Dim QuietModeChanged    As Boolean
    Dim Succeeded           As Boolean

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        UI_ClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

        If Not m_HasExcelUIStateSnapshot Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "NoSnapshot", _
                "no captured Excel UI snapshot is available"

            GoTo SafeExit
        End If

        UI_BeginQuietUIUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

'------------------------------------------------------------------------------
' RESTORE: TITLE BAR
'------------------------------------------------------------------------------
        If m_SnapshotTitleBarKnown Then
            If Not UI_TrySetTitleBarVisibleIfNeeded( _
                IsVisible:=m_SnapshotTitleBarVisible, _
                FailMsg:=Msg) Then

                UI_HandleApplyFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "TitleBar", Msg
            End If
        End If

'------------------------------------------------------------------------------
' RESTORE: RIBBON
'------------------------------------------------------------------------------
        If m_SnapshotRibbonKnown Then
            If Not UI_TrySetRibbonVisibleIfNeeded( _
                IsVisible:=m_SnapshotRibbonVisible, _
                FailMsg:=Msg) Then

                UI_HandleApplyFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "Ribbon", Msg
            End If
        End If

'------------------------------------------------------------------------------
' RESTORE: APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
        If Not UI_TrySetBooleanPropertyIfNeeded( _
            Target:=Application, _
            PropertyName:="DisplayStatusBar", _
            NewValue:=m_SnapshotStatusBarVisible, _
            FailMsg:=Msg) Then

            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "StatusBar", Msg
        End If

        If Not UI_TrySetBooleanPropertyIfNeeded( _
            Target:=Application, _
            PropertyName:="DisplayScrollBars", _
            NewValue:=m_SnapshotScrollBarsVisible, _
            FailMsg:=Msg) Then

            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "ScrollBars", Msg
        End If

        If Not UI_TrySetBooleanPropertyIfNeeded( _
            Target:=Application, _
            PropertyName:="DisplayFormulaBar", _
            NewValue:=m_SnapshotFormulaBarVisible, _
            FailMsg:=Msg) Then

            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "FormulaBar", Msg
        End If

'------------------------------------------------------------------------------
' RESTORE: WINDOW-LEVEL STATE BY OBJECT IDENTITY
'------------------------------------------------------------------------------
        For i = 1 To m_SnapshotWindowCount
            Set MatchedWindow = Nothing

            If UI_TryResolveSnapshotWindow( _
                SnapshotIndex:=i, _
                WindowOut:=MatchedWindow, _
                FailMsg:=Msg) Then

                If m_SnapshotHeadingsKnown(i) Then
                    If Not UI_TrySetBooleanPropertyIfNeeded( _
                        Target:=MatchedWindow, _
                        PropertyName:="DisplayHeadings", _
                        NewValue:=m_SnapshotHeadingsVisible(i), _
                        FailMsg:=Msg) Then

                        UI_HandleApplyFailure _
                            ProcName, LogFailures, Succeeded, FailureCount, _
                            FailureList, CaptureFailureList, _
                            "Headings [" & m_SnapshotWindowLabels(i) & "]", Msg
                    End If
                End If

                If m_SnapshotWorkbookTabsKnown(i) Then
                    If Not UI_TrySetBooleanPropertyIfNeeded( _
                        Target:=MatchedWindow, _
                        PropertyName:="DisplayWorkbookTabs", _
                        NewValue:=m_SnapshotWorkbookTabsVisible(i), _
                        FailMsg:=Msg) Then

                        UI_HandleApplyFailure _
                            ProcName, LogFailures, Succeeded, FailureCount, _
                            FailureList, CaptureFailureList, _
                            "WorkbookTabs [" & m_SnapshotWindowLabels(i) & "]", Msg
                    End If
                End If

                If m_SnapshotGridlinesKnown(i) Then
                    If Not UI_TrySetBooleanPropertyIfNeeded( _
                        Target:=MatchedWindow, _
                        PropertyName:="DisplayGridlines", _
                        NewValue:=m_SnapshotGridlinesVisible(i), _
                        FailMsg:=Msg) Then

                        UI_HandleApplyFailure _
                            ProcName, LogFailures, Succeeded, FailureCount, _
                            FailureList, CaptureFailureList, _
                            "Gridlines [" & m_SnapshotWindowLabels(i) & "]", Msg
                    End If
                End If

            Else
                UI_HandleApplyFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, _
                    "WindowIdentity [" & m_SnapshotWindowLabels(i) & "]", Msg
            End If
        Next i

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        UI_EndQuietUIUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

        UI_ResetExcelUIToSnapshot_Core = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UnexpectedDetail = UI_BuildRuntimeErrorText

        UI_HandleApplyFailure _
            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, "Unexpected", UnexpectedDetail

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
' RESET FLAGS
'------------------------------------------------------------------------------
        m_HasExcelUIStateSnapshot = False

        m_SnapshotRibbonKnown = False
        m_SnapshotTitleBarKnown = False

'------------------------------------------------------------------------------
' RESET VALUES
'------------------------------------------------------------------------------
        m_SnapshotRibbonVisible = False
        m_SnapshotStatusBarVisible = False
        m_SnapshotScrollBarsVisible = False
        m_SnapshotFormulaBarVisible = False
        m_SnapshotTitleBarVisible = False

        m_SnapshotWindowCount = 0

'------------------------------------------------------------------------------
' RELEASE WINDOW REFERENCES AND ARRAYS
'------------------------------------------------------------------------------
        Erase m_SnapshotWindows
        Erase m_SnapshotWindowLabels

        Erase m_SnapshotHeadingsKnown
        Erase m_SnapshotHeadingsVisible

        Erase m_SnapshotWorkbookTabsKnown
        Erase m_SnapshotWorkbookTabsVisible

        Erase m_SnapshotGridlinesKnown
        Erase m_SnapshotGridlinesVisible

End Sub


Private Function UI_TryResolveSnapshotWindow( _
    ByVal SnapshotIndex As Long, _
    ByRef WindowOut As Object, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                      UI_TryResolveSnapshotWindow
'------------------------------------------------------------------------------
' PURPOSE
'   Resolve one captured Excel Window by validating and returning the retained
'   Window object reference captured in the snapshot
'
' WHY
'   Re-enumerating Application.Windows may return a different COM wrapper for
'   the same live Excel window
'
'   Comparing that wrapper with the retained wrapper through the Is operator can
'   therefore reject a valid surviving window
'
'   The retained object reference is the authoritative identity and should be
'   used directly, provided it remains usable
'
' INPUTS
'   SnapshotIndex
'     1-based index into the internal snapshot arrays
'
'   WindowOut
'     Receives the retained captured Excel Window object
'
'   FailMsg
'     Receives a diagnostic reason on failure
'
' RETURNS
'   TRUE  => the retained captured Window reference remains usable
'   FALSE => the captured Window was closed or its reference is unavailable
'
' BEHAVIOR
'   - Retrieves the exact Window object retained during snapshot capture
'   - Performs a non-mutating property read to validate that the object remains
'     usable
'   - Never searches by collection index, caption, workbook name, or hWnd
'   - Never redirects captured state to a newly created replacement window
'
' ERROR POLICY
'   - Does not raise to callers
'   - Returns FALSE and populates FailMsg
'
' DEPENDENCIES
'   - UI_TryGetBooleanProperty
'   - UI_BuildRuntimeErrorText
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CapturedWindow As Object     'Exact Window object retained during capture
    Dim ProbeValue     As Boolean    'Non-mutating liveness-probe output
    Dim ProbeMsg       As String     'Diagnostic returned by the liveness probe

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TryResolveSnapshotWindow = False
        Set WindowOut = Nothing
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE SNAPSHOT INDEX
'------------------------------------------------------------------------------
        If SnapshotIndex < 1 Or SnapshotIndex > m_SnapshotWindowCount Then
            FailMsg = "snapshot index is outside the captured window range"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETRIEVE CAPTURED WINDOW REFERENCE
'------------------------------------------------------------------------------
        Set CapturedWindow = m_SnapshotWindows(SnapshotIndex)

        If CapturedWindow Is Nothing Then
            FailMsg = "captured window reference is unavailable"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' VALIDATE RETAINED WINDOW REFERENCE
'------------------------------------------------------------------------------
    'Use a non-mutating read of an existing managed property to confirm that the
    'retained Excel Window object remains usable
        If Not UI_TryGetBooleanProperty( _
            Target:=CapturedWindow, _
            PropertyName:="DisplayHeadings", _
            ValueOut:=ProbeValue, _
            FailMsg:=ProbeMsg) Then

            FailMsg = _
                "captured window is no longer open or usable; no state was applied"

            If Len(ProbeMsg) > 0 Then
                FailMsg = FailMsg & " | " & ProbeMsg
            End If

            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN EXACT CAPTURED WINDOW
'------------------------------------------------------------------------------
        Set WindowOut = CapturedWindow
        UI_TryResolveSnapshotWindow = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_BuildRuntimeErrorText
        Set WindowOut = Nothing

        Resume SafeExit

End Function



Private Function UI_BuildWindowIdentityText(ByVal TargetWindow As Object) As String

'
'==============================================================================
'                       UI_BuildWindowIdentityText
'------------------------------------------------------------------------------
' PURPOSE
'   Build a stable best-effort diagnostic label for a captured Excel Window.
'
' INPUTS
'   TargetWindow
'     Window whose identifying text should be described.
'
' RETURNS
'   A diagnostic label. This text is never used for identity matching.
'
' ERROR POLICY
'   - Does not raise.
'   - Falls back to a generic label if Excel cannot expose descriptive fields.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WorkbookName As String
    Dim WindowCaption As String

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error Resume Next

        WorkbookName = TargetWindow.Parent.Name
        WindowCaption = TargetWindow.Caption

        If Len(WorkbookName) > 0 And Len(WindowCaption) > 0 Then
            UI_BuildWindowIdentityText = WorkbookName & " :: " & WindowCaption
        ElseIf Len(WindowCaption) > 0 Then
            UI_BuildWindowIdentityText = WindowCaption
        ElseIf Len(WorkbookName) > 0 Then
            UI_BuildWindowIdentityText = WorkbookName
        Else
            UI_BuildWindowIdentityText = "captured Excel window"
        End If

End Function


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
        UI_ClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

        UI_BeginQuietUIUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
        ValidRibbon = UI_IsValidVisibility(Ribbon)
        If Not ValidRibbon Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Ribbon", _
                "invalid UIVisibility value: " & CStr(Ribbon)
        End If

        ValidStatusBar = UI_IsValidVisibility(StatusBar)
        If Not ValidStatusBar Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "StatusBar", _
                "invalid UIVisibility value: " & CStr(StatusBar)
        End If

        ValidScrollBars = UI_IsValidVisibility(ScrollBars)
        If Not ValidScrollBars Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "ScrollBars", _
                "invalid UIVisibility value: " & CStr(ScrollBars)
        End If

        ValidFormulaBar = UI_IsValidVisibility(FormulaBar)
        If Not ValidFormulaBar Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "FormulaBar", _
                "invalid UIVisibility value: " & CStr(FormulaBar)
        End If

        ValidHeadings = UI_IsValidVisibility(Headings)
        If Not ValidHeadings Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Headings", _
                "invalid UIVisibility value: " & CStr(Headings)
        End If

        ValidWorkbookTabs = UI_IsValidVisibility(WorkbookTabs)
        If Not ValidWorkbookTabs Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "WorkbookTabs", _
                "invalid UIVisibility value: " & CStr(WorkbookTabs)
        End If

        ValidGridlines = UI_IsValidVisibility(Gridlines)
        If Not ValidGridlines Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Gridlines", _
                "invalid UIVisibility value: " & CStr(Gridlines)
        End If

        ValidTitleBar = UI_IsValidVisibility(TitleBar)
        If Not ValidTitleBar Then
            UI_HandleApplyFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "TitleBar", _
                "invalid UIVisibility value: " & CStr(TitleBar)
        End If

'------------------------------------------------------------------------------
' APPLY APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
        If ValidRibbon And Ribbon <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(Ribbon)

            If Not UI_TrySetRibbonVisibleIfNeeded(ShowFlag, Msg) Then
                UI_HandleApplyFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "Ribbon", Msg
            End If
        End If

        If ValidStatusBar And StatusBar <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(StatusBar)

            If Not UI_TrySetBooleanPropertyIfNeeded( _
                Application, "DisplayStatusBar", ShowFlag, Msg) Then

                UI_HandleApplyFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "StatusBar", Msg
            End If
        End If

        If ValidScrollBars And ScrollBars <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(ScrollBars)

            If Not UI_TrySetBooleanPropertyIfNeeded( _
                Application, "DisplayScrollBars", ShowFlag, Msg) Then

                UI_HandleApplyFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "ScrollBars", Msg
            End If
        End If

        If ValidFormulaBar And FormulaBar <> UI_LeaveUnchanged Then
            ShowFlag = UI_VisibilityToBoolean(FormulaBar)

            If Not UI_TrySetBooleanPropertyIfNeeded( _
                Application, "DisplayFormulaBar", ShowFlag, Msg) Then

                UI_HandleApplyFailure _
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
                    If Not UI_TrySetBooleanPropertyIfNeeded( _
                        W, "DisplayHeadings", ShowHeadings, Msg) Then

                        UI_HandleApplyFailure _
                            ProcName, LogFailures, Succeeded, FailureCount, _
                            FailureList, CaptureFailureList, _
                            "Headings [" & W.Caption & "]", Msg
                    End If
                End If

                If DoWorkbookTabs Then
                    If Not UI_TrySetBooleanPropertyIfNeeded( _
                        W, "DisplayWorkbookTabs", ShowWorkbookTabs, Msg) Then

                        UI_HandleApplyFailure _
                            ProcName, LogFailures, Succeeded, FailureCount, _
                            FailureList, CaptureFailureList, _
                            "WorkbookTabs [" & W.Caption & "]", Msg
                    End If
                End If

                If DoGridlines Then
                    If Not UI_TrySetBooleanPropertyIfNeeded( _
                        W, "DisplayGridlines", ShowGridlines, Msg) Then

                        UI_HandleApplyFailure _
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
                UI_HandleApplyFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "TitleBar", Msg
            End If
        End If

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        UI_EndQuietUIUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

        UI_ApplyExcelUIState = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UI_HandleApplyFailure _
            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, "Unexpected", UI_BuildRuntimeErrorText

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


Private Sub UI_HandleApplyFailure( _
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
'                           UI_HandleApplyFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Record one best-effort operation failure and optionally log it.
'
' NOTES
'   The procedure name is retained for compatibility with the established
'   internal apply path, but the helper is also used by snapshot operations.
'
' ERROR POLICY
'   Does not raise.
'
' UPDATED
'   2026-07-29
'==============================================================================
'

        UI_AddFailureToResult _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:=Stage, _
            Detail:=Detail

        If LogFailures Then
            UI_LogFailure ProcName, Stage, Detail
        End If

End Sub


Private Sub UI_ClearResultBuffer( _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean)

'
'==============================================================================
'                           UI_ClearResultBuffer
'------------------------------------------------------------------------------
' PURPOSE
'   Initialize structured result buffers.
'
' ERROR POLICY
'   Does not raise.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        FailureCount = 0

        If CaptureFailureList Then
            FailureList = Empty
        End If

End Sub


Private Sub UI_AddFailureToResult( _
    ByRef Succeeded As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean, _
    ByVal Stage As String, _
    ByVal Detail As String)

'
'==============================================================================
'                          UI_AddFailureToResult
'------------------------------------------------------------------------------
' PURPOSE
'   Append one failure to the standard Boolean / count / list result contract.
'
' ERROR POLICY
'   Does not raise.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Arr() As String

'------------------------------------------------------------------------------
' UPDATE STATUS
'------------------------------------------------------------------------------
        Succeeded = False
        FailureCount = FailureCount + 1

'------------------------------------------------------------------------------
' APPEND TEXT
'------------------------------------------------------------------------------
        If CaptureFailureList Then
            If IsEmpty(FailureList) Then
                ReDim Arr(1 To 1)
            Else
                Arr = FailureList
                ReDim Preserve Arr(1 To FailureCount)
            End If

            Arr(FailureCount) = Stage & " | " & Detail
            FailureList = Arr
        End If

End Sub


Private Sub UI_BeginQuietUIUpdate( _
    ByRef OldScreenUpdating As Boolean, _
    ByRef QuietModeChanged As Boolean)

'
'==============================================================================
'                          UI_BeginQuietUIUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Enter a best-effort ScreenUpdating suppression scope.
'
' ERROR POLICY
'   Suppresses errors locally.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        On Error Resume Next

        OldScreenUpdating = Application.ScreenUpdating
        QuietModeChanged = False

        If OldScreenUpdating Then
            Application.ScreenUpdating = False
            QuietModeChanged = True
        End If

End Sub


Private Sub UI_EndQuietUIUpdate( _
    ByVal OldScreenUpdating As Boolean, _
    ByVal QuietModeChanged As Boolean)

'
'==============================================================================
'                           UI_EndQuietUIUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Restore ScreenUpdating when this module changed it.
'
' ERROR POLICY
'   Suppresses errors locally.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        On Error Resume Next

        If QuietModeChanged Then
            Application.ScreenUpdating = OldScreenUpdating
        End If

End Sub


Private Function UI_TrySetRibbonVisibleIfNeeded( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                     UI_TrySetRibbonVisibleIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Set Ribbon visibility only when required.
'
' RETURNS
'   TRUE when already correct or successfully updated.
'
' ERROR POLICY
'   Returns FALSE and FailMsg on failure.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentVisible As Boolean

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TrySetRibbonVisibleIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT
'------------------------------------------------------------------------------
        If UI_TryGetRibbonVisible(CurrentVisible, FailMsg) Then
            If CurrentVisible = IsVisible Then
                UI_TrySetRibbonVisibleIfNeeded = True
                GoTo SafeExit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
        FailMsg = vbNullString

        UI_TrySetRibbonVisibleIfNeeded = _
            UI_TrySetRibbonVisible(IsVisible, FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_BuildRuntimeErrorText

End Function


Private Function UI_TrySetBooleanPropertyIfNeeded( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByVal NewValue As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                   UI_TrySetBooleanPropertyIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Set a Boolean property only when required.
'
' RETURNS
'   TRUE when already correct or successfully updated.
'
' ERROR POLICY
'   Returns FALSE and FailMsg on failure.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentValue As Boolean

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TrySetBooleanPropertyIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT
'------------------------------------------------------------------------------
        If UI_TryGetBooleanProperty( _
            Target, PropertyName, CurrentValue, FailMsg) Then

            If CurrentValue = NewValue Then
                UI_TrySetBooleanPropertyIfNeeded = True
                GoTo SafeExit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
        FailMsg = vbNullString

        UI_TrySetBooleanPropertyIfNeeded = _
            UI_TrySetBooleanProperty(Target, PropertyName, NewValue, FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_BuildRuntimeErrorText

End Function


Private Function UI_TryGetRibbonVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                         UI_TryGetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Read current Ribbon visibility using CommandBars with an Excel4 fallback.
'
' RETURNS
'   TRUE when read succeeds.
'
' ERROR POLICY
'   Returns FALSE and FailMsg on failure.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V As Variant

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TryGetRibbonVisible = False
        IsVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' TRY COMMANDBARS
'------------------------------------------------------------------------------
        On Error Resume Next
            IsVisible = Application.CommandBars("Ribbon").Visible

        If Err.Number = 0 Then
            On Error GoTo Fail
            UI_TryGetRibbonVisible = True
            GoTo SafeExit
        End If

        Err.Clear
        On Error GoTo Fail

'------------------------------------------------------------------------------
' TRY EXCEL4 FALLBACK
'------------------------------------------------------------------------------
        On Error Resume Next
            V = Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")

        If Err.Number = 0 Then
            On Error GoTo Fail
            IsVisible = CBool(V)
            UI_TryGetRibbonVisible = True
            GoTo SafeExit
        End If

        FailMsg = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        On Error GoTo Fail

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_BuildRuntimeErrorText

End Function


Private Function UI_TryGetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByRef ValueOut As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                         UI_TryGetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Read a Boolean property through CallByName.
'
' RETURNS
'   TRUE when read succeeds.
'
' ERROR POLICY
'   Returns FALSE and FailMsg on failure.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V As Variant

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TryGetBooleanProperty = False
        ValueOut = False
        FailMsg = vbNullString

        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo SafeExit
        End If

        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        V = CallByName(Target, PropertyName, VbGet)
        ValueOut = CBool(V)
        UI_TryGetBooleanProperty = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_BuildRuntimeErrorText

End Function


Private Function UI_TrySetRibbonVisible( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                           UI_TrySetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Show or hide the Ribbon using Excel4 macro execution.
'
' RETURNS
'   TRUE on success.
'
' ERROR POLICY
'   Returns FALSE and FailMsg on failure.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MacroText As String

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TrySetRibbonVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
        If IsVisible Then
            MacroText = "Show.TOOLBAR(""Ribbon"",True)"
        Else
            MacroText = "Show.TOOLBAR(""Ribbon"",False)"
        End If

        Application.ExecuteExcel4Macro MacroText

        UI_TrySetRibbonVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_BuildRuntimeErrorText

End Function


Private Function UI_TrySetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByVal NewValue As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                           UI_TrySetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Write a Boolean property through CallByName.
'
' RETURNS
'   TRUE on success.
'
' ERROR POLICY
'   Returns FALSE and FailMsg on failure.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TrySetBooleanProperty = False
        FailMsg = vbNullString

        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo SafeExit
        End If

        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
        CallByName Target, PropertyName, VbLet, NewValue
        UI_TrySetBooleanProperty = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_BuildRuntimeErrorText

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


Private Function UI_BuildRuntimeErrorText() As String

'
'==============================================================================
'                           UI_BuildRuntimeErrorText
'------------------------------------------------------------------------------
' PURPOSE
'   Build a consistent diagnostic from the active Err object.
'
' RETURNS
'   Best-effort error number, description, source, and Erl text.
'
' ERROR POLICY
'   Suppresses formatting errors locally.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        On Error Resume Next

        UI_BuildRuntimeErrorText = _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, _
                " | Source: " & Err.Source, _
                vbNullString) & _
            IIf(Erl <> 0, _
                " | Line: " & CStr(Erl), _
                vbNullString)

End Function


Private Sub UI_LogFailure( _
    ByVal ProcName As String, _
    ByVal Stage As String, _
    ByVal Detail As String)

'
'==============================================================================
'                                UI_LogFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Write a consistent diagnostic line to the Immediate Window.
'
' ERROR POLICY
'   Suppresses logging errors locally.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        On Error Resume Next

        Debug.Print ProcName & " failed @ " & Stage & " | " & Detail

End Sub
