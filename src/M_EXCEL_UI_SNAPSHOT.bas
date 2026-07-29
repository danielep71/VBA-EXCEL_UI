Attribute VB_Name = "M_EXCEL_UI_SNAPSHOT"
'==============================================================================
'                    MODULE: M_EXCEL_UI_SNAPSHOT
'------------------------------------------------------------------------------
' PURPOSE
'   Own the in-memory Excel UI snapshot lifecycle, retained Window identities,
'   and best-effort capture / restoration orchestration.
'
' WHY
'   Snapshot state and identity resolution form a cohesive subsystem with a
'   lifecycle independent from the public facade. Isolating that subsystem keeps
'   M_EXCEL_UI focused on API compatibility and general UI application logic.
'
' INTERNAL SURFACE
'   - UI_SnapshotCaptureCore
'   - UI_SnapshotRestoreCore
'   - UI_SnapshotHasState
'   - UI_SnapshotClear
'
' BEHAVIOR
'   - Stores snapshot state in memory only.
'   - Retains exact Excel Window object references.
'   - Never restores by Application.Windows collection index.
'   - Leaves newly opened windows unchanged.
'   - Reports closed or unusable captured windows deterministically.
'   - Preserves the established ordered diagnostic contract.
'
' ERROR POLICY
'   - Internal entry points are fail-soft.
'   - Capture and restoration continue after element-level failure.
'   - No user-interface messages are displayed.
'
' DEPENDENCIES
'   - M_EXCEL_UI_RUNTIME for shared host operations and diagnostics.
'   - M_EXCEL_UI_TITLEBAR for title-bar capture and restoration.
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



Public Function UI_SnapshotCaptureCore( _
    ByVal ProcName As String, _
    ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) As Boolean

'
'==============================================================================
'                    UI_SnapshotCaptureCore
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
        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

        UI_SnapshotClear

'------------------------------------------------------------------------------
' CAPTURE: APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
        m_SnapshotStatusBarVisible = Application.DisplayStatusBar
        m_SnapshotScrollBarsVisible = Application.DisplayScrollBars
        m_SnapshotFormulaBarVisible = Application.DisplayFormulaBar

'------------------------------------------------------------------------------
' CAPTURE: RIBBON / TITLE BAR
'------------------------------------------------------------------------------
        m_SnapshotRibbonKnown = UI_RuntimeTryGetRibbonVisible( _
            IsVisible:=m_SnapshotRibbonVisible, _
            FailMsg:=Msg)

        If Not m_SnapshotRibbonKnown Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "Ribbon", Msg
        End If

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
                m_SnapshotWindowLabels(i) = UI_SnapshotBuildWindowIdentityText(W)

                m_SnapshotHeadingsKnown(i) = UI_RuntimeTryGetBooleanProperty( _
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

                m_SnapshotWorkbookTabsKnown(i) = UI_RuntimeTryGetBooleanProperty( _
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

                m_SnapshotGridlinesKnown(i) = UI_RuntimeTryGetBooleanProperty( _
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
        m_HasExcelUIStateSnapshot = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        UI_SnapshotCaptureCore = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UnexpectedDetail = UI_RuntimeBuildErrorText
        UI_SnapshotClear

        UI_RuntimeHandleFailure _
            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, "Unexpected", UnexpectedDetail

        Resume SafeExit

End Function

Public Function UI_SnapshotHasState() As Boolean

'
'==============================================================================
'                         UI_SnapshotHasState
'------------------------------------------------------------------------------
' PURPOSE
'   Return whether an explicit in-memory Excel UI snapshot is available.
'
' RETURNS
'   TRUE when a snapshot is available; otherwise FALSE.
'
' UPDATED
'   2026-07-29
'==============================================================================

        UI_SnapshotHasState = m_HasExcelUIStateSnapshot

End Function

Public Function UI_SnapshotRestoreCore( _
    ByVal ProcName As String, _
    ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) As Boolean

'
'==============================================================================
'                 UI_SnapshotRestoreCore
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
        UI_RuntimeClearResultBuffer _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList

        Succeeded = True

        On Error GoTo Fail

        If Not m_HasExcelUIStateSnapshot Then
            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "NoSnapshot", _
                "no captured Excel UI snapshot is available"

            GoTo SafeExit
        End If

        UI_RuntimeBeginQuietUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

'------------------------------------------------------------------------------
' RESTORE: TITLE BAR
'------------------------------------------------------------------------------
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
        If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
            Target:=Application, _
            PropertyName:="DisplayStatusBar", _
            NewValue:=m_SnapshotStatusBarVisible, _
            FailMsg:=Msg) Then

            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "StatusBar", Msg
        End If

        If Not UI_RuntimeTrySetBooleanPropertyIfNeeded( _
            Target:=Application, _
            PropertyName:="DisplayScrollBars", _
            NewValue:=m_SnapshotScrollBarsVisible, _
            FailMsg:=Msg) Then

            UI_RuntimeHandleFailure _
                ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                CaptureFailureList, "ScrollBars", Msg
        End If

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
        For i = 1 To m_SnapshotWindowCount
            Set MatchedWindow = Nothing

            If UI_SnapshotTryResolveWindow( _
                SnapshotIndex:=i, _
                WindowOut:=MatchedWindow, _
                FailMsg:=Msg) Then

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
                UI_RuntimeHandleFailure _
                    ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, _
                    "WindowIdentity [" & m_SnapshotWindowLabels(i) & "]", Msg
            End If
        Next i

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        UI_RuntimeEndQuietUpdate _
            OldScreenUpdating:=OldScreenUpdating, _
            QuietModeChanged:=QuietModeChanged

        UI_SnapshotRestoreCore = Succeeded
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        UnexpectedDetail = UI_RuntimeBuildErrorText

        UI_RuntimeHandleFailure _
            ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, "Unexpected", UnexpectedDetail

        Resume SafeExit

End Function

Public Sub UI_SnapshotClear()

'
'==============================================================================
'                           UI_SnapshotClear
'------------------------------------------------------------------------------
' PURPOSE
'   Remove all captured state and release retained Window object references.
'
' RETURNS
'   None.
'
' ERROR POLICY
'   Does not raise during normal operation.
'
' UPDATED
'   2026-07-29
'==============================================================================

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

Private Function UI_SnapshotTryResolveWindow( _
    ByVal SnapshotIndex As Long, _
    ByRef WindowOut As Object, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                      UI_SnapshotTryResolveWindow
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
'   - UI_RuntimeTryGetBooleanProperty
'   - UI_RuntimeBuildErrorText
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

        UI_SnapshotTryResolveWindow = False
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
        If Not UI_RuntimeTryGetBooleanProperty( _
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
        UI_SnapshotTryResolveWindow = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_RuntimeBuildErrorText
        Set WindowOut = Nothing

        Resume SafeExit

End Function

Private Function UI_SnapshotBuildWindowIdentityText(ByVal TargetWindow As Object) As String

'
'==============================================================================
'                       UI_SnapshotBuildWindowIdentityText
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
            UI_SnapshotBuildWindowIdentityText = WorkbookName & " :: " & WindowCaption
        ElseIf Len(WindowCaption) > 0 Then
            UI_SnapshotBuildWindowIdentityText = WindowCaption
        ElseIf Len(WorkbookName) > 0 Then
            UI_SnapshotBuildWindowIdentityText = WorkbookName
        Else
            UI_SnapshotBuildWindowIdentityText = "captured Excel window"
        End If

End Function
