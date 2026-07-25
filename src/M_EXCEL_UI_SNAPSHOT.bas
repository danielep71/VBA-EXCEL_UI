Attribute VB_Name = "M_EXCEL_UI_SNAPSHOT"
'==============================================================================
'                    MODULE: M_EXCEL_UI_SNAPSHOT
'------------------------------------------------------------------------------
' PURPOSE
'   Capture and restore the managed Excel UI using identity-safe Window matching
'
' WHY THIS EXISTS
'   Version 1.0.1 restored window state by collection index. Version 1.1.0
'   retains exact Window references and uses a guarded handle/caption fallback,
'   so activating or reordering windows does not redirect captured state
'
' ERROR POLICY
'   - Capture and restore are best-effort
'   - Partial operations remain usable and report deterministic failures
'
' UPDATED
'   2026-07-25
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.0
'==============================================================================
    Option Explicit
    Option Private Module

    Private m_HasSnapshot             As Boolean
    Private m_RibbonKnown             As Boolean
    Private m_RibbonVisible           As Boolean
    Private m_StatusBarVisible        As Boolean
    Private m_ScrollBarsVisible       As Boolean
    Private m_FormulaBarVisible       As Boolean
    Private m_TitleBarKnown           As Boolean
    Private m_TitleBarVisible         As Boolean
    Private m_WindowCount             As Long
    Private m_WindowRefs()            As Window
    Private m_WindowCaptions()        As String
#If VBA7 Then
    Private m_WindowHwnds()           As LongPtr
#Else
    Private m_WindowHwnds()           As Long
#End If
    Private m_WindowHwndKnown()       As Boolean
    Private m_HeadingsVisible()       As Boolean
    Private m_WorkbookTabsVisible()   As Boolean
    Private m_GridlinesVisible()      As Boolean

Public Function UI_SnapshotHasState() As Boolean
'==============================================================================
' PURPOSE
'   Return whether an in-memory snapshot is available
'==============================================================================
        UI_SnapshotHasState = m_HasSnapshot
End Function

Public Sub UI_SnapshotClear()
'==============================================================================
' PURPOSE
'   Remove all captured state from module memory
'==============================================================================
    Dim i As Long
        On Error Resume Next
        For i = 1 To m_WindowCount
            Set m_WindowRefs(i) = Nothing
        Next i
        Erase m_WindowRefs
        Erase m_WindowCaptions
        Erase m_WindowHwnds
        Erase m_WindowHwndKnown
        Erase m_HeadingsVisible
        Erase m_WorkbookTabsVisible
        Erase m_GridlinesVisible
        m_HasSnapshot = False
        m_RibbonKnown = False
        m_TitleBarKnown = False
        m_WindowCount = 0
End Sub

Public Function UI_SnapshotCapture(ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) As Boolean
'==============================================================================
' PURPOSE
'   Capture application, title-bar, Ribbon, and per-window managed UI state
'
' RETURNS
'   TRUE when every requested read succeeded; FALSE for a partial snapshot
'==============================================================================
    Dim Succeeded As Boolean
    Dim W As Window
    Dim i As Long
    Dim Msg As String
        On Error GoTo Fail
        UI_ResultClear FailureCount, FailureList, CaptureFailureList
        Succeeded = True
        UI_SnapshotClear

        m_StatusBarVisible = Application.DisplayStatusBar
        m_ScrollBarsVisible = Application.DisplayScrollBars
        m_FormulaBarVisible = Application.DisplayFormulaBar

        m_RibbonKnown = UI_HostTryGetRibbonVisible(m_RibbonVisible, Msg)
        If Not m_RibbonKnown Then
            UI_ResultHandleFailure "UI_CaptureExcelUIState", LogFailures, _
                Succeeded, FailureCount, FailureList, CaptureFailureList, _
                "RibbonSnapshot", Msg
        End If

        Msg = vbNullString
        m_TitleBarKnown = UI_TitleBarTryGetVisible(m_TitleBarVisible, Msg)
        If Not m_TitleBarKnown Then
            UI_ResultHandleFailure "UI_CaptureExcelUIState", LogFailures, _
                Succeeded, FailureCount, FailureList, CaptureFailureList, _
                "TitleBarSnapshot", Msg
        End If

        m_WindowCount = Application.Windows.Count
        If m_WindowCount > 0 Then
            ReDim m_WindowRefs(1 To m_WindowCount)
            ReDim m_WindowCaptions(1 To m_WindowCount)
            ReDim m_WindowHwnds(1 To m_WindowCount)
            ReDim m_WindowHwndKnown(1 To m_WindowCount)
            ReDim m_HeadingsVisible(1 To m_WindowCount)
            ReDim m_WorkbookTabsVisible(1 To m_WindowCount)
            ReDim m_GridlinesVisible(1 To m_WindowCount)

            i = 0
            For Each W In Application.Windows
                i = i + 1
                Set m_WindowRefs(i) = W
                m_WindowCaptions(i) = UI_HostWindowLabel(W)
                Msg = vbNullString
                m_WindowHwndKnown(i) = UI_HostTryGetWindowHwnd(W, _
                    m_WindowHwnds(i), Msg)
                If Not m_WindowHwndKnown(i) Then
                    UI_ResultHandleFailure "UI_CaptureExcelUIState", _
                        LogFailures, Succeeded, FailureCount, FailureList, _
                        CaptureFailureList, "WindowIdentity [" & _
                        m_WindowCaptions(i) & "]", Msg
                End If
                m_HeadingsVisible(i) = W.DisplayHeadings
                m_WorkbookTabsVisible(i) = W.DisplayWorkbookTabs
                m_GridlinesVisible(i) = W.DisplayGridlines
            Next W
        End If

        m_HasSnapshot = True
        UI_SnapshotCapture = Succeeded
        Exit Function
Fail:
        UI_ResultHandleFailure "UI_CaptureExcelUIState", LogFailures, _
            Succeeded, FailureCount, FailureList, CaptureFailureList, _
            "Unexpected", UI_ResultRuntimeErrorText
        m_HasSnapshot = False
        UI_SnapshotCapture = False
End Function

Public Function UI_SnapshotRestore(ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) As Boolean
'==============================================================================
' PURPOSE
'   Restore the latest snapshot by Window identity, never by collection index
'
' BEHAVIOR
'   - Exact retained object reference is preferred
'   - A unique handle plus matching captured caption is the fallback
'   - Missing captured windows and new current windows are reported
'   - New current windows are left unchanged
'==============================================================================
    Dim Succeeded As Boolean
    Dim OldScreenUpdating As Boolean
    Dim QuietModeChanged As Boolean
    Dim RestoredCurrent() As Boolean
    Dim CurrentWindowCount As Long
    Dim TargetWindow As Window
    Dim CurrentIndex As Long
    Dim i As Long
    Dim Msg As String
        On Error GoTo Fail
        UI_ResultClear FailureCount, FailureList, CaptureFailureList
        Succeeded = True

        If Not m_HasSnapshot Then
            UI_ResultHandleFailure "UI_ResetExcelUIToSnapshot", LogFailures, _
                Succeeded, FailureCount, FailureList, CaptureFailureList, _
                "NoSnapshot", "no captured Excel UI snapshot is available"
            UI_SnapshotRestore = False
            Exit Function
        End If

        CurrentWindowCount = UI_SnapshotSafeCurrentWindowCount
        If CurrentWindowCount > 0 Then
            ReDim RestoredCurrent(1 To CurrentWindowCount)
        End If
        UI_HostBeginQuietUpdate OldScreenUpdating, QuietModeChanged

        If m_TitleBarKnown Then
            If Not UI_TitleBarTrySetVisibleIfNeeded(m_TitleBarVisible, Msg) Then
                UI_ResultHandleFailure "UI_ResetExcelUIToSnapshot", _
                    LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "TitleBar", Msg
            End If
        End If

        Msg = vbNullString
        If m_RibbonKnown Then
            If Not UI_HostTrySetRibbonVisibleIfNeeded(m_RibbonVisible, Msg) Then
                UI_ResultHandleFailure "UI_ResetExcelUIToSnapshot", _
                    LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "Ribbon", Msg
            End If
        End If

        UI_SnapshotRestoreProperty Application, "DisplayStatusBar", _
            m_StatusBarVisible, "StatusBar", LogFailures, Succeeded, _
            FailureCount, FailureList, CaptureFailureList
        UI_SnapshotRestoreProperty Application, "DisplayScrollBars", _
            m_ScrollBarsVisible, "ScrollBars", LogFailures, Succeeded, _
            FailureCount, FailureList, CaptureFailureList
        UI_SnapshotRestoreProperty Application, "DisplayFormulaBar", _
            m_FormulaBarVisible, "FormulaBar", LogFailures, Succeeded, _
            FailureCount, FailureList, CaptureFailureList

        For i = 1 To m_WindowCount
            Set TargetWindow = Nothing
            CurrentIndex = 0
            If UI_SnapshotTryResolveWindow(i, TargetWindow, CurrentIndex, Msg) Then
                If CurrentIndex > 0 Then RestoredCurrent(CurrentIndex) = True
                UI_SnapshotRestoreProperty TargetWindow, "DisplayHeadings", _
                    m_HeadingsVisible(i), "Headings [" & _
                    m_WindowCaptions(i) & "]", LogFailures, Succeeded, _
                    FailureCount, FailureList, CaptureFailureList
                UI_SnapshotRestoreProperty TargetWindow, _
                    "DisplayWorkbookTabs", m_WorkbookTabsVisible(i), _
                    "WorkbookTabs [" & m_WindowCaptions(i) & "]", _
                    LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList
                UI_SnapshotRestoreProperty TargetWindow, "DisplayGridlines", _
                    m_GridlinesVisible(i), "Gridlines [" & _
                    m_WindowCaptions(i) & "]", LogFailures, Succeeded, _
                    FailureCount, FailureList, CaptureFailureList
            Else
                UI_ResultHandleFailure "UI_ResetExcelUIToSnapshot", _
                    LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "WindowMissing [" & _
                    m_WindowCaptions(i) & "]", Msg
            End If
        Next i

        UI_SnapshotReportNewWindows RestoredCurrent, CurrentWindowCount, _
            LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList

SafeExit:
        UI_HostEndQuietUpdate OldScreenUpdating, QuietModeChanged
        UI_SnapshotRestore = Succeeded
        Exit Function
Fail:
        UI_ResultHandleFailure "UI_ResetExcelUIToSnapshot", LogFailures, _
            Succeeded, FailureCount, FailureList, CaptureFailureList, _
            "Unexpected", UI_ResultRuntimeErrorText
        Resume SafeExit
End Function

Private Sub UI_SnapshotRestoreProperty(ByVal Target As Object, _
    ByVal PropertyName As String, ByVal SavedValue As Boolean, _
    ByVal Stage As String, ByVal LogFailures As Boolean, _
    ByRef Succeeded As Boolean, ByRef FailureCount As Long, _
    ByRef FailureList As Variant, ByVal CaptureFailureList As Boolean)
'==============================================================================
' PURPOSE
'   Restore one Boolean property and record a deterministic failure
'==============================================================================
    Dim Msg As String
        If Not UI_HostTrySetBooleanPropertyIfNeeded(Target, PropertyName, _
            SavedValue, Msg) Then
            UI_ResultHandleFailure "UI_ResetExcelUIToSnapshot", LogFailures, _
                Succeeded, FailureCount, FailureList, CaptureFailureList, _
                Stage, Msg
        End If
End Sub

Private Function UI_SnapshotTryResolveWindow(ByVal SnapshotIndex As Long, _
    ByRef TargetWindow As Window, ByRef CurrentIndex As Long, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Resolve one captured window without collection-index fallback
'==============================================================================
    Dim W As Window
    Dim Candidate As Window
    Dim i As Long
    Dim MatchCount As Long
#If VBA7 Then
    Dim H As LongPtr
#Else
    Dim H As Long
#End If
    Dim Msg As String
        On Error GoTo Fallback

        If Not m_WindowRefs(SnapshotIndex) Is Nothing Then
            If UI_HostIsCurrentExcelWindow(m_WindowRefs(SnapshotIndex), Msg) Then
                Set TargetWindow = m_WindowRefs(SnapshotIndex)
                CurrentIndex = UI_SnapshotCurrentIndex(TargetWindow)
                UI_SnapshotTryResolveWindow = True
                Exit Function
            End If
        End If

Fallback:
        Err.Clear
        On Error GoTo Fail
        If Not m_WindowHwndKnown(SnapshotIndex) Then
            FailMsg = "captured Window object no longer exists and no handle was captured"
            Exit Function
        End If

        i = 0
        For Each W In Application.Windows
            i = i + 1
            Msg = vbNullString
            If UI_HostTryGetWindowHwnd(W, H, Msg) Then
                If H = m_WindowHwnds(SnapshotIndex) Then
                    If UI_HostWindowLabel(W) = m_WindowCaptions(SnapshotIndex) Then
                        MatchCount = MatchCount + 1
                        Set Candidate = W
                        CurrentIndex = i
                    End If
                End If
            End If
        Next W

        If MatchCount = 1 Then
            Set TargetWindow = Candidate
            UI_SnapshotTryResolveWindow = True
        ElseIf MatchCount = 0 Then
            FailMsg = "no current window matches the captured identity"
        Else
            FailMsg = "captured identity is ambiguous across current windows"
        End If
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function

Private Function UI_SnapshotCurrentIndex(ByVal TargetWindow As Window) As Long
'==============================================================================
' PURPOSE
'   Return the current collection index for reporting only
'==============================================================================
    Dim W As Window
    Dim i As Long
        On Error Resume Next
        For Each W In Application.Windows
            i = i + 1
            If W Is TargetWindow Then
                UI_SnapshotCurrentIndex = i
                Exit Function
            End If
        Next W
End Function

Private Function UI_SnapshotSafeCurrentWindowCount() As Long
'==============================================================================
' PURPOSE
'   Read Application.Windows.Count without propagating host errors
'==============================================================================
        On Error Resume Next
        UI_SnapshotSafeCurrentWindowCount = Application.Windows.Count
End Function

Private Sub UI_SnapshotReportNewWindows(ByRef RestoredCurrent() As Boolean, _
    ByVal CurrentWindowCount As Long, ByVal LogFailures As Boolean, _
    ByRef Succeeded As Boolean, ByRef FailureCount As Long, _
    ByRef FailureList As Variant, ByVal CaptureFailureList As Boolean)
'==============================================================================
' PURPOSE
'   Report current windows that were not part of the captured snapshot
'==============================================================================
    Dim W As Window
    Dim i As Long
        On Error GoTo Fail
        For Each W In Application.Windows
            i = i + 1
            If CurrentWindowCount = 0 Then
                UI_ResultHandleFailure "UI_ResetExcelUIToSnapshot", _
                    LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "WindowAdded [" & _
                    UI_HostWindowLabel(W) & "]", _
                    "window was opened after capture and was left unchanged"
            ElseIf Not RestoredCurrent(i) Then
                UI_ResultHandleFailure "UI_ResetExcelUIToSnapshot", _
                    LogFailures, Succeeded, FailureCount, FailureList, _
                    CaptureFailureList, "WindowAdded [" & _
                    UI_HostWindowLabel(W) & "]", _
                    "window was opened after capture and was left unchanged"
            End If
        Next W
        Exit Sub
Fail:
        UI_ResultHandleFailure "UI_ResetExcelUIToSnapshot", LogFailures, _
            Succeeded, FailureCount, FailureList, CaptureFailureList, _
            "WindowAdded", UI_ResultRuntimeErrorText
End Sub
