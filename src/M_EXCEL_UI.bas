Attribute VB_Name = "M_EXCEL_UI"
Option Explicit
Option Private Module

'
'==============================================================================
'                           MODULE: M_EXCEL_UI
'------------------------------------------------------------------------------
' PURPOSE
'   Centralize visibility control for the Excel UI elements managed by this
'   module, combining:
'     - Excel object-model UI elements
'     - WinAPI-based title-bar control for the Excel main window represented
'       by Application.Hwnd
'
' WHY THIS EXISTS
'   Some workbook-driven solutions need to present Excel in a constrained,
'   kiosk-like, or application-style shell.
'
'   Excel exposes several UI elements directly through the object model
'   (Ribbon, status bar, scroll bars, formula bar, headings, workbook tabs,
'   gridlines), but it does not expose direct title-bar visibility control.
'
'   This module unifies both approaches behind a safe, explicit API so callers
'   do not need to duplicate scattered UI-handling code.
'
' PUBLIC SURFACE
'   - K_UIVisibility                Tri-state visibility enum
'   - K_SetExcelUI                  Core selective UI-state routine
'   - K_SetExcelUI_WithResult       Selective routine with structured result
'   - K_HideExcelUI                 Convenience wrapper: hide all managed UI
'   - K_ShowExcelUI                 Convenience wrapper: show all managed UI
'   - K_CaptureExcelUIState         Explicitly snapshot the current managed UI
'   - K_ResetExcelUIToSnapshot      Best-effort restore to captured UI state
'   - K_HasExcelUIStateSnapshot     Return TRUE when a snapshot exists
'   - K_ClearExcelUIStateSnapshot   Remove any captured snapshot
'
' INTERNAL SUPPORT
'   - K_ApplyExcelUIState
'   - K_HandleApplyFailure
'   - K_ClearResultBuffer
'   - K_AddFailureToResult
'   - K_BeginQuietUIUpdate
'   - K_EndQuietUIUpdate
'   - K_TrySetRibbonVisibleIfNeeded
'   - K_TrySetTitleBarVisibleIfNeeded
'   - K_TrySetBooleanPropertyIfNeeded
'   - K_TryGetRibbonVisible
'   - K_TryGetTitleBarVisible
'   - K_TryGetBooleanProperty
'   - K_TrySetTitleBarVisible
'   - K_TrySetRibbonVisible
'   - K_TrySetBooleanProperty
'   - K_TryGetWindowStyle
'   - K_TrySetWindowStyle
'   - K_TryRefreshWindowFrame
'   - K_VisibilityToBoolean
'   - K_BuildRuntimeErrorText
'   - K_LogFailure
'   - WinAPI declarations / constants
'
' BEHAVIOR
'   - Application-level elements:
'       * Ribbon
'       * Status Bar
'       * Scroll Bars
'       * Formula Bar
'
'   - Window-level elements (applied to each open Excel window):
'       * Headings
'       * Workbook Tabs
'       * Gridlines
'
'   - Title bar:
'       * applied to the Excel main window represented by Application.Hwnd
'         through WinAPI style update + non-client frame refresh
'
' ERROR POLICY
'   - Public entry points are fail-soft.
'   - Unexpected errors are logged to the Immediate Window in the fire-and-
'     forget path.
'   - Errors are not re-raised to callers.
'   - The core routine uses best-effort application:
'       * one failed UI element does not prevent later UI elements from being
'         attempted
'
' PLATFORM / COMPATIBILITY
'   - Windows only.
'   - Supports 32-bit and 64-bit Office / VBA through conditional compilation
'     and bitness-safe WinAPI wrappers.
'
' NOTES
'   - This module does NOT automatically snapshot and restore prior Excel
'     object-model UI state.
'   - K_ShowExcelUI means "show all managed UI", not "restore previous state".
'   - K_SetExcelUI is the preferred entry point for selective control.
'   - K_SetExcelUI_WithResult offers the same best-effort behavior while
'     returning structured diagnostics without a class-module dependency.
'   - Ribbon control relies on Application.ExecuteExcel4Macro.
'   - Title-bar control affects the Excel window represented by
'     Application.Hwnd, not a user-specific saved UI state.
'   - The original main-window style is snapshotted against the current window
'     handle, so the restore path can follow Application.Hwnd safely if the
'     main Excel window is recreated.
'   - The explicit snapshot / reset feature is separate from K_ShowExcelUI and
'     is best-effort for per-window restore.
'
' UPDATED
'   2026-04-11
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.0.0
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE: PUBLIC ENUMS
'------------------------------------------------------------------------------
Public Enum K_UIVisibility
    K_UI_LeaveUnchanged = -1     'Do not touch this UI element
    K_UI_Hide = 0                'Hide this UI element
    K_UI_Show = 1                'Show this UI element
End Enum

'------------------------------------------------------------------------------
' DECLARE: WIN32 / WIN64 API
'------------------------------------------------------------------------------
#If VBA7 Then

    #If Win64 Then

        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
            ByVal hWnd As LongPtr, _
            ByVal nIndex As Long) _
            As LongPtr

        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
            ByVal hWnd As LongPtr, _
            ByVal nIndex As Long, _
            ByVal dwNewLong As LongPtr) _
            As LongPtr

    #Else

        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
            ByVal hWnd As LongPtr, _
            ByVal nIndex As Long) _
            As Long

        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
            ByVal hWnd As LongPtr, _
            ByVal nIndex As Long, _
            ByVal dwNewLong As Long) _
            As Long

    #End If

    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As LongPtr, _
        ByVal hWndInsertAfter As LongPtr, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal uFlags As Long) _
        As Long

    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

    Private Declare PtrSafe Sub SetLastError Lib "kernel32" ( _
        ByVal dwErrCode As Long)

#Else

    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long) _
        As Long

    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) _
        As Long

    Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal uFlags As Long) _
        As Long

    Private Declare Function GetLastError Lib "kernel32" () As Long

    Private Declare Sub SetLastError Lib "kernel32" ( _
        ByVal dwErrCode As Long)

#End If

'------------------------------------------------------------------------------
' DECLARE: PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const GWL_STYLE          As Long = -16       'Window style index

    Private Const WS_CAPTION         As Long = &HC00000  'Caption / title bar
    Private Const WS_SYSMENU         As Long = &H80000   'System menu
    Private Const WS_THICKFRAME      As Long = &H40000   'Resizable sizing frame
    Private Const WS_MINIMIZEBOX     As Long = &H20000   'Minimize button
    Private Const WS_MAXIMIZEBOX     As Long = &H10000   'Maximize button

    Private Const SWP_NOSIZE         As Long = &H1       'Preserve current size
    Private Const SWP_NOMOVE         As Long = &H2       'Preserve current position
    Private Const SWP_NOZORDER       As Long = &H4       'Do not change Z order
    Private Const SWP_FRAMECHANGED   As Long = &H20      'Repaint non-client frame
    Private Const SWP_NOOWNERZORDER  As Long = &H200     'Do not change owner Z order

'------------------------------------------------------------------------------
' DECLARE: PRIVATE MODULE STATE
'------------------------------------------------------------------------------
#If VBA7 Then
    Private m_OriginalMainWindowStyle As LongPtr      'Snapshotted original Excel main-window style
    Private m_OriginalMainWindowHwnd  As LongPtr      'Window handle associated with the snapshotted style
#Else
    Private m_OriginalMainWindowStyle As Long         'Snapshotted original Excel main-window style
    Private m_OriginalMainWindowHwnd  As Long         'Window handle associated with the snapshotted style
#End If

Private m_HasOriginalMainWindowStyle As Boolean       'TRUE when original style has been captured

Private m_HasExcelUIStateSnapshot    As Boolean       'TRUE when an explicit snapshot exists

Private m_SnapshotRibbonKnown        As Boolean       'TRUE when Ribbon state was captured successfully
Private m_SnapshotRibbonVisible      As Boolean       'Captured Ribbon visibility
Private m_SnapshotStatusBarVisible   As Boolean       'Captured StatusBar visibility
Private m_SnapshotScrollBarsVisible  As Boolean       'Captured ScrollBars visibility
Private m_SnapshotFormulaBarVisible  As Boolean       'Captured FormulaBar visibility

Private m_SnapshotWindowCount        As Long          'Captured Application.Windows.Count
Private m_SnapshotHeadingsVisible()  As Boolean       'Captured per-window Headings visibility
Private m_SnapshotWorkbookTabsVisible() As Boolean    'Captured per-window WorkbookTabs visibility
Private m_SnapshotGridlinesVisible() As Boolean       'Captured per-window Gridlines visibility

Private m_SnapshotTitleBarKnown      As Boolean       'TRUE when TitleBar state was captured successfully
Private m_SnapshotTitleBarVisible    As Boolean       'Captured TitleBar visibility

Public Sub K_SetExcelUI( _
    Optional ByVal Ribbon As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal StatusBar As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal ScrollBars As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal FormulaBar As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal Headings As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal WorkbookTabs As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal Gridlines As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal TitleBar As K_UIVisibility = K_UI_LeaveUnchanged)

'
'==============================================================================
'                               K_SetExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the requested visibility state to the Excel UI elements managed by
'   this module.
'
' WHY THIS EXISTS
'   A Boolean-based "hide/show" routine is error-prone because omitted optional
'   arguments can accidentally imply FALSE / hidden.
'
'   This routine uses an explicit tri-state API:
'     - K_UI_Show
'     - K_UI_Hide
'     - K_UI_LeaveUnchanged
'
'   This makes the caller's intent precise and prevents accidental UI changes
'   for omitted arguments.
'
' INPUTS
'   Ribbon (optional)
'     K_UI_Show             => show Ribbon
'     K_UI_Hide             => hide Ribbon
'     K_UI_LeaveUnchanged   => do not touch Ribbon
'
'   StatusBar (optional)
'     K_UI_Show             => show status bar
'     K_UI_Hide             => hide status bar
'     K_UI_LeaveUnchanged   => do not touch status bar
'
'   ScrollBars (optional)
'     K_UI_Show             => show scroll bars
'     K_UI_Hide             => hide scroll bars
'     K_UI_LeaveUnchanged   => do not touch scroll bars
'
'   FormulaBar (optional)
'     K_UI_Show             => show formula bar
'     K_UI_Hide             => hide formula bar
'     K_UI_LeaveUnchanged   => do not touch formula bar
'
'   Headings (optional)
'     K_UI_Show             => show row / column headings in each window
'     K_UI_Hide             => hide row / column headings in each window
'     K_UI_LeaveUnchanged   => do not touch headings
'
'   WorkbookTabs (optional)
'     K_UI_Show             => show workbook tabs in each window
'     K_UI_Hide             => hide workbook tabs in each window
'     K_UI_LeaveUnchanged   => do not touch workbook tabs
'
'   Gridlines (optional)
'     K_UI_Show             => show gridlines in each window
'     K_UI_Hide             => hide gridlines in each window
'     K_UI_LeaveUnchanged   => do not touch gridlines
'
'   TitleBar (optional)
'     K_UI_Show             => show the title bar of the Excel main window
'                               represented by Application.Hwnd
'     K_UI_Hide             => hide the title bar of the Excel main window
'                               represented by Application.Hwnd
'     K_UI_LeaveUnchanged   => do not touch title bar
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Applies Ribbon / status bar / scroll bars / formula bar at Application
'     level.
'   - Applies headings / workbook tabs / gridlines to every open Excel window
'     in the current Excel instance.
'   - Applies title-bar visibility to the Excel main window represented by
'     Application.Hwnd via WinAPI.
'   - Uses best-effort processing: one failed UI element does not prevent
'     subsequent UI elements from being attempted.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Unexpected failures are written to the Immediate Window.
'   - Element-level failures are logged and processing continues.
'
' DEPENDENCIES
'   - K_ApplyExcelUIState
'
' NOTES
'   - This is the preferred entry point for selective UI control.
'   - Changes affect the current Excel instance, not only the active workbook.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FailureCount        As Long       'Internal ignored failure count
    Dim FailureList         As Variant    'Internal ignored failure list

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY STATE THROUGH INTERNAL WORKER
'------------------------------------------------------------------------------
    'Delegate the full best-effort application flow to the shared worker,
    'requesting Immediate Window logging for any failures.
        Call K_ApplyExcelUIState( _
                ProcName:="K_SetExcelUI", _
                Ribbon:=Ribbon, _
                StatusBar:=StatusBar, _
                ScrollBars:=ScrollBars, _
                FormulaBar:=FormulaBar, _
                Headings:=Headings, _
                WorkbookTabs:=WorkbookTabs, _
                Gridlines:=Gridlines, _
                TitleBar:=TitleBar, _
                LogFailures:=True, _
                FailureCount:=FailureCount, _
                FailureList:=FailureList, _
                CaptureFailureList:=False)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave quietly through the normal termination path.
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write an unexpected-procedure-level diagnostic line without interrupting
    'the caller.
        K_LogFailure "K_SetExcelUI", "Unexpected", K_BuildRuntimeErrorText

    'Exit quietly after logging.
        Resume SafeExit

End Sub

Public Sub K_HideExcelUI()

'
'==============================================================================
'                               K_HideExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Hide all Excel UI elements managed by this module.
'
' WHY THIS EXISTS
'   Some workbook-driven solutions want a simple one-call way to suppress the
'   managed Excel shell elements without specifying each element individually.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Delegates to K_SetExcelUI.
'   - Requests hidden state for:
'       * Ribbon
'       * Status Bar
'       * Scroll Bars
'       * Formula Bar
'       * Headings
'       * Workbook Tabs
'       * Gridlines
'       * Title Bar
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Unexpected failures are written to the Immediate Window.
'
' DEPENDENCIES
'   - K_SetExcelUI
'
' NOTES
'   - This is a convenience wrapper.
'   - For selective control, use K_SetExcelUI directly.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY: HIDE-ALL STATE
'------------------------------------------------------------------------------
    'Hide all managed UI elements through the central tri-state entry point.
        K_SetExcelUI _
            Ribbon:=K_UI_Hide, _
            StatusBar:=K_UI_Hide, _
            ScrollBars:=K_UI_Hide, _
            FormulaBar:=K_UI_Hide, _
            Headings:=K_UI_Hide, _
            WorkbookTabs:=K_UI_Hide, _
            Gridlines:=K_UI_Hide, _
            TitleBar:=K_UI_Hide

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave quietly through the normal termination path.
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write an unexpected-procedure-level diagnostic line without interrupting
    'the caller.
        K_LogFailure "K_HideExcelUI", "Unexpected", K_BuildRuntimeErrorText

    'Exit quietly after logging.
        Resume SafeExit

End Sub

Public Sub K_ShowExcelUI()

'
'==============================================================================
'                               K_ShowExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Show all Excel UI elements managed by this module.
'
' WHY THIS EXISTS
'   Workbook solutions that temporarily suppress the Excel shell often need a
'   single, deterministic call to restore all managed UI elements to visible
'   state.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Delegates to K_SetExcelUI.
'   - Requests visible state for:
'       * Ribbon
'       * Status Bar
'       * Scroll Bars
'       * Formula Bar
'       * Headings
'       * Workbook Tabs
'       * Gridlines
'       * Title Bar
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Unexpected failures are written to the Immediate Window.
'
' DEPENDENCIES
'   - K_SetExcelUI
'
' NOTES
'   - This means "show all managed UI".
'   - It does NOT restore a previously captured user-specific object-model UI
'     state.
'   - For selective control, use K_SetExcelUI directly.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY: SHOW-ALL STATE
'------------------------------------------------------------------------------
    'Show all managed UI elements through the central tri-state entry point.
        K_SetExcelUI _
            Ribbon:=K_UI_Show, _
            StatusBar:=K_UI_Show, _
            ScrollBars:=K_UI_Show, _
            FormulaBar:=K_UI_Show, _
            Headings:=K_UI_Show, _
            WorkbookTabs:=K_UI_Show, _
            Gridlines:=K_UI_Show, _
            TitleBar:=K_UI_Show

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave quietly through the normal termination path.
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write an unexpected-procedure-level diagnostic line without interrupting
    'the caller.
        K_LogFailure "K_ShowExcelUI", "Unexpected", K_BuildRuntimeErrorText

    'Exit quietly after logging.
        Resume SafeExit

End Sub

Public Function K_SetExcelUI_WithResult( _
    Optional ByVal Ribbon As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal StatusBar As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal ScrollBars As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal FormulaBar As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal Headings As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal WorkbookTabs As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal Gridlines As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByVal TitleBar As K_UIVisibility = K_UI_LeaveUnchanged, _
    Optional ByRef FailureCount As Long = 0, _
    Optional ByRef FailureList As Variant) As Boolean

'
'==============================================================================
'                         K_SetExcelUI_WithResult
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the requested visibility state to the Excel UI elements managed by
'   this module and return a Boolean success flag, with optional structured
'   failure details captured through ByRef outputs.
'
' WHY THIS EXISTS
'   K_SetExcelUI is the preferred fire-and-forget, fail-soft entry point for
'   callers that only need best-effort application plus Immediate Window
'   diagnostics.
'
'   Some callers, however, need structured feedback so they can:
'     - inspect whether the full operation succeeded
'     - count element-level failures
'     - enumerate the recorded failures in order
'     - surface diagnostics to higher-level orchestration or test logic
'
'   This routine provides the same best-effort behavior as K_SetExcelUI, but
'   avoids any class-module dependency by returning:
'     - a Boolean success flag
'     - FailureCount as an optional ByRef output
'     - FailureList as an optional ByRef Variant containing a 1-based String
'       array of recorded failures
'
' INPUTS
'   Ribbon (optional)
'     K_UI_Show             => show Ribbon
'     K_UI_Hide             => hide Ribbon
'     K_UI_LeaveUnchanged   => do not touch Ribbon
'
'   StatusBar (optional)
'     K_UI_Show             => show status bar
'     K_UI_Hide             => hide status bar
'     K_UI_LeaveUnchanged   => do not touch status bar
'
'   ScrollBars (optional)
'     K_UI_Show             => show scroll bars
'     K_UI_Hide             => hide scroll bars
'     K_UI_LeaveUnchanged   => do not touch scroll bars
'
'   FormulaBar (optional)
'     K_UI_Show             => show formula bar
'     K_UI_Hide             => hide formula bar
'     K_UI_LeaveUnchanged   => do not touch formula bar
'
'   Headings (optional)
'     K_UI_Show             => show row / column headings in each window
'     K_UI_Hide             => hide row / column headings in each window
'     K_UI_LeaveUnchanged   => do not touch headings
'
'   WorkbookTabs (optional)
'     K_UI_Show             => show workbook tabs in each window
'     K_UI_Hide             => hide workbook tabs in each window
'     K_UI_LeaveUnchanged   => do not touch workbook tabs
'
'   Gridlines (optional)
'     K_UI_Show             => show gridlines in each window
'     K_UI_Hide             => hide gridlines in each window
'     K_UI_LeaveUnchanged   => do not touch gridlines
'
'   TitleBar (optional)
'     K_UI_Show             => show the title bar of the Excel main window
'                               represented by Application.Hwnd
'     K_UI_Hide             => hide the title bar of the Excel main window
'                               represented by Application.Hwnd
'     K_UI_LeaveUnchanged   => do not touch title bar
'
'   FailureCount (optional, ByRef output)
'     Receives the number of recorded failures.
'
'   FailureList (optional, ByRef output)
'     When supplied, receives a 1-based String array whose entries follow the
'     format:
'         Stage & " | " & Detail
'
' RETURNS
'   TRUE  => no failures were recorded
'   FALSE => one or more failures were recorded
'
' BEHAVIOR
'   - Applies Ribbon / status bar / scroll bars / formula bar at Application
'     level.
'   - Applies headings / workbook tabs / gridlines to every open Excel window
'     in the current Excel instance.
'   - Applies title-bar visibility to the Excel main window represented by
'     Application.Hwnd via WinAPI.
'   - Uses best-effort processing: one failed UI element does not prevent
'     subsequent UI elements from being attempted.
'   - Records failures through FailureCount and, when requested, FailureList.
'
' ERROR POLICY
'   - Does NOT raise to callers for ordinary element-level failures.
'   - Returns FALSE when one or more failures were recorded.
'   - Unexpected procedure-level failures are captured as an "Unexpected"
'     failure entry and also produce a FALSE result.
'
' DEPENDENCIES
'   - K_ApplyExcelUIState
'
' NOTES
'   - This routine mirrors the best-effort semantics of K_SetExcelUI.
'   - Failure order is preserved.
'   - FailureList remains optional so callers that only need the Boolean result
'     or failure count do not need to manage an array.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean    'Overall success flag returned to the caller
    Dim CaptureFailureList  As Boolean    'TRUE when the caller supplied FailureList
    Dim InternalFailureList As Variant    'Local working failure list copied back only when requested

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Detect whether the caller supplied the optional failure-list output.
        CaptureFailureList = Not IsMissing(FailureList)

    'Initialize the public result outputs in their clean-success state.
        K_ClearResultBuffer FailureCount, InternalFailureList, CaptureFailureList
        Succeeded = True

    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY STATE THROUGH INTERNAL WORKER
'------------------------------------------------------------------------------
    'Delegate the full best-effort application flow to the shared worker,
    'requesting structured failure capture rather than Immediate Window logging.
        Succeeded = K_ApplyExcelUIState( _
                        ProcName:="K_SetExcelUI_WithResult", _
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
' COPY OPTIONAL FAILURE LIST
'------------------------------------------------------------------------------
    'Copy the internal working list back to the caller only when the optional
    'failure-list output was actually supplied.
        If CaptureFailureList Then
            FailureList = InternalFailureList
        End If

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return the overall success flag to the caller.
        K_SetExcelUI_WithResult = Succeeded

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Capture the unexpected wrapper-level failure in the structured result
    'buffers.
        K_HandleApplyFailure _
            ProcName:="K_SetExcelUI_WithResult", _
            LogFailures:=False, _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", _
            Detail:=K_BuildRuntimeErrorText

    'Copy the internal working list back to the caller only when the optional
    'failure-list output was actually supplied.
        If CaptureFailureList Then
            FailureList = InternalFailureList
        End If

    'Return the overall success flag after recording the unexpected failure.
        K_SetExcelUI_WithResult = Succeeded

    'Leave quietly through the normal termination path.
        Resume SafeExit

End Function

Public Sub K_CaptureExcelUIState()

'
'==============================================================================
'                           K_CaptureExcelUIState
'------------------------------------------------------------------------------
' PURPOSE
'   Explicitly snapshot the current managed Excel UI state so it can later be
'   restored through K_ResetExcelUIToSnapshot.
'
' WHY THIS EXISTS
'   K_ShowExcelUI intentionally means "show all managed UI", not "restore the
'   user's prior state".
'
'   Some callers need a distinct, deliberate lifecycle:
'     - capture current state
'     - apply a constrained shell
'     - restore the captured state later
'
'   This routine provides that explicit capture step.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Captures application-level state.
'   - Captures per-window state by index.
'   - Captures title-bar state on a best-effort basis.
'   - Marks the snapshot as available even when Ribbon / TitleBar state could
'     only be captured best-effort.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Logs any unexpected issue to the Immediate Window.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long      'Current window index during snapshot
    Dim Msg                 As String    'Diagnostic message from helper reads

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

'------------------------------------------------------------------------------
' CLEAR PRIOR SNAPSHOT
'------------------------------------------------------------------------------
    'Clear any prior snapshot before capturing a fresh one.
        K_ClearExcelUIStateSnapshot

'------------------------------------------------------------------------------
' CAPTURE APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Capture application-level UI state directly from Excel.
        m_SnapshotStatusBarVisible = Application.DisplayStatusBar
        m_SnapshotScrollBarsVisible = Application.DisplayScrollBars
        m_SnapshotFormulaBarVisible = Application.DisplayFormulaBar

'------------------------------------------------------------------------------
' CAPTURE RIBBON / TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Capture Ribbon state through the best-effort helper.
        m_SnapshotRibbonKnown = K_TryGetRibbonVisible(m_SnapshotRibbonVisible, Msg)
        If Not m_SnapshotRibbonKnown Then
            K_LogFailure "K_CaptureExcelUIState", "Ribbon", Msg
        End If

    'Capture TitleBar state through the best-effort helper.
        m_SnapshotTitleBarKnown = K_TryGetTitleBarVisible(m_SnapshotTitleBarVisible, Msg)
        If Not m_SnapshotTitleBarKnown Then
            K_LogFailure "K_CaptureExcelUIState", "TitleBar", Msg
        End If

'------------------------------------------------------------------------------
' CAPTURE WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Capture the current window count.
        m_SnapshotWindowCount = Application.Windows.Count

    'Allocate and fill per-window arrays only when at least one window exists.
        If m_SnapshotWindowCount > 0 Then

            'Allocate the headings array.
                ReDim m_SnapshotHeadingsVisible(1 To m_SnapshotWindowCount)

            'Allocate the workbook-tabs array.
                ReDim m_SnapshotWorkbookTabsVisible(1 To m_SnapshotWindowCount)

            'Allocate the gridlines array.
                ReDim m_SnapshotGridlinesVisible(1 To m_SnapshotWindowCount)

            'Capture the state of each current Excel window by index.
                For i = 1 To m_SnapshotWindowCount

                    'Capture Headings visibility.
                        m_SnapshotHeadingsVisible(i) = Application.Windows(i).DisplayHeadings

                    'Capture WorkbookTabs visibility.
                        m_SnapshotWorkbookTabsVisible(i) = Application.Windows(i).DisplayWorkbookTabs

                    'Capture Gridlines visibility.
                        m_SnapshotGridlinesVisible(i) = Application.Windows(i).DisplayGridlines

                Next i

        End If

'------------------------------------------------------------------------------
' MARK SNAPSHOT AVAILABLE
'------------------------------------------------------------------------------
    'Mark the snapshot as available after capture completes.
        m_HasExcelUIStateSnapshot = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point.
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Log the unexpected capture failure without interrupting callers.
        K_LogFailure "K_CaptureExcelUIState", "Unexpected", K_BuildRuntimeErrorText

    'Leave quietly after logging.
        Resume SafeExit

End Sub

Public Function K_HasExcelUIStateSnapshot() As Boolean

'
'==============================================================================
'                        K_HasExcelUIStateSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Return whether an explicit Excel UI snapshot is currently available.
'
' WHY THIS EXISTS
'   Callers may want to check whether reset-to-snapshot is meaningful before
'   attempting it.
'
' RETURNS
'   TRUE  => a snapshot is available
'   FALSE => no snapshot is available
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN SNAPSHOT AVAILABILITY
'------------------------------------------------------------------------------
    'Return whether a captured UI snapshot is currently available.
        K_HasExcelUIStateSnapshot = m_HasExcelUIStateSnapshot

End Function

Public Sub K_ResetExcelUIToSnapshot()

'
'==============================================================================
'                        K_ResetExcelUIToSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Best-effort restore the Excel UI to the most recently captured explicit
'   snapshot.
'
' WHY THIS EXISTS
'   Callers that previously used K_CaptureExcelUIState may need a distinct way
'   to restore the captured baseline rather than merely showing all managed UI.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Restores title bar and Ribbon when their snapshot states were known.
'   - Restores application-level object-model properties.
'   - Restores per-window state by common index range.
'   - Uses a quiet-update scope with ScreenUpdating where possible.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Logs any restore issue to the Immediate Window.
'   - Best-effort only, especially for per-window restore.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long      'Current window index during restore
    Dim WindowLimit         As Long      'Minimum of saved and current window counts
    Dim Msg                 As String    'Diagnostic message from helper routines
    Dim OldScreenUpdating   As Boolean   'Cached ScreenUpdating state
    Dim QuietModeChanged    As Boolean   'TRUE when ScreenUpdating was changed

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Do nothing when no explicit snapshot is available.
        If Not m_HasExcelUIStateSnapshot Then
            K_LogFailure "K_ResetExcelUIToSnapshot", "NoSnapshot", "no captured Excel UI snapshot is available"
            GoTo SafeExit
        End If

    'Enter the quiet-update scope to reduce worksheet redraw where possible.
        K_BeginQuietUIUpdate OldScreenUpdating, QuietModeChanged

'------------------------------------------------------------------------------
' RESTORE TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Restore TitleBar first when its snapshot state was captured successfully.
        If m_SnapshotTitleBarKnown Then
            If Not K_TrySetTitleBarVisibleIfNeeded(m_SnapshotTitleBarVisible, Msg) Then
                K_LogFailure "K_ResetExcelUIToSnapshot", "TitleBar", Msg
            End If
        End If

'------------------------------------------------------------------------------
' RESTORE RIBBON STATE
'------------------------------------------------------------------------------
    'Restore Ribbon when its snapshot state was captured successfully.
        If m_SnapshotRibbonKnown Then
            If Not K_TrySetRibbonVisibleIfNeeded(m_SnapshotRibbonVisible, Msg) Then
                K_LogFailure "K_ResetExcelUIToSnapshot", "Ribbon", Msg
            End If
        End If

'------------------------------------------------------------------------------
' RESTORE APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Restore StatusBar visibility best-effort.
        If Not K_TrySetBooleanPropertyIfNeeded(Application, "DisplayStatusBar", m_SnapshotStatusBarVisible, Msg) Then
            K_LogFailure "K_ResetExcelUIToSnapshot", "StatusBar", Msg
        End If

    'Restore ScrollBars visibility best-effort.
        If Not K_TrySetBooleanPropertyIfNeeded(Application, "DisplayScrollBars", m_SnapshotScrollBarsVisible, Msg) Then
            K_LogFailure "K_ResetExcelUIToSnapshot", "ScrollBars", Msg
        End If

    'Restore FormulaBar visibility best-effort.
        If Not K_TrySetBooleanPropertyIfNeeded(Application, "DisplayFormulaBar", m_SnapshotFormulaBarVisible, Msg) Then
            K_LogFailure "K_ResetExcelUIToSnapshot", "FormulaBar", Msg
        End If

'------------------------------------------------------------------------------
' RESTORE WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Restore only the common indexed window range that still exists.
        WindowLimit = Application.Windows.Count
        If m_SnapshotWindowCount < WindowLimit Then WindowLimit = m_SnapshotWindowCount

    'Restore each saved window state up to the common window count.
        For i = 1 To WindowLimit

            'Restore Headings visibility for the current saved window index.
                If Not K_TrySetBooleanPropertyIfNeeded(Application.Windows(i), "DisplayHeadings", m_SnapshotHeadingsVisible(i), Msg) Then
                    K_LogFailure "K_ResetExcelUIToSnapshot", "Headings [" & Application.Windows(i).Caption & "]", Msg
                End If

            'Restore WorkbookTabs visibility for the current saved window index.
                If Not K_TrySetBooleanPropertyIfNeeded(Application.Windows(i), "DisplayWorkbookTabs", m_SnapshotWorkbookTabsVisible(i), Msg) Then
                    K_LogFailure "K_ResetExcelUIToSnapshot", "WorkbookTabs [" & Application.Windows(i).Caption & "]", Msg
                End If

            'Restore Gridlines visibility for the current saved window index.
                If Not K_TrySetBooleanPropertyIfNeeded(Application.Windows(i), "DisplayGridlines", m_SnapshotGridlinesVisible(i), Msg) Then
                    K_LogFailure "K_ResetExcelUIToSnapshot", "Gridlines [" & Application.Windows(i).Caption & "]", Msg
                End If

        Next i

    'Log a note when the current window count differs from the captured count.
        If Application.Windows.Count <> m_SnapshotWindowCount Then
            K_LogFailure "K_ResetExcelUIToSnapshot", "WindowCount", _
                "current window count differs from captured snapshot; restore applied to common index range only"
        End If

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave the quiet-update scope and restore ScreenUpdating when needed.
        K_EndQuietUIUpdate OldScreenUpdating, QuietModeChanged

    'Normal termination point.
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Log the unexpected restore failure without interrupting callers.
        K_LogFailure "K_ResetExcelUIToSnapshot", "Unexpected", K_BuildRuntimeErrorText

    'Leave the quiet-update scope and restore ScreenUpdating when needed.
        K_EndQuietUIUpdate OldScreenUpdating, QuietModeChanged

    'Leave quietly after logging.
        Resume SafeExit

End Sub

Public Sub K_ClearExcelUIStateSnapshot()

'
'==============================================================================
'                      K_ClearExcelUIStateSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Remove any captured explicit Excel UI snapshot from module state.
'
' WHY THIS EXISTS
'   Callers may want to discard an obsolete snapshot explicitly before taking a
'   new one or before leaving a workflow.
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise to callers.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET SNAPSHOT FLAGS
'------------------------------------------------------------------------------
    'Mark the explicit UI snapshot as unavailable.
        m_HasExcelUIStateSnapshot = False

    'Reset best-effort known flags.
        m_SnapshotRibbonKnown = False
        m_SnapshotTitleBarKnown = False

'------------------------------------------------------------------------------
' RESET SNAPSHOT VALUES
'------------------------------------------------------------------------------
    'Reset application-level values.
        m_SnapshotRibbonVisible = False
        m_SnapshotStatusBarVisible = False
        m_SnapshotScrollBarsVisible = False
        m_SnapshotFormulaBarVisible = False

    'Reset title-bar value.
        m_SnapshotTitleBarVisible = False

    'Reset captured window count.
        m_SnapshotWindowCount = 0

'------------------------------------------------------------------------------
' CLEAR SNAPSHOT ARRAYS
'------------------------------------------------------------------------------
    'Clear any captured per-window arrays.
        Erase m_SnapshotHeadingsVisible
        Erase m_SnapshotWorkbookTabsVisible
        Erase m_SnapshotGridlinesVisible

End Sub

Private Function K_ApplyExcelUIState( _
    ByVal ProcName As String, _
    ByVal Ribbon As K_UIVisibility, _
    ByVal StatusBar As K_UIVisibility, _
    ByVal ScrollBars As K_UIVisibility, _
    ByVal FormulaBar As K_UIVisibility, _
    ByVal Headings As K_UIVisibility, _
    ByVal WorkbookTabs As K_UIVisibility, _
    ByVal Gridlines As K_UIVisibility, _
    ByVal TitleBar As K_UIVisibility, _
    ByVal LogFailures As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean) As Boolean

'
'==============================================================================
'                           K_ApplyExcelUIState
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the requested UI state once through a single internal worker shared
'   by both public entry points.
'
' WHY THIS EXISTS
'   The module exposes two public application paths:
'     - K_SetExcelUI
'     - K_SetExcelUI_WithResult
'
'   They are intentionally different only in how they surface failures:
'     - logging to the Immediate Window
'     - structured result buffers
'
'   Centralizing the actual UI-application logic here eliminates duplicated
'   orchestration and reduces the risk of future behavioral drift.
'
' INPUTS
'   ProcName
'     Public caller name used for failure diagnostics.
'
'   Ribbon / StatusBar / ScrollBars / FormulaBar / Headings / WorkbookTabs /
'   Gridlines / TitleBar
'     Requested tri-state UI modes.
'
'   LogFailures
'     TRUE  => write failures to the Immediate Window
'     FALSE => suppress Immediate Window logging and use only the result
'              buffers
'
'   FailureCount
'     Receives the number of recorded failures.
'
'   FailureList
'     Optional working Variant holding a 1-based String array of failures.
'
'   CaptureFailureList
'     TRUE  => populate FailureList
'     FALSE => maintain only success flag and FailureCount
'
' RETURNS
'   TRUE  => no failures were recorded
'   FALSE => one or more failures were recorded
'
' BEHAVIOR
'   - Initializes the result buffers.
'   - Applies all requested UI changes using best-effort processing.
'   - Uses ScreenUpdating suppression where possible to reduce worksheet
'     redraw flicker.
'   - Skips object-model / Ribbon / TitleBar writes when the current state can
'     be read and is already equal to the requested target.
'   - Records and optionally logs failures in insertion order.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Captures unexpected procedure-level failures as "Unexpected".
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean    'Overall success flag returned by the worker
    Dim W                   As Window     'Workbook window in current Excel instance
    Dim ShowFlag            As Boolean    'Converted Boolean visibility target
    Dim Msg                 As String     'Element-level diagnostic message

    Dim DoHeadings          As Boolean    'TRUE when headings were requested
    Dim DoWorkbookTabs      As Boolean    'TRUE when workbook tabs were requested
    Dim DoGridlines         As Boolean    'TRUE when gridlines were requested

    Dim ShowHeadings        As Boolean    'Converted headings target
    Dim ShowWorkbookTabs    As Boolean    'Converted workbook-tabs target
    Dim ShowGridlines       As Boolean    'Converted gridlines target

    Dim OldScreenUpdating   As Boolean    'Cached ScreenUpdating state
    Dim QuietModeChanged    As Boolean    'TRUE when ScreenUpdating was changed

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Initialize the result buffers in their clean-success state.
        K_ClearResultBuffer FailureCount, FailureList, CaptureFailureList
        Succeeded = True

    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Enter the quiet-update scope to reduce worksheet redraw where possible.
        K_BeginQuietUIUpdate OldScreenUpdating, QuietModeChanged

'------------------------------------------------------------------------------
' APPLY: APPLICATION-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Apply Ribbon visibility when requested.
        If Ribbon <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(Ribbon)

            'Attempt the Ribbon update only when needed and record any failure
            'without interrupting later operations.
                If Not K_TrySetRibbonVisibleIfNeeded(ShowFlag, Msg) Then
                    K_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                        CaptureFailureList, "Ribbon", Msg
                End If

        End If

    'Apply status-bar visibility when requested.
        If StatusBar <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(StatusBar)

            'Attempt the property write only when needed and record any failure
            'without interrupting later operations.
                If Not K_TrySetBooleanPropertyIfNeeded(Application, "DisplayStatusBar", ShowFlag, Msg) Then
                    K_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                        CaptureFailureList, "StatusBar", Msg
                End If

        End If

    'Apply scroll-bar visibility when requested.
        If ScrollBars <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(ScrollBars)

            'Attempt the property write only when needed and record any failure
            'without interrupting later operations.
                If Not K_TrySetBooleanPropertyIfNeeded(Application, "DisplayScrollBars", ShowFlag, Msg) Then
                    K_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                        CaptureFailureList, "ScrollBars", Msg
                End If

        End If

    'Apply formula-bar visibility when requested.
        If FormulaBar <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(FormulaBar)

            'Attempt the property write only when needed and record any failure
            'without interrupting later operations.
                If Not K_TrySetBooleanPropertyIfNeeded(Application, "DisplayFormulaBar", ShowFlag, Msg) Then
                    K_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                        CaptureFailureList, "FormulaBar", Msg
                End If

        End If

'------------------------------------------------------------------------------
' PRECOMPUTE: WINDOW-LEVEL REQUESTS
'------------------------------------------------------------------------------
    'Precompute whether each window-level property was requested.
        DoHeadings = (Headings <> K_UI_LeaveUnchanged)
        DoWorkbookTabs = (WorkbookTabs <> K_UI_LeaveUnchanged)
        DoGridlines = (Gridlines <> K_UI_LeaveUnchanged)

    'Precompute the Boolean targets only for requested properties.
        If DoHeadings Then ShowHeadings = K_VisibilityToBoolean(Headings)
        If DoWorkbookTabs Then ShowWorkbookTabs = K_VisibilityToBoolean(WorkbookTabs)
        If DoGridlines Then ShowGridlines = K_VisibilityToBoolean(Gridlines)

'------------------------------------------------------------------------------
' APPLY: WINDOW-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Process window-scoped UI only when at least one window-level element has
    'been requested for change.
        If DoHeadings Or DoWorkbookTabs Or DoGridlines Then

            'Apply the requested window-level visibility state to each open
            'Excel window in the current instance.
                For Each W In Application.Windows

                    'Apply headings visibility when requested.
                        If DoHeadings Then

                            'Attempt the property write only when needed and
                            'record any failure without interrupting later
                            'operations.
                                If Not K_TrySetBooleanPropertyIfNeeded(W, "DisplayHeadings", ShowHeadings, Msg) Then
                                    K_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                                        CaptureFailureList, "Headings [" & W.Caption & "]", Msg
                                End If

                        End If

                    'Apply workbook-tabs visibility when requested.
                        If DoWorkbookTabs Then

                            'Attempt the property write only when needed and
                            'record any failure without interrupting later
                            'operations.
                                If Not K_TrySetBooleanPropertyIfNeeded(W, "DisplayWorkbookTabs", ShowWorkbookTabs, Msg) Then
                                    K_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                                        CaptureFailureList, "WorkbookTabs [" & W.Caption & "]", Msg
                                End If

                        End If

                    'Apply gridlines visibility when requested.
                        If DoGridlines Then

                            'Attempt the property write only when needed and
                            'record any failure without interrupting later
                            'operations.
                                If Not K_TrySetBooleanPropertyIfNeeded(W, "DisplayGridlines", ShowGridlines, Msg) Then
                                    K_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                                        CaptureFailureList, "Gridlines [" & W.Caption & "]", Msg
                                End If

                        End If

                Next W

        End If

'------------------------------------------------------------------------------
' APPLY: TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Apply title-bar visibility when requested.
        If TitleBar <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(TitleBar)

            'Attempt the title-bar update only when needed and record any
            'failure without interrupting later operations.
                If Not K_TrySetTitleBarVisibleIfNeeded(ShowFlag, Msg) Then
                    K_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
                        CaptureFailureList, "TitleBar", Msg
                End If

        End If

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return the overall success flag to the caller.
        K_ApplyExcelUIState = Succeeded

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave the quiet-update scope and restore ScreenUpdating when needed.
        K_EndQuietUIUpdate OldScreenUpdating, QuietModeChanged

    'Normal termination point.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Capture the unexpected worker-level failure in the result buffers and
    'optionally log it.
        K_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, FailureList, _
            CaptureFailureList, "Unexpected", K_BuildRuntimeErrorText

    'Return the overall success flag after recording the unexpected failure.
        K_ApplyExcelUIState = Succeeded

    'Leave the quiet-update scope and restore ScreenUpdating when needed.
        K_EndQuietUIUpdate OldScreenUpdating, QuietModeChanged

    'Leave quietly through the normal termination path.
        Resume SafeExit

End Function

Private Sub K_HandleApplyFailure( _
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
'                           K_HandleApplyFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Handle one recorded failure consistently for the shared internal worker.
'
' WHY THIS EXISTS
'   The shared worker must support two public failure-surfacing modes:
'     - logging to the Immediate Window
'     - structured result capture through standard-module outputs
'
'   This helper centralizes both actions so the worker logic stays compact and
'   consistent.
'
' INPUTS
'   ProcName
'     Public caller name used for logging.
'
'   LogFailures
'     TRUE  => write the failure to the Immediate Window
'     FALSE => suppress logging
'
'   Succeeded / FailureCount / FailureList / CaptureFailureList
'     Result buffers used by the shared worker.
'
'   Stage
'     Logical stage, element, or operation associated with the failure.
'
'   Detail
'     Failure detail to append.
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
    'Record the failure into the structured result buffers.
        K_AddFailureToResult Succeeded, FailureCount, FailureList, CaptureFailureList, Stage, Detail

'------------------------------------------------------------------------------
' OPTIONAL LOGGING
'------------------------------------------------------------------------------
    'Write the failure to the Immediate Window only when requested by the
'   caller path.
        If LogFailures Then
            K_LogFailure ProcName, Stage, Detail
        End If

End Sub

Private Sub K_ClearResultBuffer( _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean)

'
'==============================================================================
'                           K_ClearResultBuffer
'------------------------------------------------------------------------------
' PURPOSE
'   Initialize the ByRef result buffers used by K_SetExcelUI_WithResult and the
'   shared internal worker.
'
' WHY THIS EXISTS
'   The standard-module result pattern used by K_SetExcelUI_WithResult needs a
'   consistent way to reset:
'     - FailureCount
'     - FailureList
'
'   Centralizing that initialization avoids duplicated setup logic and ensures
'   the function always starts from a clean result state.
'
' INPUTS
'   FailureCount
'     Receives zero on initialization.
'
'   FailureList
'     Receives Empty when the caller requested the optional list output.
'
'   CaptureFailureList
'     TRUE  => initialize FailureList
'     FALSE => leave FailureList untouched because the caller did not request it
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Reset the recorded failure count to zero.
        FailureCount = 0

    'Initialize the failure-list output only when the caller requested it.
        If CaptureFailureList Then
            FailureList = Empty
        End If

End Sub

Private Sub K_AddFailureToResult( _
    ByRef Succeeded As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean, _
    ByVal Stage As String, _
    ByVal Detail As String)

'
'==============================================================================
'                          K_AddFailureToResult
'------------------------------------------------------------------------------
' PURPOSE
'   Record a failure into the standard-module result buffers used by
'   K_SetExcelUI_WithResult.
'
' WHY THIS EXISTS
'   The module no longer depends on a dedicated result class, so failures need
'   to be accumulated through plain standard-module constructs:
'     - a Boolean success flag
'     - a Long failure count
'     - an optional 1-based String array of failure entries
'
'   This helper centralizes that logic and preserves insertion order.
'
' INPUTS
'   Succeeded
'     Set to FALSE once a failure is recorded.
'
'   FailureCount
'     Incremented for each recorded failure.
'
'   FailureList
'     Optional Variant carrying a 1-based String array of recorded failures.
'
'   CaptureFailureList
'     TRUE  => append the failure text into FailureList
'     FALSE => update only Succeeded and FailureCount
'
'   Stage
'     Logical stage, element, or operation associated with the failure.
'
'   Detail
'     Failure detail to append.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Sets Succeeded to FALSE.
'   - Increments FailureCount.
'   - When requested, appends:
'         Stage & " | " & Detail
'     to a 1-based String array stored in FailureList.
'
' ERROR POLICY
'   - Does NOT raise.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FailureText         As String    'Formatted failure entry
    Dim Arr()               As String    'Working 1-based failure array

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Build the formatted failure text once.
        FailureText = Stage & " | " & Detail

'------------------------------------------------------------------------------
' UPDATE RESULT STATUS
'------------------------------------------------------------------------------
    'Mark the overall result as unsuccessful.
        Succeeded = False

    'Increment the recorded failure count.
        FailureCount = FailureCount + 1

'------------------------------------------------------------------------------
' APPEND FAILURE TEXT
'------------------------------------------------------------------------------
    'Append the formatted failure text only when the caller requested the
    'failure-list output.
        If CaptureFailureList Then

            'Allocate the first entry when the failure list is still empty.
                If IsEmpty(FailureList) Then
                    ReDim Arr(1 To 1)

            'Otherwise, expand the existing 1-based array while preserving
            'previous entries.
                Else
                    Arr = FailureList
                    ReDim Preserve Arr(1 To FailureCount)
                End If

            'Store the new failure entry at the current 1-based position.
                Arr(FailureCount) = FailureText

            'Write the expanded array back into the Variant output.
                FailureList = Arr

        End If

End Sub

Private Sub K_BeginQuietUIUpdate( _
    ByRef OldScreenUpdating As Boolean, _
    ByRef QuietModeChanged As Boolean)

'
'==============================================================================
'                          K_BeginQuietUIUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Enter a small best-effort quiet-update scope by suppressing
'   Application.ScreenUpdating when possible.
'
' WHY THIS EXISTS
'   Many object-model UI writes can cause visible worksheet redraw.
'   Temporarily disabling ScreenUpdating reduces flicker for those surfaces,
'   even though it cannot fully suppress Ribbon or WinAPI non-client refresh.
'
' INPUTS / OUTPUTS
'   OldScreenUpdating
'     Receives the prior Application.ScreenUpdating state.
'
'   QuietModeChanged
'     Receives TRUE only when this helper actually changed ScreenUpdating.
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise.
'   - Best-effort only.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected issue while entering quiet mode.
        On Error Resume Next

    'Capture the current ScreenUpdating state.
        OldScreenUpdating = Application.ScreenUpdating

    'Initialize the changed flag.
        QuietModeChanged = False

'------------------------------------------------------------------------------
' APPLY QUIET MODE
'------------------------------------------------------------------------------
    'Disable ScreenUpdating only when it is currently enabled.
        If OldScreenUpdating Then
            Application.ScreenUpdating = False
            QuietModeChanged = True
        End If

End Sub

Private Sub K_EndQuietUIUpdate( _
    ByVal OldScreenUpdating As Boolean, _
    ByVal QuietModeChanged As Boolean)

'
'==============================================================================
'                           K_EndQuietUIUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Leave the quiet-update scope created by K_BeginQuietUIUpdate.
'
' WHY THIS EXISTS
'   ScreenUpdating should be restored only when this module actually changed it.
'
' INPUTS
'   OldScreenUpdating
'     Previously captured Application.ScreenUpdating state.
'
'   QuietModeChanged
'     TRUE when K_BeginQuietUIUpdate changed ScreenUpdating.
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise.
'   - Best-effort only.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected issue while leaving quiet mode.
        On Error Resume Next

'------------------------------------------------------------------------------
' RESTORE PRIOR STATE
'------------------------------------------------------------------------------
    'Restore ScreenUpdating only when this module previously changed it.
        If QuietModeChanged Then
            Application.ScreenUpdating = OldScreenUpdating
        End If

End Sub

Private Function K_TrySetRibbonVisibleIfNeeded( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                     K_TrySetRibbonVisibleIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to update Ribbon visibility only when the current visible state
'   differs from the requested target.
'
' WHY THIS EXISTS
'   Avoiding no-op Ribbon writes can slightly reduce visible UI churn and keeps
'   the apply path cleaner.
'
' INPUTS
'   IsVisible
'     Requested Ribbon visibility.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => Ribbon is already in the requested state or was updated
'   FALSE => Ribbon update failed
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentVisible      As Boolean   'Current Ribbon visibility when readable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize the default result state.
        K_TrySetRibbonVisibleIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT NO-OP
'------------------------------------------------------------------------------
    'When current Ribbon visibility can be read and already matches the target,
    'skip the write path entirely.
        If K_TryGetRibbonVisible(CurrentVisible, FailMsg) Then
            If CurrentVisible = IsVisible Then
                K_TrySetRibbonVisibleIfNeeded = True
                GoTo SafeExit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY RIBBON WRITE
'------------------------------------------------------------------------------
    'Clear any prior read diagnostic and attempt the actual write.
        FailMsg = vbNullString
        K_TrySetRibbonVisibleIfNeeded = K_TrySetRibbonVisible(IsVisible, FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

Private Function K_TrySetTitleBarVisibleIfNeeded( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                    K_TrySetTitleBarVisibleIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to update TitleBar visibility only when the current visible state
'   differs from the requested target.
'
' WHY THIS EXISTS
'   Avoiding no-op title-bar writes reduces unnecessary non-client frame
'   refresh attempts.
'
' INPUTS
'   IsVisible
'     Requested TitleBar visibility.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => TitleBar is already in the requested state or was updated
'   FALSE => TitleBar update failed
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentVisible      As Boolean   'Current TitleBar visibility when readable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize the default result state.
        K_TrySetTitleBarVisibleIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT NO-OP
'------------------------------------------------------------------------------
    'When current TitleBar visibility can be read and already matches the
    'target, skip the write path entirely.
        If K_TryGetTitleBarVisible(CurrentVisible, FailMsg) Then
            If CurrentVisible = IsVisible Then
                K_TrySetTitleBarVisibleIfNeeded = True
                GoTo SafeExit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY TITLE-BAR WRITE
'------------------------------------------------------------------------------
    'Clear any prior read diagnostic and attempt the actual write.
        FailMsg = vbNullString
        K_TrySetTitleBarVisibleIfNeeded = K_TrySetTitleBarVisible(IsVisible, FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

Private Function K_TrySetBooleanPropertyIfNeeded( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByVal NewValue As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                   K_TrySetBooleanPropertyIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to assign a Boolean property only when the current state differs
'   from the requested target.
'
' WHY THIS EXISTS
'   Avoiding no-op property writes reduces unnecessary redraw and keeps the UI
'   application path quieter when the property already matches the target.
'
' INPUTS
'   Target
'     Object exposing the target Boolean property.
'
'   PropertyName
'     Name of the Boolean property to assign.
'
'   NewValue
'     Requested Boolean value.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => property is already in the requested state or was updated
'   FALSE => property update failed
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentValue        As Boolean   'Current property value when readable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize the default result state.
        K_TrySetBooleanPropertyIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT NO-OP
'------------------------------------------------------------------------------
    'When current property state can be read and already matches the target,
    'skip the write path entirely.
        If K_TryGetBooleanProperty(Target, PropertyName, CurrentValue, FailMsg) Then
            If CurrentValue = NewValue Then
                K_TrySetBooleanPropertyIfNeeded = True
                GoTo SafeExit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY PROPERTY WRITE
'------------------------------------------------------------------------------
    'Clear any prior read diagnostic and attempt the actual write.
        FailMsg = vbNullString
        K_TrySetBooleanPropertyIfNeeded = K_TrySetBooleanProperty(Target, PropertyName, NewValue, FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

Private Function K_TryGetRibbonVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                         K_TryGetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read current Ribbon visibility.
'
' WHY THIS EXISTS
'   The Ribbon is not best treated as a simple direct property read, so the
'   module uses a dedicated best-effort reader.
'
' INPUTS
'   IsVisible
'     Receives current Ribbon visibility on success.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => Ribbon visibility was read successfully
'   FALSE => read failed
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant   'Fallback Excel4 macro result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize outputs and default result.
        K_TryGetRibbonVisible = False
        IsVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' TRY COMMANDBARS
'------------------------------------------------------------------------------
    'Attempt to read Ribbon visibility from the CommandBars collection.
        On Error Resume Next
            IsVisible = Application.CommandBars("Ribbon").Visible
        If Err.Number = 0 Then
            On Error GoTo Fail
            K_TryGetRibbonVisible = True
            GoTo SafeExit
        End If
        Err.Clear
        On Error GoTo Fail

'------------------------------------------------------------------------------
' TRY EXCEL4 MACRO FALLBACK
'------------------------------------------------------------------------------
    'Attempt a fallback read using an Excel4 macro.
        On Error Resume Next
            V = Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")
        If Err.Number = 0 Then
            On Error GoTo Fail
            IsVisible = CBool(V)
            K_TryGetRibbonVisible = True
            GoTo SafeExit
        End If
        FailMsg = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        On Error GoTo Fail

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

Private Function K_TryGetTitleBarVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                        K_TryGetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read current title-bar visibility for the Excel window
'   represented by Application.Hwnd.
'
' WHY THIS EXISTS
'   Title-bar state is controlled through WinAPI in this module, so the module
'   also uses a corresponding WinAPI-based read helper.
'
' INPUTS
'   IsVisible
'     Receives current title-bar visibility on success.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => title-bar visibility was read successfully
'   FALSE => read failed
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim xlHnd               As LongPtr   'Excel window handle from Application.Hwnd
    Dim StyleValue          As LongPtr   'Current window style value
#Else
    Dim xlHnd               As Long      'Excel window handle from Application.Hwnd
    Dim StyleValue          As Long      'Current window style value
#End If
    Dim LastErr             As Long      'Last Win32 error after API call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize outputs and default result.
        K_TryGetTitleBarVisible = False
        IsVisible = False
        FailMsg = vbNullString

    'Read the Excel window handle.
        xlHnd = Application.hWnd

    'Reject invalid window handle deterministically.
        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ WINDOW STYLE
'------------------------------------------------------------------------------
    'Clear last-error state before the API call.
        SetLastError 0

#If VBA7 Then
    #If Win64 Then

        'Read the current window style using the 64-bit API.
            StyleValue = GetWindowLongPtr(xlHnd, GWL_STYLE)

    #Else

        'Read the current window style using the 32-bit API under VBA7.
            StyleValue = GetWindowLong(xlHnd, GWL_STYLE)

    #End If
#Else

    'Read the current window style using the legacy 32-bit API.
        StyleValue = GetWindowLong(xlHnd, GWL_STYLE)

#End If

    'Read the Win32 last-error value immediately after the API call.
        LastErr = GetLastError

    'Treat zero + nonzero last error as failure.
        If StyleValue = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' MAP STYLE TO TITLE-BAR VISIBILITY
'------------------------------------------------------------------------------
    'Treat the caption style bit as the title-bar visibility signal.
        IsVisible = ((StyleValue And WS_CAPTION) <> 0)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after a valid style read.
        K_TryGetTitleBarVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

Private Function K_TryGetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByRef ValueOut As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                         K_TryGetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read a Boolean property from an object using CallByName.
'
' WHY THIS EXISTS
'   The module now needs read helpers both for:
'     - skip-if-already-correct write suppression
'     - explicit capture / reset behavior
'
'   A shared property reader avoids duplicated boilerplate.
'
' INPUTS
'   Target
'     Object exposing the target Boolean property.
'
'   PropertyName
'     Name of the Boolean property to read.
'
'   ValueOut
'     Receives the property value on success.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => property read succeeded
'   FALSE => property read failed
'
' ERROR POLICY
'   - Does NOT raise to callers.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant   'Late-bound property value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize outputs and default result.
        K_TryGetBooleanProperty = False
        ValueOut = False
        FailMsg = vbNullString

    'Reject invalid object input deterministically.
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo SafeExit
        End If

    'Reject empty property name deterministically.
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ PROPERTY
'------------------------------------------------------------------------------
    'Read the requested property using late-bound property access.
        V = CallByName(Target, PropertyName, VbGet)

    'Convert the result to a Boolean.
        ValueOut = CBool(V)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after property access completes.
        K_TryGetBooleanProperty = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure string without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

#If VBA7 Then
Private Function K_TrySetTitleBarVisible( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function K_TrySetTitleBarVisible( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
'                           K_TrySetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to show or hide the title bar of the Excel main window represented
'   by Application.Hwnd by updating the window style and refreshing the
'   non-client frame.
'
' WHY THIS EXISTS
'   Excel does not expose direct title-bar visibility control in the object
'   model, so the project must update the underlying window style via WinAPI.
'
'   To improve the visual result when hiding, this routine also removes the
'   sizing frame (WS_THICKFRAME). The original window style is snapshotted and
'   restored exactly when showing again.
'
' INPUTS
'   IsVisible
'     TRUE  => restore the original snapshotted Excel main-window style
'     FALSE => hide title bar / frame-related controls
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => title-bar update succeeded
'   FALSE => title-bar update failed
'
' BEHAVIOR
'   - Reads the current main-window style.
'   - Snapshots the original style against the current Application.Hwnd.
'   - Restores the exact original style when showing again.
'   - Refreshes the non-client frame after a style change.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Returns FALSE and populates FailMsg on failure.
'
' NOTES
'   - Windows only.
'   - While hidden, the Excel window is intentionally less frame-like and may
'     not be user-resizable in the normal way.
'   - This routine intentionally does NOT toggle
'     Application.DisplayFullScreen.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim xlHnd               As LongPtr   'Excel main-window handle from Application.Hwnd
    Dim CurrentStyle        As LongPtr   'Current main-window style
    Dim NewStyle            As LongPtr   'Updated main-window style
#Else
    Dim xlHnd               As Long      'Excel main-window handle from Application.Hwnd
    Dim CurrentStyle        As Long      'Current main-window style
    Dim NewStyle            As Long      'Updated main-window style
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize the default result state.
        K_TrySetTitleBarVisible = False
        FailMsg = vbNullString

    'Read the Excel main-window handle.
        xlHnd = Application.hWnd

    'Reject an invalid window handle deterministically.
        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ: CURRENT WINDOW STYLE
'------------------------------------------------------------------------------
    'Read the current main-window style using the bitness-safe wrapper.
        If Not K_TryGetWindowStyle(xlHnd, CurrentStyle, FailMsg) Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' SNAPSHOT: ORIGINAL STYLE FOR CURRENT HWND
'------------------------------------------------------------------------------
    'Snapshot the original main-window style whenever no snapshot exists yet or
    'the current Application.Hwnd differs from the one previously captured.
        If (Not m_HasOriginalMainWindowStyle) Or (m_OriginalMainWindowHwnd <> xlHnd) Then
            m_OriginalMainWindowStyle = CurrentStyle
            m_OriginalMainWindowHwnd = xlHnd
            m_HasOriginalMainWindowStyle = True
        End If

'------------------------------------------------------------------------------
' COMPUTE: UPDATED WINDOW STYLE
'------------------------------------------------------------------------------
    'Restore the exact original snapshotted style when the caller requests a
    'visible title bar.
        If IsVisible Then

            'Use the exact captured original style whenever it belongs to the
            'current window handle.
                If m_HasOriginalMainWindowStyle And m_OriginalMainWindowHwnd = xlHnd Then
                    NewStyle = m_OriginalMainWindowStyle

            'Fall back to a conservative "visible frame" composition only if
            'the original style is not available for the current handle.
                Else
                    NewStyle = CurrentStyle
                    NewStyle = NewStyle Or WS_SYSMENU
                    NewStyle = NewStyle Or WS_MINIMIZEBOX
                    NewStyle = NewStyle Or WS_MAXIMIZEBOX
                    NewStyle = NewStyle Or WS_CAPTION
                    NewStyle = NewStyle Or WS_THICKFRAME
                End If

    'Remove the frame-related bits when the caller requests a hidden title bar.
        Else
            NewStyle = CurrentStyle
            NewStyle = NewStyle And Not WS_SYSMENU
            NewStyle = NewStyle And Not WS_MINIMIZEBOX
            NewStyle = NewStyle And Not WS_MAXIMIZEBOX
            NewStyle = NewStyle And Not WS_CAPTION
            NewStyle = NewStyle And Not WS_THICKFRAME
        End If

'------------------------------------------------------------------------------
' SHORT-CIRCUIT: NO-OP
'------------------------------------------------------------------------------
    'Skip the write path entirely when no style change is required.
        If NewStyle = CurrentStyle Then
            K_TrySetTitleBarVisible = True
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' APPLY: UPDATED WINDOW STYLE
'------------------------------------------------------------------------------
    'Write the updated main-window style using the bitness-safe wrapper.
        If Not K_TrySetWindowStyle(xlHnd, NewStyle, FailMsg) Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' REFRESH: NON-CLIENT FRAME
'------------------------------------------------------------------------------
    'Force Windows to recalculate and repaint the frame after the style change.
        If Not K_TryRefreshWindowFrame(xlHnd, FailMsg) Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful only after all required steps complete.
        K_TrySetTitleBarVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

Private Function K_TrySetRibbonVisible( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                           K_TrySetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to show or hide the Ribbon using Excel4 macro execution.
'
' WHY THIS EXISTS
'   The Ribbon is not exposed through a simple Application Boolean property, so
'   a legacy but commonly used Excel4 macro call is required for direct control.
'
' INPUTS
'   IsVisible
'     TRUE  => show Ribbon
'     FALSE => hide Ribbon
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => Ribbon update succeeded
'   FALSE => Ribbon update failed
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Returns FALSE and populates FailMsg on failure.
'
' DEPENDENCIES
'   - Application.ExecuteExcel4Macro
'
' NOTES
'   - Availability may vary by Excel host / configuration.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MacroText           As String    'Excel4 macro text controlling Ribbon visibility

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize the default result state.
        K_TrySetRibbonVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' BUILD: EXCEL4 MACRO TEXT
'------------------------------------------------------------------------------
    'Build the exact macro text required to show or hide the Ribbon.
        If IsVisible Then
            MacroText = "Show.TOOLBAR(""Ribbon"",True)"
        Else
            MacroText = "Show.TOOLBAR(""Ribbon"",False)"
        End If

'------------------------------------------------------------------------------
' APPLY: RIBBON VISIBILITY
'------------------------------------------------------------------------------
    'Execute the Ribbon visibility macro through Excel's legacy macro engine.
        Application.ExecuteExcel4Macro MacroText

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful after macro execution completes.
        K_TrySetRibbonVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

Private Function K_TrySetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByVal NewValue As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                           K_TrySetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to assign a Boolean property on an object using a common,
'   best-effort helper.
'
' WHY THIS EXISTS
'   K_SetExcelUI sets several Boolean properties across different object types
'   (Application and Window). A shared helper avoids duplicating identical
'   property-write error-handling logic for each target property.
'
' INPUTS
'   Target
'     Object exposing the target Boolean property.
'
'   PropertyName
'     Name of the Boolean property to assign.
'
'   NewValue
'     Boolean value to write to the target property.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => property write succeeded
'   FALSE => property write failed
'
' BEHAVIOR
'   - Uses CallByName with vbLet to assign the property.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Returns FALSE and populates FailMsg on failure.
'
' NOTES
'   - Intended for Application / Window Boolean property writes in this module.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize the default result state.
        K_TrySetBooleanProperty = False
        FailMsg = vbNullString

    'Reject a missing target object deterministically.
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo SafeExit
        End If

    'Reject an empty property name deterministically.
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' APPLY: PROPERTY WRITE
'------------------------------------------------------------------------------
    'Assign the requested Boolean value using late-bound property assignment.
        CallByName Target, PropertyName, VbLet, NewValue

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful after the property write completes.
        K_TrySetBooleanProperty = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

#If VBA7 Then
Private Function K_TryGetWindowStyle( _
    ByVal hWnd As LongPtr, _
    ByRef StyleOut As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function K_TryGetWindowStyle( _
    ByVal hWnd As Long, _
    ByRef StyleOut As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
'                            K_TryGetWindowStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read the current GWL_STYLE value using the correct API for the
'   current VBA / Office bitness.
'
' WHY THIS EXISTS
'   GetWindowLong / GetWindowLongPtr can validly return zero, so a robust
'   wrapper should use GetLastError to distinguish "real zero" from failure.
'
' INPUTS
'   hWnd
'     Target window handle.
'
'   StyleOut
'     Receives the current GWL_STYLE value on success.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => style read succeeded
'   FALSE => style read failed
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Returns FALSE and populates FailMsg on failure.
'
' DEPENDENCIES
'   - GetWindowLong / GetWindowLongPtr
'   - GetLastError
'   - SetLastError
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastErr             As Long      'Win32 last-error value read after the API call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize the outputs and default result state.
        StyleOut = 0
        FailMsg = vbNullString
        K_TryGetWindowStyle = False

    'Reject an invalid window handle deterministically.
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ: WINDOW STYLE
'------------------------------------------------------------------------------
    'Clear the Win32 last-error state before calling the API so a valid zero
    'return can later be distinguished from failure.
        SetLastError 0

#If VBA7 Then
    #If Win64 Then

        'Read the style with the 64-bit API in 64-bit Office / VBA.
            StyleOut = GetWindowLongPtr(hWnd, GWL_STYLE)

    #Else

        'Read the style with the 32-bit API in VBA7 32-bit Office.
            StyleOut = GetWindowLong(hWnd, GWL_STYLE)

    #End If
#Else

    'Read the style with the legacy 32-bit API.
        StyleOut = GetWindowLong(hWnd, GWL_STYLE)

#End If

    'Read the Win32 last-error value immediately after the API call.
        LastErr = GetLastError

    'Treat zero + nonzero last error as an API failure.
        If StyleOut = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful after a valid style read.
        K_TryGetWindowStyle = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

#If VBA7 Then
Private Function K_TrySetWindowStyle( _
    ByVal hWnd As LongPtr, _
    ByVal NewStyle As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function K_TrySetWindowStyle( _
    ByVal hWnd As Long, _
    ByVal NewStyle As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
'                            K_TrySetWindowStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to write the GWL_STYLE value using the correct API for the current
'   VBA / Office bitness.
'
' WHY THIS EXISTS
'   SetWindowLong / SetWindowLongPtr can validly return zero, so a robust
'   wrapper should use GetLastError to distinguish "real previous zero" from
'   failure.
'
' INPUTS
'   hWnd
'     Target window handle.
'
'   NewStyle
'     New GWL_STYLE value to write.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => style write succeeded
'   FALSE => style write failed
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Returns FALSE and populates FailMsg on failure.
'
' DEPENDENCIES
'   - SetWindowLong / SetWindowLongPtr
'   - GetLastError
'   - SetLastError
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim PrevStyle           As LongPtr   'Previous style returned by the API
#Else
    Dim PrevStyle           As Long      'Previous style returned by the API
#End If
    Dim LastErr             As Long      'Win32 last-error value read after the API call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize the default result state.
        FailMsg = vbNullString
        K_TrySetWindowStyle = False

    'Reject an invalid window handle deterministically.
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' WRITE: WINDOW STYLE
'------------------------------------------------------------------------------
    'Clear the Win32 last-error state before calling the API so a valid zero
    'return can later be distinguished from failure.
        SetLastError 0

#If VBA7 Then
    #If Win64 Then

        'Write the style with the 64-bit API in 64-bit Office / VBA.
            PrevStyle = SetWindowLongPtr(hWnd, GWL_STYLE, NewStyle)

    #Else

        'Write the style with the 32-bit API in VBA7 32-bit Office.
            PrevStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)

    #End If
#Else

    'Write the style with the legacy 32-bit API.
        PrevStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)

#End If

    'Read the Win32 last-error value immediately after the API call.
        LastErr = GetLastError

    'Treat zero + nonzero last error as an API failure.
        If PrevStyle = 0 And LastErr <> 0 Then
            FailMsg = "SetWindowLong/SetWindowLongPtr failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful after a valid style write.
        K_TrySetWindowStyle = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

#If VBA7 Then
Private Function K_TryRefreshWindowFrame( _
    ByVal hWnd As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function K_TryRefreshWindowFrame( _
    ByVal hWnd As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
'                           K_TryRefreshWindowFrame
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to force Windows to repaint the non-client frame of the specified
'   window after a style change.
'
' WHY THIS EXISTS
'   Updating GWL_STYLE alone is not always visually reflected immediately.
'   SetWindowPos with SWP_FRAMECHANGED is the standard way to notify Windows
'   that the frame should be recalculated and repainted.
'
' INPUTS
'   hWnd
'     Target window handle.
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => frame refresh succeeded
'   FALSE => frame refresh failed
'
' BEHAVIOR
'   - Uses the canonical no-move / no-size / no-z-order refresh pattern.
'
' ERROR POLICY
'   - Does NOT raise to callers.
'   - Returns FALSE and populates FailMsg on failure.
'
' DEPENDENCIES
'   - SetWindowPos
'   - GetLastError
'   - SetLastError
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ApiOK               As Long      'WinAPI success flag / return code
    Dim LastErr             As Long      'Win32 last-error value read after the API call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

    'Initialize the default result state.
        FailMsg = vbNullString
        K_TryRefreshWindowFrame = False

    'Reject an invalid window handle deterministically.
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' REFRESH: NON-CLIENT FRAME
'------------------------------------------------------------------------------
    'Clear the Win32 last-error state before calling the API.
        SetLastError 0

    'Force Windows to recalculate and repaint the non-client frame without
    'moving, resizing, or reordering the target window.
        ApiOK = SetWindowPos( _
                    hWnd, _
                    0, _
                    0, _
                    0, _
                    0, _
                    0, _
                    SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED)

    'Read the Win32 last-error value immediately after the API call.
        LastErr = GetLastError

    'Reject API failure deterministically and include the Win32 error code.
        If ApiOK = 0 Then
            FailMsg = "SetWindowPos failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful after a valid frame refresh.
        K_TryRefreshWindowFrame = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path.
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising.
        FailMsg = K_BuildRuntimeErrorText

End Function

Private Function K_VisibilityToBoolean(ByVal Visibility As K_UIVisibility) As Boolean

'
'==============================================================================
'                           K_VisibilityToBoolean
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a tri-state visibility enum value into the explicit Boolean visible
'   state required by Excel properties and internal helpers.
'
' WHY THIS EXISTS
'   Public callers use K_UIVisibility values, while Excel object-model
'   properties and internal helpers require a Boolean visible / hidden state.
'
' INPUTS
'   Visibility
'     Expected values:
'       - K_UI_Hide
'       - K_UI_Show
'
' RETURNS
'   TRUE  => visible
'   FALSE => hidden
'
' BEHAVIOR
'   - K_UI_Show maps to TRUE.
'   - Any other value maps to FALSE.
'
' ERROR POLICY
'   - Does NOT raise.
'
' NOTES
'   - Callers should only invoke this helper after excluding
'     K_UI_LeaveUnchanged.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN: BOOLEAN VISIBILITY
'------------------------------------------------------------------------------
    'Convert explicit SHOW to TRUE; otherwise return FALSE.
        K_VisibilityToBoolean = (Visibility = K_UI_Show)

End Function

Private Function K_BuildRuntimeErrorText() As String

'
'==============================================================================
'                           K_BuildRuntimeErrorText
'------------------------------------------------------------------------------
' PURPOSE
'   Build a consistent runtime diagnostic string from the active Err object.
'
' WHY THIS EXISTS
'   Several procedures in this module use identical failure-text construction.
'   A shared helper avoids duplicated formatting logic and keeps diagnostics
'   consistent.
'
' RETURNS
'   A formatted diagnostic string including:
'     - Err.Number
'     - Err.Description
'     - Err.Source, when available
'     - Erl, when available
'
' ERROR POLICY
'   - Does NOT raise.
'   - Returns best-effort text.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected issue while formatting the diagnostic.
        On Error Resume Next

'------------------------------------------------------------------------------
' BUILD: RUNTIME ERROR TEXT
'------------------------------------------------------------------------------
    'Build a consistent diagnostic string from the current Err state.
        K_BuildRuntimeErrorText = _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

End Function

Private Sub K_LogFailure( _
    ByVal ProcName As String, _
    ByVal Stage As String, _
    ByVal Detail As String)

'
'==============================================================================
'                                K_LogFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Write a consistent diagnostic line to the Immediate Window.
'
' WHY THIS EXISTS
'   The module uses fail-soft behavior and needs a single place to format
'   diagnostic output consistently.
'
' INPUTS
'   ProcName
'     Procedure name associated with the failure.
'
'   Stage
'     Logical stage / element associated with the failure.
'
'   Detail
'     Failure detail to append.
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Suppresses any unexpected logging failure locally.
'
' UPDATED
'   2026-04-11
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected logging failure.
        On Error Resume Next

'------------------------------------------------------------------------------
' WRITE: DIAGNOSTIC LINE
'------------------------------------------------------------------------------
    'Write a consistent fail-soft diagnostic line to the Immediate Window.
        Debug.Print ProcName & " failed @ " & Stage & " | " & Detail

End Sub


