Attribute VB_Name = "M_EXCEL_UI"
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
'   kiosk-like, or application-style shell
'
'   Excel exposes several UI elements directly through the object model
'   (Ribbon, status bar, scroll bars, formula bar, headings, workbook tabs,
'   gridlines), but it does not expose direct title-bar visibility control
'
'   This module unifies both approaches behind a safe, explicit API so callers
'   do not need to duplicate scattered UI-handling code
'
' PUBLIC SURFACE
'   - UIVisibility                   Tri-state visibility enum
'   - UI_SetExcelUI                  Core selective UI-state routine
'   - UI_SetExcelUI_WithResult       Selective routine with structured result
'   - UI_HideExcelUI                 Convenience wrapper: hide all managed UI
'   - UI_ShowExcelUI                 Convenience wrapper: show all managed UI
'   - UI_CaptureExcelUIState         Explicitly snapshot the current managed UI
'   - UI_ResetExcelUIToSnapshot      Best-effort restore to captured UI state
'   - UI_HasExcelUIStateSnapshot     Return TRUE when a snapshot exists
'   - UI_ClearExcelUIStateSnapshot   Remove any captured snapshot
'
' INTERNAL SUPPORT
'   - UI_ApplyExcelUIState
'   - UI_HandleApplyFailure
'   - UI_ClearResultBuffer
'   - UI_AddFailureToResult
'   - UI_BeginQuietUIUpdate
'   - UI_EndQuietUIUpdate
'   - UI_TrySetRibbonVisibleIfNeeded
'   - UI_TrySetTitleBarVisibleIfNeeded
'   - UI_TrySetBooleanPropertyIfNeeded
'   - UI_TryGetRibbonVisible
'   - UI_TryGetTitleBarVisible
'   - UI_TryGetBooleanProperty
'   - UI_TrySetTitleBarVisible
'   - UI_TrySetRibbonVisible
'   - UI_TrySetBooleanProperty
'   - UI_TryGetWindowStyle
'   - UI_TrySetWindowStyle
'   - UI_TryRefreshWindowFrame
'   - UI_IsValidVisibility
'   - UI_VisibilityToBoolean
'   - UI_BuildRuntimeErrorText
'   - UI_LogFailure
'   - WinAPI declarations and constants
'
' BEHAVIOR
'   - Application-level elements:
'       * Ribbon
'       * Status Bar
'       * Scroll Bars
'       * Formula Bar
'
'   - Window-level elements applied to each open Excel window:
'       * Headings
'       * Workbook Tabs
'       * Gridlines
'
'   - Title bar:
'       * applied to the Excel main window represented by Application.Hwnd
'         through WinAPI style update and non-client frame refresh
'
' ERROR POLICY
'   - Public entry points are fail-soft
'   - Unexpected errors are logged to the Immediate Window in the fire-and-
'     forget path
'   - Errors are not re-raised to callers
'   - The core routine uses best-effort application so one failed UI element
'     does not prevent later UI elements from being attempted
'
' PLATFORM / COMPATIBILITY
'   - Windows only
'   - Supports 32-bit and 64-bit Office / VBA through conditional compilation
'     and bitness-safe WinAPI wrappers
'
' NOTES
'   - This module does NOT automatically snapshot and restore prior Excel
'     object-model UI state
'   - UI_ShowExcelUI means "show all managed UI", not "restore previous state"
'   - UI_SetExcelUI is the preferred entry point for selective control
'   - UI_SetExcelUI_WithResult offers the same best-effort behavior while
'     returning structured diagnostics without a class-module dependency
'   - Ribbon control relies on Application.ExecuteExcel4Macro
'   - Title-bar control affects the Excel window represented by
'     Application.Hwnd, not a user-specific saved UI state
'   - The original main-window style is snapshotted against the current window
'     handle, so the restore path can follow Application.Hwnd safely if the
'     main Excel window is recreated
'   - The explicit snapshot / reset feature is separate from UI_ShowExcelUI and
'     is best-effort for per-window restore
'
' UPDATED
'   2026-04-19
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.0.0
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
' DECLARE: WIN32 / WIN64 API
'------------------------------------------------------------------------------
#If VBA7 Then

    #If Win64 Then

        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias _
            "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As _
            LongPtr

        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias _
            "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal _
            dwNewLong As LongPtr) As LongPtr

    #Else

        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias _
            "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long

        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias _
            "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal _
            dwNewLong As Long) As Long

    #End If

    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As _
        LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long

    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

    Private Declare PtrSafe Sub SetLastError Lib "kernel32" (ByVal dwErrCode As _
        Long)

#Else

    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long

    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

    Private Declare Function SetWindowPos Lib "user32" ( ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As _
        Long, ByVal cy As Long, ByVal uFlags As Long) As Long

    Private Declare Function GetLastError Lib "kernel32" () As Long

    Private Declare Sub SetLastError Lib "kernel32" ( ByVal dwErrCode As Long)

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

Public Sub UI_SetExcelUI(Optional ByVal Ribbon As UIVisibility = _
    UI_LeaveUnchanged, Optional ByVal StatusBar As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal ScrollBars As UIVisibility = UI_LeaveUnchanged, Optional ByVal _
    FormulaBar As UIVisibility = UI_LeaveUnchanged, Optional ByVal Headings As _
    UIVisibility = UI_LeaveUnchanged, Optional ByVal WorkbookTabs As UIVisibility = _
    UI_LeaveUnchanged, Optional ByVal Gridlines As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal TitleBar As UIVisibility = UI_LeaveUnchanged)

'
'==============================================================================
'                               UI_SetExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the requested visibility state to the Excel UI elements managed by
'   this module
'
' WHY THIS EXISTS
'   A Boolean-based "hide/show" routine is error-prone because omitted optional
'   arguments can accidentally imply FALSE / hidden
'
'   This routine uses an explicit tri-state API:
'     - UI_Show
'     - UI_Hide
'     - UI_LeaveUnchanged
'
'   This makes the caller's intent precise and prevents accidental UI changes
'   for omitted arguments
'
' INPUTS
'   Ribbon (optional)
'     UI_Show             => show Ribbon
'     UI_Hide             => hide Ribbon
'     UI_LeaveUnchanged   => do not touch Ribbon
'
'   StatusBar (optional)
'     UI_Show             => show status bar
'     UI_Hide             => hide status bar
'     UI_LeaveUnchanged   => do not touch status bar
'
'   ScrollBars (optional)
'     UI_Show             => show scroll bars
'     UI_Hide             => hide scroll bars
'     UI_LeaveUnchanged   => do not touch scroll bars
'
'   FormulaBar (optional)
'     UI_Show             => show formula bar
'     UI_Hide             => hide formula bar
'     UI_LeaveUnchanged   => do not touch formula bar
'
'   Headings (optional)
'     UI_Show             => show row / column headings in each window
'     UI_Hide             => hide row / column headings in each window
'     UI_LeaveUnchanged   => do not touch headings
'
'   WorkbookTabs (optional)
'     UI_Show             => show workbook tabs in each window
'     UI_Hide             => hide workbook tabs in each window
'     UI_LeaveUnchanged   => do not touch workbook tabs
'
'   Gridlines (optional)
'     UI_Show             => show gridlines in each window
'     UI_Hide             => hide gridlines in each window
'     UI_LeaveUnchanged   => do not touch gridlines
'
'   TitleBar (optional)
'     UI_Show             => show the title bar of the Excel main window
'                            represented by Application.Hwnd
'     UI_Hide             => hide the title bar of the Excel main window
'                            represented by Application.Hwnd
'     UI_LeaveUnchanged   => do not touch title bar
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Applies Ribbon / status bar / scroll bars / formula bar at Application
'     level
'   - Applies headings / workbook tabs / gridlines to every open Excel window
'     in the current Excel instance
'   - Applies title-bar visibility to the Excel main window represented by
'     Application.Hwnd via WinAPI
'   - Uses best-effort processing so one failed UI element does not prevent
'     subsequent UI elements from being attempted
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Unexpected failures are written to the Immediate Window
'   - Element-level failures are logged and processing continues
'
' DEPENDENCIES
'   - UI_ApplyExcelUIState
'
' NOTES
'   - This is the preferred entry point for selective UI control
'   - Changes affect the current Excel instance, not only the active workbook
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IgnoredFailureCount As Long      'Ignored structured-result failure count
    Dim IgnoredFailureList  As Variant   'Ignored structured-result failure list

    Const PROC As String = "UI_SetExcelUI"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY STATE THROUGH INTERNAL WORKER
'------------------------------------------------------------------------------
    'Delegate the full best-effort application flow to the shared worker,
    'requesting Immediate Window logging for any failures
        UI_ApplyExcelUIState ProcName:=PROC, Ribbon:=Ribbon, _
            StatusBar:=StatusBar, ScrollBars:=ScrollBars, FormulaBar:=FormulaBar, _
            Headings:=Headings, WorkbookTabs:=WorkbookTabs, Gridlines:=Gridlines, _
            TitleBar:=TitleBar, LogFailures:=True, FailureCount:=IgnoredFailureCount, _
            FailureList:=IgnoredFailureList, CaptureFailureList:=False

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave quietly through the normal termination path
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write an unexpected-procedure-level diagnostic line without interrupting
    'the caller
        UI_LogFailure PROC, "Unexpected", UI_BuildRuntimeErrorText

    'Exit quietly after logging
        Resume SafeExit

End Sub

Public Sub UI_HideExcelUI()

'
'==============================================================================
'                               UI_HideExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Hide all Excel UI elements managed by this module
'
' WHY THIS EXISTS
'   Some workbook-driven solutions want a simple one-call way to suppress the
'   managed Excel shell elements without specifying each element individually
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Delegates to UI_SetExcelUI
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
'   - Does NOT raise to callers
'   - Unexpected failures are written to the Immediate Window
'
' DEPENDENCIES
'   - UI_SetExcelUI
'
' NOTES
'   - This is a convenience wrapper
'   - For selective control, use UI_SetExcelUI directly
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC As String = "UI_HideExcelUI"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY: HIDE-ALL STATE
'------------------------------------------------------------------------------
    'Hide all managed UI elements through the central tri-state entry point
        UI_SetExcelUI Ribbon:=UI_Hide, StatusBar:=UI_Hide, ScrollBars:=UI_Hide, _
            FormulaBar:=UI_Hide, Headings:=UI_Hide, WorkbookTabs:=UI_Hide, _
            Gridlines:=UI_Hide, TitleBar:=UI_Hide

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave quietly through the normal termination path
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write an unexpected-procedure-level diagnostic line without interrupting
    'the caller
        UI_LogFailure PROC, "Unexpected", UI_BuildRuntimeErrorText

    'Exit quietly after logging
        Resume SafeExit

End Sub

Public Sub UI_ShowExcelUI()

'
'==============================================================================
'                               UI_ShowExcelUI
'------------------------------------------------------------------------------
' PURPOSE
'   Show all Excel UI elements managed by this module
'
' WHY THIS EXISTS
'   Workbook solutions that temporarily suppress the Excel shell often need a
'   single, deterministic call to restore all managed UI elements to visible
'   state
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Delegates to UI_SetExcelUI
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
'   - Does NOT raise to callers
'   - Unexpected failures are written to the Immediate Window
'
' DEPENDENCIES
'   - UI_SetExcelUI
'
' NOTES
'   - This means "show all managed UI"
'   - It does NOT restore a previously captured user-specific object-model UI
'     state
'   - For selective control, use UI_SetExcelUI directly
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC As String = "UI_ShowExcelUI"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY: SHOW-ALL STATE
'------------------------------------------------------------------------------
    'Show all managed UI elements through the central tri-state entry point
        UI_SetExcelUI Ribbon:=UI_Show, StatusBar:=UI_Show, ScrollBars:=UI_Show, _
            FormulaBar:=UI_Show, Headings:=UI_Show, WorkbookTabs:=UI_Show, _
            Gridlines:=UI_Show, TitleBar:=UI_Show

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave quietly through the normal termination path
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write an unexpected-procedure-level diagnostic line without interrupting
    'the caller
        UI_LogFailure PROC, "Unexpected", UI_BuildRuntimeErrorText
    'Exit quietly after logging
        Resume SafeExit

End Sub

Public Function UI_SetExcelUI_WithResult(Optional ByVal Ribbon As UIVisibility = _
    UI_LeaveUnchanged, Optional ByVal StatusBar As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal ScrollBars As UIVisibility = UI_LeaveUnchanged, Optional ByVal _
    FormulaBar As UIVisibility = UI_LeaveUnchanged, Optional ByVal Headings As _
    UIVisibility = UI_LeaveUnchanged, Optional ByVal WorkbookTabs As UIVisibility = _
    UI_LeaveUnchanged, Optional ByVal Gridlines As UIVisibility = UI_LeaveUnchanged, _
    Optional ByVal TitleBar As UIVisibility = UI_LeaveUnchanged, Optional ByRef _
    FailureCount As Long = 0, Optional ByRef FailureList As Variant) As Boolean

'
'==============================================================================
'                         UI_SetExcelUI_WithResult
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the requested visibility state to the Excel UI elements managed by
'   this module and return a Boolean success flag, with optional structured
'   failure details captured through ByRef outputs
'
' WHY THIS EXISTS
'   UI_SetExcelUI is the preferred fire-and-forget, fail-soft entry point for
'   callers that only need best-effort application plus Immediate Window
'   diagnostics
'
'   Some callers, however, need structured feedback so they can:
'     - inspect whether the full operation succeeded
'     - count element-level failures
'     - enumerate the recorded failures in order
'     - surface diagnostics to higher-level orchestration or test logic
'
'   This routine provides the same best-effort behavior as UI_SetExcelUI, but
'   avoids any class-module dependency by returning:
'     - a Boolean success flag
'     - FailureCount as an optional ByRef output
'     - FailureList as an optional ByRef Variant containing a 1-based String
'       array of recorded failures
'
' INPUTS
'   Ribbon (optional)
'     UI_Show             => show Ribbon
'     UI_Hide             => hide Ribbon
'     UI_LeaveUnchanged   => do not touch Ribbon
'
'   StatusBar (optional)
'     UI_Show             => show status bar
'     UI_Hide             => hide status bar
'     UI_LeaveUnchanged   => do not touch status bar
'
'   ScrollBars (optional)
'     UI_Show             => show scroll bars
'     UI_Hide             => hide scroll bars
'     UI_LeaveUnchanged   => do not touch scroll bars
'
'   FormulaBar (optional)
'     UI_Show             => show formula bar
'     UI_Hide             => hide formula bar
'     UI_LeaveUnchanged   => do not touch formula bar
'
'   Headings (optional)
'     UI_Show             => show row / column headings in each window
'     UI_Hide             => hide row / column headings in each window
'     UI_LeaveUnchanged   => do not touch headings
'
'   WorkbookTabs (optional)
'     UI_Show             => show workbook tabs in each window
'     UI_Hide             => hide workbook tabs in each window
'     UI_LeaveUnchanged   => do not touch workbook tabs
'
'   Gridlines (optional)
'     UI_Show             => show gridlines in each window
'     UI_Hide             => hide gridlines in each window
'     UI_LeaveUnchanged   => do not touch gridlines
'
'   TitleBar (optional)
'     UI_Show             => show the title bar of the Excel main window
'                            represented by Application.Hwnd
'     UI_Hide             => hide the title bar of the Excel main window
'                            represented by Application.Hwnd
'     UI_LeaveUnchanged   => do not touch title bar
'
'   FailureCount (optional, ByRef output)
'     Receives the number of recorded failures
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
'     level
'   - Applies headings / workbook tabs / gridlines to every open Excel window
'     in the current Excel instance
'   - Applies title-bar visibility to the Excel main window represented by
'     Application.Hwnd via WinAPI
'   - Uses best-effort processing so one failed UI element does not prevent
'     subsequent UI elements from being attempted
'   - Records failures through FailureCount and, when requested, FailureList
'
' ERROR POLICY
'   - Does NOT raise to callers for ordinary element-level failures
'   - Returns FALSE when one or more failures were recorded
'   - Unexpected procedure-level failures are captured as an "Unexpected"
'     failure entry and also produce a FALSE result
'
' DEPENDENCIES
'   - UI_ApplyExcelUIState
'   - UI_HandleApplyFailure
'   - UI_ClearResultBuffer
'
' NOTES
'   - This routine mirrors the best-effort semantics of UI_SetExcelUI
'   - Failure order is preserved
'   - FailureList remains optional so callers that only need the Boolean result
'     or failure count do not need to manage an array
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean    'Overall success flag returned to the caller
    Dim CaptureFailureList  As Boolean    'TRUE when the caller supplied FailureList
    Dim InternalFailureList As Variant    'Local working failure list copied back only when requested

    Const PROC As String = "UI_SetExcelUI_WithResult"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Detect whether the caller supplied the optional failure-list output
        CaptureFailureList = Not IsMissing(FailureList)

    'Initialize the public result outputs in their clean-success state
        UI_ClearResultBuffer FailureCount, InternalFailureList, _
            CaptureFailureList
        Succeeded = True

    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY STATE THROUGH INTERNAL WORKER
'------------------------------------------------------------------------------
    'Delegate the full best-effort application flow to the shared worker,
    'requesting structured failure capture rather than Immediate Window logging
        Succeeded = UI_ApplyExcelUIState(ProcName:=PROC, Ribbon:=Ribbon, _
            StatusBar:=StatusBar, ScrollBars:=ScrollBars, FormulaBar:=FormulaBar, _
            Headings:=Headings, WorkbookTabs:=WorkbookTabs, Gridlines:=Gridlines, _
            TitleBar:=TitleBar, LogFailures:=False, FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, _
            CaptureFailureList:=CaptureFailureList)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Copy the internal working list back to the caller only when the optional
    'failure-list output was actually supplied
        If CaptureFailureList Then
            FailureList = InternalFailureList
        End If

    'Return the overall success flag to the caller
        UI_SetExcelUI_WithResult = Succeeded

    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Capture the unexpected wrapper-level failure in the structured result buffers
        UI_HandleApplyFailure ProcName:=PROC, LogFailures:=False, _
            Succeeded:=Succeeded, FailureCount:=FailureCount, _
            FailureList:=InternalFailureList, CaptureFailureList:=CaptureFailureList, _
            Stage:="Unexpected", Detail:=UI_BuildRuntimeErrorText

    'Leave quietly through the normal termination path
        Resume SafeExit

End Function

Public Sub UI_CaptureExcelUIState()

'
'==============================================================================
'                           UI_CaptureExcelUIState
'------------------------------------------------------------------------------
' PURPOSE
'   Explicitly snapshot the current managed Excel UI state so it can later be
'   restored through UI_ResetExcelUIToSnapshot
'
' WHY THIS EXISTS
'   UI_ShowExcelUI intentionally means "show all managed UI", not "restore the
'   user's prior state"
'
'   Some callers need a distinct, deliberate lifecycle:
'     - capture current state
'     - apply a constrained shell
'     - restore the captured state later
'
'   This routine provides that explicit capture step
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Captures application-level state
'   - Captures per-window state by index
'   - Captures title-bar state on a best-effort basis
'   - Marks the snapshot as available even when Ribbon / TitleBar state could
'     only be captured best-effort
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Logs any unexpected issue to the Immediate Window
'
' DEPENDENCIES
'   - UI_ClearExcelUIStateSnapshot
'   - UI_TryGetRibbonVisible
'   - UI_TryGetTitleBarVisible
'   - UI_LogFailure
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i                   As Long      'Current window index during snapshot
    Dim Msg                 As String    'Diagnostic message from helper reads

    Const PROC As String = "UI_CaptureExcelUIState"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' CLEAR PRIOR SNAPSHOT
'------------------------------------------------------------------------------
    'Clear any prior snapshot before capturing a fresh one
        UI_ClearExcelUIStateSnapshot

'------------------------------------------------------------------------------
' CAPTURE APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Capture application-level UI state directly from Excel
        m_SnapshotStatusBarVisible = Application.DisplayStatusBar
        m_SnapshotScrollBarsVisible = Application.DisplayScrollBars
        m_SnapshotFormulaBarVisible = Application.DisplayFormulaBar

'------------------------------------------------------------------------------
' CAPTURE RIBBON / TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Capture Ribbon state through the best-effort helper
        m_SnapshotRibbonKnown = UI_TryGetRibbonVisible(m_SnapshotRibbonVisible, _
            Msg)
        If Not m_SnapshotRibbonKnown Then
            UI_LogFailure PROC, "Ribbon", Msg
        End If
    'Capture TitleBar state through the best-effort helper
        m_SnapshotTitleBarKnown = _
            UI_TryGetTitleBarVisible(m_SnapshotTitleBarVisible, Msg)
        If Not m_SnapshotTitleBarKnown Then
            UI_LogFailure PROC, "TitleBar", Msg
        End If

'------------------------------------------------------------------------------
' CAPTURE WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Capture the current window count
        m_SnapshotWindowCount = Application.Windows.Count
    'Allocate and fill per-window arrays only when at least one window exists
        If m_SnapshotWindowCount > 0 Then
            'Allocate the headings array
                ReDim m_SnapshotHeadingsVisible(1 To m_SnapshotWindowCount)
            'Allocate the workbook-tabs array
                ReDim m_SnapshotWorkbookTabsVisible(1 To m_SnapshotWindowCount)
            'Allocate the gridlines array
                ReDim m_SnapshotGridlinesVisible(1 To m_SnapshotWindowCount)
            'Capture the state of each current Excel window by index
                For i = 1 To m_SnapshotWindowCount
                    'Capture Headings visibility
                        m_SnapshotHeadingsVisible(i) = _
                            Application.Windows(i).DisplayHeadings
                    'Capture WorkbookTabs visibility
                        m_SnapshotWorkbookTabsVisible(i) = _
                            Application.Windows(i).DisplayWorkbookTabs
                    'Capture Gridlines visibility
                        m_SnapshotGridlinesVisible(i) = _
                            Application.Windows(i).DisplayGridlines
                Next i
        End If

'------------------------------------------------------------------------------
' MARK SNAPSHOT AVAILABLE
'------------------------------------------------------------------------------
    'Mark the snapshot as available after capture completes
        m_HasExcelUIStateSnapshot = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Log the unexpected capture failure without interrupting callers
        UI_LogFailure PROC, "Unexpected", UI_BuildRuntimeErrorText

    'Leave quietly after logging
        Resume SafeExit

End Sub

Public Function UI_HasExcelUIStateSnapshot() As Boolean

'
'==============================================================================
'                        UI_HasExcelUIStateSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Return whether an explicit Excel UI snapshot is currently available
'
' RETURNS
'   TRUE  => a snapshot is available
'   FALSE => no snapshot is available
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN SNAPSHOT AVAILABILITY
'------------------------------------------------------------------------------
    'Return whether a captured UI snapshot is currently available
        UI_HasExcelUIStateSnapshot = m_HasExcelUIStateSnapshot

End Function

Public Sub UI_ResetExcelUIToSnapshot()

'
'==============================================================================
'                        UI_ResetExcelUIToSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Best-effort restore the Excel UI to the most recently captured explicit
'   snapshot
'
' WHY THIS EXISTS
'   Callers that previously used UI_CaptureExcelUIState may need a distinct way
'   to restore the captured baseline rather than merely showing all managed UI
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Restores title bar and Ribbon when their snapshot states were known
'   - Restores application-level object-model properties
'   - Restores per-window state by common index range
'   - Uses a quiet-update scope with ScreenUpdating where possible
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Logs any restore issue to the Immediate Window
'   - Best-effort only, especially for per-window restore
'
' DEPENDENCIES
'   - UI_BeginQuietUIUpdate
'   - UI_EndQuietUIUpdate
'   - UI_TrySetTitleBarVisibleIfNeeded
'   - UI_TrySetRibbonVisibleIfNeeded
'   - UI_TrySetBooleanPropertyIfNeeded
'   - UI_LogFailure
'
' UPDATED
'   2026-04-19
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

    Const PROC As String = "UI_ResetExcelUIToSnapshot"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Do nothing when no explicit snapshot is available
        If Not m_HasExcelUIStateSnapshot Then
            UI_LogFailure PROC, "NoSnapshot", _
                "no captured Excel UI snapshot is available"
            GoTo SafeExit
        End If

    'Enter the quiet-update scope to reduce worksheet redraw where possible
        UI_BeginQuietUIUpdate OldScreenUpdating, QuietModeChanged

'------------------------------------------------------------------------------
' RESTORE TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Restore TitleBar first when its snapshot state was captured successfully
        If m_SnapshotTitleBarKnown Then
            If Not UI_TrySetTitleBarVisibleIfNeeded(m_SnapshotTitleBarVisible, _
                Msg) Then
                UI_LogFailure PROC, "TitleBar", Msg
            End If
        End If

'------------------------------------------------------------------------------
' RESTORE RIBBON STATE
'------------------------------------------------------------------------------
    'Restore Ribbon when its snapshot state was captured successfully
        If m_SnapshotRibbonKnown Then
            If Not UI_TrySetRibbonVisibleIfNeeded(m_SnapshotRibbonVisible, Msg) _
                Then
                UI_LogFailure PROC, "Ribbon", Msg
            End If
        End If

'------------------------------------------------------------------------------
' RESTORE APPLICATION-LEVEL STATE
'------------------------------------------------------------------------------
    'Restore StatusBar visibility best-effort
        If Not UI_TrySetBooleanPropertyIfNeeded(Application, "DisplayStatusBar", _
            m_SnapshotStatusBarVisible, Msg) Then
            UI_LogFailure PROC, "StatusBar", Msg
        End If

    'Restore ScrollBars visibility best-effort
        If Not UI_TrySetBooleanPropertyIfNeeded(Application, "DisplayScrollBars", _
            m_SnapshotScrollBarsVisible, Msg) Then
            UI_LogFailure PROC, "ScrollBars", Msg
        End If

    'Restore FormulaBar visibility best-effort
        If Not UI_TrySetBooleanPropertyIfNeeded(Application, "DisplayFormulaBar", _
            m_SnapshotFormulaBarVisible, Msg) Then
            UI_LogFailure PROC, "FormulaBar", Msg
        End If

'------------------------------------------------------------------------------
' RESTORE WINDOW-LEVEL STATE
'------------------------------------------------------------------------------
    'Restore only the common indexed window range that still exists
        WindowLimit = Application.Windows.Count
        If m_SnapshotWindowCount < WindowLimit Then WindowLimit = _
            m_SnapshotWindowCount

    'Restore each saved window state up to the common window count
        For i = 1 To WindowLimit

            'Restore Headings visibility for the current saved window index
                If Not UI_TrySetBooleanPropertyIfNeeded(Application.Windows(i), _
                    "DisplayHeadings", m_SnapshotHeadingsVisible(i), Msg) Then
                    UI_LogFailure PROC, "Headings [" & _
                        Application.Windows(i).Caption & "]", Msg
                End If

            'Restore WorkbookTabs visibility for the current saved window index
                If Not UI_TrySetBooleanPropertyIfNeeded(Application.Windows(i), _
                    "DisplayWorkbookTabs", m_SnapshotWorkbookTabsVisible(i), Msg) _
                    Then
                    UI_LogFailure PROC, "WorkbookTabs [" & _
                        Application.Windows(i).Caption & "]", Msg
                End If

            'Restore Gridlines visibility for the current saved window index
                If Not UI_TrySetBooleanPropertyIfNeeded(Application.Windows(i), _
                    "DisplayGridlines", m_SnapshotGridlinesVisible(i), Msg) Then
                    UI_LogFailure PROC, "Gridlines [" & _
                        Application.Windows(i).Caption & "]", Msg
                End If

        Next i

    'Log a note when the current window count differs from the captured count
        If Application.Windows.Count <> m_SnapshotWindowCount Then
            UI_LogFailure PROC, "WindowCount", _
                "current window count differs from captured snapshot; restore applied to common index range only"
        End If

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave the quiet-update scope and restore ScreenUpdating when needed
        UI_EndQuietUIUpdate OldScreenUpdating, QuietModeChanged

    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Log the unexpected restore failure without interrupting callers
        UI_LogFailure PROC, "Unexpected", UI_BuildRuntimeErrorText

    'Leave quietly after logging
        Resume SafeExit

End Sub

Public Sub UI_ClearExcelUIStateSnapshot()

'
'==============================================================================
'                      UI_ClearExcelUIStateSnapshot
'------------------------------------------------------------------------------
' PURPOSE
'   Remove any captured explicit Excel UI snapshot from module state
'
' WHY THIS EXISTS
'   Callers may want to discard an obsolete snapshot explicitly before taking a
'   new one or before leaving a workflow
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise to callers
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' RESET SNAPSHOT FLAGS
'------------------------------------------------------------------------------
    'Mark the explicit UI snapshot as unavailable
        m_HasExcelUIStateSnapshot = False

    'Reset best-effort known flags
        m_SnapshotRibbonKnown = False
        m_SnapshotTitleBarKnown = False

'------------------------------------------------------------------------------
' RESET SNAPSHOT VALUES
'------------------------------------------------------------------------------
    'Reset application-level values
        m_SnapshotRibbonVisible = False
        m_SnapshotStatusBarVisible = False
        m_SnapshotScrollBarsVisible = False
        m_SnapshotFormulaBarVisible = False

    'Reset title-bar value
        m_SnapshotTitleBarVisible = False

    'Reset captured window count
        m_SnapshotWindowCount = 0

'------------------------------------------------------------------------------
' CLEAR SNAPSHOT ARRAYS
'------------------------------------------------------------------------------
    'Clear any captured per-window arrays
        Erase m_SnapshotHeadingsVisible
        Erase m_SnapshotWorkbookTabsVisible
        Erase m_SnapshotGridlinesVisible

End Sub

Private Function UI_ApplyExcelUIState(ByVal ProcName As String, ByVal Ribbon As _
    UIVisibility, ByVal StatusBar As UIVisibility, ByVal ScrollBars As UIVisibility, _
    ByVal FormulaBar As UIVisibility, ByVal Headings As UIVisibility, ByVal _
    WorkbookTabs As UIVisibility, ByVal Gridlines As UIVisibility, ByVal TitleBar As _
    UIVisibility, ByVal LogFailures As Boolean, ByRef FailureCount As Long, ByRef _
    FailureList As Variant, ByVal CaptureFailureList As Boolean) As Boolean

'
'==============================================================================
'                           UI_ApplyExcelUIState
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the requested Excel UI state through the shared internal worker used
'   by both public entry points
'
' WHY THIS EXISTS
'   The module exposes two public application paths:
'     - UI_SetExcelUI
'     - UI_SetExcelUI_WithResult
'
'   They are intentionally different only in how they surface failures:
'     - logging to the Immediate Window
'     - structured result buffers
'
'   Centralizing the actual UI-application logic here eliminates duplicated
'   orchestration and reduces the risk of future behavioral drift
'
' INPUTS
'   ProcName
'     Public caller name used for failure diagnostics
'
'   Ribbon / StatusBar / ScrollBars / FormulaBar / Headings / WorkbookTabs /
'   Gridlines / TitleBar
'     Requested tri-state UI modes
'
'   LogFailures
'     TRUE  => write failures to the Immediate Window
'     FALSE => suppress Immediate Window logging and use only the result
'              buffers
'
'   FailureCount
'     Receives the number of recorded failures
'
'   FailureList
'     Optional working Variant holding a 1-based String array of failures
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
'   - Initializes the result buffers
'   - Applies all requested UI changes using best-effort processing
'   - Uses ScreenUpdating suppression where possible to reduce worksheet redraw
'   - Skips object-model, Ribbon, and TitleBar writes when the current state
'     can be read and already matches the requested target
'   - Validates tri-state inputs before converting them to Boolean targets
'   - Records and optionally logs failures in insertion order
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Captures unexpected procedure-level failures as "Unexpected"
'
' DEPENDENCIES
'   - UI_ClearResultBuffer
'   - UI_BeginQuietUIUpdate
'   - UI_EndQuietUIUpdate
'   - UI_IsValidVisibility
'   - UI_VisibilityToBoolean
'   - UI_TrySetRibbonVisibleIfNeeded
'   - UI_TrySetTitleBarVisibleIfNeeded
'   - UI_TrySetBooleanPropertyIfNeeded
'   - UI_HandleApplyFailure
'   - UI_BuildRuntimeErrorText
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Succeeded           As Boolean    'Overall success flag returned by the worker
    Dim W                   As Window     'Workbook window in current Excel instance
    Dim ShowFlag            As Boolean    'Converted Boolean visibility target
    Dim Msg                 As String     'Element-level diagnostic message

    Dim ValidRibbon         As Boolean    'TRUE when Ribbon contains a valid tri-state value
    Dim ValidStatusBar      As Boolean    'TRUE when StatusBar contains a valid tri-state value
    Dim ValidScrollBars     As Boolean    'TRUE when ScrollBars contains a valid tri-state value
    Dim ValidFormulaBar     As Boolean    'TRUE when FormulaBar contains a valid tri-state value
    Dim ValidHeadings       As Boolean    'TRUE when Headings contains a valid tri-state value
    Dim ValidWorkbookTabs   As Boolean    'TRUE when WorkbookTabs contains a valid tri-state value
    Dim ValidGridlines      As Boolean    'TRUE when Gridlines contains a valid tri-state value
    Dim ValidTitleBar       As Boolean    'TRUE when TitleBar contains a valid tri-state value

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
    'Initialize the result buffers in their clean-success state
        UI_ClearResultBuffer FailureCount, FailureList, CaptureFailureList
        Succeeded = True
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail
    'Enter the quiet-update scope to reduce worksheet redraw where possible
        UI_BeginQuietUIUpdate OldScreenUpdating, QuietModeChanged

'------------------------------------------------------------------------------
' VALIDATE TRI-STATE INPUTS
'------------------------------------------------------------------------------
    'Validate the Ribbon tri-state input
        ValidRibbon = UI_IsValidVisibility(Ribbon)
        If Not ValidRibbon Then
            UI_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, _
                FailureList, CaptureFailureList, "Ribbon", _
                "invalid UIVisibility value: " & CStr(Ribbon)
        End If
    'Validate the StatusBar tri-state input
        ValidStatusBar = UI_IsValidVisibility(StatusBar)
        If Not ValidStatusBar Then
            UI_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, _
                FailureList, CaptureFailureList, "StatusBar", _
                "invalid UIVisibility value: " & CStr(StatusBar)
        End If
    'Validate the ScrollBars tri-state input
        ValidScrollBars = UI_IsValidVisibility(ScrollBars)
        If Not ValidScrollBars Then
            UI_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, _
                FailureList, CaptureFailureList, "ScrollBars", _
                "invalid UIVisibility value: " & CStr(ScrollBars)
        End If
    'Validate the FormulaBar tri-state input
        ValidFormulaBar = UI_IsValidVisibility(FormulaBar)
        If Not ValidFormulaBar Then
            UI_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, _
                FailureList, CaptureFailureList, "FormulaBar", _
                "invalid UIVisibility value: " & CStr(FormulaBar)
        End If
    'Validate the Headings tri-state input
        ValidHeadings = UI_IsValidVisibility(Headings)
        If Not ValidHeadings Then
            UI_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, _
                FailureList, CaptureFailureList, "Headings", _
                "invalid UIVisibility value: " & CStr(Headings)
        End If
    'Validate the WorkbookTabs tri-state input
        ValidWorkbookTabs = UI_IsValidVisibility(WorkbookTabs)
        If Not ValidWorkbookTabs Then
            UI_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, _
                FailureList, CaptureFailureList, "WorkbookTabs", _
                "invalid UIVisibility value: " & CStr(WorkbookTabs)
        End If
    'Validate the Gridlines tri-state input
        ValidGridlines = UI_IsValidVisibility(Gridlines)
        If Not ValidGridlines Then
            UI_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, _
                FailureList, CaptureFailureList, "Gridlines", _
                "invalid UIVisibility value: " & CStr(Gridlines)
        End If
    'Validate the TitleBar tri-state input
        ValidTitleBar = UI_IsValidVisibility(TitleBar)
        If Not ValidTitleBar Then
            UI_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, _
                FailureList, CaptureFailureList, "TitleBar", _
                "invalid UIVisibility value: " & CStr(TitleBar)
        End If

'------------------------------------------------------------------------------
' APPLY APPLICATION-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Apply Ribbon visibility when requested and valid
        If ValidRibbon Then
            If Ribbon <> UI_LeaveUnchanged Then
                'Convert the tri-state enum to the explicit Boolean state
                'expected by the lower-level helper
                    ShowFlag = UI_VisibilityToBoolean(Ribbon)
                'Attempt the Ribbon update only when needed and record any
                'failure without interrupting later operations
                    If Not UI_TrySetRibbonVisibleIfNeeded(ShowFlag, Msg) Then
                        UI_HandleApplyFailure ProcName, LogFailures, Succeeded, _
                            FailureCount, FailureList, CaptureFailureList, "Ribbon", _
                            Msg
                    End If
            End If
        End If

    'Apply StatusBar visibility when requested and valid
        If ValidStatusBar Then
            If StatusBar <> UI_LeaveUnchanged Then
                'Convert the tri-state enum to the explicit Boolean state
                'expected by the lower-level helper
                    ShowFlag = UI_VisibilityToBoolean(StatusBar)
                'Attempt the property write only when needed and record any
                'failure without interrupting later operations
                    If Not UI_TrySetBooleanPropertyIfNeeded(Application, _
                        "DisplayStatusBar", ShowFlag, Msg) Then
                        UI_HandleApplyFailure ProcName, LogFailures, Succeeded, _
                            FailureCount, FailureList, CaptureFailureList, _
                            "StatusBar", Msg
                    End If
            End If
        End If

    'Apply ScrollBars visibility when requested and valid
        If ValidScrollBars Then
            If ScrollBars <> UI_LeaveUnchanged Then
                'Convert the tri-state enum to the explicit Boolean state
                'expected by the lower-level helper
                    ShowFlag = UI_VisibilityToBoolean(ScrollBars)
                'Attempt the property write only when needed and record any
                'failure without interrupting later operations
                    If Not UI_TrySetBooleanPropertyIfNeeded(Application, _
                        "DisplayScrollBars", ShowFlag, Msg) Then
                        UI_HandleApplyFailure ProcName, LogFailures, Succeeded, _
                            FailureCount, FailureList, CaptureFailureList, _
                            "ScrollBars", Msg
                    End If
            End If
        End If
    'Apply FormulaBar visibility when requested and valid
        If ValidFormulaBar Then
            If FormulaBar <> UI_LeaveUnchanged Then
                'Convert the tri-state enum to the explicit Boolean state
                'expected by the lower-level helper
                    ShowFlag = UI_VisibilityToBoolean(FormulaBar)
                'Attempt the property write only when needed and record any
                'failure without interrupting later operations
                    If Not UI_TrySetBooleanPropertyIfNeeded(Application, _
                        "DisplayFormulaBar", ShowFlag, Msg) Then
                        UI_HandleApplyFailure ProcName, LogFailures, Succeeded, _
                            FailureCount, FailureList, CaptureFailureList, _
                            "FormulaBar", Msg
                    End If
            End If
        End If

'------------------------------------------------------------------------------
' PRECOMPUTE WINDOW-LEVEL REQUESTS
'------------------------------------------------------------------------------
    'Precompute whether each window-level property was requested
        DoHeadings = ValidHeadings And (Headings <> UI_LeaveUnchanged)
        DoWorkbookTabs = ValidWorkbookTabs And (WorkbookTabs <> _
            UI_LeaveUnchanged)
        DoGridlines = ValidGridlines And (Gridlines <> UI_LeaveUnchanged)
    'Precompute the Boolean targets only for requested properties
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
' APPLY WINDOW-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Process window-scoped UI only when at least one window-level element has
    'been requested for change
        If DoHeadings Or DoWorkbookTabs Or DoGridlines Then
            'Apply the requested window-level visibility state to each open
            'Excel window in the current instance
                For Each W In Application.Windows
                    'Apply headings visibility when requested
                        If DoHeadings Then
                            'Attempt the property write only when needed and
                            'record any failure without interrupting later
                            'operations
                                If Not UI_TrySetBooleanPropertyIfNeeded(W, _
                                    "DisplayHeadings", ShowHeadings, Msg) Then
                                    UI_HandleApplyFailure ProcName, LogFailures, _
                                        Succeeded, FailureCount, FailureList, _
                                        CaptureFailureList, "Headings [" & W.Caption _
                                        & "]", Msg
                                End If
                        End If
                    'Apply workbook-tabs visibility when requested
                        If DoWorkbookTabs Then
                            'Attempt the property write only when needed and
                            'record any failure without interrupting later
                            'operations
                                If Not UI_TrySetBooleanPropertyIfNeeded(W, _
                                    "DisplayWorkbookTabs", ShowWorkbookTabs, Msg) _
                                    Then
                                    UI_HandleApplyFailure ProcName, LogFailures, _
                                        Succeeded, FailureCount, FailureList, _
                                        CaptureFailureList, "WorkbookTabs [" & _
                                        W.Caption & "]", Msg
                                End If
                        End If
                    'Apply gridlines visibility when requested
                        If DoGridlines Then
                            'Attempt the property write only when needed and
                            'record any failure without interrupting later
                            'operations
                                If Not UI_TrySetBooleanPropertyIfNeeded(W, _
                                    "DisplayGridlines", ShowGridlines, Msg) Then
                                    UI_HandleApplyFailure ProcName, LogFailures, _
                                        Succeeded, FailureCount, FailureList, _
                                        CaptureFailureList, "Gridlines [" & _
                                        W.Caption & "]", Msg
                                End If
                        End If
                Next W
        End If

'------------------------------------------------------------------------------
' APPLY TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Apply title-bar visibility when requested and valid
        If ValidTitleBar Then
            If TitleBar <> UI_LeaveUnchanged Then
                'Convert the tri-state enum to the explicit Boolean state
                'expected by the lower-level helper
                    ShowFlag = UI_VisibilityToBoolean(TitleBar)
                'Attempt the title-bar update only when needed and record any
                'failure without interrupting later operations
                    If Not UI_TrySetTitleBarVisibleIfNeeded(ShowFlag, Msg) Then
                        UI_HandleApplyFailure ProcName, LogFailures, Succeeded, _
                            FailureCount, FailureList, CaptureFailureList, _
                            "TitleBar", Msg
                    End If
            End If
        End If

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave the quiet-update scope and restore ScreenUpdating when needed
        UI_EndQuietUIUpdate OldScreenUpdating, QuietModeChanged
    'Return the overall success flag to the caller
        UI_ApplyExcelUIState = Succeeded
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Capture the unexpected worker-level failure in the result buffers and
    'optionally log it
        UI_HandleApplyFailure ProcName, LogFailures, Succeeded, FailureCount, _
            FailureList, CaptureFailureList, "Unexpected", UI_BuildRuntimeErrorText
    'Leave quietly through the normal termination path
        Resume SafeExit

End Function

Private Function UI_IsValidVisibility(ByVal Visibility As UIVisibility) As Boolean

'
'==============================================================================
'                           UI_IsValidVisibility
'------------------------------------------------------------------------------
' PURPOSE
'   Return whether a UIVisibility argument contains one of the supported
'   tri-state enum values
'
' WHY THIS EXISTS
'   Public procedures accept UIVisibility arguments, but VBA enum-typed
'   parameters can still receive invalid numeric values at runtime
'
'   This helper centralizes defensive validation so the worker can:
'     - detect invalid values explicitly
'     - record structured failures consistently
'     - avoid silently coercing unexpected values into Boolean targets
'
' INPUTS
'   Visibility
'     Candidate UIVisibility value to validate
'
' RETURNS
'   TRUE  => value is one of:
'              * UI_LeaveUnchanged
'              * UI_Hide
'              * UI_Show
'   FALSE => value is outside the supported UIVisibility domain
'
' ERROR POLICY
'   - Does NOT raise
'
' UPDATED
'   2026-04-21
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN VALIDITY FLAG
'------------------------------------------------------------------------------
    'Return TRUE only for the three supported tri-state enum values
        UI_IsValidVisibility = _
            (Visibility = UI_LeaveUnchanged) Or _
            (Visibility = UI_Hide) Or _
            (Visibility = UI_Show)

End Function

Private Sub UI_HandleApplyFailure(ByVal ProcName As String, ByVal LogFailures As _
    Boolean, ByRef Succeeded As Boolean, ByRef FailureCount As Long, ByRef _
    FailureList As Variant, ByVal CaptureFailureList As Boolean, ByVal Stage As _
    String, ByVal Detail As String)

'
'==============================================================================
'                           UI_HandleApplyFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Handle one recorded failure consistently for the shared internal worker
'
' WHY THIS EXISTS
'   The shared worker must support two public failure-surfacing modes:
'     - logging to the Immediate Window
'     - structured result capture through standard-module outputs
'
'   This helper centralizes both actions so the worker logic stays compact and
'   consistent
'
' INPUTS
'   ProcName
'     Public caller name used for logging
'
'   LogFailures
'     TRUE  => write the failure to the Immediate Window
'     FALSE => suppress logging
'
'   Succeeded / FailureCount / FailureList / CaptureFailureList
'     Result buffers used by the shared worker
'
'   Stage
'     Logical stage, element, or operation associated with the failure
'
'   Detail
'     Failure detail to append
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise
'
' DEPENDENCIES
'   - UI_AddFailureToResult
'   - UI_LogFailure
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
    'Record the failure into the structured result buffers
        UI_AddFailureToResult Succeeded, FailureCount, FailureList, _
            CaptureFailureList, Stage, Detail

'------------------------------------------------------------------------------
' OPTIONAL LOGGING
'------------------------------------------------------------------------------
    'Write the failure to the Immediate Window only when requested by the
    'caller path
        If LogFailures Then
            UI_LogFailure ProcName, Stage, Detail
        End If

End Sub

Private Sub UI_ClearResultBuffer(ByRef FailureCount As Long, ByRef FailureList _
    As Variant, ByVal CaptureFailureList As Boolean)

'
'==============================================================================
'                           UI_ClearResultBuffer
'------------------------------------------------------------------------------
' PURPOSE
'   Initialize the result buffers used by UI_SetExcelUI_WithResult and the
'   shared internal worker
'
' WHY THIS EXISTS
'   The standard-module result pattern needs a consistent way to reset:
'     - FailureCount
'     - FailureList
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Reset the recorded failure count to zero
        FailureCount = 0

    'Initialize the failure-list output only when the caller requested it
        If CaptureFailureList Then
            FailureList = Empty
        End If

End Sub

Private Sub UI_AddFailureToResult(ByRef Succeeded As Boolean, ByRef FailureCount _
    As Long, ByRef FailureList As Variant, ByVal CaptureFailureList As Boolean, _
    ByVal Stage As String, ByVal Detail As String)

'
'==============================================================================
'                          UI_AddFailureToResult
'------------------------------------------------------------------------------
' PURPOSE
'   Record a failure into the standard-module result buffers used by
'   UI_SetExcelUI_WithResult
'
' WHY THIS EXISTS
'   The module does not depend on a dedicated result class, so failures need to
'   be accumulated through plain standard-module constructs:
'     - a Boolean success flag
'     - a Long failure count
'     - an optional 1-based String array of failure entries
'
'   This helper centralizes that logic and preserves insertion order
'
' INPUTS
'   Succeeded
'     Set to FALSE once a failure is recorded
'
'   FailureCount
'     Incremented for each recorded failure
'
'   FailureList
'     Optional Variant carrying a 1-based String array of recorded failures
'
'   CaptureFailureList
'     TRUE  => append the failure text into FailureList
'     FALSE => update only Succeeded and FailureCount
'
'   Stage
'     Logical stage, element, or operation associated with the failure
'
'   Detail
'     Failure detail to append
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Sets Succeeded to FALSE
'   - Increments FailureCount
'   - When requested, appends:
'         Stage & " | " & Detail
'     to a 1-based String array stored in FailureList
'
' ERROR POLICY
'   - Does NOT raise
'
' UPDATED
'   2026-04-19
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
    'Build the formatted failure text once
        FailureText = Stage & " | " & Detail

'------------------------------------------------------------------------------
' UPDATE RESULT STATUS
'------------------------------------------------------------------------------
    'Mark the overall result as unsuccessful
        Succeeded = False

    'Increment the recorded failure count
        FailureCount = FailureCount + 1

'------------------------------------------------------------------------------
' APPEND FAILURE TEXT
'------------------------------------------------------------------------------
    'Append the formatted failure text only when the caller requested the
    'failure-list output
        If CaptureFailureList Then

            'Allocate the first entry when the failure list is still empty
                If IsEmpty(FailureList) Then
                    ReDim Arr(1 To 1)

            'Otherwise, expand the existing 1-based array while preserving
            'previous entries
                Else
                    Arr = FailureList
                    ReDim Preserve Arr(1 To FailureCount)
                End If

            'Store the new failure entry at the current 1-based position
                Arr(FailureCount) = FailureText

            'Write the expanded array back into the Variant output
                FailureList = Arr

        End If

End Sub

Private Sub UI_BeginQuietUIUpdate(ByRef OldScreenUpdating As Boolean, ByRef _
    QuietModeChanged As Boolean)

'
'==============================================================================
'                          UI_BeginQuietUIUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Enter a small best-effort quiet-update scope by suppressing
'   Application.ScreenUpdating when possible
'
' WHY THIS EXISTS
'   Many object-model UI writes can cause visible worksheet redraw
'   Temporarily disabling ScreenUpdating reduces flicker for those surfaces,
'   even though it cannot fully suppress Ribbon or WinAPI non-client refresh
'
' INPUTS / OUTPUTS
'   OldScreenUpdating
'     Receives the prior Application.ScreenUpdating state
'
'   QuietModeChanged
'     Receives TRUE only when this helper actually changed ScreenUpdating
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise
'   - Best-effort only
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected issue while entering quiet mode
        On Error Resume Next

    'Capture the current ScreenUpdating state
        OldScreenUpdating = Application.ScreenUpdating

    'Initialize the changed flag
        QuietModeChanged = False

'------------------------------------------------------------------------------
' APPLY QUIET MODE
'------------------------------------------------------------------------------
    'Disable ScreenUpdating only when it is currently enabled
        If OldScreenUpdating Then
            Application.ScreenUpdating = False
            QuietModeChanged = True
        End If

End Sub

Private Sub UI_EndQuietUIUpdate(ByVal OldScreenUpdating As Boolean, ByVal _
    QuietModeChanged As Boolean)

'
'==============================================================================
'                           UI_EndQuietUIUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Leave the quiet-update scope created by UI_BeginQuietUIUpdate
'
' WHY THIS EXISTS
'   ScreenUpdating should be restored only when this module actually changed it
'
' INPUTS
'   OldScreenUpdating
'     Previously captured Application.ScreenUpdating state
'
'   QuietModeChanged
'     TRUE when UI_BeginQuietUIUpdate changed ScreenUpdating
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Does NOT raise
'   - Best-effort only
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected issue while leaving quiet mode
        On Error Resume Next

'------------------------------------------------------------------------------
' RESTORE PRIOR STATE
'------------------------------------------------------------------------------
    'Restore ScreenUpdating only when this module previously changed it
        If QuietModeChanged Then
            Application.ScreenUpdating = OldScreenUpdating
        End If

End Sub

Private Function UI_TrySetRibbonVisibleIfNeeded(ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                     UI_TrySetRibbonVisibleIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to update Ribbon visibility only when the current visible state
'   differs from the requested target
'
' WHY THIS EXISTS
'   Avoiding no-op Ribbon writes can slightly reduce visible UI churn and keeps
'   the apply path cleaner
'
' INPUTS
'   IsVisible
'     Requested Ribbon visibility
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => Ribbon is already in the requested state or was updated
'   FALSE => Ribbon update failed
'
' BEHAVIOR
'   - Tries to read the current Ribbon visibility
'   - Skips the write path when the current state already matches the target
'   - Falls back to the actual write path when the current state differs or
'     could not be read
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - UI_TryGetRibbonVisible
'   - UI_TrySetRibbonVisible
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentVisible      As Boolean   'Current Ribbon visibility when readable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize the default result state
        UI_TrySetRibbonVisibleIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT NO-OP
'------------------------------------------------------------------------------
    'When current Ribbon visibility can be read and already matches the target,
    'skip the write path entirely
        If UI_TryGetRibbonVisible(CurrentVisible, FailMsg) Then
            If CurrentVisible = IsVisible Then
                UI_TrySetRibbonVisibleIfNeeded = True
                GoTo SafeExit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY RIBBON WRITE
'------------------------------------------------------------------------------
    'Clear any prior read diagnostic and attempt the actual write
        FailMsg = vbNullString
        UI_TrySetRibbonVisibleIfNeeded = UI_TrySetRibbonVisible(IsVisible, _
            FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

Private Function UI_TrySetTitleBarVisibleIfNeeded(ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                    UI_TrySetTitleBarVisibleIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to update TitleBar visibility only when the current visible state
'   differs from the requested target
'
' WHY THIS EXISTS
'   Avoiding no-op title-bar writes reduces unnecessary non-client frame
'   refresh attempts
'
' INPUTS
'   IsVisible
'     Requested TitleBar visibility
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => TitleBar is already in the requested state or was updated
'   FALSE => TitleBar update failed
'
' BEHAVIOR
'   - Tries to read the current TitleBar visibility
'   - Skips the write path when the current state already matches the target
'   - Falls back to the actual write path when the current state differs or
'     could not be read
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - UI_TryGetTitleBarVisible
'   - UI_TrySetTitleBarVisible
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentVisible      As Boolean   'Current TitleBar visibility when readable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize the default result state
        UI_TrySetTitleBarVisibleIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT NO-OP
'------------------------------------------------------------------------------
    'When current TitleBar visibility can be read and already matches the
    'target, skip the write path entirely
        If UI_TryGetTitleBarVisible(CurrentVisible, FailMsg) Then
            If CurrentVisible = IsVisible Then
                UI_TrySetTitleBarVisibleIfNeeded = True
                GoTo SafeExit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY TITLE-BAR WRITE
'------------------------------------------------------------------------------
    'Clear any prior read diagnostic and attempt the actual write
        FailMsg = vbNullString
        UI_TrySetTitleBarVisibleIfNeeded = UI_TrySetTitleBarVisible(IsVisible, _
            FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

Private Function UI_TrySetBooleanPropertyIfNeeded(ByVal Target As Object, ByVal _
    PropertyName As String, ByVal NewValue As Boolean, ByRef FailMsg As String) As _
    Boolean

'
'==============================================================================
'                   UI_TrySetBooleanPropertyIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to assign a Boolean property only when the current state differs
'   from the requested target
'
' WHY THIS EXISTS
'   Avoiding no-op property writes reduces unnecessary redraw and keeps the UI
'   application path quieter when the property already matches the target
'
' INPUTS
'   Target
'     Object exposing the target Boolean property
'
'   PropertyName
'     Name of the Boolean property to assign
'
'   NewValue
'     Requested Boolean value
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => property is already in the requested state or was updated
'   FALSE => property update failed
'
' BEHAVIOR
'   - Tries to read the current property value
'   - Skips the write path when the current state already matches the target
'   - Falls back to the actual write path when the current state differs or
'     could not be read
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - UI_TryGetBooleanProperty
'   - UI_TrySetBooleanProperty
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentValue        As Boolean   'Current property value when readable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize the default result state
        UI_TrySetBooleanPropertyIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT NO-OP
'------------------------------------------------------------------------------
    'When current property state can be read and already matches the target,
    'skip the write path entirely
        If UI_TryGetBooleanProperty(Target, PropertyName, CurrentValue, FailMsg) _
            Then
            If CurrentValue = NewValue Then
                UI_TrySetBooleanPropertyIfNeeded = True
                GoTo SafeExit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY PROPERTY WRITE
'------------------------------------------------------------------------------
    'Clear any prior read diagnostic and attempt the actual write
        FailMsg = vbNullString
        UI_TrySetBooleanPropertyIfNeeded = UI_TrySetBooleanProperty(Target, _
            PropertyName, NewValue, FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

Private Function UI_TryGetRibbonVisible(ByRef IsVisible As Boolean, ByRef _
    FailMsg As String) As Boolean

'
'==============================================================================
'                         UI_TryGetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read current Ribbon visibility
'
' WHY THIS EXISTS
'   The Ribbon is not best treated as a simple direct property read, so the
'   module uses a dedicated best-effort reader
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
' BEHAVIOR
'   - First attempts CommandBars("Ribbon").Visible
'   - Falls back to an Excel4 macro read when needed
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - Application.CommandBars
'   - Application.ExecuteExcel4Macro
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant   'Fallback Excel4 macro result

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize outputs and default result
        UI_TryGetRibbonVisible = False
        IsVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' TRY COMMANDBARS
'------------------------------------------------------------------------------
    'Attempt to read Ribbon visibility from the CommandBars collection
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
' TRY EXCEL4 MACRO FALLBACK
'------------------------------------------------------------------------------
    'Attempt a fallback read using an Excel4 macro
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
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

Private Function UI_TryGetTitleBarVisible(ByRef IsVisible As Boolean, ByRef _
    FailMsg As String) As Boolean

'
'==============================================================================
'                        UI_TryGetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read current title-bar visibility for the Excel window
'   represented by Application.Hwnd
'
' WHY THIS EXISTS
'   Title-bar state is controlled through WinAPI in this module, so the module
'   also uses a corresponding WinAPI-based read helper
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
' BEHAVIOR
'   - Reads the current main-window style through the appropriate WinAPI path
'   - Treats the caption style bit as the title-bar visibility signal
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - GetWindowLong / GetWindowLongPtr
'   - GetLastError
'   - SetLastError
'
' UPDATED
'   2026-04-19
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
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize outputs and default result
        UI_TryGetTitleBarVisible = False
        IsVisible = False
        FailMsg = vbNullString

    'Read the Excel window handle
        xlHnd = Application.hWnd

    'Reject an invalid window handle deterministically
        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ WINDOW STYLE
'------------------------------------------------------------------------------
    'Clear last-error state before the API call
        SetLastError 0

#If VBA7 Then
    #If Win64 Then

        'Read the current window style using the 64-bit API
            StyleValue = GetWindowLongPtr(xlHnd, GWL_STYLE)

    #Else

        'Read the current window style using the 32-bit API under VBA7
            StyleValue = GetWindowLong(xlHnd, GWL_STYLE)

    #End If
#Else

    'Read the current window style using the legacy 32-bit API
        StyleValue = GetWindowLong(xlHnd, GWL_STYLE)

#End If

    'Read the Win32 last-error value immediately after the API call
        LastErr = GetLastError

    'Treat zero plus nonzero last error as failure
        If StyleValue = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' MAP STYLE TO TITLE-BAR VISIBILITY
'------------------------------------------------------------------------------
    'Treat the caption style bit as the title-bar visibility signal
        IsVisible = ((StyleValue And WS_CAPTION) <> 0)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after a valid style read
        UI_TryGetTitleBarVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

Private Function UI_TryGetBooleanProperty(ByVal Target As Object, ByVal _
    PropertyName As String, ByRef ValueOut As Boolean, ByRef FailMsg As String) As _
    Boolean

'
'==============================================================================
'                         UI_TryGetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read a Boolean property from an object using CallByName
'
' WHY THIS EXISTS
'   The module needs shared read helpers both for:
'     - skip-if-already-correct write suppression
'     - explicit capture / reset behavior
'
'   A shared property reader avoids duplicated boilerplate
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
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant   'Late-bound property value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize outputs and default result
        UI_TryGetBooleanProperty = False
        ValueOut = False
        FailMsg = vbNullString

    'Reject invalid object input deterministically
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo SafeExit
        End If

    'Reject empty property name deterministically
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo SafeExit
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
    'Mark success after property access completes
        UI_TryGetBooleanProperty = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure string without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

Private Function UI_TrySetTitleBarVisible(ByVal IsVisible As Boolean, ByRef _
    FailMsg As String) As Boolean

'
'==============================================================================
'                           UI_TrySetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to show or hide the title bar of the Excel main window represented
'   by Application.Hwnd by updating the window style and refreshing the
'   non-client frame
'
' WHY THIS EXISTS
'   Excel does not expose direct title-bar visibility control in the object
'   model, so the project must update the underlying window style via WinAPI
'
'   To improve the visual result when hiding, this routine also removes the
'   sizing frame (WS_THICKFRAME)
'
'   The original window style is snapshotted and restored exactly when showing
'   again
'
' INPUTS
'   IsVisible
'     TRUE  => restore the original snapshotted Excel main-window style
'     FALSE => hide title bar and related frame controls
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => title-bar update succeeded
'   FALSE => title-bar update failed
'
' BEHAVIOR
'   - Reads the current main-window style
'   - Snapshots the original style against the current Application.Hwnd
'   - Restores the exact original style when showing again
'   - Refreshes the non-client frame after a style change
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' NOTES
'   - Windows only
'   - While hidden, the Excel window is intentionally less frame-like and may
'     not be user-resizable in the normal way
'   - This routine intentionally does NOT toggle Application.DisplayFullScreen
'
' DEPENDENCIES
'   - UI_TryGetWindowStyle
'   - UI_TrySetWindowStyle
'   - UI_TryRefreshWindowFrame
'
' UPDATED
'   2026-04-19
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
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize the default result state
        UI_TrySetTitleBarVisible = False
        FailMsg = vbNullString

    'Read the Excel main-window handle
        xlHnd = Application.hWnd

    'Reject an invalid window handle deterministically
        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ: CURRENT WINDOW STYLE
'------------------------------------------------------------------------------
    'Read the current main-window style using the bitness-safe wrapper
        If Not UI_TryGetWindowStyle(xlHnd, CurrentStyle, FailMsg) Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' SNAPSHOT: ORIGINAL STYLE FOR CURRENT HWND
'------------------------------------------------------------------------------
    'Snapshot the original main-window style whenever no snapshot exists yet or
    'the current Application.Hwnd differs from the one previously captured
        If (Not m_HasOriginalMainWindowStyle) Or (m_OriginalMainWindowHwnd <> _
            xlHnd) Then
            m_OriginalMainWindowStyle = CurrentStyle
            m_OriginalMainWindowHwnd = xlHnd
            m_HasOriginalMainWindowStyle = True
        End If

'------------------------------------------------------------------------------
' COMPUTE: UPDATED WINDOW STYLE
'------------------------------------------------------------------------------
    'Restore the exact original snapshotted style when the caller requests a
    'visible title bar
        If IsVisible Then

            'Use the exact captured original style whenever it belongs to the
            'current window handle
                If m_HasOriginalMainWindowStyle And m_OriginalMainWindowHwnd = _
                    xlHnd Then
                    NewStyle = m_OriginalMainWindowStyle

            'Fall back to a conservative visible-frame composition only if the
            'original style is not available for the current handle
                Else
                    NewStyle = CurrentStyle
                    NewStyle = NewStyle Or WS_SYSMENU
                    NewStyle = NewStyle Or WS_MINIMIZEBOX
                    NewStyle = NewStyle Or WS_MAXIMIZEBOX
                    NewStyle = NewStyle Or WS_CAPTION
                    NewStyle = NewStyle Or WS_THICKFRAME
                End If

    'Remove the frame-related bits when the caller requests a hidden title bar
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
    'Skip the write path entirely when no style change is required
        If NewStyle = CurrentStyle Then
            UI_TrySetTitleBarVisible = True
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' APPLY: UPDATED WINDOW STYLE
'------------------------------------------------------------------------------
    'Write the updated main-window style using the bitness-safe wrapper
        If Not UI_TrySetWindowStyle(xlHnd, NewStyle, FailMsg) Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' REFRESH: NON-CLIENT FRAME
'------------------------------------------------------------------------------
    'Force Windows to recalculate and repaint the frame after the style change
        If Not UI_TryRefreshWindowFrame(xlHnd, FailMsg) Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful only after all required steps complete
        UI_TrySetTitleBarVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Normal termination point
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

Private Function UI_TrySetRibbonVisible(ByVal IsVisible As Boolean, ByRef _
    FailMsg As String) As Boolean

'
'==============================================================================
'                           UI_TrySetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to show or hide the Ribbon using Excel4 macro execution
'
' WHY THIS EXISTS
'   The Ribbon is not exposed through a simple Application Boolean property, so
'   a legacy but commonly used Excel4 macro call is required for direct control
'
' INPUTS
'   IsVisible
'     TRUE  => show Ribbon
'     FALSE => hide Ribbon
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => Ribbon update succeeded
'   FALSE => Ribbon update failed
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - Application.ExecuteExcel4Macro
'
' NOTES
'   - Availability may vary by Excel host / configuration
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MacroText           As String    'Excel4 macro text controlling Ribbon visibility

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize the default result state
        UI_TrySetRibbonVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' BUILD: EXCEL4 MACRO TEXT
'------------------------------------------------------------------------------
    'Build the exact macro text required to show or hide the Ribbon
        If IsVisible Then
            MacroText = "Show.TOOLBAR(""Ribbon"",True)"
        Else
            MacroText = "Show.TOOLBAR(""Ribbon"",False)"
        End If

'------------------------------------------------------------------------------
' APPLY: RIBBON VISIBILITY
'------------------------------------------------------------------------------
    'Execute the Ribbon visibility macro through Excel's legacy macro engine
        Application.ExecuteExcel4Macro MacroText

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful after macro execution completes
        UI_TrySetRibbonVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

Private Function UI_TrySetBooleanProperty(ByVal Target As Object, ByVal _
    PropertyName As String, ByVal NewValue As Boolean, ByRef FailMsg As String) As _
    Boolean

'
'==============================================================================
'                           UI_TrySetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to assign a Boolean property on an object using a common,
'   best-effort helper
'
' WHY THIS EXISTS
'   UI_SetExcelUI sets several Boolean properties across different object types
'   such as Application and Window
'
'   A shared helper avoids duplicating identical property-write error-handling
'   logic for each target property
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
'   - Uses CallByName with vbLet to assign the property
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' NOTES
'   - Intended for Application / Window Boolean property writes in this module
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize the default result state
        UI_TrySetBooleanProperty = False
        FailMsg = vbNullString

    'Reject a missing target object deterministically
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo SafeExit
        End If

    'Reject an empty property name deterministically
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' APPLY: PROPERTY WRITE
'------------------------------------------------------------------------------
    'Assign the requested Boolean value using late-bound property assignment
        CallByName Target, PropertyName, VbLet, NewValue

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful after the property write completes
        UI_TrySetBooleanProperty = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

#If VBA7 Then
Private Function UI_TryGetWindowStyle(ByVal hWnd As LongPtr, ByRef StyleOut As _
    LongPtr, ByRef FailMsg As String) As Boolean
#Else
Private Function UI_TryGetWindowStyle(ByVal hWnd As Long, ByRef StyleOut As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
'                            UI_TryGetWindowStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to read the current GWL_STYLE value using the correct API for the
'   current VBA / Office bitness
'
' WHY THIS EXISTS
'   GetWindowLong / GetWindowLongPtr can validly return zero, so a robust
'   wrapper should use GetLastError to distinguish a real zero from failure
'
' INPUTS
'   hWnd
'     Target window handle
'
'   StyleOut
'     Receives the current GWL_STYLE value on success
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => style read succeeded
'   FALSE => style read failed
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - GetWindowLong / GetWindowLongPtr
'   - GetLastError
'   - SetLastError
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastErr             As Long      'Win32 last-error value read after the API call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize the outputs and default result state
        StyleOut = 0
        FailMsg = vbNullString
        UI_TryGetWindowStyle = False

    'Reject an invalid window handle deterministically
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ: WINDOW STYLE
'------------------------------------------------------------------------------
    'Clear the Win32 last-error state before calling the API so a valid zero
    'return can later be distinguished from failure
        SetLastError 0

#If VBA7 Then
    #If Win64 Then

        'Read the style with the 64-bit API in 64-bit Office / VBA
            StyleOut = GetWindowLongPtr(hWnd, GWL_STYLE)

    #Else

        'Read the style with the 32-bit API in VBA7 32-bit Office
            StyleOut = GetWindowLong(hWnd, GWL_STYLE)

    #End If
#Else

    'Read the style with the legacy 32-bit API
        StyleOut = GetWindowLong(hWnd, GWL_STYLE)

#End If

    'Read the Win32 last-error value immediately after the API call
        LastErr = GetLastError

    'Treat zero plus nonzero last error as an API failure
        If StyleOut = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful after a valid style read
        UI_TryGetWindowStyle = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

#If VBA7 Then
Private Function UI_TrySetWindowStyle(ByVal hWnd As LongPtr, ByVal NewStyle As _
    LongPtr, ByRef FailMsg As String) As Boolean
#Else
Private Function UI_TrySetWindowStyle(ByVal hWnd As Long, ByVal NewStyle As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
'                            UI_TrySetWindowStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to write the GWL_STYLE value using the correct API for the current
'   VBA / Office bitness
'
' WHY THIS EXISTS
'   SetWindowLong / SetWindowLongPtr can validly return zero, so a robust
'   wrapper should use GetLastError to distinguish a real previous zero from
'   failure
'
' INPUTS
'   hWnd
'     Target window handle
'
'   NewStyle
'     New GWL_STYLE value to write
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => style write succeeded
'   FALSE => style write failed
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - SetWindowLong / SetWindowLongPtr
'   - GetLastError
'   - SetLastError
'
' UPDATED
'   2026-04-19
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
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize the default result state
        FailMsg = vbNullString
        UI_TrySetWindowStyle = False

    'Reject an invalid window handle deterministically
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' WRITE: WINDOW STYLE
'------------------------------------------------------------------------------
    'Clear the Win32 last-error state before calling the API so a valid zero
    'return can later be distinguished from failure
        SetLastError 0

#If VBA7 Then
    #If Win64 Then

        'Write the style with the 64-bit API in 64-bit Office / VBA
            PrevStyle = SetWindowLongPtr(hWnd, GWL_STYLE, NewStyle)

    #Else

        'Write the style with the 32-bit API in VBA7 32-bit Office
            PrevStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)

    #End If
#Else

    'Write the style with the legacy 32-bit API
        PrevStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)

#End If

    'Read the Win32 last-error value immediately after the API call
        LastErr = GetLastError

    'Treat zero plus nonzero last error as an API failure
        If PrevStyle = 0 And LastErr <> 0 Then
            FailMsg = "SetWindowLong/SetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful after a valid style write
        UI_TrySetWindowStyle = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

#If VBA7 Then
Private Function UI_TryRefreshWindowFrame(ByVal hWnd As LongPtr, ByRef FailMsg _
    As String) As Boolean
#Else
Private Function UI_TryRefreshWindowFrame(ByVal hWnd As Long, ByRef FailMsg As _
    String) As Boolean
#End If

'
'==============================================================================
'                           UI_TryRefreshWindowFrame
'------------------------------------------------------------------------------
' PURPOSE
'   Attempt to force Windows to repaint the non-client frame of the specified
'   window after a style change
'
' WHY THIS EXISTS
'   Updating GWL_STYLE alone is not always visually reflected immediately
'   SetWindowPos with SWP_FRAMECHANGED is the standard way to notify Windows
'   that the frame should be recalculated and repainted
'
' INPUTS
'   hWnd
'     Target window handle
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => frame refresh succeeded
'   FALSE => frame refresh failed
'
' BEHAVIOR
'   - Uses the canonical no-move / no-size / no-z-order refresh pattern
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - SetWindowPos
'   - GetLastError
'   - SetLastError
'
' UPDATED
'   2026-04-19
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
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize the default result state
        FailMsg = vbNullString
        UI_TryRefreshWindowFrame = False

    'Reject an invalid window handle deterministically
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' REFRESH: NON-CLIENT FRAME
'------------------------------------------------------------------------------
    'Clear the Win32 last-error state before calling the API
        SetLastError 0

    'Force Windows to recalculate and repaint the non-client frame without
    'moving, resizing, or reordering the target window
        ApiOK = SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
            SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED)

    'Read the Win32 last-error value immediately after the API call
        LastErr = GetLastError

    'Reject API failure deterministically and include the Win32 error code
        If ApiOK = 0 Then
            FailMsg = "SetWindowPos failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN: SUCCESS
'------------------------------------------------------------------------------
    'Mark the operation as successful after a valid frame refresh
        UI_TryRefreshWindowFrame = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Leave through the normal termination path
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return a descriptive failure message without raising
        FailMsg = UI_BuildRuntimeErrorText

End Function

Private Function UI_VisibilityToBoolean(ByVal Visibility As UIVisibility) As _
    Boolean

'
'==============================================================================
'                           UI_VisibilityToBoolean
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a tri-state visibility enum value into the explicit Boolean visible
'   state required by Excel properties and internal helpers
'
' WHY THIS EXISTS
'   Public callers use UIVisibility values, while Excel object-model
'   properties and internal helpers require a Boolean visible / hidden state
'
' INPUTS
'   Visibility
'     Expected values:
'       - UI_Hide
'       - UI_Show
'
' RETURNS
'   TRUE  => visible
'   FALSE => hidden
'
' BEHAVIOR
'   - UI_Show maps to TRUE
'   - Any other value maps to FALSE
'
' ERROR POLICY
'   - Does NOT raise
'
' NOTES
'   - Callers should only invoke this helper after excluding
'     UI_LeaveUnchanged
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN: BOOLEAN VISIBILITY
'------------------------------------------------------------------------------
    'Convert explicit SHOW to TRUE; otherwise return FALSE
        UI_VisibilityToBoolean = (Visibility = UI_Show)

End Function

Private Function UI_BuildRuntimeErrorText() As String

'
'==============================================================================
'                           UI_BuildRuntimeErrorText
'------------------------------------------------------------------------------
' PURPOSE
'   Build a consistent runtime diagnostic string from the active Err object
'
' WHY THIS EXISTS
'   Several procedures in this module use identical failure-text construction
'   A shared helper avoids duplicated formatting logic and keeps diagnostics
'   consistent
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
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected issue while formatting the diagnostic
        On Error Resume Next

'------------------------------------------------------------------------------
' BUILD: RUNTIME ERROR TEXT
'------------------------------------------------------------------------------
    'Build a consistent diagnostic string from the current Err state
        UI_BuildRuntimeErrorText = CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

End Function

Private Sub UI_LogFailure(ByVal ProcName As String, ByVal Stage As String, ByVal _
    Detail As String)

'
'==============================================================================
'                                UI_LogFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Write a consistent diagnostic line to the Immediate Window
'
' WHY THIS EXISTS
'   The module uses fail-soft behavior and needs a single place to format
'   diagnostic output consistently
'
' INPUTS
'   ProcName
'     Procedure name associated with the failure
'
'   Stage
'     Logical stage or element associated with the failure
'
'   Detail
'     Failure detail to append
'
' RETURNS
'   None
'
' ERROR POLICY
'   - Suppresses any unexpected logging failure locally
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected logging failure
        On Error Resume Next

'------------------------------------------------------------------------------
' WRITE: DIAGNOSTIC LINE
'------------------------------------------------------------------------------
    'Write a consistent fail-soft diagnostic line to the Immediate Window
        Debug.Print ProcName & " failed @ " & Stage & " | " & Detail

End Sub



