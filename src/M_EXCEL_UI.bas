Attribute VB_Name = "M_EXCEL_UI"
Option Explicit

'
'==============================================================================
'                    MODULE: K_UI_EXCEL_SHELL_CONTROL
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
'   - K_HideExcelUI                 Convenience wrapper: hide all managed UI
'   - K_ShowExcelUI                 Convenience wrapper: show all managed UI
'
' INTERNAL SUPPORT
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
'   - Unexpected errors are logged to the Immediate Window.
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
'   - This module does NOT snapshot and restore prior Excel object-model UI
'     state.
'   - K_ShowExcelUI means "show all managed UI", not "restore previous state".
'   - K_SetExcelUI is the preferred entry point for selective control.
'   - Ribbon control relies on Application.ExecuteExcel4Macro.
'   - Title-bar control affects the Excel window represented by
'     Application.Hwnd, not a user-specific saved UI state.
'   - The original main-window style is snapshotted once, on first title-bar
'     manipulation, so it can later be restored exactly.
'
' UPDATED
'   2026-04-09
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.0
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
#Else
    Private m_OriginalMainWindowStyle As Long         'Snapshotted original Excel main-window style
#End If

Private m_HasOriginalMainWindowStyle As Boolean       'TRUE when original style has been captured

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
'   - K_TrySetRibbonVisible
'   - K_TrySetBooleanProperty
'   - K_VisibilityToBoolean
'   - K_TrySetTitleBarVisible
'   - K_LogFailure
'
' NOTES
'   - This is the preferred entry point for selective UI control.
'   - Changes affect the current Excel instance, not only the active workbook.
'
' UPDATED
'   2026-04-09
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim W                   As Window    'Workbook window in current Excel instance
    Dim ShowFlag            As Boolean   'Converted Boolean visibility target
    Dim Msg                 As String    'Element-level diagnostic message
    Const PROC As String = "K_SetExcelUI"   'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY: APPLICATION-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Apply Ribbon visibility when requested.
        If Ribbon <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(Ribbon)

            'Attempt the Ribbon update and log any element-level failure
            'without interrupting later operations.
                If Not K_TrySetRibbonVisible(ShowFlag, Msg) Then
                    K_LogFailure PROC, "Ribbon", Msg
                End If

        End If

    'Apply status-bar visibility when requested.
        If StatusBar <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(StatusBar)

            'Attempt the property write and log any element-level failure
            'without interrupting later operations.
                If Not K_TrySetBooleanProperty(Application, "DisplayStatusBar", ShowFlag, Msg) Then
                    K_LogFailure PROC, "StatusBar", Msg
                End If

        End If

    'Apply scroll-bar visibility when requested.
        If ScrollBars <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(ScrollBars)

            'Attempt the property write and log any element-level failure
            'without interrupting later operations.
                If Not K_TrySetBooleanProperty(Application, "DisplayScrollBars", ShowFlag, Msg) Then
                    K_LogFailure PROC, "ScrollBars", Msg
                End If

        End If

    'Apply formula-bar visibility when requested.
        If FormulaBar <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(FormulaBar)

            'Attempt the property write and log any element-level failure
            'without interrupting later operations.
                If Not K_TrySetBooleanProperty(Application, "DisplayFormulaBar", ShowFlag, Msg) Then
                    K_LogFailure PROC, "FormulaBar", Msg
                End If

        End If

'------------------------------------------------------------------------------
' APPLY: WINDOW-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Process window-scoped UI only when at least one window-level element has
    'been requested for change.
        If Headings <> K_UI_LeaveUnchanged _
        Or WorkbookTabs <> K_UI_LeaveUnchanged _
        Or Gridlines <> K_UI_LeaveUnchanged Then

            'Apply the requested window-level visibility state to each open
            'Excel window in the current instance.
                For Each W In Application.Windows

                    'Apply headings visibility when requested.
                        If Headings <> K_UI_LeaveUnchanged Then

                            'Convert the tri-state enum to the explicit Boolean
                            'state expected by the lower-level helper.
                                ShowFlag = K_VisibilityToBoolean(Headings)

                            'Attempt the property write and log any element-
                            'level failure without interrupting later operations.
                                If Not K_TrySetBooleanProperty(W, "DisplayHeadings", ShowFlag, Msg) Then
                                    K_LogFailure PROC, "Headings [" & W.Caption & "]", Msg
                                End If

                        End If

                    'Apply workbook-tabs visibility when requested.
                        If WorkbookTabs <> K_UI_LeaveUnchanged Then

                            'Convert the tri-state enum to the explicit Boolean
                            'state expected by the lower-level helper.
                                ShowFlag = K_VisibilityToBoolean(WorkbookTabs)

                            'Attempt the property write and log any element-
                            'level failure without interrupting later operations.
                                If Not K_TrySetBooleanProperty(W, "DisplayWorkbookTabs", ShowFlag, Msg) Then
                                    K_LogFailure PROC, "WorkbookTabs [" & W.Caption & "]", Msg
                                End If

                        End If

                    'Apply gridlines visibility when requested.
                        If Gridlines <> K_UI_LeaveUnchanged Then

                            'Convert the tri-state enum to the explicit Boolean
                            'state expected by the lower-level helper.
                                ShowFlag = K_VisibilityToBoolean(Gridlines)

                            'Attempt the property write and log any element-
                            'level failure without interrupting later operations.
                                If Not K_TrySetBooleanProperty(W, "DisplayGridlines", ShowFlag, Msg) Then
                                    K_LogFailure PROC, "Gridlines [" & W.Caption & "]", Msg
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

            'Attempt the title-bar update and log any element-level failure
            'without interrupting the caller.
                If Not K_TrySetTitleBarVisible(ShowFlag, Msg) Then
                    K_LogFailure PROC, "TitleBar", Msg
                End If

        End If

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
        K_LogFailure PROC, "Unexpected", K_BuildRuntimeErrorText

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
'   2026-04-09
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
'   2026-04-09
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
    'Leave quietly through the normal termination path
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write an unexpected-procedure-level diagnostic line without interrupting
    'the caller.
        K_LogFailure "K_ShowExcelUI", "Unexpected", K_BuildRuntimeErrorText
    'Exit quietly after logging
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
'   - K_TrySetRibbonVisible
'   - K_TrySetBooleanProperty
'   - K_TrySetTitleBarVisible
'   - K_VisibilityToBoolean
'   - K_BuildRuntimeErrorText
'   - K_ClearResultBuffer
'   - K_AddFailureToResult
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
    Dim W                   As Window     'Workbook window in current Excel instance
    Dim ShowFlag            As Boolean    'Converted Boolean visibility target
    Dim Msg                 As String     'Element-level diagnostic message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Detect whether the caller supplied the optional failure-list output.
        CaptureFailureList = Not IsMissing(FailureList)

    'Initialize the result buffers in their clean-success state.
        Succeeded = True
        K_ClearResultBuffer FailureCount, FailureList, CaptureFailureList

    'Route unexpected runtime errors to the local failure handler.
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY: APPLICATION-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Apply Ribbon visibility when requested.
        If Ribbon <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(Ribbon)

            'Attempt the Ribbon update and record any element-level failure
            'without interrupting later operations.
                If Not K_TrySetRibbonVisible(ShowFlag, Msg) Then
                    K_AddFailureToResult Succeeded, FailureCount, FailureList, CaptureFailureList, "Ribbon", Msg
                End If

        End If

    'Apply status-bar visibility when requested.
        If StatusBar <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(StatusBar)

            'Attempt the property write and record any element-level failure
            'without interrupting later operations.
                If Not K_TrySetBooleanProperty(Application, "DisplayStatusBar", ShowFlag, Msg) Then
                    K_AddFailureToResult Succeeded, FailureCount, FailureList, CaptureFailureList, "StatusBar", Msg
                End If

        End If

    'Apply scroll-bar visibility when requested.
        If ScrollBars <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(ScrollBars)

            'Attempt the property write and record any element-level failure
            'without interrupting later operations.
                If Not K_TrySetBooleanProperty(Application, "DisplayScrollBars", ShowFlag, Msg) Then
                    K_AddFailureToResult Succeeded, FailureCount, FailureList, CaptureFailureList, "ScrollBars", Msg
                End If

        End If

    'Apply formula-bar visibility when requested.
        If FormulaBar <> K_UI_LeaveUnchanged Then

            'Convert the tri-state enum to the explicit Boolean state expected
            'by the lower-level helper.
                ShowFlag = K_VisibilityToBoolean(FormulaBar)

            'Attempt the property write and record any element-level failure
            'without interrupting later operations.
                If Not K_TrySetBooleanProperty(Application, "DisplayFormulaBar", ShowFlag, Msg) Then
                    K_AddFailureToResult Succeeded, FailureCount, FailureList, CaptureFailureList, "FormulaBar", Msg
                End If

        End If

'------------------------------------------------------------------------------
' APPLY: WINDOW-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Process window-scoped UI only when at least one window-level element has
    'been requested for change.
        If Headings <> K_UI_LeaveUnchanged _
        Or WorkbookTabs <> K_UI_LeaveUnchanged _
        Or Gridlines <> K_UI_LeaveUnchanged Then

            'Apply the requested window-level visibility state to each open
            'Excel window in the current instance.
                For Each W In Application.Windows

                    'Apply headings visibility when requested.
                        If Headings <> K_UI_LeaveUnchanged Then

                            'Convert the tri-state enum to the explicit Boolean
                            'state expected by the lower-level helper.
                                ShowFlag = K_VisibilityToBoolean(Headings)

                            'Attempt the property write and record any element-
                            'level failure without interrupting later operations.
                                If Not K_TrySetBooleanProperty(W, "DisplayHeadings", ShowFlag, Msg) Then
                                    K_AddFailureToResult Succeeded, FailureCount, FailureList, CaptureFailureList, _
                                        "Headings [" & W.Caption & "]", Msg
                                End If

                        End If

                    'Apply workbook-tabs visibility when requested.
                        If WorkbookTabs <> K_UI_LeaveUnchanged Then

                            'Convert the tri-state enum to the explicit Boolean
                            'state expected by the lower-level helper.
                                ShowFlag = K_VisibilityToBoolean(WorkbookTabs)

                            'Attempt the property write and record any element-
                            'level failure without interrupting later operations.
                                If Not K_TrySetBooleanProperty(W, "DisplayWorkbookTabs", ShowFlag, Msg) Then
                                    K_AddFailureToResult Succeeded, FailureCount, FailureList, CaptureFailureList, _
                                        "WorkbookTabs [" & W.Caption & "]", Msg
                                End If

                        End If

                    'Apply gridlines visibility when requested.
                        If Gridlines <> K_UI_LeaveUnchanged Then

                            'Convert the tri-state enum to the explicit Boolean
                            'state expected by the lower-level helper.
                                ShowFlag = K_VisibilityToBoolean(Gridlines)

                            'Attempt the property write and record any element-
                            'level failure without interrupting later operations.
                                If Not K_TrySetBooleanProperty(W, "DisplayGridlines", ShowFlag, Msg) Then
                                    K_AddFailureToResult Succeeded, FailureCount, FailureList, CaptureFailureList, _
                                        "Gridlines [" & W.Caption & "]", Msg
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

            'Attempt the title-bar update and record any element-level failure
            'without interrupting later operations.
                If Not K_TrySetTitleBarVisible(ShowFlag, Msg) Then
                    K_AddFailureToResult Succeeded, FailureCount, FailureList, CaptureFailureList, "TitleBar", Msg
                End If

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
    'Capture the unexpected procedure-level failure in the result buffers.
        K_AddFailureToResult Succeeded, FailureCount, FailureList, CaptureFailureList, "Unexpected", K_BuildRuntimeErrorText

    'Return the overall success flag after recording the unexpected failure.
        K_SetExcelUI_WithResult = Succeeded

    'Leave quietly through the normal termination path.
        Resume SafeExit

End Function

Private Sub K_ClearResultBuffer( _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean)

'
'==============================================================================
'                           K_ClearResultBuffer
'------------------------------------------------------------------------------
' PURPOSE
'   Initialize the ByRef result buffers used by K_SetExcelUI_WithResult.
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
'   - Snapshots the original style once, on first use.
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
'   2026-04-09
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

    'Snapshot the original main-window style once so SHOW can restore it
    'exactly later.
        If Not m_HasOriginalMainWindowStyle Then
            m_OriginalMainWindowStyle = CurrentStyle
            m_HasOriginalMainWindowStyle = True
        End If

'------------------------------------------------------------------------------
' COMPUTE: UPDATED WINDOW STYLE
'------------------------------------------------------------------------------
    'Restore the exact original snapshotted style when the caller requests a
    'visible title bar.
        If IsVisible Then

            'Use the exact captured original style whenever it is available.
                If m_HasOriginalMainWindowStyle Then
                    NewStyle = m_OriginalMainWindowStyle

            'Fall back to a conservative "visible frame" composition only if
            'the original style has not somehow been captured.
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
'   2026-04-09
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
'   2026-04-09
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
'   2026-04-09
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
'   2026-04-09
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
'   2026-04-09
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
'   2026-04-09
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
'   2026-04-09
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
'   2026-04-09
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



