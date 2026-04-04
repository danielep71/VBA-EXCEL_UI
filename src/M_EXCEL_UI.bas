Attribute VB_Name = "M_EXCEL_UI"
Option Explicit

'
'==============================================================================
'                         MODULE: EXCEL UI / SHELL CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Centralize visibility control for the Excel UI elements managed by this
'   module, combining:
'     - Excel object-model UI elements
'     - WinAPI-based title-bar control for the Excel window represented by
'       Application.Hwnd
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
'   - K_TrySetWindowStyle
'   - K_TryGetWindowStyle
'   - K_TryRefreshWindowFrame
'   - K_VisibilityToBoolean
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
'       * applied to the Excel window represented by Application.Hwnd through
'         WinAPI style update + non-client frame refresh
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
'   - This module does NOT snapshot and restore prior UI state.
'   - K_ShowExcelUI means "show all managed UI", not "restore previous state".
'   - K_SetExcelUI is the preferred entry point for selective control.
'   - Ribbon control relies on Application.ExecuteExcel4Macro.
'   - Title-bar control affects the Excel window represented by
'     Application.Hwnd, not a user-specific saved UI state.
'
' UPDATED
'   2026-04-04
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

Private Const WS_THICKFRAME          As Long = &H40000  'Resizable sizing frame

#If VBA7 Then
    Private m_OriginalMainWindowStyle As LongPtr        'Snapshotted original Excel window style
#Else
    Private m_OriginalMainWindowStyle As Long           'Snapshotted original Excel window style
#End If

Private m_HasOriginalMainWindowStyle As Boolean         'TRUE when original style has been captured


'------------------------------------------------------------------------------
' DECLARE: CONSTANTS
'------------------------------------------------------------------------------
Private Const GWL_STYLE              As Long = -16       'Window style index

Private Const WS_CAPTION             As Long = &HC00000  'Caption / title bar
Private Const WS_SYSMENU             As Long = &H80000   'System menu
Private Const WS_MAXIMIZEBOX         As Long = &H10000   'Maximize button
Private Const WS_MINIMIZEBOX         As Long = &H20000   'Minimize button

Private Const SWP_NOSIZE             As Long = &H1       'Preserve current size
Private Const SWP_NOMOVE             As Long = &H2       'Preserve current position
Private Const SWP_NOZORDER           As Long = &H4       'Do not change Z order
Private Const SWP_FRAMECHANGED       As Long = &H20      'Repaint non-client frame
Private Const SWP_NOOWNERZORDER      As Long = &H200     'Do not change owner Z order

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
'   A Boolean-based "hide" routine is error-prone because omitted optional
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
'     K_UI_Show             => show the title bar of the Excel window
'                               represented by Application.Hwnd
'     K_UI_Hide             => hide the title bar of the Excel window
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
'   - Applies title-bar visibility to the Excel window represented by
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
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim W                   As Window    'Workbook window in current Excel instance
    Dim ShowFlag            As Boolean   'Converted Boolean visibility target
    Dim Msg                 As String    'Immediate Window diagnostic message
    Const PROC As String = "K_SetExcelUI"    'Procedure name for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY APPLICATION-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Apply Ribbon visibility when requested
        If Ribbon <> K_UI_LeaveUnchanged Then

            'Convert tri-state enum to explicit Boolean visible state
                ShowFlag = K_VisibilityToBoolean(Ribbon)

            'Attempt Ribbon update and log element failure without stopping
                If Not K_TrySetRibbonVisible(ShowFlag, Msg) Then
                    K_LogFailure PROC, "Ribbon", Msg
                End If

        End If

    'Apply status-bar visibility when requested
        If StatusBar <> K_UI_LeaveUnchanged Then

            'Convert tri-state enum to explicit Boolean visible state
                ShowFlag = K_VisibilityToBoolean(StatusBar)

            'Attempt status-bar update and log element failure without stopping
                If Not K_TrySetBooleanProperty(Application, "DisplayStatusBar", ShowFlag, Msg) Then
                    K_LogFailure PROC, "StatusBar", Msg
                End If

        End If

    'Apply scroll-bar visibility when requested
        If ScrollBars <> K_UI_LeaveUnchanged Then

            'Convert tri-state enum to explicit Boolean visible state
                ShowFlag = K_VisibilityToBoolean(ScrollBars)

            'Attempt scroll-bar update and log element failure without stopping
                If Not K_TrySetBooleanProperty(Application, "DisplayScrollBars", ShowFlag, Msg) Then
                    K_LogFailure PROC, "ScrollBars", Msg
                End If

        End If

    'Apply formula-bar visibility when requested
        If FormulaBar <> K_UI_LeaveUnchanged Then

            'Convert tri-state enum to explicit Boolean visible state
                ShowFlag = K_VisibilityToBoolean(FormulaBar)

            'Attempt formula-bar update and log element failure without stopping
                If Not K_TrySetBooleanProperty(Application, "DisplayFormulaBar", ShowFlag, Msg) Then
                    K_LogFailure PROC, "FormulaBar", Msg
                End If

        End If

'------------------------------------------------------------------------------
' APPLY WINDOW-LEVEL UI STATE
'------------------------------------------------------------------------------
    'Apply window-scoped settings only when at least one such element is
    'requested for change
        If Headings <> K_UI_LeaveUnchanged _
        Or WorkbookTabs <> K_UI_LeaveUnchanged _
        Or Gridlines <> K_UI_LeaveUnchanged Then

            'Apply requested visibility states to each open Excel window
                For Each W In Application.Windows

                    'Apply headings visibility when requested
                        If Headings <> K_UI_LeaveUnchanged Then

                            'Convert tri-state enum to explicit Boolean visible state
                                ShowFlag = K_VisibilityToBoolean(Headings)

                            'Attempt headings update and log element failure without stopping
                                If Not K_TrySetBooleanProperty(W, "DisplayHeadings", ShowFlag, Msg) Then
                                    K_LogFailure PROC, "Headings [" & W.Caption & "]", Msg
                                End If

                        End If

                    'Apply workbook-tabs visibility when requested
                        If WorkbookTabs <> K_UI_LeaveUnchanged Then

                            'Convert tri-state enum to explicit Boolean visible state
                                ShowFlag = K_VisibilityToBoolean(WorkbookTabs)

                            'Attempt workbook-tabs update and log element failure without stopping
                                If Not K_TrySetBooleanProperty(W, "DisplayWorkbookTabs", ShowFlag, Msg) Then
                                    K_LogFailure PROC, "WorkbookTabs [" & W.Caption & "]", Msg
                                End If

                        End If

                    'Apply gridlines visibility when requested
                        If Gridlines <> K_UI_LeaveUnchanged Then

                            'Convert tri-state enum to explicit Boolean visible state
                                ShowFlag = K_VisibilityToBoolean(Gridlines)

                            'Attempt gridlines update and log element failure without stopping
                                If Not K_TrySetBooleanProperty(W, "DisplayGridlines", ShowFlag, Msg) Then
                                    K_LogFailure PROC, "Gridlines [" & W.Caption & "]", Msg
                                End If

                        End If

                Next W

        End If

'------------------------------------------------------------------------------
' APPLY TITLE-BAR STATE
'------------------------------------------------------------------------------
    'Apply title-bar visibility when requested
        If TitleBar <> K_UI_LeaveUnchanged Then

            'Convert tri-state enum to explicit Boolean visible state
                ShowFlag = K_VisibilityToBoolean(TitleBar)

            'Attempt title-bar update and log element failure without stopping
                If Not K_TrySetTitleBarVisible(ShowFlag, Msg) Then
                    K_LogFailure PROC, "TitleBar", Msg
                End If

        End If

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
    'Write a diagnostic line without interrupting callers
        K_LogFailure PROC, "Unexpected", _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

    'Exit quietly after logging
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
'   Some workbook solutions want a simple one-call way to suppress the managed
'   Excel shell elements without specifying each element individually.
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
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY HIDE-ALL STATE
'------------------------------------------------------------------------------
    'Hide all managed UI elements
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
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write a diagnostic line without interrupting callers
        K_LogFailure "K_HideExcelUI", "Unexpected", _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

    'Exit quietly after logging
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
'   - It does NOT restore a previously captured user-specific state.
'   - For selective control, use K_SetExcelUI directly.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

'------------------------------------------------------------------------------
' APPLY SHOW-ALL STATE
'------------------------------------------------------------------------------
    'Show all managed UI elements
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
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Write a diagnostic line without interrupting callers
        K_LogFailure "K_ShowExcelUI", "Unexpected", _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
            IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

    'Exit quietly after logging
        Resume SafeExit

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
'   Attempt to show or hide the title bar of the Excel window represented by
'   Application.Hwnd by updating the window style and refreshing the
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
'     TRUE  => restore the original snapshotted Excel window style
'     FALSE => hide title bar / system controls / sizing frame
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE.
'
' RETURNS
'   TRUE  => title-bar update succeeded
'   FALSE => title-bar update failed
'
' NOTES
'   - Windows-only.
'   - While hidden, the Excel window is intentionally less frame-like and may
'     not be user-resizable in the normal way.
'   - This routine intentionally does NOT toggle Application.DisplayFullScreen.
'
' UPDATED
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim xlHnd               As LongPtr   'Excel window handle from Application.Hwnd
    Dim CurrentStyle        As LongPtr   'Current main-window style
    Dim NewStyle            As LongPtr   'Updated main-window style
#Else
    Dim xlHnd               As Long      'Excel window handle from Application.Hwnd
    Dim CurrentStyle        As Long      'Current main-window style
    Dim NewStyle            As Long      'Updated main-window style
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize default failure result
        K_TrySetTitleBarVisible = False
        FailMsg = vbNullString

    'Read the Excel window handle
        xlHnd = Application.hWnd

    'Reject invalid window handle deterministically
        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ CURRENT WINDOW STYLE
'------------------------------------------------------------------------------
    'Read the current window style using the bitness-safe helper
        If Not K_TryGetWindowStyle(xlHnd, CurrentStyle, FailMsg) Then
            GoTo SafeExit
        End If

    'Snapshot the original window style once so SHOW can restore it exactly
        If Not m_HasOriginalMainWindowStyle Then
            m_OriginalMainWindowStyle = CurrentStyle
            m_HasOriginalMainWindowStyle = True
        End If

'------------------------------------------------------------------------------
' COMPUTE UPDATED WINDOW STYLE
'------------------------------------------------------------------------------
    'Restore the exact original snapshotted style when showing
        If IsVisible Then

            'Use the original captured style when available
                If m_HasOriginalMainWindowStyle Then
                    NewStyle = m_OriginalMainWindowStyle
                Else
                    NewStyle = CurrentStyle
                    NewStyle = NewStyle Or WS_SYSMENU
                    NewStyle = NewStyle Or WS_MAXIMIZEBOX
                    NewStyle = NewStyle Or WS_MINIMIZEBOX
                    NewStyle = NewStyle Or WS_CAPTION
                    NewStyle = NewStyle Or WS_THICKFRAME
                End If

        Else

            'Start from the current style and remove frame-related bits
                NewStyle = CurrentStyle
                NewStyle = NewStyle And Not WS_SYSMENU
                NewStyle = NewStyle And Not WS_MAXIMIZEBOX
                NewStyle = NewStyle And Not WS_MINIMIZEBOX
                NewStyle = NewStyle And Not WS_CAPTION
                NewStyle = NewStyle And Not WS_THICKFRAME

        End If

'------------------------------------------------------------------------------
' SHORT-CIRCUIT NO-OP
'------------------------------------------------------------------------------
    'Skip the write when no style change is required
        If NewStyle = CurrentStyle Then
            K_TrySetTitleBarVisible = True
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' WRITE UPDATED WINDOW STYLE
'------------------------------------------------------------------------------
    'Write the updated style using the bitness-safe helper
        If Not K_TrySetWindowStyle(xlHnd, NewStyle, FailMsg) Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' REFRESH NON-CLIENT FRAME
'------------------------------------------------------------------------------
    'Force Windows to repaint the frame without moving or resizing the window
        If Not K_TryRefreshWindowFrame(xlHnd, FailMsg) Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after all operations complete
        K_TrySetTitleBarVisible = True

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
        FailMsg = CStr(Err.Number) & ": " & Err.Description & _
                  IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
                  IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

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
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MacroText           As String    'Excel4 macro text for Ribbon visibility

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize default failure result
        K_TrySetRibbonVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' BUILD MACRO
'------------------------------------------------------------------------------
    'Build the Ribbon visibility macro text explicitly
        If IsVisible Then
            MacroText = "Show.TOOLBAR(""Ribbon"",True)"
        Else
            MacroText = "Show.TOOLBAR(""Ribbon"",False)"
        End If

'------------------------------------------------------------------------------
' EXECUTE MACRO
'------------------------------------------------------------------------------
    'Execute the Ribbon visibility macro
        Application.ExecuteExcel4Macro MacroText

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after macro execution completes
        K_TrySetRibbonVisible = True

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
        FailMsg = CStr(Err.Number) & ": " & Err.Description & _
                  IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
                  IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

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
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize default failure result
        K_TrySetBooleanProperty = False
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
' APPLY PROPERTY WRITE
'------------------------------------------------------------------------------
    'Assign the requested Boolean value using late-bound property assignment
        CallByName Target, PropertyName, VbLet, NewValue

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after property assignment completes
        K_TrySetBooleanProperty = True

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
        FailMsg = CStr(Err.Number) & ": " & Err.Description & _
                  IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
                  IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

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
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastErr             As Long      'Last Win32 error after API call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize outputs and default result
        StyleOut = 0
        FailMsg = vbNullString
        K_TryGetWindowStyle = False

    'Reject invalid input deterministically
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ STYLE
'------------------------------------------------------------------------------
    'Clear last-error state before the API call
        SetLastError 0

#If VBA7 Then
    #If Win64 Then

        'Read the window style using the 64-bit API
            StyleOut = GetWindowLongPtr(hWnd, GWL_STYLE)

    #Else

        'Read the window style using the 32-bit API under VBA7
            StyleOut = GetWindowLong(hWnd, GWL_STYLE)

    #End If
#Else

    'Read the window style using the legacy 32-bit API
        StyleOut = GetWindowLong(hWnd, GWL_STYLE)

#End If

    'Read the Win32 last-error value immediately after the API call
        LastErr = GetLastError

    'Treat zero + nonzero last error as failure
        If StyleOut = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after a valid style read
        K_TryGetWindowStyle = True

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
        FailMsg = CStr(Err.Number) & ": " & Err.Description & _
                  IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
                  IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

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
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim PrevStyle           As LongPtr   'Previous style returned by API
#Else
    Dim PrevStyle           As Long      'Previous style returned by API
#End If
    Dim LastErr             As Long      'Last Win32 error after API call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize default result
        FailMsg = vbNullString
        K_TrySetWindowStyle = False

    'Reject invalid input deterministically
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' WRITE STYLE
'------------------------------------------------------------------------------
    'Clear last-error state before the API call
        SetLastError 0

#If VBA7 Then
    #If Win64 Then

        'Write the window style using the 64-bit API
            PrevStyle = SetWindowLongPtr(hWnd, GWL_STYLE, NewStyle)

    #Else

        'Write the window style using the 32-bit API under VBA7
            PrevStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)

    #End If
#Else

    'Write the window style using the legacy 32-bit API
        PrevStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)

#End If

    'Read the Win32 last-error value immediately after the API call
        LastErr = GetLastError

    'Treat zero + nonzero last error as failure
        If PrevStyle = 0 And LastErr <> 0 Then
            FailMsg = "SetWindowLong/SetWindowLongPtr failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after a valid style write
        K_TrySetWindowStyle = True

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
        FailMsg = CStr(Err.Number) & ": " & Err.Description & _
                  IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
                  IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

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
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ApiOK               As Long      'WinAPI success flag / return code
    Dim LastErr             As Long      'Last Win32 error after API call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the local failure handler
        On Error GoTo Fail

    'Initialize default result
        FailMsg = vbNullString
        K_TryRefreshWindowFrame = False

    'Reject invalid input deterministically
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' REFRESH FRAME
'------------------------------------------------------------------------------
    'Clear last-error state before the API call
        SetLastError 0

    'Force Windows to recalculate and repaint the non-client frame without
    'moving, resizing, or reordering the window
        ApiOK = SetWindowPos( _
                    hWnd, _
                    0, _
                    0, _
                    0, _
                    0, _
                    0, _
                    SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED)

    'Read the Win32 last-error value immediately after the API call
        LastErr = GetLastError

    'Reject API failure deterministically and include the Win32 error code
        If ApiOK = 0 Then
            FailMsg = "SetWindowPos failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Mark success after a valid frame refresh
        K_TryRefreshWindowFrame = True

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
        FailMsg = CStr(Err.Number) & ": " & Err.Description & _
                  IIf(Len(Err.Source) > 0, " | Source: " & Err.Source, vbNullString) & _
                  IIf(Erl <> 0, " | Line: " & CStr(Erl), vbNullString)

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
'   properties require a Boolean visible / hidden state.
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
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN BOOLEAN VISIBILITY
'------------------------------------------------------------------------------
    'Convert explicit "show" to TRUE; otherwise return FALSE
        K_VisibilityToBoolean = (Visibility = K_UI_Show)

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
'   2026-04-04
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect callers from any unexpected logging failure
        On Error Resume Next

'------------------------------------------------------------------------------
' WRITE DIAGNOSTIC LINE
'------------------------------------------------------------------------------
    'Write a consistent diagnostic line to the Immediate Window
        Debug.Print ProcName & " failed @ " & Stage & " | " & Detail

End Sub

