Attribute VB_Name = "M_EXCEL_UI_TITLEBAR"
'==============================================================================
'                      MODULE: M_EXCEL_UI_TITLEBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Own all WinAPI, state, and style-merging responsibilities required to show
'   or hide the Excel main-window title bar.
'
' WHY
'   Isolating the non-client-frame subsystem keeps M_EXCEL_UI focused on its
'   public facade and prevents WinAPI details and mutable handle-specific state
'   from being mixed with object-model and snapshot orchestration.
'
' INTERNAL SURFACE
'   - UI_TryGetTitleBarVisible
'   - UI_TrySetTitleBarVisibleIfNeeded
'   - UI_InternalMergeTitleBarStyleBits
'
' BEHAVIOR
'   - Owns only WS_CAPTION, WS_SYSMENU, WS_THICKFRAME, WS_MINIMIZEBOX, and
'     WS_MAXIMIZEBOX.
'   - Captures owned bits per Application.Hwnd.
'   - Merges only owned bits into the current GWL_STYLE value.
'   - Preserves unrelated style changes.
'   - Refreshes the non-client frame only after an actual style write.
'
' ERROR POLICY
'   - Internal entry points are fail-soft and return FALSE plus diagnostic text.
'   - No user-interface messages are displayed.
'
' PLATFORM / COMPATIBILITY
'   - Windows only.
'   - Supports 32-bit and 64-bit Office / VBA through conditional compilation.
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
' DECLARE: WIN32 / WIN64 API
'------------------------------------------------------------------------------
#If VBA7 Then

    #If Win64 Then

        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias _
            "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As _
            LongPtr

        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias _
            "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, _
            ByVal dwNewLong As LongPtr) As LongPtr

    #Else

        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias _
            "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long

        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias _
            "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, _
            ByVal dwNewLong As Long) As Long

    #End If

    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As LongPtr, _
        ByVal hWndInsertAfter As LongPtr, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal uFlags As Long) As Long

    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

    Private Declare PtrSafe Sub SetLastError Lib "kernel32" ( _
        ByVal dwErrCode As Long)

#Else

    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long) As Long

    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, _
        ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long

    Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal uFlags As Long) As Long

    Private Declare Function GetLastError Lib "kernel32" () As Long

    Private Declare Sub SetLastError Lib "kernel32" ( _
        ByVal dwErrCode As Long)

#End If

'------------------------------------------------------------------------------
' DECLARE: PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const GWL_STYLE          As Long = -16

    Private Const WS_CAPTION         As Long = &HC00000
    Private Const WS_SYSMENU         As Long = &H80000
    Private Const WS_THICKFRAME      As Long = &H40000
    Private Const WS_MINIMIZEBOX     As Long = &H20000
    Private Const WS_MAXIMIZEBOX     As Long = &H10000

    'Exact GWL_STYLE bits owned by title-bar control.
    Private Const TITLEBAR_OWNED_STYLE_MASK As Long = &HCF0000

    Private Const SWP_NOSIZE         As Long = &H1
    Private Const SWP_NOMOVE         As Long = &H2
    Private Const SWP_NOZORDER       As Long = &H4
    Private Const SWP_FRAMECHANGED   As Long = &H20
    Private Const SWP_NOOWNERZORDER  As Long = &H200

'------------------------------------------------------------------------------
' DECLARE: PRIVATE MODULE STATE
'------------------------------------------------------------------------------
#If VBA7 Then
    Private m_OriginalMainWindowOwnedStyleBits As LongPtr
    Private m_OriginalMainWindowHwnd           As LongPtr
#Else
    Private m_OriginalMainWindowOwnedStyleBits As Long
    Private m_OriginalMainWindowHwnd           As Long
#End If

    Private m_HasOriginalMainWindowOwnedStyleBits As Boolean


Public Function UI_TrySetTitleBarVisibleIfNeeded( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                    UI_TrySetTitleBarVisibleIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Apply the requested title-bar state through the owned-style-bit worker.
'
' WHY
'   Title-bar visibility alone is not sufficient for no-op detection because
'   another owned frame bit may also require restoration. The worker computes
'   the exact merged style and short-circuits only when no owned bit would
'   change.
'
' RETURNS
'   TRUE when the current owned bits already match or were successfully updated.
'
' ERROR POLICY
'   Returns FALSE and FailMsg on failure.
'
' DEPENDENCIES
'   - UI_TrySetTitleBarVisible
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TrySetTitleBarVisibleIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
        UI_TrySetTitleBarVisibleIfNeeded = _
            UI_TrySetTitleBarVisible(IsVisible, FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_TitleBarBuildRuntimeErrorText
        Resume SafeExit

End Function

Public Function UI_TryGetTitleBarVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                        UI_TryGetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Read title-bar visibility from the Application.Hwnd window style.
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
#If VBA7 Then
    Dim xlHnd      As LongPtr
    Dim StyleValue As LongPtr
#Else
    Dim xlHnd      As Long
    Dim StyleValue As Long
#End If

    Dim LastErr As Long

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TryGetTitleBarVisible = False
        IsVisible = False
        FailMsg = vbNullString

        xlHnd = Application.hWnd

        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ STYLE
'------------------------------------------------------------------------------
        SetLastError 0

#If VBA7 Then
    #If Win64 Then
        StyleValue = GetWindowLongPtr(xlHnd, GWL_STYLE)
    #Else
        StyleValue = GetWindowLong(xlHnd, GWL_STYLE)
    #End If
#Else
        StyleValue = GetWindowLong(xlHnd, GWL_STYLE)
#End If

        LastErr = GetLastError

        If StyleValue = 0 And LastErr <> 0 Then
            FailMsg = _
                "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)

            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' RETURN
'------------------------------------------------------------------------------
        IsVisible = ((StyleValue And WS_CAPTION) <> 0)
        UI_TryGetTitleBarVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_TitleBarBuildRuntimeErrorText

End Function

#If VBA7 Then
Public Function UI_InternalMergeTitleBarStyleBits( _
    ByVal CurrentStyle As LongPtr, _
    ByVal OwnedStyleBits As LongPtr) As LongPtr
#Else
Public Function UI_InternalMergeTitleBarStyleBits( _
    ByVal CurrentStyle As Long, _
    ByVal OwnedStyleBits As Long) As Long
#End If

'
'==============================================================================
'                 UI_InternalMergeTitleBarStyleBits
'------------------------------------------------------------------------------
' PURPOSE
'   Merge the title-bar style bits owned by this module into a current
'   GWL_STYLE value without altering unrelated style bits.
'
' WHY
'   The merge policy is deliberately isolated from the WinAPI write so it can
'   be validated deterministically without depending on Windows normalizing
'   particular style bits on Excel's top-level window.
'
' INPUTS
'   CurrentStyle
'     Current GWL_STYLE value whose unrelated bits must be preserved.
'
'   OwnedStyleBits
'     Desired values for TITLEBAR_OWNED_STYLE_MASK. Any bits outside that mask
'     are ignored defensively.
'
' RETURNS
'   CurrentStyle with only TITLEBAR_OWNED_STYLE_MASK replaced.
'
' BEHAVIOR
'   - Clears only TITLEBAR_OWNED_STYLE_MASK from CurrentStyle.
'   - Applies only TITLEBAR_OWNED_STYLE_MASK from OwnedStyleBits.
'   - Preserves every unrelated bit exactly.
'
' ERROR POLICY
'   - Does not raise.
'
' NOTES
'   - Public only for same-project regression access.
'   - Option Private Module prevents exposure to external VBA projects.
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' MERGE OWNED BITS
'------------------------------------------------------------------------------
        UI_InternalMergeTitleBarStyleBits = _
            (CurrentStyle And Not TITLEBAR_OWNED_STYLE_MASK) Or _
            (OwnedStyleBits And TITLEBAR_OWNED_STYLE_MASK)

End Function

Private Function UI_TrySetTitleBarVisible( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                           UI_TrySetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Show or hide the title bar of the Excel main window represented by
'   Application.Hwnd.
'
' WHY
'   Restoring an entire previously captured GWL_STYLE value can overwrite
'   unrelated style changes made later by Excel, another add-in, or caller code.
'
' RETURNS
'   TRUE on success.
'
' BEHAVIOR
'   - Owns only TITLEBAR_OWNED_STYLE_MASK:
'       * WS_CAPTION
'       * WS_SYSMENU
'       * WS_THICKFRAME
'       * WS_MINIMIZEBOX
'       * WS_MAXIMIZEBOX
'   - Captures only those owned bits on first use for the current
'     Application.Hwnd.
'   - Recaptures owned bits when Application.Hwnd changes.
'   - Hiding clears only the owned bits from the current style.
'   - Showing merges the captured owned bits into the current style.
'   - Preserves every unrelated current style bit.
'   - Refreshes the non-client frame only after an actual style write.
'
' ERROR POLICY
'   Returns FALSE and FailMsg on failure.
'
' DEPENDENCIES
'   - UI_TryGetWindowStyle
'   - UI_InternalMergeTitleBarStyleBits
'   - UI_TrySetWindowStyle
'   - UI_TryRefreshWindowFrame
'
' UPDATED
'   2026-07-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim xlHnd        As LongPtr
    Dim CurrentStyle As LongPtr
    Dim NewStyle     As LongPtr
#Else
    Dim xlHnd        As Long
    Dim CurrentStyle As Long
    Dim NewStyle     As Long
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TrySetTitleBarVisible = False
        FailMsg = vbNullString

        xlHnd = Application.hWnd

        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ CURRENT STYLE
'------------------------------------------------------------------------------
        If Not UI_TryGetWindowStyle(xlHnd, CurrentStyle, FailMsg) Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' CAPTURE OWNED STYLE BITS FOR THE CURRENT HANDLE
'------------------------------------------------------------------------------
        If (Not m_HasOriginalMainWindowOwnedStyleBits) Or _
            (m_OriginalMainWindowHwnd <> xlHnd) Then

            m_OriginalMainWindowOwnedStyleBits = _
                CurrentStyle And TITLEBAR_OWNED_STYLE_MASK

            m_OriginalMainWindowHwnd = xlHnd
            m_HasOriginalMainWindowOwnedStyleBits = True
        End If

'------------------------------------------------------------------------------
' COMPUTE NEW STYLE
'------------------------------------------------------------------------------
    'Showing restores only the captured owned bits. Hiding supplies zero owned
    'bits. The helper preserves every unrelated bit from CurrentStyle.
        If IsVisible Then
            NewStyle = UI_InternalMergeTitleBarStyleBits( _
                CurrentStyle:=CurrentStyle, _
                OwnedStyleBits:=m_OriginalMainWindowOwnedStyleBits)
        Else
            NewStyle = UI_InternalMergeTitleBarStyleBits( _
                CurrentStyle:=CurrentStyle, _
                OwnedStyleBits:=0)
        End If

'------------------------------------------------------------------------------
' SHORT-CIRCUIT
'------------------------------------------------------------------------------
        If NewStyle = CurrentStyle Then
            UI_TrySetTitleBarVisible = True
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' WRITE AND REFRESH
'------------------------------------------------------------------------------
        If Not UI_TrySetWindowStyle(xlHnd, NewStyle, FailMsg) Then
            GoTo SafeExit
        End If

        If Not UI_TryRefreshWindowFrame(xlHnd, FailMsg) Then
            GoTo SafeExit
        End If

        UI_TrySetTitleBarVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_TitleBarBuildRuntimeErrorText
        Resume SafeExit

End Function

#If VBA7 Then
Private Function UI_TryGetWindowStyle( _
    ByVal hWnd As LongPtr, _
    ByRef StyleOut As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function UI_TryGetWindowStyle( _
    ByVal hWnd As Long, _
    ByRef StyleOut As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
'                            UI_TryGetWindowStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Read GWL_STYLE using the correct API for Office bitness.
'
' RETURNS
'   TRUE on success.
'
' ERROR POLICY
'   Uses GetLastError to distinguish a valid zero from failure.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastErr As Long

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TryGetWindowStyle = False
        StyleOut = 0
        FailMsg = vbNullString

        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        SetLastError 0

#If VBA7 Then
    #If Win64 Then
        StyleOut = GetWindowLongPtr(hWnd, GWL_STYLE)
    #Else
        StyleOut = GetWindowLong(hWnd, GWL_STYLE)
    #End If
#Else
        StyleOut = GetWindowLong(hWnd, GWL_STYLE)
#End If

        LastErr = GetLastError

        If StyleOut = 0 And LastErr <> 0 Then
            FailMsg = _
                "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)

            GoTo SafeExit
        End If

        UI_TryGetWindowStyle = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_TitleBarBuildRuntimeErrorText

End Function

#If VBA7 Then
Private Function UI_TrySetWindowStyle( _
    ByVal hWnd As LongPtr, _
    ByVal NewStyle As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function UI_TrySetWindowStyle( _
    ByVal hWnd As Long, _
    ByVal NewStyle As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
'                            UI_TrySetWindowStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Write GWL_STYLE using the correct API for Office bitness.
'
' RETURNS
'   TRUE on success.
'
' ERROR POLICY
'   Uses GetLastError to distinguish a valid zero from failure.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim PrevStyle As LongPtr
#Else
    Dim PrevStyle As Long
#End If

    Dim LastErr As Long

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TrySetWindowStyle = False
        FailMsg = vbNullString

        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        SetLastError 0

#If VBA7 Then
    #If Win64 Then
        PrevStyle = SetWindowLongPtr(hWnd, GWL_STYLE, NewStyle)
    #Else
        PrevStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)
    #End If
#Else
        PrevStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)
#End If

        LastErr = GetLastError

        If PrevStyle = 0 And LastErr <> 0 Then
            FailMsg = _
                "SetWindowLong/SetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)

            GoTo SafeExit
        End If

        UI_TrySetWindowStyle = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_TitleBarBuildRuntimeErrorText

End Function

#If VBA7 Then
Private Function UI_TryRefreshWindowFrame( _
    ByVal hWnd As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function UI_TryRefreshWindowFrame( _
    ByVal hWnd As Long, _
    ByRef FailMsg As String) As Boolean
#End If

'
'==============================================================================
'                           UI_TryRefreshWindowFrame
'------------------------------------------------------------------------------
' PURPOSE
'   Recalculate and repaint the non-client frame after a style change.
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
    Dim ApiOK   As Long
    Dim LastErr As Long

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_TryRefreshWindowFrame = False
        FailMsg = vbNullString

        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' REFRESH
'------------------------------------------------------------------------------
        SetLastError 0

        ApiOK = SetWindowPos( _
            hWnd, _
            0, _
            0, _
            0, _
            0, _
            0, _
            SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or _
                SWP_NOOWNERZORDER Or SWP_FRAMECHANGED)

        LastErr = GetLastError

        If ApiOK = 0 Then
            FailMsg = "SetWindowPos failed; GetLastError=" & CStr(LastErr)
            GoTo SafeExit
        End If

        UI_TryRefreshWindowFrame = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_TitleBarBuildRuntimeErrorText

End Function

Private Function UI_TitleBarBuildRuntimeErrorText() As String

'
'==============================================================================
'                    UI_TitleBarBuildRuntimeErrorText
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
'   2026-07-29
'==============================================================================
'

        On Error Resume Next

        UI_TitleBarBuildRuntimeErrorText = _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, _
                " | Source: " & Err.Source, _
                vbNullString) & _
            IIf(Erl <> 0, _
                " | Line: " & CStr(Erl), _
                vbNullString)

End Function
