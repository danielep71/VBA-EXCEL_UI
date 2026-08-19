Attribute VB_Name = "M_EXCEL_UI_TITLEBAR"
Option Explicit
Option Private Module

'==============================================================================
' M_EXCEL_UI_TITLEBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Owns every WinAPI declaration, mutable frame state and style-merging rule
'   required to show or hide the title bar of the Excel main window represented
'   by Application.Hwnd.
'
' WHY THIS EXISTS
'   Excel provides no object-model control over title-bar visibility, so the
'   only route is a GWL_STYLE update on the top-level window. That makes this
'   the single operating-system-sensitive part of the package.
'
'   Isolating it keeps M_EXCEL_UI free of WinAPI declarations and handle-scoped
'   mutable state, and confines the blast radius of any Windows or Office
'   behavior change to one module. It also lets the style-merge policy be tested
'   as pure arithmetic, with no window and no host required.
'
' INTERNAL SURFACE
'   - UI_TryGetTitleBarVisible
'   - UI_TrySetTitleBarVisibleIfNeeded
'   - UI_InternalMergeTitleBarStyleBits
'   - UI_InternalResetTitleBarBaseline
'
' OWNERSHIP MODEL
'   This module claims exactly five GWL_STYLE bits and nothing else:
'
'       WS_CAPTION      &HC00000    caption bar
'       WS_SYSMENU      &H80000     system menu
'       WS_THICKFRAME   &H40000     sizing frame
'       WS_MINIMIZEBOX  &H20000     minimize box
'       WS_MAXIMIZEBOX  &H10000     maximize box
'
'   Their union is TITLEBAR_OWNED_STYLE_MASK (&HCF0000).
'
'   Restoring a whole previously captured GWL_STYLE value would overwrite
'   unrelated style changes made later by Excel, another add-in or caller code.
'   Every write therefore merges only the owned bits into the CURRENT style and
'   leaves every other bit exactly as found.
'
' DESIGN PRINCIPLES
'   - The merge policy is a pure function, deliberately separated from the
'     WinAPI write so it can be validated deterministically.
'   - Owned bits are captured once per Application.Hwnd and re-captured only
'     when the handle changes.
'   - The non-client frame is refreshed only after a style write actually
'     occurred, never after a no-op.
'   - Entry points are fail-soft and report through a ByRef FailMsg.
'
' ERROR POLICY
'   - Internal entry points return False plus diagnostic text.
'   - A zero API return is disambiguated from a genuine failure before being
'     reported.
'   - No user-interface message is displayed.
'
' DEPENDENCIES
'   None. This module deliberately has no project-module dependency, which is
'   why it carries its own error-text builder rather than calling the one in
'   M_EXCEL_UI_RUNTIME.
'
' PLATFORM / COMPATIBILITY
'   - Windows only.
'   - Supports 32-bit and 64-bit Office through conditional compilation:
'       VBA7 + Win64  => GetWindowLongPtr / SetWindowLongPtr
'       VBA7 + Win32  => GetWindowLong / SetWindowLong with LongPtr handles
'       pre-VBA7      => GetWindowLong / SetWindowLong with Long handles
'
' NOTES
'   - Title-bar behavior can vary with Excel version, window state, Windows
'     desktop-composition settings and other loaded add-ins. It is documented
'     as best effort throughout.
'   - The merge arithmetic relies on GWL_STYLE never setting bit 31 on Excel's
'     main window, which holds because that window is not WS_POPUP. Were it
'     set, the Long mask would sign-extend when widened to LongPtr.
'   - Module state is lost on a VBA project reset while the window style
'     itself survives, because the style belongs to the running Excel process.
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
' WIN32 / WIN64 API DECLARATIONS
'==============================================================================

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

'==============================================================================
' PRIVATE CONSTANTS
'==============================================================================

'Window-style index passed to the GetWindowLong / SetWindowLong family.
Private Const GWL_STYLE                 As Long = -16

'The five frame bits this module claims. Nothing outside this set is written.
Private Const WS_CAPTION                As Long = &HC00000
Private Const WS_SYSMENU                As Long = &H80000
Private Const WS_THICKFRAME             As Long = &H40000
Private Const WS_MINIMIZEBOX            As Long = &H20000
Private Const WS_MAXIMIZEBOX            As Long = &H10000

'Union of the five owned bits above, held as a literal so the merge helper
'needs a single mask rather than five OR-ed constants on every call.
Private Const TITLEBAR_OWNED_STYLE_MASK As Long = &HCF0000

'Owned bits assumed when a show is requested and no baseline was ever captured
'for the current handle. Restoring the full owned frame is the only safe
'assumption in that case: the alternative is to re-apply the current hidden
'bits, which silently leaves the title bar hidden and reports success.
Private Const TITLEBAR_DEFAULT_STYLE_BITS As Long = &HCF0000

'SetWindowPos flags. Only the frame is recalculated: position, size, Z-order
'and owner Z-order are all left untouched.
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOZORDER              As Long = &H4
Private Const SWP_FRAMECHANGED          As Long = &H20
Private Const SWP_NOOWNERZORDER         As Long = &H200

'==============================================================================
' PRIVATE MODULE STATE
'==============================================================================

'Owned style bits captured for the current Excel main window, plus the handle
'they belong to. Held only in memory and lost on a VBA project reset.
#If VBA7 Then
    Private m_OriginalMainWindowOwnedStyleBits As LongPtr
    Private m_OriginalMainWindowHwnd           As LongPtr
#Else
    Private m_OriginalMainWindowOwnedStyleBits As Long
    Private m_OriginalMainWindowHwnd           As Long
#End If

'True once owned bits have been captured for m_OriginalMainWindowHwnd.
Private m_HasOriginalMainWindowOwnedStyleBits  As Boolean


Public Function UI_TrySetTitleBarVisibleIfNeeded( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_TrySetTitleBarVisibleIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the requested title-bar state through the owned-style-bit worker.
'
' WHY THIS EXISTS
'   Title-bar visibility alone is not a sufficient basis for no-op detection,
'   because another owned frame bit may still require restoration while
'   WS_CAPTION already matches. The worker computes the exact merged style and
'   short-circuits only when no owned bit would change, so the decision is made
'   on the full owned set rather than on one bit.
'
' INPUTS
'   IsVisible
'     Requested title-bar visibility.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => owned bits already match, or were successfully updated.
'     False => the update was attempted and failed.
'
' BEHAVIOR
'   - Delegates to UI_TrySetTitleBarVisible.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_TrySetTitleBarVisible
'   - UI_TitleBarBuildRuntimeErrorText
'
' CALLED FROM
'   - M_EXCEL_UI
'   - M_EXCEL_UI_SNAPSHOT
'
' NOTES
'   The IfNeeded decision itself lives in the worker; this entry point exists to
'   present the same naming shape as the M_EXCEL_UI_RUNTIME helpers.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the worker reports otherwise
        UI_TrySetTitleBarVisibleIfNeeded = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Delegate the owned-bit merge, write and frame refresh
        UI_TrySetTitleBarVisibleIfNeeded = _
            UI_TrySetTitleBarVisible(IsVisible, FailMsg)

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_TitleBarBuildRuntimeErrorText
        Resume Safe_Exit

End Function


Public Sub UI_InternalResetTitleBarBaseline()
'
'==============================================================================
' UI_InternalResetTitleBarBaseline
'------------------------------------------------------------------------------
' PURPOSE
'   Discards the captured owned-bit baseline, returning this module to the
'   state it holds before its first title-bar call for a given window handle.
'
' WHY THIS EXISTS
'   The show-recovery defect only appears when a show is requested while no
'   baseline has been captured and the frame is already hidden. That is the
'   state left by a VBA project reset, because the window style belongs to the
'   Excel process and survives, while this module's state does not.
'
'   A regression case cannot reach that state on its own. VBA offers no
'   supported way to clear another module's private variables, and any earlier
'   title-bar operation in the same session, including the round-trip case that
'   runs before it, captures a baseline first. Without this entry point the
'   guarding case silently exercises the ordinary path and passes whether or
'   not the defect is present.
'
'   Exposing a deliberate seam is therefore the honest option. It follows the
'   precedent already set by UI_InternalMergeTitleBarStyleBits, which is Public
'   for the same reason.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Clears the captured owned bits, the associated handle, and the captured
'     flag.
'   - Touches no window style. The live frame is left exactly as it is.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - M_EXCEL_UI_REGRESSION_TESTS
'
' NOTES
'   - Public only for same-project regression access. Option Private Module
'     prevents exposure to external VBA projects.
'   - Not part of the supported public API. Production code must never call it:
'     doing so discards the frame the next show would restore.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR CAPTURED BASELINE
'------------------------------------------------------------------------------
    'Forget the owned bits, the handle they belong to, and the captured flag
        m_OriginalMainWindowOwnedStyleBits = 0
        m_OriginalMainWindowHwnd = 0
        m_HasOriginalMainWindowOwnedStyleBits = False

End Sub


Public Function UI_TryGetTitleBarVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_TryGetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Reads title-bar visibility from the Application.Hwnd window style.
'
' WHY THIS EXISTS
'   The snapshot engine must record whether the title bar was visible at capture
'   time. There is no object-model property to read, so visibility is inferred
'   from the presence of WS_CAPTION in the live window style.
'
' INPUTS
'   IsVisible
'     ByRef. Receives True when WS_CAPTION is set.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the style was read and IsVisible is meaningful.
'     False => the handle or the style read was unusable.
'
' BEHAVIOR
'   - Validates the Excel main-window handle.
'   - Reads GWL_STYLE using the API matching the host bitness.
'   - Reports visibility from the WS_CAPTION bit alone.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_TitleBarBuildRuntimeErrorText
'
' CALLED FROM
'   - M_EXCEL_UI_SNAPSHOT
'
' NOTES
'   WS_CAPTION is the visibility signal; the other four owned bits travel with
'   it but do not participate in this decision.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim xlHnd               As LongPtr         'Excel main-window handle
    Dim StyleValue          As LongPtr         'Live GWL_STYLE value
#Else
    Dim xlHnd               As Long            'Excel main-window handle
    Dim StyleValue          As Long            'Live GWL_STYLE value
#End If

    Dim LastErr             As Long            'Win32 last-error after the read

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the style has been read
        UI_TryGetTitleBarVisible = False

    'Initialize the output and the failure message buffer
        IsVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Resolve and validate the Excel main-window handle
        xlHnd = Application.hWnd

        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' READ STYLE
'------------------------------------------------------------------------------
    'Clear the thread last-error so a zero return can be disambiguated
        SetLastError 0

    'Read GWL_STYLE using the API matching the host bitness
#If VBA7 Then
    #If Win64 Then
        StyleValue = GetWindowLongPtr(xlHnd, GWL_STYLE)
    #Else
        StyleValue = GetWindowLong(xlHnd, GWL_STYLE)
    #End If
#Else
        StyleValue = GetWindowLong(xlHnd, GWL_STYLE)
#End If

    'Capture the last-error immediately after the call
        LastErr = GetLastError

    'A zero return is only a failure when the last-error also reports one
        If StyleValue = 0 And LastErr <> 0 Then
            FailMsg = _
                "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)

            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Visibility is carried by the caption bit alone
        IsVisible = ((StyleValue And WS_CAPTION) <> 0)
        UI_TryGetTitleBarVisible = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_TitleBarBuildRuntimeErrorText

End Function


#If VBA7 Then
Public Function UI_InternalMergeTitleBarStyleBits( _
    ByVal CurrentStyle As LongPtr, _
    ByVal OwnedStyleBits As LongPtr) _
    As LongPtr
#Else
Public Function UI_InternalMergeTitleBarStyleBits( _
    ByVal CurrentStyle As Long, _
    ByVal OwnedStyleBits As Long) _
    As Long
#End If
'
'==============================================================================
' UI_InternalMergeTitleBarStyleBits
'------------------------------------------------------------------------------
' PURPOSE
'   Merges the title-bar style bits owned by this module into a current
'   GWL_STYLE value without altering any unrelated style bit.
'
' WHY THIS EXISTS
'   The merge policy is the whole correctness argument for title-bar control, so
'   it is deliberately isolated from the WinAPI write. As pure arithmetic it can
'   be validated deterministically by the regression harness, with no window, no
'   host and no dependence on Windows normalizing particular bits on Excel's
'   top-level window.
'
' INPUTS
'   CurrentStyle
'     Live GWL_STYLE value whose unrelated bits must be preserved exactly.
'
'   OwnedStyleBits
'     Desired values for TITLEBAR_OWNED_STYLE_MASK. Bits outside that mask are
'     ignored defensively, so a caller cannot widen this module's ownership by
'     passing a richer value.
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
'   - Does not raise. The operation is pure arithmetic on two style values.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - UI_TrySetTitleBarVisible
'   - M_EXCEL_UI_REGRESSION_TESTS
'
' NOTES
'   - Public only for same-project regression access. Option Private Module
'     keeps it out of the cross-project automation namespace.
'   - Passing zero owned bits is the hide case; passing captured bits is the
'     show case. The helper itself has no notion of show or hide.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' MERGE OWNED BITS
'------------------------------------------------------------------------------
    'Keep every unrelated bit, then overlay only the owned bits requested
        UI_InternalMergeTitleBarStyleBits = _
            (CurrentStyle And Not TITLEBAR_OWNED_STYLE_MASK) Or _
            (OwnedStyleBits And TITLEBAR_OWNED_STYLE_MASK)

End Function


Private Function UI_TrySetTitleBarVisible( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_TrySetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Shows or hides the title bar of the Excel main window represented by
'   Application.Hwnd.
'
' WHY THIS EXISTS
'   Restoring an entire previously captured GWL_STYLE value would overwrite
'   unrelated style changes made later by Excel, another add-in or caller code.
'   This routine therefore reads the live style on every call and rewrites only
'   the five bits this module claims.
'
' INPUTS
'   IsVisible
'     Requested title-bar visibility.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => owned bits already match, or were successfully updated.
'     False => the read, the write or the frame refresh failed.
'
' BEHAVIOR
'   - Validates the Excel main-window handle.
'   - Reads the live GWL_STYLE value.
'   - Captures the owned bits on first use for the current handle, and
'     re-captures them when Application.Hwnd changes.
'   - Hiding supplies zero owned bits; showing supplies the captured bits.
'   - Short-circuits when the merged style equals the current style.
'   - Writes the style and refreshes the non-client frame only after an actual
'     change.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on the first failing step.
'
' DEPENDENCIES
'   - UI_TryGetWindowStyle
'   - UI_InternalMergeTitleBarStyleBits
'   - UI_TrySetWindowStyle
'   - UI_TryRefreshWindowFrame
'   - UI_TitleBarBuildRuntimeErrorText
'
' CALLED FROM
'   - UI_TrySetTitleBarVisibleIfNeeded
'
' NOTES
'   The captured baseline lives only in module memory. A VBA project reset
'   discards it while the window style itself survives in the running Excel
'   process, so the two can disagree after a reset.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim xlHnd               As LongPtr         'Excel main-window handle
    Dim CurrentStyle        As LongPtr         'Live GWL_STYLE value
    Dim NewStyle            As LongPtr         'Merged GWL_STYLE value to write
    Dim RestoreBits         As LongPtr         'Owned bits a show will re-apply
#Else
    Dim xlHnd               As Long            'Excel main-window handle
    Dim CurrentStyle        As Long            'Live GWL_STYLE value
    Dim NewStyle            As Long            'Merged GWL_STYLE value to write
    Dim RestoreBits         As Long            'Owned bits a show will re-apply
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until every step has completed
        UI_TrySetTitleBarVisible = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Resolve and validate the Excel main-window handle
        xlHnd = Application.hWnd

        If xlHnd = 0 Then
            FailMsg = "invalid Excel window handle"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' READ CURRENT STYLE
'------------------------------------------------------------------------------
    'Read the live style; unrelated bits will be preserved from this value
        If Not UI_TryGetWindowStyle(xlHnd, CurrentStyle, FailMsg) Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' CAPTURE OWNED STYLE BITS FOR THE CURRENT HANDLE
'------------------------------------------------------------------------------
    'Capture on first use, and again whenever the main window handle changes
        If (Not m_HasOriginalMainWindowOwnedStyleBits) Or _
            (m_OriginalMainWindowHwnd <> xlHnd) Then

            m_OriginalMainWindowOwnedStyleBits = _
                CurrentStyle And TITLEBAR_OWNED_STYLE_MASK

            m_OriginalMainWindowHwnd = xlHnd
            m_HasOriginalMainWindowOwnedStyleBits = True
        End If

    'Take the baseline a show would re-apply
        RestoreBits = m_OriginalMainWindowOwnedStyleBits

'------------------------------------------------------------------------------
' COMPUTE NEW STYLE
'------------------------------------------------------------------------------
    'Showing restores the captured owned bits; hiding supplies none. Either way
    'the helper preserves every unrelated bit from CurrentStyle.
        If IsVisible Then

            'A show must never re-apply an all-zero baseline. That happens when
            'the first title-bar call after a VBA project reset is a show while
            'the frame is already hidden: the capture above then records zero
            'owned bits, the merge becomes a no-op, and the short circuit below
            'reports success while the title bar stays hidden. Falling back to
            'the full owned frame keeps UI_ShowExcelUI a real recovery path.
                If RestoreBits = 0 Then
                    RestoreBits = TITLEBAR_DEFAULT_STYLE_BITS
                End If

            NewStyle = UI_InternalMergeTitleBarStyleBits( _
                CurrentStyle:=CurrentStyle, _
                OwnedStyleBits:=RestoreBits)
        Else
            NewStyle = UI_InternalMergeTitleBarStyleBits( _
                CurrentStyle:=CurrentStyle, _
                OwnedStyleBits:=0)
        End If

'------------------------------------------------------------------------------
' SHORT-CIRCUIT
'------------------------------------------------------------------------------
    'Skip the write and the frame refresh when no owned bit would change
        If NewStyle = CurrentStyle Then
            UI_TrySetTitleBarVisible = True
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' WRITE AND REFRESH
'------------------------------------------------------------------------------
    'Write the merged style
        If Not UI_TrySetWindowStyle(xlHnd, NewStyle, FailMsg) Then
            GoTo Safe_Exit
        End If

    'Recalculate and repaint the non-client frame
        If Not UI_TryRefreshWindowFrame(xlHnd, FailMsg) Then
            GoTo Safe_Exit
        End If

    'Report success
        UI_TrySetTitleBarVisible = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_TitleBarBuildRuntimeErrorText
        Resume Safe_Exit

End Function


#If VBA7 Then
Private Function UI_TryGetWindowStyle( _
    ByVal hWnd As LongPtr, _
    ByRef StyleOut As LongPtr, _
    ByRef FailMsg As String) _
    As Boolean
#Else
Private Function UI_TryGetWindowStyle( _
    ByVal hWnd As Long, _
    ByRef StyleOut As Long, _
    ByRef FailMsg As String) _
    As Boolean
#End If
'
'==============================================================================
' UI_TryGetWindowStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Reads GWL_STYLE using the API matching the host bitness.
'
' WHY THIS EXISTS
'   Keeping the bitness branch in one place means the callers work with a single
'   Boolean contract and never repeat the conditional compilation.
'
' INPUTS
'   hWnd
'     Target window handle.
'
'   StyleOut
'     ByRef. Receives the style value. Zero when the read fails.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the style was read.
'     False => the handle was invalid or the API reported a failure.
'
' BEHAVIOR
'   - Rejects a zero handle explicitly.
'   - Clears the thread last-error before the call so that a zero return can be
'     told apart from a genuine failure.
'
' ERROR POLICY
'   - Does not raise.
'   - Uses GetLastError to distinguish a valid zero from a failure.
'
' DEPENDENCIES
'   - UI_TitleBarBuildRuntimeErrorText
'
' CALLED FROM
'   - UI_TrySetTitleBarVisible
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastErr             As Long            'Win32 last-error after the read

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the style has been read
        UI_TryGetWindowStyle = False

    'Initialize the output and the failure message buffer
        StyleOut = 0
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the window handle
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' READ STYLE
'------------------------------------------------------------------------------
    'Clear the thread last-error so a zero return can be disambiguated
        SetLastError 0

    'Read GWL_STYLE using the API matching the host bitness
#If VBA7 Then
    #If Win64 Then
        StyleOut = GetWindowLongPtr(hWnd, GWL_STYLE)
    #Else
        StyleOut = GetWindowLong(hWnd, GWL_STYLE)
    #End If
#Else
        StyleOut = GetWindowLong(hWnd, GWL_STYLE)
#End If

    'Capture the last-error immediately after the call
        LastErr = GetLastError

    'A zero return is only a failure when the last-error also reports one
        If StyleOut = 0 And LastErr <> 0 Then
            FailMsg = _
                "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)

            GoTo Safe_Exit
        End If

    'Report success
        UI_TryGetWindowStyle = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_TitleBarBuildRuntimeErrorText

End Function


#If VBA7 Then
Private Function UI_TrySetWindowStyle( _
    ByVal hWnd As LongPtr, _
    ByVal NewStyle As LongPtr, _
    ByRef FailMsg As String) _
    As Boolean
#Else
Private Function UI_TrySetWindowStyle( _
    ByVal hWnd As Long, _
    ByVal NewStyle As Long, _
    ByRef FailMsg As String) _
    As Boolean
#End If
'
'==============================================================================
' UI_TrySetWindowStyle
'------------------------------------------------------------------------------
' PURPOSE
'   Writes GWL_STYLE using the API matching the host bitness.
'
' WHY THIS EXISTS
'   SetWindowLong returns the PREVIOUS style, so zero is an ambiguous result:
'   it means either "the previous style was zero" or "the call failed". The
'   last-error is cleared beforehand so the two can be told apart.
'
' INPUTS
'   hWnd
'     Target window handle.
'
'   NewStyle
'     Merged style value to write.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the style was written.
'     False => the handle was invalid or the API reported a failure.
'
' BEHAVIOR
'   - Rejects a zero handle explicitly.
'   - Clears the thread last-error before the call.
'   - Treats a zero return as a failure only when the last-error agrees.
'
' ERROR POLICY
'   - Does not raise.
'   - Uses GetLastError to distinguish a valid zero from a failure.
'
' DEPENDENCIES
'   - UI_TitleBarBuildRuntimeErrorText
'
' CALLED FROM
'   - UI_TrySetTitleBarVisible
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim PrevStyle           As LongPtr         'Style value replaced by the write
#Else
    Dim PrevStyle           As Long            'Style value replaced by the write
#End If

    Dim LastErr             As Long            'Win32 last-error after the write

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the style has been written
        UI_TrySetWindowStyle = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the window handle
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' WRITE STYLE
'------------------------------------------------------------------------------
    'Clear the thread last-error so a zero return can be disambiguated
        SetLastError 0

    'Write GWL_STYLE using the API matching the host bitness
#If VBA7 Then
    #If Win64 Then
        PrevStyle = SetWindowLongPtr(hWnd, GWL_STYLE, NewStyle)
    #Else
        PrevStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)
    #End If
#Else
        PrevStyle = SetWindowLong(hWnd, GWL_STYLE, NewStyle)
#End If

    'Capture the last-error immediately after the call
        LastErr = GetLastError

    'A zero previous style is only a failure when the last-error agrees
        If PrevStyle = 0 And LastErr <> 0 Then
            FailMsg = _
                "SetWindowLong/SetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)

            GoTo Safe_Exit
        End If

    'Report success
        UI_TrySetWindowStyle = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_TitleBarBuildRuntimeErrorText

End Function


#If VBA7 Then
Private Function UI_TryRefreshWindowFrame( _
    ByVal hWnd As LongPtr, _
    ByRef FailMsg As String) _
    As Boolean
#Else
Private Function UI_TryRefreshWindowFrame( _
    ByVal hWnd As Long, _
    ByRef FailMsg As String) _
    As Boolean
#End If
'
'==============================================================================
' UI_TryRefreshWindowFrame
'------------------------------------------------------------------------------
' PURPOSE
'   Recalculates and repaints the non-client frame after a style change.
'
' WHY THIS EXISTS
'   Writing GWL_STYLE alone does not make Windows re-measure the non-client
'   area. Without an explicit SWP_FRAMECHANGED the caption can remain drawn
'   after it has already been removed from the style.
'
' INPUTS
'   hWnd
'     Target window handle.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the frame was recalculated.
'     False => the handle was invalid or SetWindowPos failed.
'
' BEHAVIOR
'   - Rejects a zero handle explicitly.
'   - Requests a frame change while suppressing move, size, Z-order and owner
'     Z-order effects, so only the frame is affected.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_TitleBarBuildRuntimeErrorText
'
' CALLED FROM
'   - UI_TrySetTitleBarVisible
'
' NOTES
'   SetWindowPos returns a BOOL, so zero is unambiguously a failure here and no
'   last-error disambiguation is required.
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ApiOK               As Long            'SetWindowPos BOOL result
    Dim LastErr             As Long            'Win32 last-error after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the frame has been recalculated
        UI_TryRefreshWindowFrame = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Validate the window handle
        If hWnd = 0 Then
            FailMsg = "invalid window handle"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' REFRESH FRAME
'------------------------------------------------------------------------------
    'Clear the thread last-error so the diagnostic reports this call
        SetLastError 0

    'Request a frame change only; position, size and Z-order are untouched
        ApiOK = SetWindowPos( _
            hWnd, _
            0, _
            0, _
            0, _
            0, _
            0, _
            SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or _
                SWP_NOOWNERZORDER Or SWP_FRAMECHANGED)

    'Capture the last-error immediately after the call
        LastErr = GetLastError

    'SetWindowPos returns a BOOL, so zero is unambiguously a failure
        If ApiOK = 0 Then
            FailMsg = "SetWindowPos failed; GetLastError=" & CStr(LastErr)
            GoTo Safe_Exit
        End If

    'Report success
        UI_TryRefreshWindowFrame = True

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Report the unexpected runtime error as the failure reason
        FailMsg = UI_TitleBarBuildRuntimeErrorText

End Function


Private Function UI_TitleBarBuildRuntimeErrorText() _
    As String
'
'==============================================================================
' UI_TitleBarBuildRuntimeErrorText
'------------------------------------------------------------------------------
' PURPOSE
'   Builds a consistent diagnostic string from the active Err object.
'
' WHY THIS EXISTS
'   Functionally identical to UI_RuntimeBuildErrorText, but duplicated here on
'   purpose: it is what keeps M_EXCEL_UI_TITLEBAR free of any project-module
'   dependency, so the WinAPI subsystem can be reasoned about and tested on its
'   own. The duplication is the price of that isolation and is intentional.
'
' RETURNS
'   String
'     Best-effort error number, description, source and Erl text.
'
' ERROR POLICY
'   - Suppresses formatting errors locally.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - This module
'
' UPDATED
'   2026-08-18
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber           As Long            'Err.Number captured on entry
    Dim ErrDescription      As String          'Err.Description captured on entry
    Dim ErrSource           As String          'Err.Source captured on entry
    Dim ErrLine             As Long            'Erl captured on entry

'------------------------------------------------------------------------------
' CAPTURE ERR STATE
'------------------------------------------------------------------------------
    'Read the Err object BEFORE any On Error statement. Any form of On Error
    'resets Err, so protecting this routine first would blank the very values
    'it exists to report and every diagnostic would read "0: ".
        ErrNumber = Err.Number
        ErrDescription = Err.Description
        ErrSource = Err.Source
        ErrLine = Erl

'------------------------------------------------------------------------------
' BUILD DIAGNOSTIC TEXT
'------------------------------------------------------------------------------
    'Never let diagnostic formatting raise inside an error handler
        On Error Resume Next

    'Compose number, description, and the optional source and line fragments
        UI_TitleBarBuildRuntimeErrorText = _
            CStr(ErrNumber) & ": " & ErrDescription & _
            IIf(Len(ErrSource) > 0, _
                " | Source: " & ErrSource, _
                vbNullString) & _
            IIf(ErrLine <> 0, _
                " | Line: " & CStr(ErrLine), _
                vbNullString)

End Function


