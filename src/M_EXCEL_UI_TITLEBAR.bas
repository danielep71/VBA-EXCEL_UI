Attribute VB_Name = "M_EXCEL_UI_TITLEBAR"
Option Explicit
Option Private Module

'==============================================================================
' M_EXCEL_UI_TITLEBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Owns every WinAPI declaration, mutable frame state and style-merging rule
'   required to show or hide the title bar of a specified top-level Excel
'   window.
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
'   Explicit-target entry points, used by the snapshot engine:
'   - UI_TryGetActiveTitleBarHwnd
'   - UI_TryGetTitleBarVisibleForHwnd
'   - UI_TrySetTitleBarVisibleForHwndIfNeeded
'   - UI_InternalIsTitleBarFrameAlive
'
'   Active-window wrappers, kept for callers that have no target of their own:
'   - UI_TryGetTitleBarVisible
'   - UI_TrySetTitleBarVisibleIfNeeded
'
'   Regression seams:
'   - UI_InternalMergeTitleBarStyleBits
'   - UI_InternalResetTitleBarBaseline
'   - UI_InternalResetTitleBarBaselineForHwnd
'   - UI_InternalIsFrameRefreshPending
'   - UI_InternalInjectFrameRefreshFailure
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
' TARGET MODEL
'   Under the Single Document Interface used by modern Excel, each workbook
'   window is its own top-level window and Application.Hwnd returns whichever
'   one is active at the moment it is read. A module that resolves the target
'   on every call therefore cannot promise that a value read from one window is
'   written back to the same window.
'
'   Every operation in this module consequently takes an explicit handle. The
'   caller decides which frame it means and holds that decision for the whole
'   capture/restore round trip. The no-argument wrappers resolve the active
'   window once and are documented as active-window operations, which is a
'   contract this module can actually keep.
'
' FRAME STATE REGISTRY
'   Frame state is keyed by handle rather than held as one value for the
'   process:
'
'       hWnd            the top-level window the entry belongs to
'       OwnedStyleBits  the baseline a show will re-apply
'       HasBaseline     whether OwnedStyleBits has ever been captured
'       ComponentHidden whether THIS component is the reason the frame is hidden
'       RefreshPending  whether a style write succeeded but its frame refresh
'                       did not, and must be retried before anything else
'
'   ComponentHidden is what makes the baseline self-healing. While the
'   component does not own a hidden state, the live owned bits are the truth and
'   the baseline is recaptured from them on every call, so a legitimate frame
'   change made by Excel or another add-in is adopted rather than overwritten.
'   Once the component has hidden the frame, the live bits are the component's
'   own zeros and the stored baseline is the only surviving record of what to
'   restore, so it is left alone.
'
' DESIGN PRINCIPLES
'   - The merge policy is a pure function, deliberately separated from the
'     WinAPI write so it can be validated deterministically.
'   - Frame state is per handle. Operating on a second window never destroys
'     the state of the first.
'   - A style write and its frame refresh are treated as one unit of work. If
'     the refresh fails the debt is recorded and paid before the next call is
'     allowed to conclude that there is nothing to do.
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
'   - The frame-state registry is bounded by compaction, not by growth alone.
'     Entries whose handle no longer passes IsWindow are reclaimed before the
'     registry is extended, so a long session that opens and closes many
'     workbook windows does not accumulate dead entries.
'   - Conditional compilation appears inside the frame-state Type so that one
'     declaration serves both handle widths. The alternative, declaring the
'     whole Type twice, duplicates the member list and lets the two copies
'     drift apart.
'
' UPDATED
'   2026-08-19 - Replaced the singleton frame baseline with a per-handle
'                registry, added the explicit-target entry points, and made a
'                failed frame refresh a recorded debt rather than a silent
'                one. Fixes ICR-UI-P1-01, ICR-UI-P2-04 and ICR-UI-P2-03.
'   2026-08-18 - Reformatted to the project house style. No behavior change.
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.1
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

    Private Declare PtrSafe Function IsWindow Lib "user32" ( _
        ByVal hWnd As LongPtr) As Long

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

    Private Declare Function IsWindow Lib "user32" ( _
        ByVal hWnd As Long) As Long

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

'Slots added each time the frame-state registry has to grow. Excel sessions
'rarely hold more than a handful of workbook windows, so a small step keeps the
'registry compact without reallocating on every new window.
Private Const FRAME_STATE_GROWTH_SLOTS  As Long = 8

'==============================================================================
' PRIVATE TYPES
'==============================================================================

'One registry entry per top-level window this module has operated on. See the
'FRAME STATE REGISTRY section of the module header for the invariants.
Private Type tTitleBarFrameState
#If VBA7 Then
    hWnd                As LongPtr             'Window the entry belongs to
    OwnedStyleBits      As LongPtr             'Baseline a show re-applies
#Else
    hWnd                As Long                'Window the entry belongs to
    OwnedStyleBits      As Long                'Baseline a show re-applies
#End If
    HasBaseline         As Boolean             'OwnedStyleBits ever captured
    ComponentHidden     As Boolean             'This component hid the frame
    RefreshPending      As Boolean             'Frame refresh still owed
End Type

'==============================================================================
' PRIVATE MODULE STATE
'==============================================================================

'Frame state per top-level window. Held only in memory and lost on a VBA
'project reset, while the window styles themselves survive in the Excel process.
Private m_FrameStates()                 As tTitleBarFrameState
Private m_FrameStateCount               As Long

'Regression seam. When True the next frame refresh reports failure without
'calling SetWindowPos, so the refresh-debt path can be exercised without a
'way to make Windows fail on demand.
Private m_InjectFrameRefreshFailure     As Boolean


#If VBA7 Then
Public Function UI_TryGetActiveTitleBarHwnd( _
    ByRef HwndOut As LongPtr, _
    ByRef FailMsg As String) _
    As Boolean
#Else
Public Function UI_TryGetActiveTitleBarHwnd( _
    ByRef HwndOut As Long, _
    ByRef FailMsg As String) _
    As Boolean
#End If
'
'==============================================================================
' UI_TryGetActiveTitleBarHwnd
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves and validates the top-level window handle Excel currently reports
'   as active.
'
' WHY THIS EXISTS
'   Under the Single Document Interface, Application.Hwnd is not a property of
'   the process: it is a property of whichever workbook window happens to be
'   active when it is read. Any caller that must read a frame now and write it
'   back later has to resolve the handle ONCE and keep it, or the two ends of
'   the operation can address different windows.
'
'   Exposing the resolution step is what lets the snapshot engine do exactly
'   that. It is deliberately separate from the read and write helpers so that
'   the caller, not this module, owns the decision about which frame is meant.
'
' INPUTS
'   HwndOut
'     ByRef. Receives the active top-level window handle. Zero on failure.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => HwndOut holds a handle that currently passes IsWindow.
'     False => Excel reported no usable handle.
'
' BEHAVIOR
'   - Reads Application.Hwnd once.
'   - Rejects a zero handle and a handle that is not a live window.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_InternalIsTitleBarFrameAlive
'   - UI_TitleBarBuildRuntimeErrorText
'
' CALLED FROM
'   - M_EXCEL_UI_SNAPSHOT
'   - This module
'
' NOTES
'   The handle is valid only for as long as that window is open. Callers that
'   retain it must re-probe with UI_InternalIsTitleBarFrameAlive before use.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until a live handle has been resolved
        UI_TryGetActiveTitleBarHwnd = False

    'Initialize the output and the failure message buffer
        HwndOut = 0
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' RESOLVE HANDLE
'------------------------------------------------------------------------------
    'Read the host handle exactly once
        HwndOut = Application.hWnd

'------------------------------------------------------------------------------
' VALIDATE HANDLE
'------------------------------------------------------------------------------
    'Reject a handle that is absent or no longer refers to a window
        If Not UI_InternalIsTitleBarFrameAlive(HwndOut) Then
            HwndOut = 0
            FailMsg = "invalid Excel window handle"
            GoTo Safe_Exit
        End If

    'Report success
        UI_TryGetActiveTitleBarHwnd = True

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
        HwndOut = 0
        FailMsg = UI_TitleBarBuildRuntimeErrorText
        Resume Safe_Exit

End Function


#If VBA7 Then
Public Function UI_TryGetTitleBarVisibleForHwnd( _
    ByVal TargetHwnd As LongPtr, _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
#Else
Public Function UI_TryGetTitleBarVisibleForHwnd( _
    ByVal TargetHwnd As Long, _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
#End If
'
'==============================================================================
' UI_TryGetTitleBarVisibleForHwnd
'------------------------------------------------------------------------------
' PURPOSE
'   Reads title-bar visibility from the window style of an explicitly supplied
'   top-level window.
'
' WHY THIS EXISTS
'   The snapshot engine must be able to say which window a captured Boolean
'   came from. Reading through Application.Hwnd cannot support that claim,
'   because the answer changes with the active window.
'
' INPUTS
'   TargetHwnd
'     Top-level window to read.
'
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
'   - Rejects a handle that no longer refers to a window.
'   - Reads GWL_STYLE using the API matching the host bitness.
'   - Reports visibility from the WS_CAPTION bit alone.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_InternalIsTitleBarFrameAlive
'   - UI_TryGetWindowStyle
'   - UI_TitleBarBuildRuntimeErrorText
'
' CALLED FROM
'   - M_EXCEL_UI_SNAPSHOT
'   - UI_TryGetTitleBarVisible
'
' NOTES
'   WS_CAPTION is the visibility signal; the other four owned bits travel with
'   it but do not participate in this decision.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim StyleValue          As LongPtr         'Live GWL_STYLE value
#Else
    Dim StyleValue          As Long            'Live GWL_STYLE value
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the style has been read
        UI_TryGetTitleBarVisibleForHwnd = False

    'Initialize the output and the failure message buffer
        IsVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a handle that is absent or no longer refers to a window
        If Not UI_InternalIsTitleBarFrameAlive(TargetHwnd) Then
            FailMsg = "target window is not available; hWnd=" & CStr(TargetHwnd)
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' READ STYLE
'------------------------------------------------------------------------------
    'Read the live style through the shared bitness-aware helper
        If Not UI_TryGetWindowStyle(TargetHwnd, StyleValue, FailMsg) Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Visibility is carried by the caption bit alone
        IsVisible = ((StyleValue And WS_CAPTION) <> 0)
        UI_TryGetTitleBarVisibleForHwnd = True

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


Public Function UI_TryGetTitleBarVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_TryGetTitleBarVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Reads title-bar visibility from the top-level window Excel currently
'   reports as active.
'
' WHY THIS EXISTS
'   Retained as an active-window convenience wrapper for callers that have no
'   target of their own. It is deliberately no longer used by the snapshot
'   engine, which resolves and retains its own handle.
'
' INPUTS
'   IsVisible
'     ByRef. Receives True when WS_CAPTION is set on the active window.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the style was read and IsVisible is meaningful.
'     False => no active handle was available, or the read failed.
'
' BEHAVIOR
'   - Resolves the active handle, then delegates to the explicit-target read.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_TryGetActiveTitleBarHwnd
'   - UI_TryGetTitleBarVisibleForHwnd
'
' CALLED FROM
'   - M_EXCEL_UI
'
' NOTES
'   The value describes the active window at the instant of the call and must
'   not be stored for later restoration. Use the explicit-target entry points
'   for anything that spans two operations.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim TargetHwnd          As LongPtr         'Active top-level window handle
#Else
    Dim TargetHwnd          As Long            'Active top-level window handle
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Assume failure until the delegate reports otherwise
        UI_TryGetTitleBarVisible = False

    'Initialize the output and the failure message buffer
        IsVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' RESOLVE ACTIVE TARGET
'------------------------------------------------------------------------------
    'Resolve the active handle once and hand it to the explicit-target read
        If Not UI_TryGetActiveTitleBarHwnd(TargetHwnd, FailMsg) Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Delegate to the single implementation of the read
        UI_TryGetTitleBarVisible = UI_TryGetTitleBarVisibleForHwnd( _
            TargetHwnd:=TargetHwnd, _
            IsVisible:=IsVisible, _
            FailMsg:=FailMsg)

End Function


#If VBA7 Then
Public Function UI_TrySetTitleBarVisibleForHwndIfNeeded( _
    ByVal TargetHwnd As LongPtr, _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
#Else
Public Function UI_TrySetTitleBarVisibleForHwndIfNeeded( _
    ByVal TargetHwnd As Long, _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
#End If
'
'==============================================================================
' UI_TrySetTitleBarVisibleForHwndIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the requested title-bar state to an explicitly supplied top-level
'   window through the owned-style-bit worker.
'
' WHY THIS EXISTS
'   Title-bar visibility alone is not a sufficient basis for no-op detection,
'   because another owned frame bit may still require restoration while
'   WS_CAPTION already matches, and because a previous call may have left a
'   frame refresh outstanding. The worker evaluates both before deciding that
'   there is nothing to do.
'
' INPUTS
'   TargetHwnd
'     Top-level window to update.
'
'   IsVisible
'     Requested title-bar visibility.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the owned bits and the frame are in the requested state.
'     False => the update was attempted and failed.
'
' BEHAVIOR
'   - Delegates to UI_TrySetTitleBarVisibleForHwndWorker.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_TrySetTitleBarVisibleForHwndWorker
'   - UI_TitleBarBuildRuntimeErrorText
'
' CALLED FROM
'   - M_EXCEL_UI_SNAPSHOT
'   - UI_TrySetTitleBarVisibleIfNeeded
'
' NOTES
'   The IfNeeded decision itself lives in the worker; this entry point exists to
'   present the same naming shape as the M_EXCEL_UI_RUNTIME helpers.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until the worker reports otherwise
        UI_TrySetTitleBarVisibleForHwndIfNeeded = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Delegate the owned-bit merge, write and frame refresh
        UI_TrySetTitleBarVisibleForHwndIfNeeded = _
            UI_TrySetTitleBarVisibleForHwndWorker( _
                TargetHwnd:=TargetHwnd, _
                IsVisible:=IsVisible, _
                FailMsg:=FailMsg)

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


Public Function UI_TrySetTitleBarVisibleIfNeeded( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'==============================================================================
' UI_TrySetTitleBarVisibleIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the requested title-bar state to the top-level window Excel
'   currently reports as active.
'
' WHY THIS EXISTS
'   Retained as an active-window convenience wrapper for the fire-and-forget
'   show/hide entry points, whose documented scope IS the active window. The
'   snapshot engine deliberately does not use it: restoring through whichever
'   window happens to be active is the defect this module was reworked to
'   remove.
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
'     True  => the owned bits and the frame are in the requested state.
'     False => no active handle was available, or the update failed.
'
' BEHAVIOR
'   - Resolves the active handle, then delegates to the explicit-target write.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on failure.
'
' DEPENDENCIES
'   - UI_TryGetActiveTitleBarHwnd
'   - UI_TrySetTitleBarVisibleForHwndIfNeeded
'
' CALLED FROM
'   - M_EXCEL_UI
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim TargetHwnd          As LongPtr         'Active top-level window handle
#Else
    Dim TargetHwnd          As Long            'Active top-level window handle
#End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Assume failure until the delegate reports otherwise
        UI_TrySetTitleBarVisibleIfNeeded = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' RESOLVE ACTIVE TARGET
'------------------------------------------------------------------------------
    'Resolve the active handle once and hand it to the explicit-target write
        If Not UI_TryGetActiveTitleBarHwnd(TargetHwnd, FailMsg) Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Delegate to the single implementation of the write
        UI_TrySetTitleBarVisibleIfNeeded = _
            UI_TrySetTitleBarVisibleForHwndIfNeeded( _
                TargetHwnd:=TargetHwnd, _
                IsVisible:=IsVisible, _
                FailMsg:=FailMsg)

End Function


#If VBA7 Then
Public Function UI_InternalIsTitleBarFrameAlive( _
    ByVal TargetHwnd As LongPtr) _
    As Boolean
#Else
Public Function UI_InternalIsTitleBarFrameAlive( _
    ByVal TargetHwnd As Long) _
    As Boolean
#End If
'
'==============================================================================
' UI_InternalIsTitleBarFrameAlive
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a retained handle still refers to a live window.
'
' WHY THIS EXISTS
'   A captured handle outlives the window it names. Windows is free to reuse a
'   handle value once the original window is destroyed, so a snapshot that
'   retained only a handle could restore state into an unrelated window and
'   report success. Probing before every use turns that silent misdirection
'   into a reported failure.
'
' INPUTS
'   TargetHwnd
'     Handle to probe.
'
' RETURNS
'   Boolean
'     True  => the handle currently refers to a window.
'     False => the handle is zero, destroyed, or could not be probed.
'
' BEHAVIOR
'   - Treats a zero handle as not alive without calling the API.
'
' ERROR POLICY
'   - Does not raise. Any unexpected error is reported as not alive.
'
' DEPENDENCIES
'   None.
'
' CALLED FROM
'   - M_EXCEL_UI_SNAPSHOT
'   - This module
'
' NOTES
'   Handle reuse means a True result proves a window exists, not that it is the
'   same window. The snapshot engine therefore pairs this probe with a retained
'   Window object, which cannot be recycled in the same way.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume not alive until the probe says otherwise
        UI_InternalIsTitleBarFrameAlive = False

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'A zero handle never refers to a window, so do not call the API for it
        If TargetHwnd = 0 Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' PROBE HANDLE
'------------------------------------------------------------------------------
    'IsWindow returns a BOOL, so any non-zero result means the window exists
        UI_InternalIsTitleBarFrameAlive = (IsWindow(TargetHwnd) <> 0)

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
    'A probe that cannot complete must never claim the window is usable
        UI_InternalIsTitleBarFrameAlive = False

End Function


Public Sub UI_InternalResetTitleBarBaseline()
'
'==============================================================================
' UI_InternalResetTitleBarBaseline
'------------------------------------------------------------------------------
' PURPOSE
'   Discards the entire frame-state registry, returning this module to the
'   state it holds before its first title-bar call of the session.
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
' RETURNS
'   None.
'
' BEHAVIOR
'   - Discards every registry entry, including any outstanding refresh debt.
'   - Touches no window style. Live frames are left exactly as they are.
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
'     doing so discards the frame every pending show would restore.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' CLEAR REGISTRY
'------------------------------------------------------------------------------
    'Release every entry and reset the count, leaving live windows untouched
        Erase m_FrameStates
        m_FrameStateCount = 0

    'Drop any armed regression seam so one test cannot leak into the next
        m_InjectFrameRefreshFailure = False

End Sub


#If VBA7 Then
Public Sub UI_InternalResetTitleBarBaselineForHwnd( _
    ByVal TargetHwnd As LongPtr)
#Else
Public Sub UI_InternalResetTitleBarBaselineForHwnd( _
    ByVal TargetHwnd As Long)
#End If
'
'==============================================================================
' UI_InternalResetTitleBarBaselineForHwnd
'------------------------------------------------------------------------------
' PURPOSE
'   Discards the frame state held for one window, leaving every other entry in
'   the registry intact.
'
' WHY THIS EXISTS
'   The multi-window regression cases need to simulate "this module has never
'   seen window B" while window A still holds a captured baseline. Clearing the
'   whole registry would destroy the very state those cases exist to assert on.
'
' INPUTS
'   TargetHwnd
'     Window whose entry is discarded. An unknown handle is ignored.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Removes the matching entry and closes the gap left behind.
'   - Touches no window style.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - UI_FrameStateIndexForHwnd
'
' CALLED FROM
'   - M_EXCEL_UI_REGRESSION_TESTS
'
' NOTES
'   Public only for same-project regression access, on the same basis as
'   UI_InternalResetTitleBarBaseline.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Slot                As Long            'Registry index for TargetHwnd
    Dim ShiftIdx            As Long            'Cursor closing the gap left

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

'------------------------------------------------------------------------------
' LOCATE ENTRY
'------------------------------------------------------------------------------
    'Look the handle up without creating an entry for it
        Slot = UI_FrameStateIndexForHwnd(TargetHwnd, False)

        If Slot < 1 Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' REMOVE ENTRY
'------------------------------------------------------------------------------
    'Shift the tail down over the removed entry, then drop the last slot
        For ShiftIdx = Slot To m_FrameStateCount - 1
            m_FrameStates(ShiftIdx) = m_FrameStates(ShiftIdx + 1)
        Next ShiftIdx

        m_FrameStateCount = m_FrameStateCount - 1

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'A seam must never raise into the harness that called it
        Resume Safe_Exit

End Sub


#If VBA7 Then
Public Function UI_InternalIsFrameRefreshPending( _
    ByVal TargetHwnd As LongPtr) _
    As Boolean
#Else
Public Function UI_InternalIsFrameRefreshPending( _
    ByVal TargetHwnd As Long) _
    As Boolean
#End If
'
'==============================================================================
' UI_InternalIsFrameRefreshPending
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a style write succeeded for a window while its frame
'   refresh did not, leaving a repaint owed.
'
' WHY THIS EXISTS
'   The refresh-debt behavior is invisible from the outside: a caller sees a
'   failure, and the retry happens inside the next call. Without a seam the
'   regression case can only assert that the second call succeeded, which it
'   would also do if the debt had simply been forgotten.
'
' INPUTS
'   TargetHwnd
'     Window to query. An unknown handle reports False.
'
' RETURNS
'   Boolean
'     True  => a frame refresh is owed for this window.
'     False => no debt is recorded, or the window is unknown.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - UI_FrameStateIndexForHwnd
'
' CALLED FROM
'   - M_EXCEL_UI_REGRESSION_TESTS
'
' NOTES
'   Public only for same-project regression access. Not part of the supported
'   public API.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Slot                As Long            'Registry index for TargetHwnd

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume no debt until an entry says otherwise
        UI_InternalIsFrameRefreshPending = False

'------------------------------------------------------------------------------
' READ ENTRY
'------------------------------------------------------------------------------
    'Look the handle up without creating an entry for it
        Slot = UI_FrameStateIndexForHwnd(TargetHwnd, False)

        If Slot < 1 Then
            GoTo Safe_Exit
        End If

        UI_InternalIsFrameRefreshPending = m_FrameStates(Slot).RefreshPending

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
    'A seam must never raise into the harness that called it
        UI_InternalIsFrameRefreshPending = False

End Function


Public Sub UI_InternalInjectFrameRefreshFailure( _
    ByVal FailNextRefresh As Boolean)
'
'==============================================================================
' UI_InternalInjectFrameRefreshFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Arms or disarms a one-shot failure of the next non-client frame refresh.
'
' WHY THIS EXISTS
'   The transactional defect this module now guards against requires a style
'   write to succeed and its SetWindowPos refresh to fail. There is no
'   supported way to make Windows fail that call on demand, so the boundary
'   itself has to provide the seam. Without it the recovery path can be
'   reasoned about but never executed, which is indistinguishable from not
'   having written it.
'
' INPUTS
'   FailNextRefresh
'     True arms the seam; False disarms it.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - The armed state is consumed by the next refresh attempt and is not
'     re-armed, so a test cannot leave the module permanently broken by
'     omitting its own cleanup.
'   - No window is touched while the seam is armed.
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
'   - Production code must never call it.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' SET SEAM
'------------------------------------------------------------------------------
    'Record the armed state for the next refresh attempt to consume
        m_InjectFrameRefreshFailure = FailNextRefresh

End Sub


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
'   - UI_TrySetTitleBarVisibleForHwndWorker
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


#If VBA7 Then
Private Function UI_TrySetTitleBarVisibleForHwndWorker( _
    ByVal TargetHwnd As LongPtr, _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
#Else
Private Function UI_TrySetTitleBarVisibleForHwndWorker( _
    ByVal TargetHwnd As Long, _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
#End If
'
'==============================================================================
' UI_TrySetTitleBarVisibleForHwndWorker
'------------------------------------------------------------------------------
' PURPOSE
'   Shows or hides the title bar of one explicitly supplied top-level window.
'
' WHY THIS EXISTS
'   Restoring an entire previously captured GWL_STYLE value would overwrite
'   unrelated style changes made later by Excel, another add-in or caller code.
'   This routine therefore reads the live style on every call and rewrites only
'   the five bits this module claims, for the one window it was given.
'
' INPUTS
'   TargetHwnd
'     Top-level window to update.
'
'   IsVisible
'     Requested title-bar visibility.
'
'   FailMsg
'     ByRef diagnostic message. Empty on success.
'
' RETURNS
'   Boolean
'     True  => the owned bits and the frame are in the requested state.
'     False => the read, the write or the frame refresh failed.
'
' BEHAVIOR
'   - Rejects a handle that no longer refers to a window.
'   - Resolves the registry entry for the window, creating it on first contact.
'   - Settles any outstanding frame refresh BEFORE evaluating the no-op case.
'   - Refreshes the baseline from the live style while this component does not
'     own a hidden state for the window.
'   - Hiding supplies zero owned bits; showing supplies the stored baseline.
'   - Short-circuits when no owned bit would change and nothing is owed.
'   - Records a refresh debt when the style write succeeds and the repaint
'     does not.
'
' ERROR POLICY
'   - Does not raise.
'   - Returns False and populates FailMsg on the first failing step.
'
' DEPENDENCIES
'   - UI_InternalIsTitleBarFrameAlive
'   - UI_FrameStateIndexForHwnd
'   - UI_TryGetWindowStyle
'   - UI_InternalMergeTitleBarStyleBits
'   - UI_TrySetWindowStyle
'   - UI_TryRefreshWindowFrame
'   - UI_TitleBarBuildRuntimeErrorText
'
' CALLED FROM
'   - UI_TrySetTitleBarVisibleForHwndIfNeeded
'
' NOTES
'   Registry state lives only in module memory. A VBA project reset discards it
'   while the window styles themselves survive in the running Excel process, so
'   the two can disagree after a reset. The zero-baseline fallback below is what
'   keeps a show a real recovery path in that situation.
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
#If VBA7 Then
    Dim CurrentStyle        As LongPtr         'Live GWL_STYLE value
    Dim NewStyle            As LongPtr         'Merged GWL_STYLE value to write
    Dim RestoreBits         As LongPtr         'Owned bits a show will re-apply
#Else
    Dim CurrentStyle        As Long            'Live GWL_STYLE value
    Dim NewStyle            As Long            'Merged GWL_STYLE value to write
    Dim RestoreBits         As Long            'Owned bits a show will re-apply
#End If

    Dim Slot                As Long            'Registry index for TargetHwnd

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume failure until every step has completed
        UI_TrySetTitleBarVisibleForHwndWorker = False

    'Initialize the failure message buffer
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Refuse to write into a handle that no longer names a window. Reporting the
    'miss is the whole point: silently writing through whatever handle is to
    'hand is the failure mode this module was reworked to remove.
        If Not UI_InternalIsTitleBarFrameAlive(TargetHwnd) Then
            FailMsg = "target window is not available; hWnd=" & CStr(TargetHwnd)
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' RESOLVE REGISTRY ENTRY
'------------------------------------------------------------------------------
    'Obtain the per-window entry, creating it on first contact with this frame
        Slot = UI_FrameStateIndexForHwnd(TargetHwnd, True)

        If Slot < 1 Then
            FailMsg = "unable to allocate title-bar frame state"
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' READ CURRENT STYLE
'------------------------------------------------------------------------------
    'Read the live style; unrelated bits will be preserved from this value
        If Not UI_TryGetWindowStyle(TargetHwnd, CurrentStyle, FailMsg) Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' SETTLE OUTSTANDING FRAME REFRESH
'------------------------------------------------------------------------------
    'A previous call wrote the style but could not repaint the frame. Pay that
    'debt before the no-op test below: the style already matches the request, so
    'the short circuit would otherwise report success over a frame Windows has
    'never re-measured.
        If m_FrameStates(Slot).RefreshPending Then

            If Not UI_TryRefreshWindowFrame(TargetHwnd, FailMsg) Then
                GoTo Safe_Exit
            End If

            m_FrameStates(Slot).RefreshPending = False
        End If

'------------------------------------------------------------------------------
' REFRESH BASELINE
'------------------------------------------------------------------------------
    'While this component does not own a hidden state, the live owned bits are
    'the truth: adopt them, so a legitimate frame change made by Excel or
    'another add-in survives the next hide and show. Once the component has
    'hidden the frame the live bits are its own zeros, and the stored baseline
    'is the only surviving record of what a show must restore.
        If Not m_FrameStates(Slot).ComponentHidden Then

            m_FrameStates(Slot).OwnedStyleBits = _
                CurrentStyle And TITLEBAR_OWNED_STYLE_MASK

            m_FrameStates(Slot).HasBaseline = True
        End If

    'Take the baseline a show would re-apply
        RestoreBits = m_FrameStates(Slot).OwnedStyleBits

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
    'Skip the write and the frame refresh when no owned bit would change. Any
    'refresh debt was already settled above, so this really is a no-op.
        If NewStyle = CurrentStyle Then
            m_FrameStates(Slot).ComponentHidden = Not IsVisible
            UI_TrySetTitleBarVisibleForHwndWorker = True
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' WRITE STYLE
'------------------------------------------------------------------------------
    'Write the merged style
        If Not UI_TrySetWindowStyle(TargetHwnd, NewStyle, FailMsg) Then
            GoTo Safe_Exit
        End If

    'The style is committed, so this component now owns the resulting state
    'whether or not the repaint below succeeds.
        m_FrameStates(Slot).ComponentHidden = Not IsVisible

'------------------------------------------------------------------------------
' REFRESH FRAME
'------------------------------------------------------------------------------
    'Recalculate and repaint the non-client frame. A failure here leaves the
    'window in a state Windows has not re-measured, so record the debt before
    'reporting it: the next call must retry the repaint rather than conclude
    'from the matching style bits that there is nothing left to do.
        If Not UI_TryRefreshWindowFrame(TargetHwnd, FailMsg) Then
            m_FrameStates(Slot).RefreshPending = True
            GoTo Safe_Exit
        End If

    'Report success
        UI_TrySetTitleBarVisibleForHwndWorker = True

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
Private Function UI_FrameStateIndexForHwnd( _
    ByVal TargetHwnd As LongPtr, _
    ByVal CreateIfMissing As Boolean) _
    As Long
#Else
Private Function UI_FrameStateIndexForHwnd( _
    ByVal TargetHwnd As Long, _
    ByVal CreateIfMissing As Boolean) _
    As Long
#End If
'
'==============================================================================
' UI_FrameStateIndexForHwnd
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the registry index holding the frame state for one window,
'   optionally creating the entry.
'
' WHY THIS EXISTS
'   Frame state used to be a single handle-and-value pair for the whole
'   process. Under the Single Document Interface that meant operating on a
'   second workbook window destroyed the baseline captured for the first, and
'   a show could then restore the wrong frame or none at all. Keying the state
'   by handle is what makes per-window ownership possible.
'
'   A linear scan is deliberate. An Excel session holds a handful of workbook
'   windows, so the cost is trivial and a Collection or Dictionary would add a
'   dependency and a string-keying step for no measurable gain.
'
' INPUTS
'   TargetHwnd
'     Window to look up.
'
'   CreateIfMissing
'     True to append an entry when the handle is unknown; False to report the
'     miss without mutating the registry.
'
' RETURNS
'   Long
'     1-based registry index, or -1 when the handle is unknown and was not
'     created, or when allocation failed.
'
' BEHAVIOR
'   - Scans the live entries for a matching handle.
'   - Reclaims entries whose window no longer exists before extending the
'     registry, so a long session cannot accumulate dead slots.
'   - Initializes a new entry with no baseline, no ownership and no debt.
'
' ERROR POLICY
'   - Does not raise. Any unexpected error is reported as -1.
'
' DEPENDENCIES
'   - UI_CompactFrameStates
'
' CALLED FROM
'   - UI_TrySetTitleBarVisibleForHwndWorker
'   - UI_InternalResetTitleBarBaselineForHwnd
'   - UI_InternalIsFrameRefreshPending
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScanIdx             As Long            'Cursor over the live entries
    Dim Capacity            As Long            'Slots currently allocated

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Assume the handle is unknown until the scan or the append says otherwise
        UI_FrameStateIndexForHwnd = -1

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'A zero handle never names a window and must never be given an entry
        If TargetHwnd = 0 Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' SCAN EXISTING ENTRIES
'------------------------------------------------------------------------------
    'Return the first entry whose handle matches
        For ScanIdx = 1 To m_FrameStateCount

            If m_FrameStates(ScanIdx).hWnd = TargetHwnd Then
                UI_FrameStateIndexForHwnd = ScanIdx
                GoTo Safe_Exit
            End If

        Next ScanIdx

'------------------------------------------------------------------------------
' STOP WHEN NOT CREATING
'------------------------------------------------------------------------------
    'A pure lookup reports the miss and leaves the registry untouched
        If Not CreateIfMissing Then
            GoTo Safe_Exit
        End If

'------------------------------------------------------------------------------
' RECLAIM DEAD ENTRIES
'------------------------------------------------------------------------------
    'Drop entries whose window has closed before growing the registry
        UI_CompactFrameStates

'------------------------------------------------------------------------------
' ENSURE CAPACITY
'------------------------------------------------------------------------------
    'Allocate on first use, then grow in fixed steps as windows are added
        If m_FrameStateCount = 0 Then
            ReDim m_FrameStates(1 To FRAME_STATE_GROWTH_SLOTS)
        Else
            Capacity = UBound(m_FrameStates)

            If m_FrameStateCount >= Capacity Then
                ReDim Preserve m_FrameStates( _
                    1 To Capacity + FRAME_STATE_GROWTH_SLOTS)
            End If
        End If

'------------------------------------------------------------------------------
' APPEND ENTRY
'------------------------------------------------------------------------------
    'Initialize the new entry explicitly rather than relying on array defaults
        m_FrameStateCount = m_FrameStateCount + 1

        m_FrameStates(m_FrameStateCount).hWnd = TargetHwnd
        m_FrameStates(m_FrameStateCount).OwnedStyleBits = 0
        m_FrameStates(m_FrameStateCount).HasBaseline = False
        m_FrameStates(m_FrameStateCount).ComponentHidden = False
        m_FrameStates(m_FrameStateCount).RefreshPending = False

        UI_FrameStateIndexForHwnd = m_FrameStateCount

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
    'A registry that cannot be extended must report the miss, not a bad index
        UI_FrameStateIndexForHwnd = -1

End Function


Private Sub UI_CompactFrameStates()
'
'==============================================================================
' UI_CompactFrameStates
'------------------------------------------------------------------------------
' PURPOSE
'   Removes registry entries whose window no longer exists.
'
' WHY THIS EXISTS
'   Keying state by handle trades one stale singleton for a collection that
'   would otherwise grow for the life of the Excel session, one entry per
'   workbook window ever touched. Reclaiming closed windows at the moment the
'   registry would have to grow bounds it by the number of windows actually
'   open rather than by the number ever opened.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Compacts survivors toward the front of the array, preserving order.
'   - Reduces the live count; the allocated capacity is deliberately kept, so
'     a session that repeatedly opens and closes windows does not reallocate.
'   - Discards any refresh debt held for a window that has closed, which is
'     correct: there is no longer a frame to repaint.
'
' ERROR POLICY
'   - Does not raise.
'
' DEPENDENCIES
'   - UI_InternalIsTitleBarFrameAlive
'
' CALLED FROM
'   - UI_FrameStateIndexForHwnd
'
' UPDATED
'   2026-08-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ReadIdx             As Long            'Cursor over existing entries
    Dim KeepCount           As Long            'Entries retained so far

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to the error handler
        On Error GoTo Err_Handler

    'Nothing to compact before the registry has been allocated
        If m_FrameStateCount = 0 Then
            GoTo Safe_Exit
        End If

        KeepCount = 0

'------------------------------------------------------------------------------
' COMPACT SURVIVORS
'------------------------------------------------------------------------------
    'Copy every entry whose window still exists toward the front of the array
        For ReadIdx = 1 To m_FrameStateCount

            If UI_InternalIsTitleBarFrameAlive(m_FrameStates(ReadIdx).hWnd) Then

                KeepCount = KeepCount + 1

                If KeepCount <> ReadIdx Then
                    m_FrameStates(KeepCount) = m_FrameStates(ReadIdx)
                End If

            End If

        Next ReadIdx

    'Adopt the retained count; capacity above it is left allocated for reuse
        m_FrameStateCount = KeepCount

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
Safe_Exit:
    'Exit before the error-handler block
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Leave the registry exactly as it was rather than truncating it on error
        Resume Safe_Exit

End Sub


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
'   - UI_TrySetTitleBarVisibleForHwndWorker
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
'   - UI_TrySetTitleBarVisibleForHwndWorker
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
'   - UI_TrySetTitleBarVisibleForHwndWorker
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
' CONSUME REGRESSION SEAM
'------------------------------------------------------------------------------
    'When armed, report failure without touching the window. The seam is
    'one-shot, so a test that forgets to disarm it cannot leave the module
    'permanently unable to repaint a frame.
        If m_InjectFrameRefreshFailure Then
            m_InjectFrameRefreshFailure = False
            FailMsg = "SetWindowPos failure injected by the regression harness"
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


