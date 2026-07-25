Attribute VB_Name = "M_EXCEL_UI_TITLEBAR"
'==============================================================================
'                    MODULE: M_EXCEL_UI_TITLEBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Own bitness-safe WinAPI title-bar control for Application.Hwnd
'
' WHY THIS EXISTS
'   Version 1.1.0 restores only the frame-style bits controlled by this project.
'   Unrelated style changes made by Excel or another component are preserved
'
' ERROR POLICY
'   - Helpers are best-effort and never intentionally raise to callers
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

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias _
            "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) _
            As LongPtr
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
        ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, _
        ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
        ByVal uFlags As Long) As Long
    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
    Private Declare PtrSafe Sub SetLastError Lib "kernel32" ( _
        ByVal dwErrCode As Long)
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
        ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Private Declare Function GetLastError Lib "kernel32" () As Long
    Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
#End If

    Private Const GWL_STYLE         As Long = -16
    Private Const WS_CAPTION        As Long = &HC00000
    Private Const WS_SYSMENU        As Long = &H80000
    Private Const WS_THICKFRAME     As Long = &H40000
    Private Const WS_MINIMIZEBOX    As Long = &H20000
    Private Const WS_MAXIMIZEBOX    As Long = &H10000
    Private Const UI_OWNED_MASK     As Long = &HCF0000
    Private Const SWP_NOSIZE        As Long = &H1
    Private Const SWP_NOMOVE        As Long = &H2
    Private Const SWP_NOZORDER      As Long = &H4
    Private Const SWP_FRAMECHANGED  As Long = &H20
    Private Const SWP_NOOWNERZORDER As Long = &H200

#If VBA7 Then
    Private m_OwnerHwnd          As LongPtr
    Private m_OriginalOwnedBits  As LongPtr
#Else
    Private m_OwnerHwnd          As Long
    Private m_OriginalOwnedBits  As Long
#End If
    Private m_HasOwnership       As Boolean

Public Function UI_TitleBarTryGetVisible(ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Read title-bar visibility from the WS_CAPTION style bit
'==============================================================================
#If VBA7 Then
    Dim H As LongPtr
    Dim StyleValue As LongPtr
#Else
    Dim H As Long
    Dim StyleValue As Long
#End If
        On Error GoTo Fail
        H = Application.hWnd
        If H = 0 Then
            FailMsg = "invalid Excel window handle"
            Exit Function
        End If
        If Not UI_TitleBarTryGetStyle(H, StyleValue, FailMsg) Then Exit Function
        IsVisible = ((StyleValue And WS_CAPTION) <> 0)
        UI_TitleBarTryGetVisible = True
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function

Public Function UI_TitleBarTrySetVisibleIfNeeded(ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Skip the WinAPI write when current visibility already matches the target
'==============================================================================
    Dim CurrentVisible As Boolean
        If UI_TitleBarTryGetVisible(CurrentVisible, FailMsg) Then
            If CurrentVisible = IsVisible Then
                If IsVisible Then UI_TitleBarClearOwnership
                UI_TitleBarTrySetVisibleIfNeeded = True
                Exit Function
            End If
        End If
        FailMsg = vbNullString
        UI_TitleBarTrySetVisibleIfNeeded = UI_TitleBarTrySetVisible( _
            IsVisible, FailMsg)
End Function

Public Function UI_TitleBarTrySetVisible(ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Hide or show only the window-style bits owned by this project
'
' BEHAVIOR
'   - Hide captures the owned bits once for the current Application.Hwnd
'   - Show merges captured owned bits into the current full style
'   - Unowned bits are never restored from a stale full-style snapshot
'==============================================================================
#If VBA7 Then
    Dim H As LongPtr
    Dim CurrentStyle As LongPtr
    Dim NewStyle As LongPtr
#Else
    Dim H As Long
    Dim CurrentStyle As Long
    Dim NewStyle As Long
#End If
        On Error GoTo Fail
        FailMsg = vbNullString
        H = Application.hWnd
        If H = 0 Then
            FailMsg = "invalid Excel window handle"
            Exit Function
        End If
        If Not UI_TitleBarTryGetStyle(H, CurrentStyle, FailMsg) Then Exit Function

        If IsVisible Then
            NewStyle = CurrentStyle And Not UI_OWNED_MASK
            If m_HasOwnership And m_OwnerHwnd = H Then
                NewStyle = NewStyle Or m_OriginalOwnedBits
            Else
                NewStyle = NewStyle Or UI_OWNED_MASK
            End If
        Else
            If (Not m_HasOwnership) Or m_OwnerHwnd <> H Then
                m_OwnerHwnd = H
                m_OriginalOwnedBits = CurrentStyle And UI_OWNED_MASK
                m_HasOwnership = True
            End If
            NewStyle = CurrentStyle And Not UI_OWNED_MASK
        End If

        If NewStyle <> CurrentStyle Then
            If Not UI_TitleBarTrySetStyle(H, NewStyle, FailMsg) Then Exit Function
            If Not UI_TitleBarTryRefreshFrame(H, FailMsg) Then Exit Function
        End If

        If IsVisible Then UI_TitleBarClearOwnership
        UI_TitleBarTrySetVisible = True
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function

Public Sub UI_TitleBarClearOwnership()
'==============================================================================
' PURPOSE
'   Discard the in-memory owned-bit baseline
'==============================================================================
        m_HasOwnership = False
        m_OwnerHwnd = 0
        m_OriginalOwnedBits = 0
End Sub

#If VBA7 Then
Private Function UI_TitleBarTryGetStyle(ByVal H As LongPtr, _
    ByRef StyleOut As LongPtr, ByRef FailMsg As String) As Boolean
#Else
Private Function UI_TitleBarTryGetStyle(ByVal H As Long, ByRef StyleOut As Long, _
    ByRef FailMsg As String) As Boolean
#End If
'==============================================================================
' PURPOSE
'   Read GWL_STYLE and distinguish a valid zero from API failure
'==============================================================================
    Dim LastErr As Long
        On Error GoTo Fail
        SetLastError 0
#If VBA7 Then
    #If Win64 Then
        StyleOut = GetWindowLongPtr(H, GWL_STYLE)
    #Else
        StyleOut = GetWindowLong(H, GWL_STYLE)
    #End If
#Else
        StyleOut = GetWindowLong(H, GWL_STYLE)
#End If
        LastErr = GetLastError
        If StyleOut = 0 And LastErr <> 0 Then
            FailMsg = "GetWindowLong/GetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)
            Exit Function
        End If
        UI_TitleBarTryGetStyle = True
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function

#If VBA7 Then
Private Function UI_TitleBarTrySetStyle(ByVal H As LongPtr, _
    ByVal NewStyle As LongPtr, ByRef FailMsg As String) As Boolean
#Else
Private Function UI_TitleBarTrySetStyle(ByVal H As Long, ByVal NewStyle As Long, _
    ByRef FailMsg As String) As Boolean
#End If
'==============================================================================
' PURPOSE
'   Write GWL_STYLE and distinguish a valid zero return from API failure
'==============================================================================
#If VBA7 Then
    Dim PreviousStyle As LongPtr
#Else
    Dim PreviousStyle As Long
#End If
    Dim LastErr As Long
        On Error GoTo Fail
        SetLastError 0
#If VBA7 Then
    #If Win64 Then
        PreviousStyle = SetWindowLongPtr(H, GWL_STYLE, NewStyle)
    #Else
        PreviousStyle = SetWindowLong(H, GWL_STYLE, NewStyle)
    #End If
#Else
        PreviousStyle = SetWindowLong(H, GWL_STYLE, NewStyle)
#End If
        LastErr = GetLastError
        If PreviousStyle = 0 And LastErr <> 0 Then
            FailMsg = "SetWindowLong/SetWindowLongPtr failed; GetLastError=" & _
                CStr(LastErr)
            Exit Function
        End If
        UI_TitleBarTrySetStyle = True
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function

#If VBA7 Then
Private Function UI_TitleBarTryRefreshFrame(ByVal H As LongPtr, _
    ByRef FailMsg As String) As Boolean
#Else
Private Function UI_TitleBarTryRefreshFrame(ByVal H As Long, _
    ByRef FailMsg As String) As Boolean
#End If
'==============================================================================
' PURPOSE
'   Recalculate the non-client frame without moving or resizing the window
'==============================================================================
    Dim ApiOK As Long
        On Error GoTo Fail
        SetLastError 0
        ApiOK = SetWindowPos(H, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or _
            SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED)
        If ApiOK = 0 Then
            FailMsg = "SetWindowPos failed; GetLastError=" & _
                CStr(GetLastError)
            Exit Function
        End If
        UI_TitleBarTryRefreshFrame = True
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function
