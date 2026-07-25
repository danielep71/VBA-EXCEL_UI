Attribute VB_Name = "M_EXCEL_UI_HOST"
'==============================================================================
'                       MODULE: M_EXCEL_UI_HOST
'------------------------------------------------------------------------------
' PURPOSE
'   Isolate Excel object-model, Ribbon, quiet-update, and Window helpers
'
' ERROR POLICY
'   - Helpers are best-effort and never intentionally raise to callers
'   - Boolean helpers return FALSE and populate FailMsg on failure
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

Public Sub UI_HostBeginQuietUpdate(ByRef OldScreenUpdating As Boolean, _
    ByRef QuietModeChanged As Boolean)
'==============================================================================
' PURPOSE
'   Suppress ScreenUpdating while preserving the caller's prior state
'==============================================================================
        On Error Resume Next
        OldScreenUpdating = Application.ScreenUpdating
        QuietModeChanged = False
        If OldScreenUpdating Then
            Application.ScreenUpdating = False
            QuietModeChanged = (Err.Number = 0)
            Err.Clear
        End If
End Sub

Public Sub UI_HostEndQuietUpdate(ByVal OldScreenUpdating As Boolean, _
    ByVal QuietModeChanged As Boolean)
'==============================================================================
' PURPOSE
'   Restore ScreenUpdating only when this project changed it
'==============================================================================
        On Error Resume Next
        If QuietModeChanged Then Application.ScreenUpdating = OldScreenUpdating
End Sub

Public Function UI_HostTryGetBooleanProperty(ByVal Target As Object, _
    ByVal PropertyName As String, ByRef ValueOut As Boolean, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Read a Boolean property through CallByName
'==============================================================================
    Dim V As Variant
        On Error GoTo Fail
        FailMsg = vbNullString
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            Exit Function
        End If
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            Exit Function
        End If
        V = CallByName(Target, PropertyName, VbGet)
        ValueOut = CBool(V)
        UI_HostTryGetBooleanProperty = True
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function

Public Function UI_HostTrySetBooleanProperty(ByVal Target As Object, _
    ByVal PropertyName As String, ByVal NewValue As Boolean, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Write a Boolean property through CallByName
'==============================================================================
        On Error GoTo Fail
        FailMsg = vbNullString
        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            Exit Function
        End If
        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            Exit Function
        End If
        CallByName Target, PropertyName, VbLet, NewValue
        UI_HostTrySetBooleanProperty = True
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function

Public Function UI_HostTrySetBooleanPropertyIfNeeded(ByVal Target As Object, _
    ByVal PropertyName As String, ByVal NewValue As Boolean, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Skip a Boolean write when a readable current value already matches
'==============================================================================
    Dim CurrentValue As Boolean
        If UI_HostTryGetBooleanProperty(Target, PropertyName, CurrentValue, _
            FailMsg) Then
            If CurrentValue = NewValue Then
                UI_HostTrySetBooleanPropertyIfNeeded = True
                Exit Function
            End If
        End If
        FailMsg = vbNullString
        UI_HostTrySetBooleanPropertyIfNeeded = UI_HostTrySetBooleanProperty( _
            Target, PropertyName, NewValue, FailMsg)
End Function

Public Function UI_HostTryGetRibbonVisible(ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Read Ribbon visibility using CommandBars with an Excel 4 fallback
'==============================================================================
    Dim V As Variant
    Dim FirstFailure As String
        On Error GoTo Fail
        FailMsg = vbNullString
        On Error Resume Next
        IsVisible = Application.CommandBars("Ribbon").Visible
        If Err.Number = 0 Then
            On Error GoTo Fail
            UI_HostTryGetRibbonVisible = True
            Exit Function
        End If
        FirstFailure = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        V = Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")
        If Err.Number = 0 Then
            On Error GoTo Fail
            IsVisible = CBool(V)
            UI_HostTryGetRibbonVisible = True
            Exit Function
        End If
        FailMsg = "CommandBars read failed (" & FirstFailure & _
            "); Excel 4 fallback failed (" & CStr(Err.Number) & ": " & _
            Err.Description & ")"
        Err.Clear
        On Error GoTo Fail
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function

Public Function UI_HostTrySetRibbonVisible(ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Show or hide the Ribbon through a fixed Excel 4 macro command
'==============================================================================
    Dim MacroText As String
        On Error GoTo Fail
        FailMsg = vbNullString
        If IsVisible Then
            MacroText = "Show.TOOLBAR(""Ribbon"",True)"
        Else
            MacroText = "Show.TOOLBAR(""Ribbon"",False)"
        End If
        Application.ExecuteExcel4Macro MacroText
        UI_HostTrySetRibbonVisible = True
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function

Public Function UI_HostTrySetRibbonVisibleIfNeeded(ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Change Ribbon visibility only when its readable state differs
'==============================================================================
    Dim CurrentVisible As Boolean
        If UI_HostTryGetRibbonVisible(CurrentVisible, FailMsg) Then
            If CurrentVisible = IsVisible Then
                UI_HostTrySetRibbonVisibleIfNeeded = True
                Exit Function
            End If
        End If
        FailMsg = vbNullString
        UI_HostTrySetRibbonVisibleIfNeeded = UI_HostTrySetRibbonVisible( _
            IsVisible, FailMsg)
End Function

Public Function UI_HostIsCurrentExcelWindow(ByVal TargetWindow As Window, _
    ByRef FailMsg As String) As Boolean
'==============================================================================
' PURPOSE
'   Verify exact Window-object membership in Application.Windows
'==============================================================================
    Dim W As Window
        On Error GoTo Fail
        FailMsg = vbNullString
        If TargetWindow Is Nothing Then
            FailMsg = "target window is Nothing"
            Exit Function
        End If
        For Each W In Application.Windows
            If W Is TargetWindow Then
                UI_HostIsCurrentExcelWindow = True
                Exit Function
            End If
        Next W
        FailMsg = "target window does not belong to this Excel instance"
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function

Public Function UI_HostWindowLabel(ByVal TargetWindow As Window) As String
'==============================================================================
' PURPOSE
'   Build a best-effort label for diagnostics
'==============================================================================
        On Error Resume Next
        UI_HostWindowLabel = TargetWindow.Caption
        If Len(UI_HostWindowLabel) = 0 Then _
            UI_HostWindowLabel = "Unnamed Excel window"
End Function

#If VBA7 Then
Public Function UI_HostTryGetWindowHwnd(ByVal TargetWindow As Window, _
    ByRef HwndOut As LongPtr, ByRef FailMsg As String) As Boolean
#Else
Public Function UI_HostTryGetWindowHwnd(ByVal TargetWindow As Window, _
    ByRef HwndOut As Long, ByRef FailMsg As String) As Boolean
#End If
'==============================================================================
' PURPOSE
'   Read a Window handle for snapshot identity matching
'==============================================================================
        On Error GoTo Fail
        FailMsg = vbNullString
        If TargetWindow Is Nothing Then
            FailMsg = "target window is Nothing"
            Exit Function
        End If
        HwndOut = TargetWindow.hWnd
        If HwndOut = 0 Then
            FailMsg = "window handle is zero"
            Exit Function
        End If
        UI_HostTryGetWindowHwnd = True
        Exit Function
Fail:
        FailMsg = UI_ResultRuntimeErrorText
End Function
