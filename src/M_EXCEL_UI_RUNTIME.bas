Attribute VB_Name = "M_EXCEL_UI_RUNTIME"
'==============================================================================
'                     MODULE: M_EXCEL_UI_RUNTIME
'------------------------------------------------------------------------------
' PURPOSE
'   Provide shared fail-soft runtime services used by the public facade and the
'   snapshot engine.
'
' WHY
'   Both M_EXCEL_UI and M_EXCEL_UI_SNAPSHOT require identical result buffering,
'   diagnostics, Ribbon access, Boolean property access, and quiet-update scopes.
'   Centralizing those operations avoids duplicated logic and circular module
'   dependencies.
'
' INTERNAL SURFACE
'   - UI_RuntimeHandleFailure
'   - UI_RuntimeClearResultBuffer
'   - UI_RuntimeBeginQuietUpdate
'   - UI_RuntimeEndQuietUpdate
'   - UI_RuntimeTrySetRibbonVisibleIfNeeded
'   - UI_RuntimeTrySetBooleanPropertyIfNeeded
'   - UI_RuntimeTryGetRibbonVisible
'   - UI_RuntimeTryGetBooleanProperty
'   - UI_RuntimeBuildErrorText
'   - UI_RuntimeLogFailure
'
' ERROR POLICY
'   - Entry points are fail-soft.
'   - No user-interface messages are displayed.
'   - Diagnostics preserve the established ordered "Stage | Detail" contract.
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

Public Sub UI_RuntimeHandleFailure( _
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
'                           UI_RuntimeHandleFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Record one best-effort operation failure and optionally log it.
'
' NOTES
'   The procedure name is retained for compatibility with the established
'   internal apply path, but the helper is also used by snapshot operations.
'
' ERROR POLICY
'   Does not raise.
'
' UPDATED
'   2026-07-29
'==============================================================================
'

        UI_RuntimeAddFailure _
            Succeeded:=Succeeded, _
            FailureCount:=FailureCount, _
            FailureList:=FailureList, _
            CaptureFailureList:=CaptureFailureList, _
            Stage:=Stage, _
            Detail:=Detail

        If LogFailures Then
            UI_RuntimeLogFailure ProcName, Stage, Detail
        End If

End Sub

Public Sub UI_RuntimeClearResultBuffer( _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean)

'
'==============================================================================
'                           UI_RuntimeClearResultBuffer
'------------------------------------------------------------------------------
' PURPOSE
'   Initialize structured result buffers.
'
' ERROR POLICY
'   Does not raise.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        FailureCount = 0

        If CaptureFailureList Then
            FailureList = Empty
        End If

End Sub

Private Sub UI_RuntimeAddFailure( _
    ByRef Succeeded As Boolean, _
    ByRef FailureCount As Long, _
    ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean, _
    ByVal Stage As String, _
    ByVal Detail As String)

'
'==============================================================================
'                          UI_RuntimeAddFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Append one failure to the standard Boolean / count / list result contract.
'
' ERROR POLICY
'   Does not raise.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Arr() As String

'------------------------------------------------------------------------------
' UPDATE STATUS
'------------------------------------------------------------------------------
        Succeeded = False
        FailureCount = FailureCount + 1

'------------------------------------------------------------------------------
' APPEND TEXT
'------------------------------------------------------------------------------
        If CaptureFailureList Then
            If IsEmpty(FailureList) Then
                ReDim Arr(1 To 1)
            Else
                Arr = FailureList
                ReDim Preserve Arr(1 To FailureCount)
            End If

            Arr(FailureCount) = Stage & " | " & Detail
            FailureList = Arr
        End If

End Sub

Public Sub UI_RuntimeBeginQuietUpdate( _
    ByRef OldScreenUpdating As Boolean, _
    ByRef QuietModeChanged As Boolean)

'
'==============================================================================
'                          UI_RuntimeBeginQuietUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Enter a best-effort ScreenUpdating suppression scope.
'
' ERROR POLICY
'   Suppresses errors locally.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        On Error Resume Next

        OldScreenUpdating = Application.ScreenUpdating
        QuietModeChanged = False

        If OldScreenUpdating Then
            Application.ScreenUpdating = False
            QuietModeChanged = True
        End If

End Sub

Public Sub UI_RuntimeEndQuietUpdate( _
    ByVal OldScreenUpdating As Boolean, _
    ByVal QuietModeChanged As Boolean)

'
'==============================================================================
'                           UI_RuntimeEndQuietUpdate
'------------------------------------------------------------------------------
' PURPOSE
'   Restore ScreenUpdating when this module changed it.
'
' ERROR POLICY
'   Suppresses errors locally.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        On Error Resume Next

        If QuietModeChanged Then
            Application.ScreenUpdating = OldScreenUpdating
        End If

End Sub

Public Function UI_RuntimeTrySetRibbonVisibleIfNeeded( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                     UI_RuntimeTrySetRibbonVisibleIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Set Ribbon visibility only when required.
'
' RETURNS
'   TRUE when already correct or successfully updated.
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
    Dim CurrentVisible As Boolean

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_RuntimeTrySetRibbonVisibleIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT
'------------------------------------------------------------------------------
        If UI_RuntimeTryGetRibbonVisible(CurrentVisible, FailMsg) Then
            If CurrentVisible = IsVisible Then
                UI_RuntimeTrySetRibbonVisibleIfNeeded = True
                GoTo SafeExit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
        FailMsg = vbNullString

        UI_RuntimeTrySetRibbonVisibleIfNeeded = _
            UI_RuntimeTrySetRibbonVisible(IsVisible, FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_RuntimeBuildErrorText

End Function

Public Function UI_RuntimeTrySetBooleanPropertyIfNeeded( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByVal NewValue As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                   UI_RuntimeTrySetBooleanPropertyIfNeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Set a Boolean property only when required.
'
' RETURNS
'   TRUE when already correct or successfully updated.
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
    Dim CurrentValue As Boolean

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_RuntimeTrySetBooleanPropertyIfNeeded = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' SHORT-CIRCUIT
'------------------------------------------------------------------------------
        If UI_RuntimeTryGetBooleanProperty( _
            Target, PropertyName, CurrentValue, FailMsg) Then

            If CurrentValue = NewValue Then
                UI_RuntimeTrySetBooleanPropertyIfNeeded = True
                GoTo SafeExit
            End If
        End If

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
        FailMsg = vbNullString

        UI_RuntimeTrySetBooleanPropertyIfNeeded = _
            UI_RuntimeTrySetBooleanProperty(Target, PropertyName, NewValue, FailMsg)

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_RuntimeBuildErrorText

End Function

Public Function UI_RuntimeTryGetRibbonVisible( _
    ByRef IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                         UI_RuntimeTryGetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Read current Ribbon visibility using CommandBars with an Excel4 fallback.
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
    Dim V As Variant

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_RuntimeTryGetRibbonVisible = False
        IsVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' TRY COMMANDBARS
'------------------------------------------------------------------------------
        On Error Resume Next
            IsVisible = Application.CommandBars("Ribbon").Visible

        If Err.Number = 0 Then
            On Error GoTo Fail
            UI_RuntimeTryGetRibbonVisible = True
            GoTo SafeExit
        End If

        Err.Clear
        On Error GoTo Fail

'------------------------------------------------------------------------------
' TRY EXCEL4 FALLBACK
'------------------------------------------------------------------------------
        On Error Resume Next
            V = Application.ExecuteExcel4Macro("Get.ToolBar(7,""Ribbon"")")

        If Err.Number = 0 Then
            On Error GoTo Fail
            IsVisible = CBool(V)
            UI_RuntimeTryGetRibbonVisible = True
            GoTo SafeExit
        End If

        FailMsg = CStr(Err.Number) & ": " & Err.Description
        Err.Clear
        On Error GoTo Fail

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_RuntimeBuildErrorText

End Function

Public Function UI_RuntimeTryGetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByRef ValueOut As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                         UI_RuntimeTryGetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Read a Boolean property through CallByName.
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
    Dim V As Variant

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_RuntimeTryGetBooleanProperty = False
        ValueOut = False
        FailMsg = vbNullString

        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo SafeExit
        End If

        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        V = CallByName(Target, PropertyName, VbGet)
        ValueOut = CBool(V)
        UI_RuntimeTryGetBooleanProperty = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_RuntimeBuildErrorText

End Function

Private Function UI_RuntimeTrySetRibbonVisible( _
    ByVal IsVisible As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                           UI_RuntimeTrySetRibbonVisible
'------------------------------------------------------------------------------
' PURPOSE
'   Show or hide the Ribbon using Excel4 macro execution.
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
    Dim MacroText As String

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_RuntimeTrySetRibbonVisible = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
        If IsVisible Then
            MacroText = "Show.TOOLBAR(""Ribbon"",True)"
        Else
            MacroText = "Show.TOOLBAR(""Ribbon"",False)"
        End If

        Application.ExecuteExcel4Macro MacroText

        UI_RuntimeTrySetRibbonVisible = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_RuntimeBuildErrorText

End Function

Private Function UI_RuntimeTrySetBooleanProperty( _
    ByVal Target As Object, _
    ByVal PropertyName As String, _
    ByVal NewValue As Boolean, _
    ByRef FailMsg As String) As Boolean

'
'==============================================================================
'                           UI_RuntimeTrySetBooleanProperty
'------------------------------------------------------------------------------
' PURPOSE
'   Write a Boolean property through CallByName.
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
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        UI_RuntimeTrySetBooleanProperty = False
        FailMsg = vbNullString

        If Target Is Nothing Then
            FailMsg = "target object is Nothing"
            GoTo SafeExit
        End If

        If Len(PropertyName) = 0 Then
            FailMsg = "property name is empty"
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
        CallByName Target, PropertyName, VbLet, NewValue
        UI_RuntimeTrySetBooleanProperty = True

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        FailMsg = UI_RuntimeBuildErrorText

End Function

Public Function UI_RuntimeBuildErrorText() As String

'
'==============================================================================
'                           UI_RuntimeBuildErrorText
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
'   2026-07-25
'==============================================================================
'

        On Error Resume Next

        UI_RuntimeBuildErrorText = _
            CStr(Err.Number) & ": " & Err.Description & _
            IIf(Len(Err.Source) > 0, _
                " | Source: " & Err.Source, _
                vbNullString) & _
            IIf(Erl <> 0, _
                " | Line: " & CStr(Erl), _
                vbNullString)

End Function

Public Sub UI_RuntimeLogFailure( _
    ByVal ProcName As String, _
    ByVal Stage As String, _
    ByVal Detail As String)

'
'==============================================================================
'                                UI_RuntimeLogFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Write a consistent diagnostic line to the Immediate Window.
'
' ERROR POLICY
'   Suppresses logging errors locally.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        On Error Resume Next

        Debug.Print ProcName & " failed @ " & Stage & " | " & Detail

End Sub
