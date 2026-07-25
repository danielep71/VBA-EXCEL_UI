Attribute VB_Name = "M_EXCEL_UI_RESULTS"
'==============================================================================
'                    MODULE: M_EXCEL_UI_RESULTS
'------------------------------------------------------------------------------
' PURPOSE
'   Centralize ordered structured failures and Immediate Window diagnostics
'
' BEHAVIOR
'   - Optional failure lists are 1-based String arrays stored in a Variant
'   - Entries use deterministic Stage | Detail formatting
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

Public Sub UI_ResultClear(ByRef FailureCount As Long, ByRef FailureList As _
    Variant, ByVal CaptureFailureList As Boolean)
'==============================================================================
' PURPOSE
'   Initialize structured-result outputs in a clean-success state
'==============================================================================
        FailureCount = 0
        If CaptureFailureList Then FailureList = Empty
End Sub

Public Sub UI_ResultAdd(ByRef Succeeded As Boolean, ByRef FailureCount As Long, _
    ByRef FailureList As Variant, ByVal CaptureFailureList As Boolean, _
    ByVal Stage As String, ByVal Detail As String)
'==============================================================================
' PURPOSE
'   Append one ordered structured failure
'==============================================================================
    Dim Entries() As String
        Succeeded = False
        FailureCount = FailureCount + 1
        If Not CaptureFailureList Then Exit Sub
        If IsEmpty(FailureList) Then
            ReDim Entries(1 To 1)
        Else
            Entries = FailureList
            ReDim Preserve Entries(1 To FailureCount)
        End If
        Entries(FailureCount) = Stage & " | " & Detail
        FailureList = Entries
End Sub

Public Sub UI_ResultHandleFailure(ByVal ProcName As String, _
    ByVal LogFailures As Boolean, ByRef Succeeded As Boolean, _
    ByRef FailureCount As Long, ByRef FailureList As Variant, _
    ByVal CaptureFailureList As Boolean, ByVal Stage As String, _
    ByVal Detail As String)
'==============================================================================
' PURPOSE
'   Record one failure and optionally log it
'==============================================================================
        UI_ResultAdd Succeeded, FailureCount, FailureList, CaptureFailureList, _
            Stage, Detail
        If LogFailures Then UI_ResultLogFailure ProcName, Stage, Detail
End Sub

Public Function UI_ResultRuntimeErrorText() As String
'==============================================================================
' PURPOSE
'   Build a best-effort diagnostic from the active Err object
'==============================================================================
        On Error Resume Next
        UI_ResultRuntimeErrorText = CStr(Err.Number) & ": " & Err.Description
        If Len(Err.Source) > 0 Then
            UI_ResultRuntimeErrorText = UI_ResultRuntimeErrorText & _
                " | Source: " & Err.Source
        End If
        If Erl <> 0 Then
            UI_ResultRuntimeErrorText = UI_ResultRuntimeErrorText & _
                " | Line: " & CStr(Erl)
        End If
End Function

Public Sub UI_ResultLogFailure(ByVal ProcName As String, ByVal Stage As String, _
    ByVal Detail As String)
'==============================================================================
' PURPOSE
'   Write one consistent diagnostic line to the Immediate Window
'==============================================================================
        On Error Resume Next
        Debug.Print ProcName & " failed @ " & Stage & " | " & Detail
End Sub
