Attribute VB_Name = "M_EXCEL_UI_SNAPSHOT_ID_TESTS"
'==============================================================================
'              MODULE: M_EXCEL_UI_SNAPSHOT_IDENTITY_TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Provide focused regression coverage for v1.1.0 identity-safe per-window
'   snapshot restoration.
'
' WHY
'   The v1.0.1 implementation restored window-level state by
'   Application.Windows collection index. Closing one captured window and
'   opening another could therefore redirect saved state to the replacement
'   window occupying the same collection position.
'
' PUBLIC SURFACE
'   - Test_EXCEL_UI_RunSnapshotIdentity
'
' TEST SCOPE
'   - exact captured-window restoration
'   - closed captured-window handling
'   - replacement-window non-interference
'   - object-identity matching independent of caption and collection position
'
' STATE MANAGEMENT
'   - Refuses to run when an explicit EXCEL_UI snapshot already exists.
'   - Creates and closes only temporary additional windows for ThisWorkbook.
'   - Restores the original managed state of the anchor window.
'
' ERROR POLICY
'   - Raises after best-effort cleanup when an assertion or unexpected failure
'     occurs.
'
' COMPATIBILITY
'   - Windows Excel.
'   - Requires M_EXCEL_UI v1.1.0 in the same VBA project.
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

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' TEST CONSTANTS
'------------------------------------------------------------------------------
    Private Const TEST_ERR_BASE As Long = vbObjectError + 4810


Public Sub Test_EXCEL_UI_RunSnapshotIdentity()

'
'==============================================================================
'                  Test_EXCEL_UI_RunSnapshotIdentity
'------------------------------------------------------------------------------
' PURPOSE
'   Verify that reset targets the exact captured Excel Window object and never a
'   newly created replacement window.
'
' WHY
'   A replacement window can occupy the same Application.Windows index as a
'   captured window that was closed after snapshot creation. Index-based restore
'   can then silently apply the wrong state to the replacement.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a temporary captured window.
'   - Establishes different baselines on the anchor and captured windows.
'   - Captures the EXCEL_UI snapshot.
'   - Closes the captured window.
'   - Creates a replacement window and assigns a sentinel state.
'   - Resets to snapshot.
'   - Verifies the anchor restored and the replacement remained unchanged.
'
' ERROR POLICY
'   - Raises after best-effort cleanup.
'
' DEPENDENCIES
'   - UI_CaptureExcelUIState
'   - UI_ResetExcelUIToSnapshot
'   - UI_ClearExcelUIStateSnapshot
'
' UPDATED
'   2026-07-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorWindow      As Window
    Dim CapturedWindow    As Window
    Dim ReplacementWindow As Window

    Dim SavedHeadings     As Boolean
    Dim SavedWorkbookTabs As Boolean
    Dim SavedGridlines    As Boolean

    Dim CapturedIndex     As Long
    Dim ReplacementIndex  As Long

    Dim HasFailure        As Boolean
    Dim FailNumber        As Long
    Dim FailSource        As String
    Dim FailDescription   As String

    Const PROC As String = "Test_EXCEL_UI_RunSnapshotIdentity"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo Fail

        Debug.Print PROC & " | START | Identity-safe snapshot test started"

        If UI_HasExcelUIStateSnapshot Then
            Err.Raise _
                TEST_ERR_BASE + 1, _
                PROC, _
                "an explicit EXCEL_UI snapshot already exists; clear or restore it before running this destructive test"
        End If

        Set AnchorWindow = ActiveWindow

        If AnchorWindow Is Nothing Then
            Err.Raise _
                TEST_ERR_BASE + 2, _
                PROC, _
                "no active Excel window is available"
        End If

        If Not (AnchorWindow.Parent Is ThisWorkbook) Then
            ThisWorkbook.Activate
            Set AnchorWindow = ActiveWindow
        End If

        SavedHeadings = AnchorWindow.DisplayHeadings
        SavedWorkbookTabs = AnchorWindow.DisplayWorkbookTabs
        SavedGridlines = AnchorWindow.DisplayGridlines

'------------------------------------------------------------------------------
' CREATE CAPTURED WINDOW
'------------------------------------------------------------------------------
        Set CapturedWindow = ThisWorkbook.NewWindow

        CapturedIndex = TST_FindApplicationWindowIndex(CapturedWindow)

        If CapturedIndex = 0 Then
            Err.Raise _
                TEST_ERR_BASE + 3, _
                PROC, _
                "the temporary captured window could not be resolved in Application.Windows"
        End If

'------------------------------------------------------------------------------
' ESTABLISH DISTINCT CAPTURE BASELINES
'------------------------------------------------------------------------------
        AnchorWindow.DisplayHeadings = True
        AnchorWindow.DisplayWorkbookTabs = False
        AnchorWindow.DisplayGridlines = True

        CapturedWindow.DisplayHeadings = False
        CapturedWindow.DisplayWorkbookTabs = True
        CapturedWindow.DisplayGridlines = False

        UI_CaptureExcelUIState

        TST_AssertTrue _
            UI_HasExcelUIStateSnapshot, _
            PROC & ".SnapshotAvailable"

'------------------------------------------------------------------------------
' MUTATE ANCHOR AND REPLACE CAPTURED WINDOW
'------------------------------------------------------------------------------
        AnchorWindow.DisplayHeadings = False
        AnchorWindow.DisplayWorkbookTabs = True
        AnchorWindow.DisplayGridlines = False

        CapturedWindow.Close
        Set CapturedWindow = Nothing

        Set ReplacementWindow = ThisWorkbook.NewWindow

        ReplacementIndex = TST_FindApplicationWindowIndex(ReplacementWindow)

        If ReplacementIndex = 0 Then
            Err.Raise _
                TEST_ERR_BASE + 4, _
                PROC, _
                "the replacement window could not be resolved in Application.Windows"
        End If

        'Sentinel state must remain unchanged because this window did not exist
        'when the snapshot was captured.
        ReplacementWindow.DisplayHeadings = True
        ReplacementWindow.DisplayWorkbookTabs = False
        ReplacementWindow.DisplayGridlines = True

        Debug.Print PROC & " | INFO | Captured index=" & _
            CStr(CapturedIndex) & "; replacement index=" & _
            CStr(ReplacementIndex)

'------------------------------------------------------------------------------
' RESET AND ASSERT IDENTITY-SAFE BEHAVIOR
'------------------------------------------------------------------------------
        UI_ResetExcelUIToSnapshot

        TST_AssertWindowState _
            TargetWindow:=AnchorWindow, _
            ExpectedHeadings:=True, _
            ExpectedWorkbookTabs:=False, _
            ExpectedGridlines:=True, _
            AssertionName:=PROC & ".AnchorRestored"

        TST_AssertWindowState _
            TargetWindow:=ReplacementWindow, _
            ExpectedHeadings:=True, _
            ExpectedWorkbookTabs:=False, _
            ExpectedGridlines:=True, _
            AssertionName:=PROC & ".ReplacementUnchanged"

        Debug.Print PROC & _
            " | PASS | Captured window identity restored without touching replacement"

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        On Error Resume Next

        UI_ClearExcelUIStateSnapshot

        TST_SafeCloseWindow ReplacementWindow
        TST_SafeCloseWindow CapturedWindow

        If Not AnchorWindow Is Nothing Then
            AnchorWindow.DisplayHeadings = SavedHeadings
            AnchorWindow.DisplayWorkbookTabs = SavedWorkbookTabs
            AnchorWindow.DisplayGridlines = SavedGridlines
            AnchorWindow.Activate
        End If

        If HasFailure Then
            Err.Raise FailNumber, FailSource, FailDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
        HasFailure = True
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description

        Resume SafeExit

End Sub


Private Function TST_FindApplicationWindowIndex( _
    ByVal TargetWindow As Window) As Long

'
'==============================================================================
'                    TST_FindApplicationWindowIndex
'------------------------------------------------------------------------------
' PURPOSE
'   Return the current Application.Windows index of an exact Window object.
'
' RETURNS
'   1-based index when found; otherwise zero.
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
    Dim i As Long

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo SafeExit

        If TargetWindow Is Nothing Then
            GoTo SafeExit
        End If

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        For i = 1 To Application.Windows.Count
            If Application.Windows(i) Is TargetWindow Then
                TST_FindApplicationWindowIndex = i
                Exit Function
            End If
        Next i

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
        TST_FindApplicationWindowIndex = 0

End Function


Private Sub TST_AssertWindowState( _
    ByVal TargetWindow As Window, _
    ByVal ExpectedHeadings As Boolean, _
    ByVal ExpectedWorkbookTabs As Boolean, _
    ByVal ExpectedGridlines As Boolean, _
    ByVal AssertionName As String)

'
'==============================================================================
'                       TST_AssertWindowState
'------------------------------------------------------------------------------
' PURPOSE
'   Assert the three managed window-level Boolean properties.
'
' ERROR POLICY
'   Raises on mismatch.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        If TargetWindow Is Nothing Then
            Err.Raise _
                TEST_ERR_BASE + 10, _
                AssertionName, _
                "target window is Nothing"
        End If

        If TargetWindow.DisplayHeadings <> ExpectedHeadings Then
            Err.Raise _
                TEST_ERR_BASE + 11, _
                AssertionName, _
                "DisplayHeadings mismatch"
        End If

        If TargetWindow.DisplayWorkbookTabs <> ExpectedWorkbookTabs Then
            Err.Raise _
                TEST_ERR_BASE + 12, _
                AssertionName, _
                "DisplayWorkbookTabs mismatch"
        End If

        If TargetWindow.DisplayGridlines <> ExpectedGridlines Then
            Err.Raise _
                TEST_ERR_BASE + 13, _
                AssertionName, _
                "DisplayGridlines mismatch"
        End If

End Sub


Private Sub TST_AssertTrue( _
    ByVal ActualValue As Boolean, _
    ByVal AssertionName As String)

'
'==============================================================================
'                             TST_AssertTrue
'------------------------------------------------------------------------------
' PURPOSE
'   Raise when a Boolean assertion is FALSE.
'
' ERROR POLICY
'   Raises on mismatch.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        If Not ActualValue Then
            Err.Raise _
                TEST_ERR_BASE + 20, _
                AssertionName, _
                "expected TRUE but received FALSE"
        End If

End Sub


Private Sub TST_SafeCloseWindow(ByRef TargetWindow As Window)

'
'==============================================================================
'                         TST_SafeCloseWindow
'------------------------------------------------------------------------------
' PURPOSE
'   Close and release one temporary Excel Window on a best-effort basis.
'
' ERROR POLICY
'   Suppresses cleanup errors locally.
'
' UPDATED
'   2026-07-25
'==============================================================================
'

        On Error Resume Next

        If Not TargetWindow Is Nothing Then
            TargetWindow.Close
        End If

        Set TargetWindow = Nothing

End Sub
