VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_UIResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'==============================================================================
'                               CLASS: C_UIResult
'------------------------------------------------------------------------------
' PURPOSE
'   Lightweight result container for best-effort UI operations that may
'   partially succeed while collecting one or more non-fatal failures.
'
' WHY THIS EXISTS
'   Some public UI routines intentionally follow fail-soft behavior:
'     - they continue applying later UI changes even if one element fails
'     - they may need to return a structured outcome to the caller instead of
'       only writing diagnostics to the Immediate Window
'
'   This class provides a minimal, reusable container that tracks:
'     - whether the overall operation remained fully successful
'     - how many element-level failures were recorded
'     - the text of each recorded failure
'
'   The design is intentionally simple so it can be used as an internal helper
'   without adding unnecessary abstraction or dependencies.
'
' PUBLIC SURFACE
'   - Succeeded                    TRUE if no failures were recorded
'   - FailureCount                 Number of recorded failures
'   - Failure(Index)               Retrieve a specific failure entry
'   - AddFailure                   Append a failure entry and mark result failed
'
' BEHAVIOR
'   - New instances initialize as successful.
'   - Each AddFailure call:
'       * marks the result as unsuccessful
'       * appends a formatted failure entry to the internal collection
'   - Failure(Index) uses 1-based indexing, consistent with VBA Collection.
'
' ERROR POLICY
'   - Class_Initialize is straightforward and expected to succeed.
'   - Failure(Index) performs explicit bounds validation and raises a clear
'     runtime error for invalid indexes.
'   - AddFailure does not suppress unexpected errors because invalid internal
'     state should not fail silently in this helper class.
'
' DEPENDENCIES
'   - VBA.Collection
'
' NOTES
'   - This class is intended as a lightweight internal result object, not as a
'     general-purpose logging framework.
'   - Failure text is stored exactly as provided by AddFailure, using the
'     format:
'         Stage & " | " & Detail
'
' UPDATED
'   2026-04-09
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.0.0
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE: PRIVATE STATE
'------------------------------------------------------------------------------
Private m_Succeeded         As Boolean       'TRUE when no failures have been recorded
Private m_Failures          As Collection    'Recorded failure messages

Private Sub Class_Initialize()

'
'==============================================================================
'                            Class_Initialize
'------------------------------------------------------------------------------
' PURPOSE
'   Initialize the result object to its default successful state.
'
' WHY THIS EXISTS
'   A newly created result object should represent "no failures recorded yet"
'   until a caller explicitly adds one.
'
' BEHAVIOR
'   - Sets Succeeded to TRUE.
'   - Creates the internal failure collection.
'
' ERROR POLICY
'   - Does not implement special error handling.
'   - Any unexpected initialization failure should surface normally because it
'     indicates a fundamental object-construction problem.
'
' UPDATED
'   2026-04-09
'==============================================================================
'

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Initialize the object in the default fully successful state.
        m_Succeeded = True

    'Create the internal collection that will store failure messages.
        Set m_Failures = New Collection

End Sub

Public Property Get Succeeded() As Boolean

'
'==============================================================================
'                                Succeeded
'------------------------------------------------------------------------------
' PURPOSE
'   Return whether the overall result is still fully successful.
'
' WHY THIS EXISTS
'   Callers often need a quick Boolean check before deciding whether to inspect
'   detailed failure entries.
'
' RETURNS
'   TRUE  => no failures have been recorded
'   FALSE => at least one failure has been recorded
'
' UPDATED
'   2026-04-09
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN: SUCCESS FLAG
'------------------------------------------------------------------------------
    'Return the current overall success flag.
        Succeeded = m_Succeeded

End Property

Public Sub AddFailure( _
    ByVal Stage As String, _
    ByVal Detail As String)

'
'==============================================================================
'                                AddFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Record a failure entry and mark the overall result as unsuccessful.
'
' WHY THIS EXISTS
'   Best-effort callers need a simple way to accumulate element-level failure
'   diagnostics while continuing execution.
'
' INPUTS
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
'   - Appends a single formatted entry to the internal collection:
'       Stage & " | " & Detail
'
' ERROR POLICY
'   - Does not suppress unexpected errors.
'   - The class should fail loudly if its internal storage is unavailable or
'     invalid.
'
' UPDATED
'   2026-04-09
'==============================================================================
'

'------------------------------------------------------------------------------
' UPDATE: OVERALL SUCCESS STATE
'------------------------------------------------------------------------------
    'Mark the overall result as unsuccessful because at least one failure has
    'now been recorded.
        m_Succeeded = False

'------------------------------------------------------------------------------
' APPEND: FAILURE ENTRY
'------------------------------------------------------------------------------
    'Store the formatted failure entry in insertion order.
        m_Failures.Add Stage & " | " & Detail

End Sub

Public Property Get FailureCount() As Long

'
'==============================================================================
'                               FailureCount
'------------------------------------------------------------------------------
' PURPOSE
'   Return the number of recorded failure entries.
'
' WHY THIS EXISTS
'   Callers may need to iterate the recorded failures or determine whether any
'   detailed diagnostics were captured.
'
' RETURNS
'   Number of entries currently stored in the internal failure collection.
'
' UPDATED
'   2026-04-09
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN: FAILURE COUNT
'------------------------------------------------------------------------------
    'Return the number of stored failure entries.
        FailureCount = m_Failures.Count

End Property

Public Property Get Failure(ByVal Index As Long) As String

'
'==============================================================================
'                                 Failure
'------------------------------------------------------------------------------
' PURPOSE
'   Return a specific recorded failure entry by 1-based index.
'
' WHY THIS EXISTS
'   Callers may need deterministic access to the detailed failure text captured
'   during a best-effort operation.
'
' INPUTS
'   Index
'     1-based failure index, consistent with VBA Collection indexing.
'
' RETURNS
'   The requested failure entry as text.
'
' BEHAVIOR
'   - Valid indexes are 1 through FailureCount.
'   - Raises a clear runtime error when Index is outside the valid range.
'
' ERROR POLICY
'   - Raises error 5 ("Invalid procedure call or argument") with an explicit
'     message when Index is invalid.
'
' UPDATED
'   2026-04-09
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE: INPUT INDEX
'------------------------------------------------------------------------------
    'Reject indexes outside the valid Collection range deterministically.
        If Index < 1 Or Index > m_Failures.Count Then
            Err.Raise 5, "C_UIResult.Failure", _
                "Index is out of range. Valid range is 1 to " & CStr(m_Failures.Count) & "."
        End If

'------------------------------------------------------------------------------
' RETURN: REQUESTED FAILURE ENTRY
'------------------------------------------------------------------------------
    'Return the requested stored failure text.
        Failure = CStr(m_Failures(Index))

End Property

