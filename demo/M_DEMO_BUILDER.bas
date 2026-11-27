Attribute VB_Name = "M_DEMO_BUILDER"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: M_DEMO_BUILDER
'==============================================================================
' PURPOSE
'   Provides reusable helper routines to build, reset, and format DEMO sheets
'   consistently.
'
' WHY THIS EXISTS
'   DEMO / showcase workbooks often require repeated setup logic for:
'     - worksheet creation / retrieval
'     - sheet reset before rebuild
'     - title / subtitle bands
'     - section headers
'     - label and input formatting
'     - output formatting
'     - borders
'     - named ranges
'     - validation lists
'     - helper-list storage
'     - placeholder DEMO buttons
'     - table / log setup
'     - reading and writing Forms and ActiveX check boxes
'
'   Centralizing that logic in one module keeps DEMO-sheet construction:
'     - more consistent
'     - easier to maintain
'     - easier to reuse across projects
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   - Runs nothing at load time; every routine is a procedure that demo code
'     calls, apart from the cleanup probe routines, which only
'     M_DEMO_CLEANUP_TESTS calls
'   - Groups the helpers by job:
'       * sheet lifecycle: DEMO_Sheet_BuildTemplate, DEMO_Sheet_Reset,
'         DEMO_Sheet_GetOrCreate and the helpers they share
'       * fast mode: DEMO_FastMode_Begin, DEMO_FastMode_End and
'         DEMO_FastMode_EndPreservingError
'       * formatting, borders, validation lists, workbook names and tables
'       * shape buttons (DEMO_Btn_*), the status-bar progress message
'         (DEMO_SB_SetProgress) and Forms / ActiveX check boxes (DEMO_CB_*)
'       * the cleanup probe and its one-shot fault suppliers
'   - The only state the code reads or writes is the cleanup probe: the armed
'     flag, the hit counter of each probe site and the armed fault bits
'
' ERROR POLICY
'   - Build, format, validation, table, button and check-box builder helpers
'     raise errors normally, as vbObjectError + 2xxx with the procedure name as
'     the source
'   - DEMO_Sheet_BuildTemplate, DEMO_Sheet_Reset and DEMO_FastMode_Begin run one
'     cleanup pass and then re-raise the original error; a failed cleanup step
'     is appended to its description as a "Cleanup failure: " line
'   - DEMO_FastMode_End raises after trying every restoration; the cleanup
'     helpers (DEMO_FastMode_Restore, DEMO_Sheet_DeleteOwned,
'     DEMO_FastMode_EndPreservingError), DEMO_CB_TryGetState and
'     DEMO_CB_TrySetState do not raise
'   - DEMO_Btn_PlayFeedback suppresses any error from DEMO_Btn_Click
'
' DEPENDENCIES
'   - Excel object model: Application, Workbook, Worksheet, Range, Shape,
'     ListObject, CheckBox and OLEObject
'   - Office shape and state constants (MsoTriState, msoShapeRoundedRectangle
'     and the like)
'   - No other project module and no WinAPI declaration
'   - M_DEMO_CLEANUP_TESTS calls the DEMO_InternalCleanupProbe* routines and
'     mirrors the probe-site, fault-bit and fast-mode-bit constants
'
' NOTES
'   - Option Private Module keeps the public routines out of other VBA projects
'   - The injected-fault error numbers are vbObjectError + 2190 to + 2196; the
'     probe and the fault suppliers do nothing until
'     DEMO_InternalCleanupProbeBegin arms them
'   - mResetWorkbookName, mResetSheetName and mResetShapeName are declared but
'     no code uses them
'
' DESIGN
'   - This module is workbook-oriented and defaults to ThisWorkbook for build
'     operations unless a workbook is explicitly passed by the caller.
'   - Formatting helpers are kept generic enough to be reusable outside one
'     specific DEMO sheet.
'   - Button creation supports an optional action macro so the helper is not
'     hard-wired to one single placeholder routine.
'   - Fast-mode helpers centralize temporary Application-state changes used
'     during sheet-building operations. They own five settings only:
'     ScreenUpdating, EnableEvents, DisplayAlerts, Calculation and Cursor.
'   - DEMO_Sheet_BuildTemplate and DEMO_Sheet_Reset carry a cleanup probe and
'     one-shot fault seams, and so do the fast-mode writes. They are inert
'     until M_DEMO_CLEANUP_TESTS arms them through
'     DEMO_InternalCleanupProbeBegin.
'
' AUTHOR
'   Daniele Penza
'
' VERSION
'   1.1.0
'
' HISTORY
'   2026-11-15 - DEMO_Sheet_BuildTemplate keeps manual page breaks, and both
'                procedures write DisplayPageBreaks only when it is True, so
'                a demo build does not wait for the printer driver (#187).
'   2026-11-14 - DEMO_Sheet_BuildTemplate validates its layout before any
'                reset or creation; a sheet created by a failed call is
'                removed (#74).
'   2026-11-14 - Check-box state helpers DEMO_CB_TryGetState and
'                DEMO_CB_TrySetState, dispatching on the control type (#73).
'   2026-11-14 - Fast-mode rollback, entry cursor and reported restoration
'                failures (#70).
'   2026-11-06 - Cleanup probe and one-shot fault seams for the #71 harness.
'   2026-09-10 - Single-pass demo cleanup and owned protection restoration.
'
' UPDATED
'   2026-11-15
'==============================================================================

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    Public Const COLOR_TITLE        As Long = 0 'RGB(0, 0, 0)
    Public Const COLOR_SUBTITLE     As Long = 6568980
    Public Const COLOR_HEADER       As Long = 6568980
    Public Const COLOR_SUBHEADER    As Long = 13145700
    Public Const COLOR_INPUT        As Long = 13172735
    Public Const COLOR_LABEL_BACK   As Long = 15787740
    Public Const COLOR_LABEL_FONT   As Long = 0
    'Public Const COLOR_BUTTON       As Long = &HFF660A
    Public Const COLOR_BUTTON   As Long = 13145700
    Public Const COLOR_TABLE    As Long = 16774635
    Public Const COLOR_HELP     As Long = 6697728
    'Buttons color tokens
    Public Const BTN_PRIMARY_N  As Long = &HFF660A  ' #0A66FF (BGR in VBA)
    Public Const BTN_PRIMARY_H  As Long = &HB34700  ' #0047B3
    Public Const BTN_PRIMARY_P  As Long = &H853000  ' #003085
    Public Const WHITE          As Long = &HFFFFFF

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const DEFAULT_BOOL_WS       As String = "DEMO"
    Private Const DEFAULT_BOOLLIST_NAME As String = "BoolList"

    'Cleanup probe sites; mirrored by M_DEMO_CLEANUP_TESTS
    Private Const DEMO_PROBE_RESET_HANDLER      As Long = 1     'Reset Clean_Fail entries
    Private Const DEMO_PROBE_RESET_CLEANUP      As Long = 2     'Reset Clean_Exit passes
    Private Const DEMO_PROBE_RESET_PROTECT      As Long = 3     'Reset Protect statements reached
    Private Const DEMO_PROBE_RESET_SELECTION    As Long = 4     'Reset EnableSelection restores reached
    Private Const DEMO_PROBE_BUILD_HANDLER      As Long = 5     'BuildTemplate Clean_Fail entries
    Private Const DEMO_PROBE_BUILD_CLEANUP      As Long = 6     'BuildTemplate Clean_Exit passes
    Private Const DEMO_PROBE_FM_ENTRY           As Long = 7     'Fast-mode entry writes reached
    Private Const DEMO_PROBE_FM_RESTORE         As Long = 8     'Fast-mode restoration writes reached
    Private Const DEMO_PROBE_PAGE_BREAKS        As Long = 9     'ResetAllPageBreaks calls reached
    Private Const DEMO_PROBE_DISPLAY_BREAKS     As Long = 10    'DisplayPageBreaks writes reached
    Private Const DEMO_PROBE_SITE_COUNT         As Long = 10    'Highest probe site

    'One-shot cleanup fault bits; mirrored by M_DEMO_CLEANUP_TESTS
    Private Const DEMO_FAULT_RESET_BODY         As Long = 1     'Operation error after unprotect
    Private Const DEMO_FAULT_RESET_PROTECT      As Long = 2     'Reprotect statement fails
    Private Const DEMO_FAULT_RESET_SELECTION    As Long = 4     'EnableSelection restore fails
    Private Const DEMO_FAULT_BUILD_FORMULABAR   As Long = 8     'BuildTemplate DisplayFormulaBar write fails
    Private Const DEMO_FAULT_SHEET_NAME         As Long = 16    'GetOrCreate naming of a new sheet fails
    Private Const DEMO_FAULT_FM_ENTRY_UNIT      As Long = 256   'Times a DEMO_FM_* bit: its entry write fails
    Private Const DEMO_FAULT_FM_RESTORE_UNIT    As Long = 8192  'Times a DEMO_FM_* bit: its restoration fails

    'Fast-mode settings, one bit each; mirrored by M_DEMO_CLEANUP_TESTS
    Private Const DEMO_FM_SCREENUPDATING    As Long = 1     'Application.ScreenUpdating
    Private Const DEMO_FM_ENABLEEVENTS      As Long = 2     'Application.EnableEvents
    Private Const DEMO_FM_DISPLAYALERTS     As Long = 4     'Application.DisplayAlerts
    Private Const DEMO_FM_CALCULATION       As Long = 8     'Application.Calculation
    Private Const DEMO_FM_CURSOR            As Long = 16    'Application.Cursor

    'Check-box kinds resolved by DEMO_CB_ResolveControl
    Private Const DEMO_CB_KIND_NONE     As Long = 0 'Missing or refused
    Private Const DEMO_CB_KIND_FORMS    As Long = 1 'Forms check box
    Private Const DEMO_CB_KIND_ACTIVEX  As Long = 2 'ActiveX check box
    Private Const DEMO_CB_MIXED_STATE   As String = "check box state is mixed: neither checked nor unchecked"

'------------------------------------------------------------------------------
' PRIVATE VARIABLES
'------------------------------------------------------------------------------
    Private mResetWorkbookName  As String
    Private mResetSheetName     As String
    Private mResetShapeName     As String

    Private mCleanupProbeArmed                              As Boolean  'Count probe sites
    Private mCleanupProbeCounts(1 To DEMO_PROBE_SITE_COUNT) As Long     'Hits per probe site
    Private mCleanupFaults                                  As Long     'Armed one-shot fault bits

'------------------------------------------------------------------------------
' PUBLIC TYPES
'------------------------------------------------------------------------------
    Public Type tDEMOFastModeState
        ScreenUpdating  As Boolean       'Saved Application.ScreenUpdating
        EnableEvents    As Boolean       'Saved Application.EnableEvents
        DisplayAlerts   As Boolean       'Saved Application.DisplayAlerts
        Calculation     As XlCalculation 'Saved Application.Calculation
        Cursor          As XlMousePointer 'Saved Application.Cursor
        Captured        As Boolean       'TRUE once all five entry values were read
        Applied         As Long          'DEMO_FM_* bits written and not yet restored
    End Type

'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
    Private Type tDEMOProtectionState
        DrawingObjects As Boolean
        Contents As Boolean
        Scenarios As Boolean
        UserInterfaceOnly As Boolean
        AllowFormattingCells As Boolean
        AllowFormattingColumns As Boolean
        AllowFormattingRows As Boolean
        AllowInsertingColumns As Boolean
        AllowInsertingRows As Boolean
        AllowInsertingHyperlinks As Boolean
        AllowDeletingColumns As Boolean
        AllowDeletingRows As Boolean
        AllowSorting As Boolean
        AllowFiltering As Boolean
        AllowUsingPivotTables As Boolean
        EnableSelection As Long
    End Type

    Private Type tButtonAppearance
        FillVisible         As MsoTriState      'Original fill visibility
        FillColor           As Long             'Original fill color

        LineVisible         As MsoTriState      'Original line visibility
        LineColor           As Long             'Original line color
        LineWeight          As Single           'Original line weight

        TextColor           As Long             'Original text color
        TextBold            As MsoTriState      'Original text bold flag
        TextSize            As Single           'Original text size

        ShadowVisible       As MsoTriState      'Original shadow visibility
        ShadowBlur          As Single           'Original shadow blur
        ShadowOffsetX       As Single           'Original shadow X offset
        ShadowOffsetY       As Single           'Original shadow Y offset

        TopPos              As Double           'Original top position
        LeftPos             As Double           'Original left position
    End Type

'------------------------------------------------------------------------------
' ENUMS
'------------------------------------------------------------------------------
    Public Enum DEMOInputValidationKind
        DEMOInputValidationNone = 0
        DEMOInputValidationList = 1
        DEMOInputValidationNumeric = 2
        DEMOInputValidationBoolean = 3
    End Enum


Public Sub DEMO_Sheet_BuildTemplate( _
    ByVal WS_Name As String, Optional ByVal Title As String = "Title", _
    Optional ByVal SubTitle As String = "SubTitle", Optional ByVal IsFrozenPane As Boolean = True, _
    Optional ByVal FreezeAtRow As Long = 3, Optional ByVal TargetWorkbook As Variant, _
    Optional ByVal LeftMarginColumns As String = "A:B", Optional ByVal LeftMarginColumnWidth As Double = 2, _
    Optional ByVal ContentColumns As String = "C:M", Optional ByVal ContentColumnWidth As Double = 20, _
    Optional ByVal SeparatorColumns As String = "N:N", Optional ByVal SeparatorColumnWidth As Double = 2, _
    Optional ByVal TitleRowHeight As Double = 24, Optional ByVal SubTitleRowHeight As Double = 28, _
    Optional ByVal BodyRowFrom As Long = 3, Optional ByVal BodyRowTo As Long = 1000, _
    Optional ByVal BodyRowHeight As Double = 20, Optional ByVal HideColumnsFrom As String = "XFD", _
    Optional ByVal HideRowsFrom As Long = 1048576, Optional ByVal RestrictScrollAreaToVisible As Boolean = False, _
    Optional ByVal HideGridlines As Boolean = True, Optional ByVal ShowHeadings As Boolean = True, _
    Optional ByVal ShowFormulaBar As Boolean = True, Optional ByVal ShowHorizontalScrollBar As Boolean = True, _
    Optional ByVal ShowVerticalScrollBar As Boolean = True, Optional ByVal ZoomPercent As Long = 0)
'
'------------------------------------------------------------------------------
'                             BUILD DEMO TEMPLATE
'------------------------------------------------------------------------------
' PURPOSE
'   Creates or rebuilds a DEMO worksheet with a standard title/subtitle layout
'   and configurable base formatting
'
' WHY THIS EXISTS
'   DEMO sheets should start from a consistent visual and structural baseline.
'   This routine centralizes worksheet retrieval, reset, title/subtitle band
'   creation, base formatting, configurable sizing, and view normalization
'
' INPUTS
'   WS_Name
'     Name of the worksheet to create or rebuild
'
'   Title / SubTitle (optional)
'     Main title and subtitle for the header bands
'
'   IsFrozenPane / FreezeAtRow (optional)
'     Freeze-pane behavior
'
'   TargetWorkbook (optional)
'     Workbook to build into
'     When omitted, defaults to ThisWorkbook
'
'   LeftMarginColumns / LeftMarginColumnWidth (optional)
'     Left-margin column block and width
'
'   ContentColumns / ContentColumnWidth (optional)
'     Main content block and width
'
'   SeparatorColumns (optional)
'     Retained for backward compatibility
'     The effective separator column is derived automatically as the next
'     column after ContentColumns
'
'   SeparatorColumnWidth (optional)
'     Width applied to the derived separator column
'
'   TitleRowHeight / SubTitleRowHeight (optional)
'     Heights for rows 1 and 2
'
'   BodyRowFrom / BodyRowTo / BodyRowHeight (optional)
'     Body-row range and applied height
'
'   HideColumnsFrom (optional)
'     Retained for backward compatibility
'     The effective first hidden column is derived automatically as the column
'     after the derived separator column
'
'   HideRowsFrom (optional)
'     First row to hide through the final worksheet row
'
'   RestrictScrollAreaToVisible (optional)
'     TRUE  => restrict worksheet scroll area to the visible block
'     FALSE => leave worksheet scroll area unrestricted
'
'   HideGridlines / ShowHeadings / ShowFormulaBar / ShowHorizontalScrollBar /
'   ShowVerticalScrollBar / ZoomPercent (optional)
'     View settings applied only when the target workbook is active
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Validates every layout input first: the worksheet name, the content and
'     left-margin columns, the body, freeze and hidden rows, widths, heights
'     and zoom. Checks that need the worksheet grid read its limits from the
'     target or another worksheet of the workbook. An invalid request raises
'     before any worksheet is created or reset.
'   - Uses the requested workbook or defaults to ThisWorkbook
'   - Refuses a name already used by a chart or other non-worksheet sheet
'   - Gets or creates the requested worksheet
'   - Resets the sheet before rebuilding it, keeping its manual page breaks
'   - Derives the separator column from ContentColumns
'   - Derives the first hidden column from the separator column
'   - Writes title/subtitle bands from column B to the last content column
'   - Applies configurable row/column sizing and view normalization
'
' ERROR POLICY
'   Raises errors normally. When the call created the worksheet and then
'   fails, it deletes that worksheet before re-raising; a failed deletion is
'   appended to the description as a "Cleanup failure: " line.
'
' DEPENDENCIES
'   - DEMO_Sheet_FindWorksheet
'   - DEMO_Sheet_GetOrCreate
'   - DEMO_Sheet_Reset
'   - DEMO_Sheet_HidePageBreaks
'   - DEMO_Sheet_DeleteOwned
'   - DEMO_SheetNameProblem
'   - DEMO_ParseColumnSpec
'   - DEMO_Write_BandHeader
'   - DEMO_ColumnLetter
'   - DEMO_CleanupProbeHit
'
' NOTES
'   - Window-level settings can only be applied safely through the active window
'   - SeparatorColumns and HideColumnsFrom are retained in the signature only
'     for backward compatibility
'   - The cleanup probe counts handler entries and cleanup passes only while
'     M_DEMO_CLEANUP_TESTS has armed it; an armed one-shot fault can fail the
'     DisplayFormulaBar write
'
' DESTRUCTIVE SCOPE
'   A rebuild of an existing worksheet is destructive by design: once the
'   inputs are valid, DEMO_Sheet_Reset deletes its shapes and tables and
'   clears its cells, formats, validation, comments, sizes, hidden rows and
'   columns and tab colour. A failure after that point is reported, but
'   nothing is rolled back; only a worksheet this call created is removed.
'
' PAGE BREAKS
'   The template adds no manual page breaks, so it does not remove any: it
'   resets the sheet with RemovePageBreaks:=False and keeps manual breaks
'   that were already there. ResetAllPageBreaks makes Excel lay the sheet out
'   through the default printer driver, which stalls the build on a slow,
'   offline or network printer. Checking for manual breaks first would not
'   avoid it: HPageBreaks, VPageBreaks and Range.PageBreak also report
'   automatic breaks, which Excel computes from that layout. Page-break
'   indicators are hidden only when they are shown (DEMO_Sheet_HidePageBreaks).
'
' VIEW OWNERSHIP
'   The window settings and DisplayFormulaBar are applied as the requested
'   demo view and are not restored afterwards. DisplayFormulaBar is
'   application-wide: it changes every window of this Excel instance. A
'   DEMO_FastMode_Begin scope around this call does not capture or restore it.
'
' HISTORY
'   2026-11-15 - Manual page breaks kept; indicators hidden only when shown
'                (#187).
'   2026-11-14 - Layout validated before any reset or creation; a worksheet
'                created by a failed call is removed (#74).
'   2026-11-14 - Formula-bar ownership documented; fault seam (#70).
'   2026-11-06
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate"

    Dim WB              As Workbook     'Target workbook
    Dim WS              As Worksheet    'Target worksheet
    Dim LimitSheet      As Worksheet    'Worksheet the grid limits are read from
    Dim CreatedSheet    As Boolean      'TRUE when this call created WS
    Dim NameProblem     As String       'Why WS_Name is not a valid sheet name
    Dim CleanupFailures As String       'Failed removal of a created worksheet

    Dim MaxRow  As Long 'Last worksheet row
    Dim MaxCol  As Long 'Last worksheet column

    Dim ContentText             As String   'Trimmed ContentColumns text
    Dim ContentFirstColIndex    As Long     'Resolved first content-column index
    Dim ContentLastColIndex     As Long     'Resolved last content-column index
    Dim LeftFirstColIndex       As Long     'Resolved first left-margin column, 0 when blank
    Dim LeftLastColIndex        As Long     'Resolved last left-margin column, 0 when blank

    Dim EffectiveSeparatorColIndex  As Long     'Derived separator-column index
    Dim EffectiveHideFromColIndex   As Long     'Derived first hidden-column index
    Dim EffectiveSeparatorColumns   As String   'Derived separator-column address
    Dim EffectiveHideColumnsFrom    As String   'Derived first hidden-column text

    Dim LastVisibleRow      As Long     'Last visible worksheet row
    Dim LastVisibleCol      As Long     'Last visible worksheet column
    Dim ScrollAreaAddress   As String   'Resolved worksheet ScrollArea text

    Dim TitleBandRange      As Range    'Resolved title-band range
    Dim SubTitleBandRange   As Range    'Resolved subtitle-band range

    Dim SavedErrNum     As Long     'Captured error number
    Dim SavedErrSrc     As String   'Captured error source
    Dim SavedErrDesc    As String   'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Every check that needs no workbook runs here, before anything is created
    'or reset, so an invalid request leaves the workbook as it was

    'Reject an empty worksheet name
        If Len(Trim$(WS_Name)) = 0 Then
            Err.Raise vbObjectError + 2000, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "Worksheet name cannot be blank."
        End If

    'Reject a name Excel would refuse for a sheet
        NameProblem = DEMO_SheetNameProblem(WS_Name)
        If Len(NameProblem) > 0 Then
            Err.Raise vbObjectError + 2027, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "Worksheet name is not valid: " & NameProblem
        End If

    'Reject invalid freeze-row requests when pane freezing is enabled
        If IsFrozenPane Then
            If FreezeAtRow < 2 Then
                Err.Raise vbObjectError + 2013, _
                          "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                          "FreezeAtRow must be >= 2 when IsFrozenPane = True."
            End If
        End If

    'Reject invalid body-row ranges
        If BodyRowFrom < 1 Or BodyRowTo < BodyRowFrom Then
            Err.Raise vbObjectError + 2014, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "Invalid body-row range."
        End If

    'Reject invalid width settings
        If LeftMarginColumnWidth <= 0# Or _
           ContentColumnWidth <= 0# Or _
           SeparatorColumnWidth <= 0# Then
            Err.Raise vbObjectError + 2015, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "Column widths must be positive."
        End If

    'Reject invalid row-height settings
        If TitleRowHeight <= 0# Or _
           SubTitleRowHeight <= 0# Or _
           BodyRowHeight <= 0# Then
            Err.Raise vbObjectError + 2016, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "Row heights must be positive."
        End If

    'Reject invalid zoom values when explicitly provided
        If ZoomPercent <> 0 Then
            If ZoomPercent < 10 Or ZoomPercent > 400 Then
                Err.Raise vbObjectError + 2017, _
                          "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                          "ZoomPercent must be 0 or between 10 and 400."
            End If
        End If

    'Reject a blank content-column specification
        ContentText = Trim$(ContentColumns)
        If Len(ContentText) = 0 Then
            Err.Raise vbObjectError + 2021, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "ContentColumns cannot be blank."
        End If

    'Reject content columns that are not one column or one column range;
    'reversed ranges come back in ascending order
        If Not DEMO_ParseColumnSpec(ContentText, ContentFirstColIndex, ContentLastColIndex) Then
            Err.Raise vbObjectError + 2022, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "ContentColumns must resolve to one column or one contiguous column range."
        End If

    'Reject content-column blocks that end before column B
        If ContentLastColIndex < 2 Then
            Err.Raise vbObjectError + 2024, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "ContentColumns must extend at least through column B."
        End If

    'Reject left-margin columns that are neither blank nor one column range
        If Len(Trim$(LeftMarginColumns)) > 0 Then
            If Not DEMO_ParseColumnSpec(LeftMarginColumns, LeftFirstColIndex, LeftLastColIndex) Then
                Err.Raise vbObjectError + 2020, _
                          "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                          "LeftMarginColumns must be blank, one column or one contiguous column range."
            End If
        End If

    'Reject row-hide requests before the first row
        If HideRowsFrom < 1 Then
            Err.Raise vbObjectError + 2019, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "HideRowsFrom must be between 1 and the worksheet last row."
        End If

    'Enable structured cleanup on failure
        On Error GoTo CleanFail

    'Use the requested workbook when supplied; otherwise default to ThisWorkbook
        If IsMissing(TargetWorkbook) Or IsEmpty(TargetWorkbook) Then
            Set WB = ThisWorkbook
        ElseIf IsObject(TargetWorkbook) Then
            Set WB = TargetWorkbook

            If WB Is Nothing Then
                Set WB = ThisWorkbook
            End If
        Else
            Err.Raise vbObjectError + 2018, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "TargetWorkbook must be a Workbook object when supplied."
        End If

'------------------------------------------------------------------------------
' LOCATE TARGET AND WORKSHEET LIMITS
'------------------------------------------------------------------------------
    'Find an existing worksheet; a chart or other sheet with the name is refused
        Set WS = DEMO_Sheet_FindWorksheet(WB, WS_Name, "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate")

    'Every worksheet of a workbook has the same grid; read it from the target,
    'else from another worksheet. Only a workbook without any worksheet makes
    'this call create the target first; cleanup removes it if a check fails.
        If Not WS Is Nothing Then
            Set LimitSheet = WS
        ElseIf WB.Worksheets.Count > 0 Then
            Set LimitSheet = WB.Worksheets(1)
        Else
            Set WS = DEMO_Sheet_GetOrCreate(WB, WS_Name)
            CreatedSheet = True
            Set LimitSheet = WS
        End If

    'Capture the last available worksheet row and column
        MaxRow = LimitSheet.Rows.Count
        MaxCol = LimitSheet.Columns.Count

'------------------------------------------------------------------------------
' VALIDATE AGAINST WORKSHEET LIMITS
'------------------------------------------------------------------------------
    'Reject content-column blocks that do not leave room for separator + hidden block
        If ContentLastColIndex > MaxCol - 2 Then
            Err.Raise vbObjectError + 2023, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "ContentColumns must leave room for a separator column and at least one hidden trailing column."
        End If

    'Reject left-margin columns beyond the last worksheet column
        If LeftLastColIndex > MaxCol Then
            Err.Raise vbObjectError + 2020, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "LeftMarginColumns must be blank, one column or one contiguous column range."
        End If

    'Reject body rows beyond the last worksheet row
        If BodyRowTo > MaxRow Then
            Err.Raise vbObjectError + 2014, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "Invalid body-row range."
        End If

    'Reject a freeze row beyond the last worksheet row
        If IsFrozenPane Then
            If FreezeAtRow > MaxRow Then
                Err.Raise vbObjectError + 2013, _
                          "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                          "FreezeAtRow must not exceed the worksheet last row."
            End If
        End If

    'Reject row-hide requests beyond the last worksheet row
        If HideRowsFrom > MaxRow Then
            Err.Raise vbObjectError + 2019, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate", _
                      "HideRowsFrom must be between 1 and the worksheet last row."
        End If

'------------------------------------------------------------------------------
' CREATE OR RESET SHEET
'------------------------------------------------------------------------------
    'Create the worksheet when it does not exist yet; this call then owns it
        If WS Is Nothing Then
            Set WS = DEMO_Sheet_GetOrCreate(WB, WS_Name)
            CreatedSheet = True
        End If

    'Reset the worksheet to a clean reusable state before rebuilding it; from
    'here on an existing worksheet is rebuilt destructively. Manual page
    'breaks are kept: removing them waits for the printer driver (#187)
        DEMO_Sheet_Reset WS, RemovePageBreaks:=False

    'Hide page-break indicators, writing the property only when they are shown
        DEMO_Sheet_HidePageBreaks WS

'------------------------------------------------------------------------------
' DERIVE SEPARATOR / HIDDEN COLUMN LOGIC
'------------------------------------------------------------------------------
    'Derive the separator column as the next column after the content block
        EffectiveSeparatorColIndex = ContentLastColIndex + 1

    'Derive the first hidden column as the column after the separator
        EffectiveHideFromColIndex = EffectiveSeparatorColIndex + 1

    'Resolve the effective separator-column address
        EffectiveSeparatorColumns = DEMO_ColumnLetter(EffectiveSeparatorColIndex) & ":" & _
                                    DEMO_ColumnLetter(EffectiveSeparatorColIndex)

    'Resolve the effective first hidden-column text
        EffectiveHideColumnsFrom = DEMO_ColumnLetter(EffectiveHideFromColIndex)

'------------------------------------------------------------------------------
' APPLY GLOBAL SHEET FORMATTING
'------------------------------------------------------------------------------
    'Apply the standard tab color
        WS.Tab.Color = COLOR_SUBTITLE

    'Apply base cell formatting
        With WS.Cells
            .Interior.Pattern = xlNone
            .Font.Name = "Aptos"
            .Font.Size = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
        End With

'------------------------------------------------------------------------------
' APPLY COLUMN / ROW LAYOUT
'------------------------------------------------------------------------------
    'Apply the configured left-margin column widths
        If LeftFirstColIndex > 0 Then
            WS.Range(WS.Columns(LeftFirstColIndex), WS.Columns(LeftLastColIndex)).ColumnWidth = _
                LeftMarginColumnWidth
        End If

    'Apply the configured main content-block widths
        WS.Range(WS.Columns(ContentFirstColIndex), WS.Columns(ContentLastColIndex)).ColumnWidth = _
            ContentColumnWidth

    'Apply the derived separator-column width
        WS.Columns(EffectiveSeparatorColumns).ColumnWidth = SeparatorColumnWidth

    'Apply the configured title-row height
        WS.Rows(1).RowHeight = TitleRowHeight

    'Apply the configured subtitle-row height
        WS.Rows(2).RowHeight = SubTitleRowHeight

    'Apply the configured body-row height range
        WS.Rows(CStr(BodyRowFrom) & ":" & CStr(BodyRowTo)).RowHeight = BodyRowHeight

'------------------------------------------------------------------------------
' RESOLVE TITLE / SUBTITLE BANDS
'------------------------------------------------------------------------------
    'Resolve the title-band range from column B to the last content column
        Set TitleBandRange = WS.Range("B1:" & DEMO_ColumnLetter(ContentLastColIndex) & "1")

    'Resolve the subtitle-band range from column B to the last content column
        Set SubTitleBandRange = WS.Range("B2:" & DEMO_ColumnLetter(ContentLastColIndex) & "2")

'------------------------------------------------------------------------------
' FORMAT TITLE / SUBTITLE BANDS
'------------------------------------------------------------------------------
    'Write and format the title band
        DEMO_Write_BandHeader _
            TargetRange:=TitleBandRange, _
            HeaderText:=Title, _
            FillColor:=COLOR_TITLE, _
            FontColor:=RGB(255, 192, 0), _
            FontSize:=14, _
            IsBold:=True, _
            IsBorder:=True, _
            IsCentered:=False

    'Write and format the subtitle band
        DEMO_Write_BandHeader _
            TargetRange:=SubTitleBandRange, _
            HeaderText:=SubTitle, _
            FillColor:=COLOR_SUBTITLE, _
            FontColor:=RGB(255, 255, 255), _
            FontSize:=16, _
            IsBold:=True, _
            IsBorder:=True, _
            IsCentered:=False

'------------------------------------------------------------------------------
' HIDE TRAILING ROWS / COLUMNS
'------------------------------------------------------------------------------
    'Hide all rows from HideRowsFrom to the final worksheet row when requested
        If HideRowsFrom < MaxRow Then
            WS.Rows(CStr(HideRowsFrom) & ":" & CStr(MaxRow)).Hidden = True
        End If

    'Hide all columns from the derived first hidden column to the final worksheet column
        If EffectiveHideFromColIndex <= MaxCol Then
            WS.Range(WS.Columns(EffectiveHideFromColIndex), WS.Columns(MaxCol)).EntireColumn.Hidden = True
        End If

'------------------------------------------------------------------------------
' APPLY SCROLL AREA
'------------------------------------------------------------------------------
    'Compute the last visible worksheet row
        If HideRowsFrom < MaxRow Then
            LastVisibleRow = HideRowsFrom - 1
        Else
            LastVisibleRow = MaxRow
        End If

    'Compute the last visible worksheet column
        LastVisibleCol = EffectiveHideFromColIndex - 1

    'Apply or clear the worksheet scroll area
        If RestrictScrollAreaToVisible Then
            ScrollAreaAddress = WS.Range(WS.Cells(1, 1), WS.Cells(LastVisibleRow, LastVisibleCol)).Address
            WS.ScrollArea = ScrollAreaAddress
        Else
            WS.ScrollArea = vbNullString
        End If

'------------------------------------------------------------------------------
' APPLY VIEW SETTINGS
'------------------------------------------------------------------------------
    'Apply window-level view settings only when the target workbook is active
        If WB Is ActiveWorkbook Then

        'Activate the target sheet in the current window
            WS.Activate

        'Apply active-window and application-level view settings
            With ActiveWindow
                .DisplayGridlines = Not HideGridlines
                .DisplayHeadings = ShowHeadings
                .DisplayHorizontalScrollBar = ShowHorizontalScrollBar
                .DisplayVerticalScrollBar = ShowVerticalScrollBar

                If ZoomPercent <> 0 Then
                    .Zoom = ZoomPercent
                End If

                .FreezePanes = False
                .SplitColumn = 0

                If IsFrozenPane Then
                    .SplitRow = FreezeAtRow - 1
                    WS.Range("A" & CStr(FreezeAtRow)).Select
                    .FreezePanes = True
                Else
                    .SplitRow = 0
                End If
            End With

        'Apply formula-bar visibility, an intended demo view change that no
        'fast-mode scope captures or restores
            Application.DisplayFormulaBar = DEMO_FormulaBarValue(ShowFormulaBar)
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanExit:
    'Re-raise the captured error after cleanup when needed
        On Error GoTo 0
        DEMO_CleanupProbeHit DEMO_PROBE_BUILD_CLEANUP

    'Remove a worksheet this call created before it failed; keep the error
        If SavedErrNum <> 0 And CreatedSheet Then
            CleanupFailures = DEMO_Sheet_DeleteOwned(WS)
            If Len(CleanupFailures) > 0 Then
                SavedErrDesc = SavedErrDesc & vbCrLf & "Cleanup failure: " & CleanupFailures
            End If
        End If

        If SavedErrNum <> 0 Then
            Err.Raise SavedErrNum, SavedErrSrc, SavedErrDesc
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CleanFail:
    'Capture the original error before cleanup
        SavedErrNum = Err.Number
        SavedErrSrc = Err.Source
        SavedErrDesc = Err.Description

    'Count the handler entry for the regression probe
        DEMO_CleanupProbeHit DEMO_PROBE_BUILD_HANDLER

    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub

Private Function DEMO_ColumnLetter( _
    ByVal ColIndex As Long) _
    As String
'
'------------------------------------------------------------------------------
'                        CONVERT INDEX TO COLUMN LETTER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Excel column letter for a 1-based column index
'
' WHY THIS EXISTS
'   Some layout logic derives column positions numerically and then needs a
'   stable A1-style column label
'
' INPUTS
'   ColIndex
'     1-based worksheet column index
'
' RETURNS
'   String
'     Excel column letter text
'
' BEHAVIOR
'   - Raises vbObjectError + 2025 when ColIndex is less than 1
'   - Converts the index by repeated base-26 division, prepending one letter
'     per step (1 is A, 26 is Z, 27 is AA)
'
' ERROR POLICY
'   Raises when ColIndex is less than 1
'
' DEPENDENCIES
'   - VBA Chr$ only
'
' NOTES
'   - The index is not compared with the last worksheet column; the caller keeps
'     it inside the grid
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_ColumnLetter"

    Dim N   As Long     'Working column index
    Dim R   As Long     'Current base-26 remainder
    Dim S   As String   'Accumulated column text

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject invalid column indices
        If ColIndex < 1 Then
            Err.Raise vbObjectError + 2025, _
                      "M_DEMO_BUILDER.DEMO_ColumnLetter", _
                      "Column index must be >= 1."
        End If

'------------------------------------------------------------------------------
' CONVERT INDEX TO LETTER
'------------------------------------------------------------------------------
    'Initialize the working index
        N = ColIndex

    'Convert the numeric column index to Excel letter notation
        Do While N > 0
            R = (N - 1) Mod 26
            S = Chr$(65 + R) & S
            N = (N - 1) \ 26
        Loop

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the resolved column letter
        DEMO_ColumnLetter = S

End Function

Public Sub DEMO_FastMode_Begin( _
    ByRef StateOut As tDEMOFastModeState)
'
'------------------------------------------------------------------------------
'                               BEGIN FAST MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Captures the current Excel Application state and switches Excel into a
'   reduced-noise fast mode suitable for sheet-building operations.
'
' WHY THIS EXISTS
'   DEMO-sheet builders often repeat the same Application-state pattern:
'     - save current state
'     - reduce UI noise
'     - restore later
'
'   Centralizing that pattern reduces duplication and lowers the risk of
'   incomplete cleanup.
'
' INPUTS
'   StateOut
'     Output structure that receives the saved Application state and the
'     settings this scope now owns.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Clears StateOut.Captured and StateOut.Applied first
'   - Captures all five entry values before the first write:
'       * ScreenUpdating
'       * EnableEvents
'       * DisplayAlerts
'       * Calculation
'       * Cursor
'   - Sets StateOut.Captured once all five are read
'   - Marks each setting in StateOut.Applied before writing it, so a write
'     that fails partway is still restored
'   - Applies ScreenUpdating False, EnableEvents False, DisplayAlerts False,
'     manual calculation and the wait cursor
'
' ERROR POLICY
'   - A failed capture raises before anything is written
'   - A failed write rolls back every setting already marked, each attempted
'     independently, then raises the original error with its number and
'     source unchanged. A failed rollback appends a line that starts with
'     "Cleanup failure: Fast-mode rollback failed: " to the description.
'   - After a failed Begin, StateOut.Applied is 0, so DEMO_FastMode_End has
'     nothing left to restore
'
' DEPENDENCIES
'   - Application.ScreenUpdating, EnableEvents, DisplayAlerts, Calculation and
'     Cursor
'   - DEMO_FastModeValue
'   - DEMO_FastMode_Restore
'
' NOTES
'   - Callers end the scope with DEMO_FastMode_End or
'     DEMO_FastMode_EndPreservingError on their cleanup path
'
' OWNERSHIP
'   Fast mode owns only these five temporary settings. The window and
'   Application view settings that DEMO_Sheet_BuildTemplate applies, such as
'   DisplayFormulaBar, are intended demo view changes; fast mode neither
'   captures nor restores them.
'
' REGRESSION SEAMS
'   Every write passes its value through DEMO_FastModeValue, which counts the
'   write and fails it once while M_DEMO_CLEANUP_TESTS has armed a fault for
'   that setting. Disarmed, the value passes through unchanged.
'
' HISTORY
'   2026-11-14 - Rollback of a partial entry and the entry cursor (#70).
'   2026-04-19
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_FastMode_Begin"

    Dim RollbackFailures    As String   'Settings the rollback could not restore
    Dim SavedErrNumber      As Long     'Captured error number
    Dim SavedErrSource      As String   'Captured error source
    Dim SavedErrDescription As String   'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Own nothing until every entry value is known
        StateOut.Captured = False
        StateOut.Applied = 0

    'Capture the current Excel Application state; a failed read writes nothing
        StateOut.ScreenUpdating = Application.ScreenUpdating
        StateOut.EnableEvents = Application.EnableEvents
        StateOut.DisplayAlerts = Application.DisplayAlerts
        StateOut.Calculation = Application.Calculation
        StateOut.Cursor = Application.Cursor
        StateOut.Captured = True

    'Roll back a partial entry on failure
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' APPLY FAST MODE
'------------------------------------------------------------------------------
    'Mark each setting before writing it, then reduce UI noise
        StateOut.Applied = StateOut.Applied Or DEMO_FM_SCREENUPDATING
        Application.ScreenUpdating = DEMO_FastModeValue(DEMO_FM_SCREENUPDATING, False, False)

        StateOut.Applied = StateOut.Applied Or DEMO_FM_ENABLEEVENTS
        Application.EnableEvents = DEMO_FastModeValue(DEMO_FM_ENABLEEVENTS, False, False)

        StateOut.Applied = StateOut.Applied Or DEMO_FM_DISPLAYALERTS
        Application.DisplayAlerts = DEMO_FastModeValue(DEMO_FM_DISPLAYALERTS, False, False)

        StateOut.Applied = StateOut.Applied Or DEMO_FM_CALCULATION
        Application.Calculation = DEMO_FastModeValue(DEMO_FM_CALCULATION, False, xlCalculationManual)

        StateOut.Applied = StateOut.Applied Or DEMO_FM_CURSOR
        Application.Cursor = DEMO_FastModeValue(DEMO_FM_CURSOR, False, xlWait)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanExit:
    'Restore what was already written, then re-raise the original error
        On Error GoTo 0
        If SavedErrNumber <> 0 Then
            RollbackFailures = DEMO_FastMode_Restore(StateOut)
            If Len(RollbackFailures) > 0 Then
                SavedErrDescription = SavedErrDescription & vbCrLf & _
                    "Cleanup failure: Fast-mode rollback failed: " & RollbackFailures
            End If
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CleanFail:
    'Capture the original error before the rollback
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description

        Resume CleanExit

End Sub

Public Sub DEMO_FastMode_End( _
    ByRef StateIn As tDEMOFastModeState)
'
'------------------------------------------------------------------------------
'                                END FAST MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Restores the Application settings a DEMO_FastMode_Begin scope applied.
'
' WHY THIS EXISTS
'   This is the cleanup companion to DEMO_FastMode_Begin.
'
' INPUTS
'   StateIn
'     The state DEMO_FastMode_Begin filled.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Restores only the settings marked in StateIn.Applied, in reverse order:
'     Cursor, Calculation, DisplayAlerts, EnableEvents, ScreenUpdating
'   - Attempts every marked setting, whatever an earlier one did
'   - Clears StateIn.Applied, so a second call restores nothing
'   - Does nothing when Begin failed or was never called
'
' ERROR POLICY
'   Raises vbObjectError + 2026 after every attempt when any restoration
'   failed. The description names each failed setting with its error number,
'   source and description.
'
' DEPENDENCIES
'   - DEMO_FastMode_Restore
'
' NOTES
'   None
'
' HISTORY
'   2026-11-14 - Every restoration attempted; failures reported (#70).
'   2026-04-19
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_FastMode_End"

    Dim Failures    As String   'Settings that could not be restored

'------------------------------------------------------------------------------
' RESTORE SAVED STATE
'------------------------------------------------------------------------------
    'Attempt every owned restoration, then report the failures together
        Failures = DEMO_FastMode_Restore(StateIn)

        If Len(Failures) > 0 Then
            Err.Raise vbObjectError + 2026, _
                      "M_DEMO_BUILDER.DEMO_FastMode_End", _
                      "Fast-mode restore failed: " & Failures
        End If

End Sub

Public Sub DEMO_FastMode_EndPreservingError( _
    ByRef StateIn As tDEMOFastModeState, _
    ByRef ErrNumber As Long, _
    ByRef ErrSource As String, _
    ByRef ErrDescription As String)
'
'------------------------------------------------------------------------------
'                      END FAST MODE PRESERVING AN ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Ends fast mode on a caller's cleanup path and folds a restoration failure
'   into the caller's saved error instead of discarding it.
'
' WHY THIS EXISTS
'   A cleanup path that re-raises a saved error would lose a failed
'   restoration, or let it replace the operation error. This routine keeps both
'   (added for #70).
'
' INPUTS
'   StateIn
'     The state DEMO_FastMode_Begin filled.
'
'   ErrNumber / ErrSource / ErrDescription
'     The caller's saved operation error, or 0 when there is none
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Calls DEMO_FastMode_End once
'   - With a saved operation error, keeps its number and source and appends
'     vbCrLf & "Cleanup failure: <number> | <source> | <description>" of the
'     restoration failure to its description
'   - Without one, the restoration failure becomes the saved error
'
' ERROR POLICY
'   Does not raise. The caller re-raises the saved error after disabling its
'   own handler.
'
' DEPENDENCIES
'   - DEMO_FastMode_End
'
' NOTES
'   - The caller's three error values are ByRef, so they are updated in place
'   - On Error Resume Next covers only the DEMO_FastMode_End call
'
' HISTORY
'   2026-11-14 - Added for #70.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_FastMode_EndPreservingError"

    Dim CleanupErrNumber        As Long     'Restoration error number
    Dim CleanupErrSource        As String   'Restoration error source
    Dim CleanupErrDescription   As String   'Restoration error description

'------------------------------------------------------------------------------
' END FAST MODE
'------------------------------------------------------------------------------
    'Attempt the restoration once and capture its failure
        On Error Resume Next
        DEMO_FastMode_End StateIn
        CleanupErrNumber = Err.Number
        CleanupErrSource = Err.Source
        CleanupErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0

'------------------------------------------------------------------------------
' MERGE THE FAILURE
'------------------------------------------------------------------------------
    'Keep the operation error first; report the restoration failure with it
        If CleanupErrNumber <> 0 Then
            If ErrNumber = 0 Then
                ErrNumber = CleanupErrNumber
                ErrSource = CleanupErrSource
                ErrDescription = CleanupErrDescription
            Else
                ErrDescription = ErrDescription & vbCrLf & _
                    "Cleanup failure: " & CStr(CleanupErrNumber) & _
                    " | " & CleanupErrSource & " | " & CleanupErrDescription
            End If
        End If

End Sub

Private Function DEMO_FastMode_Restore( _
    ByRef State As tDEMOFastModeState) _
    As String
'
'------------------------------------------------------------------------------
'                       RESTORE OWNED FAST-MODE SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes back every setting marked in State.Applied, each in its own guarded
'   attempt, and returns the failures.
'
' WHY THIS EXISTS
'   DEMO_FastMode_End and the rollback of DEMO_FastMode_Begin restore the same
'   settings. Both go through this routine, so each attempt is guarded the same
'   way (added for #70).
'
' INPUTS
'   State
'     Fast-mode state; State.Applied names the settings to restore and is
'     cleared afterwards
'
' RETURNS
'   String
'     Empty when every attempt succeeded; otherwise one
'     "<Setting>: <number> | <source> | <description>" entry per failed
'     setting, joined by " || "
'
' BEHAVIOR
'   - Restores in reverse entry order: Cursor, Calculation, DisplayAlerts,
'     EnableEvents, ScreenUpdating
'   - A failed attempt does not stop the later ones
'   - Clears State.Applied afterwards: the scope gives up its ownership once
'     every restoration has been attempted
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   - Application.Cursor, Calculation, DisplayAlerts, EnableEvents and
'     ScreenUpdating
'   - DEMO_FastModeValue
'   - DEMO_FastMode_AddFailure
'
' NOTES
'   - On Error Resume Next is on for the whole restoration and switched off
'     before State.Applied is cleared
'   - Each write passes through DEMO_FastModeValue, so an armed restoration
'     fault fails that attempt inside the guarded step
'
' HISTORY
'   2026-11-14 - Added for #70.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_FastMode_Restore"

    Dim Failures        As String   'Collected failures
    Dim FailNumber      As Long     'Error number of one attempt
    Dim FailSource      As String   'Error source of one attempt
    Dim FailDescription As String   'Error description of one attempt

'------------------------------------------------------------------------------
' RESTORE EACH OWNED SETTING
'------------------------------------------------------------------------------
        On Error Resume Next

    'Restore the cursor
        If (State.Applied And DEMO_FM_CURSOR) <> 0 Then
            Err.Clear
            Application.Cursor = DEMO_FastModeValue(DEMO_FM_CURSOR, True, State.Cursor)
            FailNumber = Err.Number
            FailSource = Err.Source
            FailDescription = Err.Description
            Failures = DEMO_FastMode_AddFailure(Failures, DEMO_FM_CURSOR, FailNumber, FailSource, FailDescription)
        End If

    'Restore the calculation mode
        If (State.Applied And DEMO_FM_CALCULATION) <> 0 Then
            Err.Clear
            Application.Calculation = DEMO_FastModeValue(DEMO_FM_CALCULATION, True, State.Calculation)
            FailNumber = Err.Number
            FailSource = Err.Source
            FailDescription = Err.Description
            Failures = DEMO_FastMode_AddFailure(Failures, DEMO_FM_CALCULATION, FailNumber, FailSource, FailDescription)
        End If

    'Restore alerts
        If (State.Applied And DEMO_FM_DISPLAYALERTS) <> 0 Then
            Err.Clear
            Application.DisplayAlerts = DEMO_FastModeValue(DEMO_FM_DISPLAYALERTS, True, State.DisplayAlerts)
            FailNumber = Err.Number
            FailSource = Err.Source
            FailDescription = Err.Description
            Failures = DEMO_FastMode_AddFailure(Failures, DEMO_FM_DISPLAYALERTS, FailNumber, FailSource, _
                FailDescription)
        End If

    'Restore events
        If (State.Applied And DEMO_FM_ENABLEEVENTS) <> 0 Then
            Err.Clear
            Application.EnableEvents = DEMO_FastModeValue(DEMO_FM_ENABLEEVENTS, True, State.EnableEvents)
            FailNumber = Err.Number
            FailSource = Err.Source
            FailDescription = Err.Description
            Failures = DEMO_FastMode_AddFailure(Failures, DEMO_FM_ENABLEEVENTS, FailNumber, FailSource, FailDescription)
        End If

    'Restore screen updating last, once everything else is back
        If (State.Applied And DEMO_FM_SCREENUPDATING) <> 0 Then
            Err.Clear
            Application.ScreenUpdating = DEMO_FastModeValue(DEMO_FM_SCREENUPDATING, True, State.ScreenUpdating)
            FailNumber = Err.Number
            FailSource = Err.Source
            FailDescription = Err.Description
            Failures = DEMO_FastMode_AddFailure(Failures, DEMO_FM_SCREENUPDATING, FailNumber, FailSource, _
                FailDescription)
        End If

        Err.Clear
        On Error GoTo 0

'------------------------------------------------------------------------------
' RELEASE OWNERSHIP
'------------------------------------------------------------------------------
        State.Applied = 0
        DEMO_FastMode_Restore = Failures

End Function

Private Function DEMO_FastMode_AddFailure( _
    ByVal Failures As String, _
    ByVal Setting As Long, _
    ByVal FailNumber As Long, _
    ByVal FailSource As String, _
    ByVal FailDescription As String) _
    As String
'
'------------------------------------------------------------------------------
'                       APPEND FAST-MODE RESTORE FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one failed restoration to the failure list; returns the list
'   unchanged when FailNumber is 0.
'
' ERROR POLICY
'   Does not raise
'
' HISTORY
'   2026-11-14 - Added for #70.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

        If FailNumber = 0 Then
            DEMO_FastMode_AddFailure = Failures
        Else
            If Len(Failures) > 0 Then
                Failures = Failures & " || "
            End If
            DEMO_FastMode_AddFailure = Failures & DEMO_FastModeSettingName(Setting) & ": " & _
                CStr(FailNumber) & " | " & FailSource & " | " & FailDescription
        End If

End Function

Private Function DEMO_FastModeSettingName( _
    ByVal Setting As Long) _
    As String
'
'------------------------------------------------------------------------------
'                           NAME A FAST-MODE SETTING
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the Application property name of one DEMO_FM_* setting bit.
'
' HISTORY
'   2026-11-14 - Added for #70.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

        Select Case Setting
            Case DEMO_FM_SCREENUPDATING
                DEMO_FastModeSettingName = "ScreenUpdating"
            Case DEMO_FM_ENABLEEVENTS
                DEMO_FastModeSettingName = "EnableEvents"
            Case DEMO_FM_DISPLAYALERTS
                DEMO_FastModeSettingName = "DisplayAlerts"
            Case DEMO_FM_CALCULATION
                DEMO_FastModeSettingName = "Calculation"
            Case DEMO_FM_CURSOR
                DEMO_FastModeSettingName = "Cursor"
            Case Else
                DEMO_FastModeSettingName = "Setting " & CStr(Setting)
        End Select

End Function

Private Function DEMO_FastModeValue( _
    ByVal Setting As Long, _
    ByVal Restoring As Boolean, _
    ByVal Value As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                         SUPPLY FAST-MODE WRITE VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Supplies the value of one fast-mode entry or restoration write.
'
' WHY THIS EXISTS
'   No host condition makes an Application setter fail on demand. Raising
'   while the value is evaluated fails that exact write, so the guarded
'   rollback and restoration paths run unchanged and Excel is not called.
'
' INPUTS
'   Setting
'     One DEMO_FM_* setting bit
'
'   Restoring
'     TRUE for a restoration write, FALSE for an entry write
'
'   Value
'     The value to write
'
' RETURNS
'   Variant
'     Value, unchanged
'
' BEHAVIOR
'   - Counts the write: DEMO_PROBE_FM_RESTORE for a restoration,
'     DEMO_PROBE_FM_ENTRY for an entry write
'   - Resolves the fault bit as Setting times DEMO_FAULT_FM_RESTORE_UNIT or
'     DEMO_FAULT_FM_ENTRY_UNIT
'   - When that bit is armed, clears it and raises the injected fault
'   - Otherwise returns Value unchanged
'
' ERROR POLICY
'   Raises once while M_DEMO_CLEANUP_TESTS has armed a fault for this setting
'   and phase: vbObjectError + 2193 for an entry write, vbObjectError + 2194
'   for a restoration write
'
' DEPENDENCIES
'   - DEMO_CleanupProbeHit
'   - DEMO_FastModeSettingName
'   - mCleanupFaults, the armed fault bits
'
' NOTES
'   - Inert unless M_DEMO_CLEANUP_TESTS has armed a fault through
'     DEMO_InternalCleanupProbeBegin
'
' HISTORY
'   2026-11-14 - Added for #70.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_FastModeValue"

    Dim FaultBit    As Long 'Fault bit of this setting and phase

'------------------------------------------------------------------------------
' COUNT AND RAISE ARMED FAULT
'------------------------------------------------------------------------------
    'Count the write and resolve its fault bit
        If Restoring Then
            DEMO_CleanupProbeHit DEMO_PROBE_FM_RESTORE
            FaultBit = Setting * DEMO_FAULT_FM_RESTORE_UNIT
        Else
            DEMO_CleanupProbeHit DEMO_PROBE_FM_ENTRY
            FaultBit = Setting * DEMO_FAULT_FM_ENTRY_UNIT
        End If

    'Fail the write once when armed
        If (mCleanupFaults And FaultBit) <> 0 Then
            mCleanupFaults = mCleanupFaults And Not FaultBit
            If Restoring Then
                Err.Raise vbObjectError + 2194, _
                          "M_DEMO_BUILDER.DEMO_FastMode.InjectedRestoreFault", _
                          "Injected fast-mode restore fault: " & DEMO_FastModeSettingName(Setting) & "."
            Else
                Err.Raise vbObjectError + 2193, _
                          "M_DEMO_BUILDER.DEMO_FastMode.InjectedEntryFault", _
                          "Injected fast-mode entry fault: " & DEMO_FastModeSettingName(Setting) & "."
            End If
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Pass the value through unchanged
        DEMO_FastModeValue = Value

End Function

Public Sub DEMO_Set_RangeBorder( _
    ByVal Target As Range, _
    Optional ByVal BorderColor As Long = 0, _
    Optional ByVal BorderWeight As XlBorderWeight = xlThin, _
    Optional ByVal IncludeInside As Boolean = True, _
    Optional ByVal InsideColor As Long = vbWhite)
'
'------------------------------------------------------------------------------
'                               SET RANGE BORDER
'------------------------------------------------------------------------------
' PURPOSE
'   Applies a consistent border format to a target range.
'
' WHY THIS EXISTS
'   Border formatting is often repeated across DEMO-sheet setup and reporting
'   routines.
'
'   This helper centralizes the logic so callers can apply:
'     - outside borders only, or
'     - outside + inside borders
'
'   using one small reusable routine.
'
' INPUTS
'   Target
'     The range to format.
'
'   BorderColor (optional)
'     The outside-border color as a VBA Long.
'     Default = 0 => black => RGB(0, 0, 0)
'
'   BorderWeight (optional)
'     The Excel border weight to apply.
'     Default = xlThin
'
'   IncludeInside (optional)
'     TRUE  => apply inside horizontal and vertical borders where applicable
'     FALSE => apply outside borders only
'
'   InsideColor (optional)
'     The inside-border color as a VBA Long.
'     Default = vbWhite
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits quietly if Target Is Nothing
'   - Applies left / top / bottom / right borders
'   - Optionally applies inside horizontal and vertical borders
'   - Uses xlContinuous as the line style
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - Excel Range object model
'   - XlBordersIndex constants
'   - XlBorderWeight constants
'
' NOTES
'   - Inside borders are meaningful only for multi-cell ranges
'   - BorderColor = 0 corresponds to black
'   - InsideColor defaults to white so the caller can visually separate inner
'     grid lines from the outer frame when desired
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no target range was supplied
        If Target Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' APPLY OUTSIDE BORDERS
'------------------------------------------------------------------------------
    'Apply the left border
        With Target.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = BorderWeight
            .Color = BorderColor
        End With

    'Apply the top border
        With Target.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = BorderWeight
            .Color = BorderColor
        End With

    'Apply the bottom border
        With Target.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = BorderWeight
            .Color = BorderColor
        End With

    'Apply the right border
        With Target.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = BorderWeight
            .Color = BorderColor
        End With

'------------------------------------------------------------------------------
' APPLY INSIDE BORDERS
'------------------------------------------------------------------------------
    'Apply inside borders only when requested
        If IncludeInside Then

        'Apply the inside horizontal borders
            With Target.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = BorderWeight
                .Color = InsideColor
            End With

        'Apply the inside vertical borders
            With Target.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = BorderWeight
                .Color = InsideColor
            End With

        End If

End Sub

Public Function DEMO_Sheet_GetOrCreate( _
    ByVal WB As Workbook, _
    ByVal SheetName As String) _
    As Worksheet
'
'------------------------------------------------------------------------------
'                             GET OR CREATE SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns an existing worksheet by name, or creates it if missing.
'
' WHY THIS EXISTS
'   The DEMO builder needs repeatable access to a known set of sheets while
'   remaining safe to run on a fresh workbook or a partially prepared workbook.
'
' INPUTS
'   WB
'     Target workbook.
'
'   SheetName
'     Required worksheet name.
'
' RETURNS
'   Worksheet
'     Existing or newly created worksheet.
'
' BEHAVIOR
'   - Refuses a name Excel would not accept for a sheet before creating
'     anything
'   - Searches the workbook for the requested worksheet name; a chart or other
'     non-worksheet sheet with that name is refused
'   - Creates a new worksheet if the name is not found
'   - Renames the new sheet to the requested name; when the rename fails, the
'     new sheet is deleted before the error is re-raised, so no unnamed sheet
'     is left behind
'
' ERROR POLICY
'   Raises errors normally. A failed deletion after a failed rename is
'   appended to the rename error's description as a "Cleanup failure: " line.
'
' DEPENDENCIES
'   - DEMO_SheetNameProblem
'   - DEMO_Sheet_FindWorksheet
'   - DEMO_SheetNameValue
'   - DEMO_Sheet_DeleteOwned
'   - Workbook.Worksheets.Add and Worksheet.Name
'
' NOTES
'   - Names are compared without regard to case, as Excel does
'   - The new worksheet is added after the last worksheet, or after the last
'     sheet of a workbook that has no worksheet
'
' REGRESSION SEAMS
'   The new name passes through DEMO_SheetNameValue, which fails the rename
'   once while M_DEMO_CLEANUP_TESTS has armed that fault.
'
' HISTORY
'   2026-11-14 - Name checked first; a sheet whose rename fails is removed
'                (#74).
'   2026-04-19
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Sheet_GetOrCreate"

    Dim WS                  As Worksheet    'Existing or new worksheet
    Dim NameProblem         As String       'Why SheetName is not a valid sheet name
    Dim NameErrNumber       As Long         'Rename error number
    Dim NameErrSource       As String       'Rename error source
    Dim NameErrDescription  As String       'Rename error description
    Dim CleanupFailures     As String       'Failed removal of the new sheet

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing workbook reference
        If WB Is Nothing Then
            Err.Raise vbObjectError + 2001, _
                      "M_DEMO_BUILDER.DEMO_Sheet_GetOrCreate", _
                      "Workbook reference cannot be Nothing."
        End If

    'Reject a blank worksheet name
        If Len(Trim$(SheetName)) = 0 Then
            Err.Raise vbObjectError + 2002, _
                      "M_DEMO_BUILDER.DEMO_Sheet_GetOrCreate", _
                      "Worksheet name cannot be blank."
        End If

    'Reject a name Excel would refuse for a sheet
        NameProblem = DEMO_SheetNameProblem(SheetName)
        If Len(NameProblem) > 0 Then
            Err.Raise vbObjectError + 2027, _
                      "M_DEMO_BUILDER.DEMO_Sheet_GetOrCreate", _
                      "Worksheet name is not valid: " & NameProblem
        End If

'------------------------------------------------------------------------------
' SEARCH EXISTING SHEETS
'------------------------------------------------------------------------------
    'Return an existing worksheet; a non-worksheet sheet with the name raises
        Set WS = DEMO_Sheet_FindWorksheet(WB, SheetName, "M_DEMO_BUILDER.DEMO_Sheet_GetOrCreate")
        If Not WS Is Nothing Then
            Set DEMO_Sheet_GetOrCreate = WS
            Exit Function
        End If

'------------------------------------------------------------------------------
' CREATE SHEET
'------------------------------------------------------------------------------
    'Create a new worksheet after the last worksheet, or after the last sheet
    'of a workbook that has no worksheet
        If WB.Worksheets.Count > 0 Then
            Set WS = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
        Else
            Set WS = WB.Worksheets.Add(After:=WB.Sheets(WB.Sheets.Count))
        End If

'------------------------------------------------------------------------------
' NAME SHEET
'------------------------------------------------------------------------------
    'Assign the requested name; on failure delete the new sheet, then re-raise
        On Error Resume Next
        WS.Name = DEMO_SheetNameValue(SheetName)
        NameErrNumber = Err.Number
        NameErrSource = Err.Source
        NameErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0

        If NameErrNumber <> 0 Then
            CleanupFailures = DEMO_Sheet_DeleteOwned(WS)
            If Len(CleanupFailures) > 0 Then
                NameErrDescription = NameErrDescription & vbCrLf & "Cleanup failure: " & CleanupFailures
            End If
            Err.Raise NameErrNumber, NameErrSource, NameErrDescription
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the existing or newly created worksheet
        Set DEMO_Sheet_GetOrCreate = WS

End Function

Private Function DEMO_Sheet_FindWorksheet( _
    ByVal WB As Workbook, _
    ByVal SheetName As String, _
    ByVal ErrSource As String) _
    As Worksheet
'
'------------------------------------------------------------------------------
'                            FIND WORKSHEET BY NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the worksheet of WB named SheetName, or Nothing when no sheet has
'   that name
'
' WHY THIS EXISTS
'   Worksheet and chart-sheet names share one namespace. Looking only at
'   WB.Worksheets would report a chart sheet's name as free, and creating a
'   worksheet under it would then fail at the rename.
'
' INPUTS
'   WB
'     Workbook whose sheets are searched
'
'   SheetName
'     Name to look for, compared without regard to case
'
'   ErrSource
'     Source of the error raised for a name held by another sheet type, so
'     the error names the public procedure that refused the request
'
' ERROR POLICY
'   Raises vbObjectError + 2028 when a chart or other non-worksheet sheet
'   already has the name
'
' HISTORY
'   2026-11-14 - Added for #74.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Sheet_FindWorksheet"

    Dim Sh  As Object   'Sheet of any type

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Compare names the way Excel does: without regard to case
        For Each Sh In WB.Sheets
            If StrComp(Sh.Name, SheetName, vbTextCompare) = 0 Then
                If TypeOf Sh Is Worksheet Then
                    Set DEMO_Sheet_FindWorksheet = Sh
                    Exit Function
                End If
                Err.Raise vbObjectError + 2028, ErrSource, _
                          "The name is already used by a sheet that is not a worksheet: " & SheetName
            End If
        Next Sh

End Function

Private Function DEMO_Sheet_DeleteOwned( _
    ByVal WS As Worksheet) _
    As String
'
'------------------------------------------------------------------------------
'                           DELETE CREATED WORKSHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes a worksheet that the calling invocation created, on its failure
'   path
'
' WHY THIS EXISTS
'   The failure paths of DEMO_Sheet_GetOrCreate and DEMO_Sheet_BuildTemplate
'   remove a worksheet they created without losing their original error. The
'   delete is guarded and its failures come back as text (added for #74).
'
' INPUTS
'   WS
'     Worksheet to delete; Nothing makes the routine return an empty string
'     without doing anything
'
' RETURNS
'   String
'     Empty on success; otherwise "<number> | <source> | <description>" of the
'     failed Delete, and of a failed DisplayAlerts write-back after " || "
'
' BEHAVIOR
'   - Turns DisplayAlerts off for the Delete only, so Excel does not ask for
'     confirmation, and writes back the value it read
'   - Attempts the Delete once; Excel refuses to delete the last visible sheet
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   - Application.DisplayAlerts
'   - Worksheet.Delete
'
' NOTES
'   - Callers append a non-empty result to their saved error as a
'     "Cleanup failure: " line
'
' HISTORY
'   2026-11-14 - Added for #74.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Sheet_DeleteOwned"

    Dim Failures        As String   'Collected failures
    Dim AlertsOn        As Boolean  'DisplayAlerts before the Delete
    Dim AlertsRead      As Boolean  'TRUE when AlertsOn was read
    Dim FailNumber      As Long     'Error number of one step
    Dim FailSource      As String   'Error source of one step
    Dim FailDescription As String   'Error description of one step

'------------------------------------------------------------------------------
' DELETE
'------------------------------------------------------------------------------
        If WS Is Nothing Then Exit Function

        On Error Resume Next

    'Suppress the confirmation prompt for this Delete only
        Err.Clear
        AlertsOn = Application.DisplayAlerts
        FailNumber = Err.Number
        AlertsRead = (FailNumber = 0)

        Err.Clear
        Application.DisplayAlerts = False

    'Delete the owned worksheet once
        Err.Clear
        WS.Delete
        FailNumber = Err.Number
        FailSource = Err.Source
        FailDescription = Err.Description
        If FailNumber <> 0 Then
            Failures = CStr(FailNumber) & " | " & FailSource & " | " & FailDescription
        End If

    'Write back the alert setting that was read
        If AlertsRead Then
            Err.Clear
            Application.DisplayAlerts = AlertsOn
            FailNumber = Err.Number
            FailSource = Err.Source
            FailDescription = Err.Description
            If FailNumber <> 0 Then
                If Len(Failures) > 0 Then
                    Failures = Failures & " || "
                End If
                Failures = Failures & "DisplayAlerts: " & CStr(FailNumber) & _
                    " | " & FailSource & " | " & FailDescription
            End If
        End If

        Err.Clear
        On Error GoTo 0

        DEMO_Sheet_DeleteOwned = Failures

End Function

Private Function DEMO_SheetNameProblem( _
    ByVal SheetName As String) _
    As String
'
'------------------------------------------------------------------------------
'                            CHECK SHEET NAME RULES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns why Excel would refuse SheetName as a sheet name, or an empty
'   string when the name passes these rules
'
' WHY THIS EXISTS
'   A name Excel would refuse is rejected before a worksheet is created, so a
'   failed call cannot leave an unnamed sheet behind (added for #74).
'
' INPUTS
'   SheetName
'     Candidate worksheet name
'
' RETURNS
'   String
'     Empty when SheetName passes; otherwise the first problem found, such as
'     "it is blank", "it is longer than 31 characters" or "it contains the
'     character <c>"
'
' BEHAVIOR
'   Checks the documented sheet-name rules: not blank, at most 31 characters,
'   none of : \ / ? * [ ], no apostrophe at either end, and not History.
'   Uniqueness is checked separately against the workbook.
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   - VBA string functions only
'
' NOTES
'   - The History comparison ignores case; the reserved characters are matched
'     with a binary compare
'
' HISTORY
'   2026-11-14 - Added for #74.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_SheetNameProblem"

    Dim Pos As Long     'Character position
    Dim Ch  As String   'Character at Pos

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If Len(SheetName) = 0 Then
            DEMO_SheetNameProblem = "it is blank"
            Exit Function
        End If

        If Len(SheetName) > 31 Then
            DEMO_SheetNameProblem = "it is longer than 31 characters"
            Exit Function
        End If

        If Left$(SheetName, 1) = "'" Or Right$(SheetName, 1) = "'" Then
            DEMO_SheetNameProblem = "it begins or ends with an apostrophe"
            Exit Function
        End If

        If StrComp(SheetName, "History", vbTextCompare) = 0 Then
            DEMO_SheetNameProblem = "History is reserved by Excel"
            Exit Function
        End If

        For Pos = 1 To Len(SheetName)
            Ch = Mid$(SheetName, Pos, 1)
            If InStr(1, ":\/?*[]", Ch, vbBinaryCompare) > 0 Then
                DEMO_SheetNameProblem = "it contains the character " & Ch
                Exit Function
            End If
        Next Pos

End Function

Private Function DEMO_ParseColumnSpec( _
    ByVal Spec As String, _
    ByRef FirstCol As Long, _
    ByRef LastCol As Long) _
    As Boolean
'
'------------------------------------------------------------------------------
'                          PARSE COLUMN SPECIFICATION
'------------------------------------------------------------------------------
' PURPOSE
'   Parses "C" or "C:M" (letters only, case-insensitive, each part optionally
'   prefixed with $) into 1-based column indexes, without touching a worksheet
'
' WHY THIS EXISTS
'   DEMO_Sheet_BuildTemplate checks its column specifications before it creates
'   or resets anything, so the parse works on text alone (added for #74).
'
' INPUTS
'   Spec
'     Column text such as "C" or "C:M"
'
'   FirstCol
'     Receives the first column index, or 0 when Spec is not valid
'
'   LastCol
'     Receives the last column index, or 0 when Spec is not valid
'
' RETURNS
'   Boolean
'     TRUE with FirstCol <= LastCol, a reversed range being swapped; FALSE
'     with both set to 0 when Spec is not one column or one column range
'
' BEHAVIOR
'   - Sets both outputs to 0 and splits the trimmed Spec on ":"
'   - One part gives both columns; two parts give the first and the last; any
'     other count returns FALSE
'   - Resolves each part with DEMO_ColumnIndex and returns FALSE when either is
'     not a column
'   - Swaps a reversed range so FirstCol is not greater than LastCol
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   - DEMO_ColumnIndex
'
' NOTES
'   The indexes are not checked against the worksheet's last column here;
'   the caller does that once it knows the grid.
'
' HISTORY
'   2026-11-14 - Added for #74.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_ParseColumnSpec"

    Dim Parts   As Variant  'Spec split on ":"
    Dim Swap    As Long     'Swap buffer

'------------------------------------------------------------------------------
' PARSE
'------------------------------------------------------------------------------
        FirstCol = 0
        LastCol = 0

        Parts = Split(Trim$(Spec), ":")

        Select Case UBound(Parts)
            Case 0
                FirstCol = DEMO_ColumnIndex(Parts(0))
                LastCol = FirstCol
            Case 1
                FirstCol = DEMO_ColumnIndex(Parts(0))
                LastCol = DEMO_ColumnIndex(Parts(1))
            Case Else
                Exit Function
        End Select

        If FirstCol = 0 Or LastCol = 0 Then
            FirstCol = 0
            LastCol = 0
            Exit Function
        End If

        If LastCol < FirstCol Then
            Swap = FirstCol
            FirstCol = LastCol
            LastCol = Swap
        End If

        DEMO_ParseColumnSpec = True

End Function

Private Function DEMO_ColumnIndex( _
    ByVal Part As String) _
    As Long
'
'------------------------------------------------------------------------------
'                       CONVERT COLUMN LETTERS TO INDEX
'------------------------------------------------------------------------------
' PURPOSE
'   Converts one to three column letters, optionally prefixed with $, to a
'   1-based column index; returns 0 for anything else
'
' WHY THIS EXISTS
'   DEMO_ParseColumnSpec needs a conversion from column letters to an index that
'   does not touch a worksheet (added for #74).
'
' INPUTS
'   Part
'     One to three column letters, optionally prefixed with $
'
' RETURNS
'   Long
'     1-based column index, or 0 when Part is empty, longer than three letters
'     or holds a character outside A to Z
'
' BEHAVIOR
'   - Trims and upper-cases Part and drops one leading $
'   - Returns 0 when 1 to 3 letters do not remain
'   - Accumulates the index in base 26, returning 0 at the first character that
'     is not A to Z
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   - VBA string functions only
'
' NOTES
'   - Up to ZZZ is accepted, which can lie beyond the last worksheet column;
'     callers check the grid
'
' HISTORY
'   2026-11-14 - Added for #74.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_ColumnIndex"

    Dim Letters As String   'Normalized letters
    Dim Pos     As Long     'Letter position
    Dim Code    As Long     'Character code at Pos
    Dim Index   As Long     'Accumulated index

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
        Letters = UCase$(Trim$(Part))
        If Left$(Letters, 1) = "$" Then
            Letters = Mid$(Letters, 2)
        End If

        If Len(Letters) < 1 Or Len(Letters) > 3 Then Exit Function

        For Pos = 1 To Len(Letters)
            Code = Asc(Mid$(Letters, Pos, 1))
            If Code < 65 Or Code > 90 Then Exit Function
            Index = Index * 26 + (Code - 64)
        Next Pos

        DEMO_ColumnIndex = Index

End Function

Private Function DEMO_SheetNameValue( _
    ByVal SheetName As String) _
    As String
'
'------------------------------------------------------------------------------
'                            SUPPLY NEW SHEET NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Supplies the name DEMO_Sheet_GetOrCreate assigns to a sheet it created
'
' WHY THIS EXISTS
'   Names that would fail the rename are refused before the sheet is created,
'   so no valid request makes the rename fail on demand. Raising while the
'   value is evaluated fails that exact statement and exercises the removal
'   of the new sheet.
'
' RETURNS
'   String
'     SheetName, unchanged
'
' ERROR POLICY
'   Raises vbObjectError + 2196 once when the naming fault is armed
'
' HISTORY
'   2026-11-14 - Added for #74.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RAISE ARMED FAULT
'------------------------------------------------------------------------------
    'Fail the rename once when armed
        If (mCleanupFaults And DEMO_FAULT_SHEET_NAME) <> 0 Then
            mCleanupFaults = mCleanupFaults And Not DEMO_FAULT_SHEET_NAME
            Err.Raise vbObjectError + 2196, _
                      "M_DEMO_BUILDER.DEMO_Sheet_GetOrCreate.InjectedNameFault", _
                      "Injected worksheet naming fault."
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Pass the name through unchanged
        DEMO_SheetNameValue = SheetName

End Function

Public Sub DEMO_Sheet_Reset( _
    Optional ByVal WS_In As Variant, _
    Optional ByVal DeleteShapes As Boolean = True, _
    Optional ByVal DeleteTables As Boolean = True, _
    Optional ByVal UnmergeCells As Boolean = True, _
    Optional ByVal ClearValidation As Boolean = True, _
    Optional ByVal ClearCells As Boolean = True, _
    Optional ByVal ClearCommentsAndNotes As Boolean = True, _
    Optional ByVal ResetWrapText As Boolean = True, _
    Optional ByVal ResetColumnWidths As Boolean = True, _
    Optional ByVal ResetRowHeights As Boolean = True, _
    Optional ByVal UnhideRows As Boolean = True, _
    Optional ByVal UnhideColumns As Boolean = True, _
    Optional ByVal ResetTabColor As Boolean = True, _
    Optional ByVal RemovePageBreaks As Boolean = True, _
    Optional ByVal ProtectPassword As String = vbNullString, _
    Optional ByVal ReProtectAtEnd As Boolean = False)
'
'------------------------------------------------------------------------------
'                                 RESET SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Clears a worksheet to a clean, reusable state for DEMO-sheet rebuilding
'
' WHY THIS EXISTS
'   Re-running the DEMO builder should rebuild the layout predictably without
'   requiring manual deletion of old shapes, formatting, content, tables, or
'   merged cells
'
'   This version supports:
'     - selective reset steps
'     - safer collection deletion
'     - optional protected-sheet handling
'     - optional defaulting to ActiveSheet when no worksheet is supplied
'     - optional row / column unhiding
'
' INPUTS
'   WS_In (optional)
'     Worksheet to reset
'
'     When omitted, the routine defaults to ActiveSheet
'
'   DeleteShapes (optional)
'     TRUE  => delete all worksheet shapes
'     FALSE => leave shapes unchanged
'
'   DeleteTables (optional)
'     TRUE  => delete all worksheet ListObjects
'     FALSE => leave tables unchanged
'
'   UnmergeCells (optional)
'     TRUE  => unmerge all cells
'     FALSE => leave merged ranges unchanged
'
'   ClearValidation (optional)
'     TRUE  => delete all data-validation rules
'     FALSE => leave validation unchanged
'
'   ClearCells (optional)
'     TRUE  => clear cell values, formulas, and formatting
'     FALSE => leave cell content/format unchanged
'
'   ClearCommentsAndNotes (optional)
'     TRUE  => explicitly clear comments and notes on a best-effort basis
'     FALSE => leave them unchanged
'
'   ResetWrapText (optional)
'     TRUE  => set WrapText = False for all cells
'     FALSE => leave wrapping unchanged
'
'   ResetColumnWidths (optional)
'     TRUE  => reset all column widths to Excel's neutral default
'     FALSE => leave column widths unchanged
'
'   ResetRowHeights (optional)
'     TRUE  => reset all row heights to Excel's neutral default
'     FALSE => leave row heights unchanged
'
'   UnhideRows (optional)
'     TRUE  => unhide all worksheet rows
'     FALSE => preserve row hidden state
'
'   UnhideColumns (optional)
'     TRUE  => unhide all worksheet columns
'     FALSE => preserve column hidden state
'
'   ResetTabColor (optional)
'     TRUE  => clear the worksheet tab color
'     FALSE => leave tab color unchanged
'
'   RemovePageBreaks (optional)
'     TRUE  => remove manual page breaks and hide indicators on a best-effort basis
'     FALSE => leave page-break settings unchanged
'
'     Removing manual page breaks makes Excel lay the sheet out through the
'     default printer driver, which waits on a slow, offline or network
'     printer. Pass FALSE for a sheet known to have no manual breaks. The
'     indicators are hidden only when they are shown.
'
'   ProtectPassword (optional)
'     Password used to temporarily unprotect and optionally re-protect the sheet
'
'   ReProtectAtEnd (optional)
'     TRUE  => re-protect the sheet at the end when it was successfully unprotected
'     FALSE => leave the sheet unprotected after reset
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Resolves the target: WS_In, or ActiveSheet when WS_In is omitted or
'     Nothing; refuses a sheet that is not a worksheet
'   - Captures whether the sheet is protected, every Protect option and
'     EnableSelection
'   - Unprotects with ProtectPassword; raises vbObjectError + 2100 when the
'     sheet is protected and no password was supplied
'   - Applies each requested step in this order: deletes shapes, deletes
'     tables, unmerges cells, clears validation, clears cells, clears comments
'     and notes, turns wrap text off, unhides rows, unhides columns, resets
'     column widths, resets row heights, resets the tab color, and removes
'     manual page breaks and hides the indicators
'   - Calls the armed operation-fault seam
'   - Cleanup pass: when this call unprotected the sheet and ReProtectAtEnd is
'     TRUE, restores the Protect options and EnableSelection, one attempt each,
'     then re-raises the saved error
'
' ERROR POLICY
'   Raises errors normally except for explicitly marked best-effort cleanup
'   steps guarded with On Error Resume Next
'
'   Protection cleanup errors are reported, not silently ignored. An existing
'   operation error keeps its number/source and original description prefix;
'   a separate cleanup diagnostic is appended. Otherwise cleanup failure is
'   raised directly. There is no retry or watchdog for a blocked Excel call.
'
' DEPENDENCIES
'   - Worksheet Shapes, ListObjects, Cells, Rows, Columns, Tab, Protect,
'     Unprotect, EnableSelection and ResetAllPageBreaks
'   - DEMO_Sheet_HidePageBreaks
'   - DEMO_CleanupProbeHit
'   - DEMO_CleanupOperationFault
'   - DEMO_CleanupProtectPassword
'   - DEMO_CleanupSelectionValue
'
' NOTES
'   - Clearing validation, comments and notes and removing page breaks are
'     guarded with On Error Resume Next and are best effort
'   - The handler captures the error and resumes at CleanExit, so cleanup runs
'     once whether or not the body failed
'   - DEMO_Sheet_BuildTemplate calls it with RemovePageBreaks:=False (#187)
'
' PROTECTION CONTRACT
'   Captures Protect options and EnableSelection before unprotecting. Cleanup
'   attempts each restoration once. This does not roll back deleted contents,
'   shapes, tables or AllowEditRanges definitions; host verification is required.
'   The supplied password is reused; no password is read back from Excel.
'
' REGRESSION SEAMS
'   While M_DEMO_CLEANUP_TESTS has armed the cleanup probe, handler entries,
'   cleanup passes and restoration attempts are counted, and an armed one-shot
'   fault can raise an operation error after the protection steps, fail the
'   Protect statement while its arguments are evaluated, or fail the
'   EnableSelection restoration the same way. Disarmed, every seam is inert:
'   the password and selection values pass through unchanged. The
'   ResetAllPageBreaks call and the DisplayPageBreaks write are counted too.
'
' HISTORY
'   2026-11-15 - DisplayPageBreaks written only when True; page-break calls
'                counted by the probe (#187).
'   2026-11-06
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Sheet_Reset"

    Dim WS                      As Worksheet    'Resolved worksheet to reset
    Dim i                       As Long         'Reverse loop index
    Dim WasProtected            As Boolean      'TRUE when the sheet was protected on entry
    Dim DidUnprotect            As Boolean      'This invocation owns the unprotect
    Dim EntryProtection         As tDEMOProtectionState
    Dim CleanupErrNumber        As Long
    Dim CleanupErrSource        As String
    Dim CleanupErrDescription   As String
    Dim SelectionErrNumber      As Long
    Dim SelectionErrSource      As String
    Dim SelectionErrDescription As String
    Dim SavedErrNumber          As Long         'Captured error number
    Dim SavedErrSource          As String       'Captured error source
    Dim SavedErrDescription     As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' RESOLVE TARGET WORKSHEET
'------------------------------------------------------------------------------
    'Use ActiveSheet when no worksheet argument was supplied
        If IsMissing(WS_In) Or IsEmpty(WS_In) Then

        'Reject non-worksheet active sheets such as chart sheets
            If Not TypeOf ActiveSheet Is Worksheet Then
                Err.Raise vbObjectError + 2101, _
                          "M_DEMO_BUILDER.DEMO_Sheet_Reset", _
                          "ActiveSheet is not a worksheet."
            End If

        'Resolve the active worksheet
            Set WS = ActiveSheet

        ElseIf IsObject(WS_In) Then

        'Default Nothing object arguments to ActiveSheet as well
            If WS_In Is Nothing Then

            'Reject non-worksheet active sheets such as chart sheets
                If Not TypeOf ActiveSheet Is Worksheet Then
                    Err.Raise vbObjectError + 2102, _
                              "M_DEMO_BUILDER.DEMO_Sheet_Reset", _
                              "ActiveSheet is not a worksheet."
                End If

            'Resolve the active worksheet
                Set WS = ActiveSheet

            ElseIf TypeOf WS_In Is Worksheet Then

            'Use the supplied worksheet
                Set WS = WS_In

            Else
                Err.Raise vbObjectError + 2103, _
                          "M_DEMO_BUILDER.DEMO_Sheet_Reset", _
                          "WS_In must be a Worksheet when supplied."
            End If

        Else
            Err.Raise vbObjectError + 2104, _
                      "M_DEMO_BUILDER.DEMO_Sheet_Reset", _
                      "WS_In must be omitted or be a Worksheet object."
        End If

'------------------------------------------------------------------------------
' CAPTURE PROTECTION STATE
'------------------------------------------------------------------------------
    'Capture whether the sheet is protected on entry
        WasProtected = WS.ProtectContents Or WS.ProtectDrawingObjects Or WS.ProtectScenarios

    'Capture every Protect option before changing protection; getters may fail.
        If WasProtected Then
            EntryProtection.DrawingObjects = WS.ProtectDrawingObjects
            EntryProtection.Contents = WS.ProtectContents
            EntryProtection.Scenarios = WS.ProtectScenarios
            EntryProtection.UserInterfaceOnly = WS.ProtectionMode
            EntryProtection.EnableSelection = WS.EnableSelection
            With WS.Protection
                EntryProtection.AllowFormattingCells = .AllowFormattingCells
                EntryProtection.AllowFormattingColumns = .AllowFormattingColumns
                EntryProtection.AllowFormattingRows = .AllowFormattingRows
                EntryProtection.AllowInsertingColumns = .AllowInsertingColumns
                EntryProtection.AllowInsertingRows = .AllowInsertingRows
                EntryProtection.AllowInsertingHyperlinks = .AllowInsertingHyperlinks
                EntryProtection.AllowDeletingColumns = .AllowDeletingColumns
                EntryProtection.AllowDeletingRows = .AllowDeletingRows
                EntryProtection.AllowSorting = .AllowSorting
                EntryProtection.AllowFiltering = .AllowFiltering
                EntryProtection.AllowUsingPivotTables = .AllowUsingPivotTables
            End With
        End If

'------------------------------------------------------------------------------
' OPTIONAL UNPROTECT
'------------------------------------------------------------------------------
    'Temporarily unprotect the sheet when protection is active
        If WasProtected Then
            If Len(ProtectPassword) > 0 Then
                WS.Unprotect Password:=ProtectPassword
                DidUnprotect = True
            Else
                Err.Raise vbObjectError + 2100, _
                          "M_DEMO_BUILDER.DEMO_Sheet_Reset", _
                          "Worksheet is protected and no password was supplied."
            End If
        End If

'------------------------------------------------------------------------------
' REMOVE SHAPES
'------------------------------------------------------------------------------
    'Delete all shapes currently present on the worksheet when requested
        If DeleteShapes Then
            For i = WS.Shapes.Count To 1 Step -1
                WS.Shapes(i).Delete
            Next i
        End If

'------------------------------------------------------------------------------
' REMOVE TABLES
'------------------------------------------------------------------------------
    'Delete all ListObjects currently present on the worksheet when requested
        If DeleteTables Then
            For i = WS.ListObjects.Count To 1 Step -1
                WS.ListObjects(i).Delete
            Next i
        End If

'------------------------------------------------------------------------------
' UNMERGE CELLS
'------------------------------------------------------------------------------
    'Remove any existing merges when requested
        If UnmergeCells Then
            WS.Cells.UnMerge
        End If

'------------------------------------------------------------------------------
' CLEAR VALIDATION
'------------------------------------------------------------------------------
    'Delete any data-validation rules from the sheet when requested
        If ClearValidation Then
            On Error Resume Next
            WS.Cells.Validation.Delete
            On Error GoTo CleanFail
        End If

'------------------------------------------------------------------------------
' CLEAR CELLS
'------------------------------------------------------------------------------
    'Clear values, formulas, formats, and general cell state when requested
        If ClearCells Then
            WS.Cells.Clear
        End If

'------------------------------------------------------------------------------
' CLEAR COMMENTS / NOTES
'------------------------------------------------------------------------------
    'Explicitly clear comments and notes when requested
        If ClearCommentsAndNotes Then
            On Error Resume Next
            WS.Cells.ClearComments
            WS.Cells.ClearNotes
            On Error GoTo CleanFail
        End If

'------------------------------------------------------------------------------
' RESET CELL TEXT LAYOUT
'------------------------------------------------------------------------------
    'Remove word wrap from all cells when requested
        If ResetWrapText Then
            WS.Cells.WrapText = False
        End If

'------------------------------------------------------------------------------
' UNHIDE ROWS / COLUMNS
'------------------------------------------------------------------------------
    'Unhide all worksheet rows when requested
        If UnhideRows Then
            WS.Rows.Hidden = False
        End If

    'Unhide all worksheet columns when requested
        If UnhideColumns Then
            WS.Columns.Hidden = False
        End If

'------------------------------------------------------------------------------
' RESET VIEW-LIKE SHEET PROPERTIES
'------------------------------------------------------------------------------
    'Reset column widths to a neutral default when requested
        If ResetColumnWidths Then
            WS.Cells.ColumnWidth = 8.43
        End If

    'Reset row heights to a neutral default when requested
        If ResetRowHeights Then
            WS.Cells.RowHeight = 15
        End If

    'Reset the worksheet tab color when requested
        If ResetTabColor Then
            WS.Tab.ColorIndex = xlColorIndexNone
        End If

'------------------------------------------------------------------------------
' REMOVE PAGE BREAKS
'------------------------------------------------------------------------------
    'Remove page breaks on a best-effort basis when requested
        If RemovePageBreaks Then
            On Error Resume Next

        'Remove any manual page breaks; Excel lays the sheet out through
        'the printer driver to do it
            DEMO_CleanupProbeHit DEMO_PROBE_PAGE_BREAKS
            WS.ResetAllPageBreaks

            On Error GoTo CleanFail

        'Hide page-break indicators, writing the property only when shown
            DEMO_Sheet_HidePageBreaks WS
        End If

'------------------------------------------------------------------------------
' REGRESSION FAULT SEAM
'------------------------------------------------------------------------------
    'Raise the armed one-shot operation fault; inert unless a harness armed it
        DEMO_CleanupOperationFault

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanExit:
    'Cleanup is single-pass and only restores an unprotect owned by this call.
        On Error Resume Next
        DEMO_CleanupProbeHit DEMO_PROBE_RESET_CLEANUP
        If DidUnprotect And ReProtectAtEnd Then
            Err.Clear
            WS.Protect Password:=DEMO_CleanupProtectPassword(ProtectPassword), _
                DrawingObjects:=EntryProtection.DrawingObjects, _
                Contents:=EntryProtection.Contents, _
                Scenarios:=EntryProtection.Scenarios, _
                UserInterfaceOnly:=EntryProtection.UserInterfaceOnly, _
                AllowFormattingCells:=EntryProtection.AllowFormattingCells, _
                AllowFormattingColumns:=EntryProtection.AllowFormattingColumns, _
                AllowFormattingRows:=EntryProtection.AllowFormattingRows, _
                AllowInsertingColumns:=EntryProtection.AllowInsertingColumns, _
                AllowInsertingRows:=EntryProtection.AllowInsertingRows, _
                AllowInsertingHyperlinks:=EntryProtection.AllowInsertingHyperlinks, _
                AllowDeletingColumns:=EntryProtection.AllowDeletingColumns, _
                AllowDeletingRows:=EntryProtection.AllowDeletingRows, _
                AllowSorting:=EntryProtection.AllowSorting, _
                AllowFiltering:=EntryProtection.AllowFiltering, _
                AllowUsingPivotTables:=EntryProtection.AllowUsingPivotTables
            CleanupErrNumber = Err.Number
            CleanupErrSource = Err.Source
            CleanupErrDescription = Err.Description

            Err.Clear
            WS.EnableSelection = DEMO_CleanupSelectionValue(EntryProtection.EnableSelection)
            SelectionErrNumber = Err.Number
            SelectionErrSource = Err.Source
            SelectionErrDescription = Err.Description
        End If
        On Error GoTo 0

    'Retain both cleanup diagnostics without replacing the operation error.
        If SelectionErrNumber <> 0 Then
            If CleanupErrNumber = 0 Then
                CleanupErrNumber = SelectionErrNumber
                CleanupErrSource = SelectionErrSource
                CleanupErrDescription = SelectionErrDescription
            Else
                CleanupErrDescription = CleanupErrDescription & _
                    " | EnableSelection: " & CStr(SelectionErrNumber) & _
                    " | " & SelectionErrSource & " | " & SelectionErrDescription
            End If
        End If
        If CleanupErrNumber <> 0 Then
            If SavedErrNumber = 0 Then
                SavedErrNumber = CleanupErrNumber
                SavedErrSource = CleanupErrSource
                SavedErrDescription = CleanupErrDescription
            Else
                SavedErrDescription = SavedErrDescription & vbCrLf & _
                    "Cleanup failure: " & CStr(CleanupErrNumber) & _
                    " | " & CleanupErrSource & " | " & CleanupErrDescription
            End If
        End If

    'Re-raise the captured error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CleanFail:
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description

    'Count the handler entry for the regression probe
        DEMO_CleanupProbeHit DEMO_PROBE_RESET_HANDLER

    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub

Private Sub DEMO_Sheet_HidePageBreaks( _
    ByVal WS As Worksheet)
'
'------------------------------------------------------------------------------
'                          HIDE PAGE-BREAK INDICATORS
'------------------------------------------------------------------------------
' PURPOSE
'   Hides the page-break indicators of a worksheet when they are shown.
'
' WHY THIS EXISTS
'   Writing Worksheet.DisplayPageBreaks can make Excel lay the sheet out
'   through the default printer driver, even when the value does not change,
'   and a slow, offline or network printer stalls that write. A new worksheet
'   starts with the indicators hidden, so the write is needed only when they
'   are shown (#187).
'
' INPUTS
'   WS
'     Worksheet whose indicators are hidden
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Best effort: does not raise. A read that fails leaves the indicators as
'   they are, and nothing is written.
'
' REGRESSION SEAMS
'   While M_DEMO_CLEANUP_TESTS has armed the cleanup probe, each write is
'   counted before it is attempted.
'
' HISTORY
'   2026-11-15 - Added for #187.
'
' UPDATED
'   2026-11-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Sheet_HidePageBreaks"

    Dim Shown   As Boolean  'Indicators shown on entry

'------------------------------------------------------------------------------
' HIDE
'------------------------------------------------------------------------------
    'Read first; a failed read leaves Shown False, so nothing is written
        On Error Resume Next
        Shown = WS.DisplayPageBreaks
        If Shown Then
            DEMO_CleanupProbeHit DEMO_PROBE_DISPLAY_BREAKS
            WS.DisplayPageBreaks = False
        End If
        On Error GoTo 0

End Sub

Public Sub DEMO_InternalCleanupProbeBegin( _
    ByVal FaultMask As Long)
'
'------------------------------------------------------------------------------
'                              ARM CLEANUP PROBE
'------------------------------------------------------------------------------
' PURPOSE
'   Zeroes the cleanup probe counters, starts counting and arms one-shot
'   cleanup faults for the next DEMO_Sheet_Reset / DEMO_Sheet_BuildTemplate
'   calls.
'
' WHY THIS EXISTS
'   The single-pass cleanup of both procedures is only observable from inside
'   them: a caller cannot tell how often a handler ran, and no worksheet
'   condition makes Protect or the EnableSelection restoration fail on demand.
'   Without this seam the #71 failure paths could be reasoned about but never
'   executed in Excel.
'
' INPUTS
'   FaultMask
'     Sum of the fault bits to arm, or 0 for counting only
'       1  DEMO_Sheet_Reset raises an operation error after its protection
'          steps, inside its guarded body
'       2  The cleanup Protect statement fails while its arguments are
'          evaluated, so Excel's Protect is not called
'       4  The cleanup EnableSelection restoration fails the same way
'       8  The DisplayFormulaBar write of DEMO_Sheet_BuildTemplate fails while
'          its value is evaluated
'       16 The rename of a sheet DEMO_Sheet_GetOrCreate has just created
'          fails the same way
'       256 times a DEMO_FM_* setting bit (1, 2, 4, 8, 16)
'          The fast-mode entry write of that setting fails
'       8192 times a DEMO_FM_* setting bit
'          The fast-mode restoration write of that setting fails, in
'          DEMO_FastMode_End or in the rollback of DEMO_FastMode_Begin
'     Unrecognised bits are stored and never fire.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Each armed fault fires once and is then cleared
'   - Counting continues until DEMO_InternalCleanupProbeEnd
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   - DEMO_PROBE_SITE_COUNT
'   - mCleanupProbeCounts, mCleanupFaults and mCleanupProbeArmed
'
' NOTES
'   - Public only so another module of the demo project can reach it. Option
'     Private Module keeps it out of other VBA projects. Demo code must never
'     call it.
'
' CALLED FROM
'   - M_DEMO_CLEANUP_TESTS
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_InternalCleanupProbeBegin"

    Dim Site    As Long 'Probe-site iterator

'------------------------------------------------------------------------------
' SET SEAM
'------------------------------------------------------------------------------
    'Zero every probe counter
        For Site = 1 To DEMO_PROBE_SITE_COUNT
            mCleanupProbeCounts(Site) = 0
        Next Site

    'Arm counting and the requested one-shot faults
        mCleanupFaults = FaultMask
        mCleanupProbeArmed = True

End Sub

Public Function DEMO_InternalCleanupProbeCount( _
    ByVal Site As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                           READ CLEANUP PROBE COUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns how often one cleanup probe site was reached since the last
'   DEMO_InternalCleanupProbeBegin.
'
' WHY THIS EXISTS
'   The counters are module-private. M_DEMO_CLEANUP_TESTS reads them through
'   this routine to check how often each cleanup site ran (#71 harness).
'
' INPUTS
'   Site
'     1  DEMO_Sheet_Reset handler entries
'     2  DEMO_Sheet_Reset cleanup passes
'     3  DEMO_Sheet_Reset cleanup Protect statements reached
'     4  DEMO_Sheet_Reset cleanup EnableSelection restorations reached
'     5  DEMO_Sheet_BuildTemplate handler entries
'     6  DEMO_Sheet_BuildTemplate cleanup passes
'     7  Fast-mode entry writes reached
'     8  Fast-mode restoration writes reached
'     9  DEMO_Sheet_Reset ResetAllPageBreaks calls reached
'     10 DisplayPageBreaks writes reached (DEMO_Sheet_HidePageBreaks)
'
' RETURNS
'   Long
'     The count, or -1 for an unknown site
'
' BEHAVIOR
'   - Returns the counter of Site when it lies between 1 and
'     DEMO_PROBE_SITE_COUNT
'   - Returns -1 for any other Site
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   - mCleanupProbeCounts
'
' NOTES
'   - Counters change only while the probe is armed, and stay readable after
'     DEMO_InternalCleanupProbeEnd
'
' CALLED FROM
'   - M_DEMO_CLEANUP_TESTS
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the counter of a known site; report an unknown site as -1
        If Site >= 1 And Site <= DEMO_PROBE_SITE_COUNT Then
            DEMO_InternalCleanupProbeCount = mCleanupProbeCounts(Site)
        Else
            DEMO_InternalCleanupProbeCount = -1
        End If

End Function

Public Sub DEMO_InternalCleanupProbeEnd()

'
'------------------------------------------------------------------------------
'                             DISARM CLEANUP PROBE
'------------------------------------------------------------------------------
' PURPOSE
'   Stops counting and disarms every cleanup fault that has not fired.
'
' WHY THIS EXISTS
'   A fault left armed would fire inside an unrelated demo build later in the
'   same session. The counters are kept so the last run can still be read.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Clears the armed flag, so the probe sites stop counting
'   - Clears every pending fault bit
'   - Leaves the counters as they are
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   - mCleanupProbeArmed and mCleanupFaults
'
' NOTES
'   - Call it after a probe run; the counters stay readable through
'     DEMO_InternalCleanupProbeCount
'
' CALLED FROM
'   - M_DEMO_CLEANUP_TESTS
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR SEAM
'------------------------------------------------------------------------------
    'Disarm counting and every pending fault
        mCleanupProbeArmed = False
        mCleanupFaults = 0

End Sub

Private Sub DEMO_CleanupProbeHit( _
    ByVal Site As Long)
'
'------------------------------------------------------------------------------
'                           COUNT CLEANUP PROBE HIT
'------------------------------------------------------------------------------
' PURPOSE
'   Counts one pass through a cleanup probe site while the probe is armed.
'
' INPUTS
'   Site
'     One of the DEMO_PROBE_* sites
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Does not raise and does not touch Err, so it is safe inside handlers and
'   cleanup scopes
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Count only while armed and only for a known site
        If mCleanupProbeArmed Then
            If Site >= 1 And Site <= DEMO_PROBE_SITE_COUNT Then
                mCleanupProbeCounts(Site) = mCleanupProbeCounts(Site) + 1
            End If
        End If

End Sub

Private Sub DEMO_CleanupOperationFault()

'
'------------------------------------------------------------------------------
'                         RAISE ARMED OPERATION FAULT
'------------------------------------------------------------------------------
' PURPOSE
'   Raises the armed DEMO_Sheet_Reset operation fault once, and disarms it.
'
' WHY THIS EXISTS
'   Once the sheet is unprotected, no worksheet condition makes the reset body
'   fail on demand. This point stands in for any such failure inside the
'   guarded body.
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Raises vbObjectError + 2190 once when the operation fault is armed;
'   otherwise does nothing. The caller's handler receives it as an ordinary
'   operation failure.
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RAISE ARMED FAULT
'------------------------------------------------------------------------------
    'Consume the fault before raising so it fires exactly once
        If (mCleanupFaults And DEMO_FAULT_RESET_BODY) <> 0 Then
            mCleanupFaults = mCleanupFaults And Not DEMO_FAULT_RESET_BODY
            Err.Raise vbObjectError + 2190, _
                      "M_DEMO_BUILDER.DEMO_Sheet_Reset.InjectedOperationFault", _
                      "Injected operation fault after the protection steps."
        End If

End Sub

Private Function DEMO_CleanupProtectPassword( _
    ByVal Password As String) _
    As String
'
'------------------------------------------------------------------------------
'                       SUPPLY CLEANUP PROTECT PASSWORD
'------------------------------------------------------------------------------
' PURPOSE
'   Supplies the Password argument of the cleanup Protect statement.
'
' WHY THIS EXISTS
'   No worksheet condition makes the reprotection fail on demand. Raising
'   while the Protect arguments are evaluated fails that exact statement, so
'   its immediate error capture runs unchanged and Excel's Protect is not
'   called.
'
' INPUTS
'   Password
'     The password supplied to DEMO_Sheet_Reset
'
' RETURNS
'   String
'     Password, unchanged
'
' ERROR POLICY
'   Raises vbObjectError + 2191 once when the reprotect fault is armed
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COUNT AND RAISE ARMED FAULT
'------------------------------------------------------------------------------
    'Count the Protect statement and fail it once when armed
        DEMO_CleanupProbeHit DEMO_PROBE_RESET_PROTECT

        If (mCleanupFaults And DEMO_FAULT_RESET_PROTECT) <> 0 Then
            mCleanupFaults = mCleanupFaults And Not DEMO_FAULT_RESET_PROTECT
            Err.Raise vbObjectError + 2191, _
                      "M_DEMO_BUILDER.DEMO_Sheet_Reset.InjectedProtectFault", _
                      "Injected reprotection fault."
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Pass the password through unchanged
        DEMO_CleanupProtectPassword = Password

End Function

Private Function DEMO_CleanupSelectionValue( _
    ByVal SelectionPolicy As Long) _
    As Long
'
'------------------------------------------------------------------------------
'                        SUPPLY CLEANUP SELECTION VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Supplies the value of the cleanup EnableSelection restoration.
'
' WHY THIS EXISTS
'   As for DEMO_CleanupProtectPassword: raising while the value is evaluated
'   fails that exact assignment without a worksheet condition.
'
' INPUTS
'   SelectionPolicy
'     The EnableSelection value captured on entry
'
' RETURNS
'   Long
'     SelectionPolicy, unchanged
'
' ERROR POLICY
'   Raises vbObjectError + 2192 once when the selection fault is armed
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COUNT AND RAISE ARMED FAULT
'------------------------------------------------------------------------------
    'Count the restoration and fail it once when armed
        DEMO_CleanupProbeHit DEMO_PROBE_RESET_SELECTION

        If (mCleanupFaults And DEMO_FAULT_RESET_SELECTION) <> 0 Then
            mCleanupFaults = mCleanupFaults And Not DEMO_FAULT_RESET_SELECTION
            Err.Raise vbObjectError + 2192, _
                      "M_DEMO_BUILDER.DEMO_Sheet_Reset.InjectedSelectionFault", _
                      "Injected EnableSelection restoration fault."
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Pass the selection policy through unchanged
        DEMO_CleanupSelectionValue = SelectionPolicy

End Function

Private Function DEMO_FormulaBarValue( _
    ByVal ShowFormulaBar As Boolean) _
    As Boolean
'
'------------------------------------------------------------------------------
'                           SUPPLY FORMULA-BAR VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Supplies the value of the DisplayFormulaBar write of
'   DEMO_Sheet_BuildTemplate.
'
' WHY THIS EXISTS
'   No host condition makes that write fail on demand. Raising while the value
'   is evaluated fails that exact statement inside the guarded body, so the
'   template's cleanup path runs unchanged and the formula bar is not touched.
'
' RETURNS
'   Boolean
'     ShowFormulaBar, unchanged
'
' ERROR POLICY
'   Raises vbObjectError + 2195 once when the formula-bar fault is armed
'
' HISTORY
'   2026-11-14 - Added for #70.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RAISE ARMED FAULT
'------------------------------------------------------------------------------
    'Fail the write once when armed
        If (mCleanupFaults And DEMO_FAULT_BUILD_FORMULABAR) <> 0 Then
            mCleanupFaults = mCleanupFaults And Not DEMO_FAULT_BUILD_FORMULABAR
            Err.Raise vbObjectError + 2195, _
                      "M_DEMO_BUILDER.DEMO_Sheet_BuildTemplate.InjectedFormulaBarFault", _
                      "Injected DisplayFormulaBar fault."
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Pass the value through unchanged
        DEMO_FormulaBarValue = ShowFormulaBar

End Function

Public Sub DEMO_Format_Labels( _
    ByVal TargetRange As Range)
'
'------------------------------------------------------------------------------
'                                FORMAT LABELS
'------------------------------------------------------------------------------
' PURPOSE
'   Applies a standard label format to a target range.
'
' WHY THIS EXISTS
'   DEMO sheets often contain repeated label blocks that should share a common
'   visual treatment.
'
' INPUTS
'   TargetRange
'     Range to format as a label block.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits quietly if TargetRange Is Nothing
'   - Applies fill, alignment, and bold formatting
'   - Applies borders
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - DEMO_Set_RangeBorder
'   - Range.Interior, Font, VerticalAlignment and HorizontalAlignment
'
' NOTES
'   - Colors come from COLOR_LABEL_BACK and COLOR_LABEL_FONT; the border is
'     black, thin and includes the inside lines
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no target range was supplied
        If TargetRange Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' APPLY FORMAT
'------------------------------------------------------------------------------
    'Apply the standard label formatting
        With TargetRange
            .Interior.Color = COLOR_LABEL_BACK
            .Font.Color = COLOR_LABEL_FONT
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
            .Font.Bold = True
        End With

'------------------------------------------------------------------------------
' APPLY BORDER
'------------------------------------------------------------------------------
    'Apply borders to the label block
        DEMO_Set_RangeBorder TargetRange, RGB(0, 0, 0), xlThin, True

End Sub

Public Sub DEMO_Format_InputCell( _
    ByVal TargetRange As Range)
'
'------------------------------------------------------------------------------
'                              FORMAT INPUT CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Formats one or more DEMO control-input cells consistently.
'
' WHY THIS EXISTS
'   DEMO sheets should clearly distinguish user-editable control cells from
'   labels and output areas.
'
' INPUTS
'   TargetRange
'     Range of input cells to format.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits quietly if TargetRange Is Nothing
'   - Applies light input fill
'   - Applies centered alignment
'   - Applies borders
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - DEMO_Set_RangeBorder
'   - Range.Interior, VerticalAlignment and HorizontalAlignment
'
' NOTES
'   - The fill is COLOR_INPUT; the border is black, thin and includes the
'     inside lines
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no target range was supplied
        If TargetRange Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' APPLY FORMAT
'------------------------------------------------------------------------------
    'Apply the standard control-input formatting
        With TargetRange
            .Interior.Color = COLOR_INPUT
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With

'------------------------------------------------------------------------------
' APPLY BORDER
'------------------------------------------------------------------------------
    'Apply borders to the input range
        DEMO_Set_RangeBorder TargetRange, RGB(0, 0, 0), xlThin, True

End Sub

Public Sub DEMO_Format_OutputCell( _
    ByVal TargetRange As Range)
'
'------------------------------------------------------------------------------
'                              FORMAT OUTPUT CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Formats one or more output/result cells consistently.
'
' WHY THIS EXISTS
'   DEMO sheets often contain read-only output or result areas that should be
'   visually distinct from labels and editable input cells.
'
' INPUTS
'   TargetRange
'     Range of output cells to format.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits quietly if TargetRange Is Nothing
'   - Applies a neutral output style
'   - Applies borders
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - DEMO_Set_RangeBorder
'   - Range.Interior, Font, VerticalAlignment and HorizontalAlignment
'
' NOTES
'   - The style is no fill, black regular font and left alignment; the border
'     is black, thin and includes the inside lines
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no target range was supplied
        If TargetRange Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' APPLY FORMAT
'------------------------------------------------------------------------------
    'Apply the standard output formatting
        With TargetRange
            .Interior.Pattern = xlNone
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlLeft
        End With

'------------------------------------------------------------------------------
' APPLY BORDER
'------------------------------------------------------------------------------
    'Apply borders to the output range
        DEMO_Set_RangeBorder TargetRange, RGB(0, 0, 0), xlThin, True

End Sub

Public Sub DEMO_Write_BandHeader( _
    ByVal TargetRange As Range, _
    ByVal HeaderText As String, _
    Optional ByVal FillColor As Long = 6568980, _
    Optional ByVal FontColor As Long = vbWhite, _
    Optional ByVal FontSize As Double = 11, _
    Optional ByVal IsBold As Boolean = True, _
    Optional ByVal IsBorder As Boolean = True, _
    Optional ByVal IsCentered As Boolean = True)
'
'------------------------------------------------------------------------------
'                              WRITE BAND HEADER
'------------------------------------------------------------------------------
' PURPOSE
'   Writes and formats one header band without merged cells.
'
' WHY THIS EXISTS
'   DEMO sheets often use title, subtitle, or banner ranges spanning multiple
'   columns. This helper applies a consistent visual band while preserving the
'   underlying cell structure.
'
' INPUTS
'   TargetRange
'     Single-row range to format as a header band.
'
'   HeaderText
'     Text to display in the band.
'
'   FillColor (optional)
'     Background fill color for the band.
'
'   FontColor (optional)
'     Font color for the band text.
'
'   FontSize (optional)
'     Font size for the band text.
'
'   IsBold (optional)
'     TRUE  => bold text
'     FALSE => regular text
'
'   IsBorder (optional)
'     TRUE  => apply an outside border
'     FALSE => do not apply a border
'
'   IsCentered (optional)
'     TRUE  => center the text across the band (the default)
'     FALSE => align the text to the left
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits quietly if TargetRange Is Nothing
'   - Ensures the range is not merged
'   - Clears prior contents from the full target range
'   - Writes the requested text into the first cell only
'   - Applies fill / font formatting
'   - Aligns the text to the left, or centers it across the band when
'     IsCentered is TRUE
'   - Optionally applies a border
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - DEMO_Set_RangeBorder
'   - Range.UnMerge, ClearContents, Interior, Font, HorizontalAlignment and
'     VerticalAlignment
'
' NOTES
'   - The band is never merged; centering uses xlCenterAcrossSelection
'   - The border is outside only: DEMO_Set_RangeBorder is called with
'     IncludeInside FALSE
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no target range was supplied
        If TargetRange Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' RESET RANGE STATE
'------------------------------------------------------------------------------
    'Ensure the target range is not merged
        TargetRange.UnMerge
    'Clear any prior contents from the full target range
        TargetRange.ClearContents

'------------------------------------------------------------------------------
' WRITE HEADER TEXT
'------------------------------------------------------------------------------
    'Write the header text into the first cell only
        TargetRange.Cells(1, 1).Value = HeaderText

'------------------------------------------------------------------------------
' APPLY FORMAT
'------------------------------------------------------------------------------
    'Apply the requested band formatting
        With TargetRange
            .Interior.Color = FillColor
            .Font.Color = FontColor
            .Font.Bold = IsBold
            .Font.Size = FontSize
            If IsCentered Then
                .HorizontalAlignment = xlCenterAcrossSelection
            Else
                .HorizontalAlignment = xlLeft
            End If
            .VerticalAlignment = xlCenter
        End With

'------------------------------------------------------------------------------
' APPLY BORDER
'------------------------------------------------------------------------------
    'Apply the border when requested
        If IsBorder Then DEMO_Set_RangeBorder TargetRange, RGB(0, 0, 0), xlThin, False

End Sub

Public Sub DEMO_Apply_ValidationList( _
    ByVal TargetCell As Range, _
    ByVal ListText As String)
'
'------------------------------------------------------------------------------
'                            APPLY VALIDATION LIST
'------------------------------------------------------------------------------
' PURPOSE
'   Applies an in-cell dropdown validation list to a target cell or range.
'
' WHY THIS EXISTS
'   DEMO-sheet inputs are easier to use and less error-prone when users can pick
'   values from a controlled dropdown instead of typing freely.
'
' INPUTS
'   TargetCell
'     Cell or range that should receive the validation list.
'
'   ListText
'     Validation list source.
'     This can be:
'       - a comma-separated literal list, or
'       - a formula such as =BoolList
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits quietly if TargetCell Is Nothing
'   - Rejects a blank ListText
'   - Deletes any prior validation
'   - Adds a list-type validation
'   - Enables the dropdown, input message, and error alert
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - Range.Validation: Delete, Add, IgnoreBlank, InCellDropdown, ShowInput
'     and ShowError
'
' NOTES
'   - ListText is passed to the validation as Formula1 without change; the
'     alert style is xlValidAlertStop
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no target range was supplied
        If TargetCell Is Nothing Then Exit Sub

    'Reject a blank validation source
        If Len(Trim$(ListText)) = 0 Then
            Err.Raise vbObjectError + 2003, _
                      "M_DEMO_BUILDER.DEMO_Apply_ValidationList", _
                      "Validation list source cannot be blank."
        End If

'------------------------------------------------------------------------------
' RESET VALIDATION
'------------------------------------------------------------------------------
    'Delete any prior validation from the target range
        TargetCell.Validation.Delete

'------------------------------------------------------------------------------
' APPLY VALIDATION
'------------------------------------------------------------------------------
    'Apply the requested dropdown validation list
        TargetCell.Validation.Add Type:=xlValidateList, _
                                  AlertStyle:=xlValidAlertStop, _
                                  Operator:=xlBetween, _
                                  Formula1:=ListText

'------------------------------------------------------------------------------
' CONFIGURE VALIDATION
'------------------------------------------------------------------------------
    'Allow blanks
        TargetCell.Validation.IgnoreBlank = True

    'Show the in-cell dropdown arrow
        TargetCell.Validation.InCellDropdown = True

    'Enable the input message
        TargetCell.Validation.ShowInput = True

    'Enable the error alert
        TargetCell.Validation.ShowError = True

End Sub

Public Sub DEMO_Apply_NumericValidation( _
    ByVal Target As Range, _
    ByVal MinValue As Double, _
    ByVal MaxValue As Double, _
    Optional ByVal AllowDecimals As Boolean = True)
'
'------------------------------------------------------------------------------
'                           APPLY NUMERIC VALIDATION
'------------------------------------------------------------------------------
' PURPOSE
'   Applies a Data Validation rule requiring numeric values between MinValue
'   and MaxValue inclusive.
'
' WHY THIS EXISTS
'   Numeric validation is often needed in setup and control-panel ranges.
'   This helper centralizes the logic so callers can choose:
'     - decimal validation, or
'     - whole-number validation
'
'   using one reusable routine.
'
' INPUTS
'   Target
'     The target cell or range to validate.
'
'   MinValue
'     Minimum accepted numeric value (inclusive).
'
'   MaxValue
'     Maximum accepted numeric value (inclusive).
'
'   AllowDecimals (optional)
'     TRUE  => allow decimal values
'     FALSE => allow whole numbers only
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits quietly if Target Is Nothing
'   - Swaps MinValue / MaxValue if they are passed in reversed order
'   - Removes any existing validation from the target
'   - Applies either:
'       * xlValidateDecimal, or
'       * xlValidateWholeNumber
'   - Enables input and error alerts
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - Excel Range.Validation object model
'
' NOTES
'   - The bounds are inclusive
'   - When AllowDecimals = False, decimal values are rejected
'   - Formula1 / Formula2 are passed as strings to the Validation object
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Apply_NumericValidation"

    Dim Lo              As Double   'Normalized lower bound
    Dim Hi              As Double   'Normalized upper bound
    Dim ValidationType  As Long     'Validation type to apply

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no target range was supplied
        If Target Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' NORMALIZE BOUNDS
'------------------------------------------------------------------------------
    'Store the lower and upper bounds in ascending order
        If MinValue <= MaxValue Then
            Lo = MinValue
            Hi = MaxValue
        Else
            Lo = MaxValue
            Hi = MinValue
        End If

'------------------------------------------------------------------------------
' SELECT VALIDATION TYPE
'------------------------------------------------------------------------------
    'Choose decimal or whole-number validation
        If AllowDecimals Then
            ValidationType = xlValidateDecimal
        Else
            ValidationType = xlValidateWholeNumber
        End If

'------------------------------------------------------------------------------
' RESET VALIDATION
'------------------------------------------------------------------------------
    'Remove any existing validation rule from the target
        Target.Validation.Delete

'------------------------------------------------------------------------------
' APPLY VALIDATION
'------------------------------------------------------------------------------
    'Apply numeric validation between the normalized bounds
        Target.Validation.Add Type:=ValidationType, _
                              AlertStyle:=xlValidAlertStop, _
                              Operator:=xlBetween, _
                              Formula1:=CStr(Lo), _
                              Formula2:=CStr(Hi)

'------------------------------------------------------------------------------
' CONFIGURE ALERTS
'------------------------------------------------------------------------------
    'Allow blanks
        Target.Validation.IgnoreBlank = True

    'Enable the input message
        Target.Validation.ShowInput = True

    'Enable the error alert
        Target.Validation.ShowError = True

End Sub

Public Sub DEMO_Create_BoolList( _
    Optional ByVal WS_Name As String = DEFAULT_BOOL_WS)
'
'------------------------------------------------------------------------------
'                               CREATE BOOL LIST
'------------------------------------------------------------------------------
' PURPOSE
'   Creates a workbook-level named range "BoolList" containing the literal text
'   values "TRUE" and "FALSE" for use in Data Validation lists.
'
' WHY THIS EXISTS
'   Excel logical values are localized by host language settings:
'     - TRUE / FALSE
'     - VERO / FALSO
'     - etc.
'
'   Therefore a validation list built from Boolean formulas is not stable across
'   international environments.
'
'   This routine avoids that problem by storing the values as plain text.
'
' INPUTS
'   WS_Name (optional)
'     Name of the worksheet where the helper list should be created.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Uses ThisWorkbook as the target workbook
'   - Gets or creates the requested worksheet
'   - Writes the helper-list label in AA1
'   - Writes the literal text values "TRUE" and "FALSE" in AA2:AA3
'   - Formats the helper cells as text
'   - Recreates the workbook-level name "BoolList"
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - DEMO_Sheet_GetOrCreate
'   - DEMO_Hide_HelperColumns
'
' NOTES
'   - This creates a text list, not a Boolean list
'   - That is intentional so the dropdown remains locale-independent
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Create_BoolList"

    Dim WB  As Workbook     'Target workbook
    Dim WS  As Worksheet    'Target worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Use the workbook that contains this module
        Set WB = ThisWorkbook

    'Get or create the target worksheet
        Set WS = DEMO_Sheet_GetOrCreate(WB, WS_Name)

'------------------------------------------------------------------------------
' WRITE HELPER LIST
'------------------------------------------------------------------------------
    'Write the helper-list label
        WS.Range("AA1").Value = DEFAULT_BOOLLIST_NAME

    'Format the helper-list cells as text
        WS.Range("AA2:AA3").NumberFormat = "@"

    'Write the literal TRUE/FALSE text values
        WS.Range("AA2").Value = "TRUE"
        WS.Range("AA3").Value = "FALSE"

'------------------------------------------------------------------------------
' RECREATE WORKBOOK NAME
'------------------------------------------------------------------------------
    'Delete the existing workbook-level name if already present
        On Error Resume Next
        WB.Names(DEFAULT_BOOLLIST_NAME).Delete
        On Error GoTo 0

    'Create the workbook-level name pointing to the helper-list range
        WB.Names.Add Name:=DEFAULT_BOOLLIST_NAME, _
                     RefersTo:="='" & WS.Name & "'!$AA$2:$AA$3"

'------------------------------------------------------------------------------
' HIDE HELPER AREA
'------------------------------------------------------------------------------
    'Hide the helper columns used to store the list
        DEMO_Hide_HelperColumns WS, "AA:AZ"

End Sub

Public Sub DEMO_Apply_BoolValidation( _
    ByVal Target As Range, _
    Optional ByVal WS_Name As String = DEFAULT_BOOL_WS, _
    Optional ByVal ListName As String = DEFAULT_BOOLLIST_NAME)
'
'------------------------------------------------------------------------------
'                            APPLY BOOL VALIDATION
'------------------------------------------------------------------------------
' PURPOSE
'   Applies a locale-independent TRUE/FALSE validation list to a target range.
'
' WHY THIS EXISTS
'   Boolean dropdowns built from native Excel logical values can be localized by
'   host language. This helper ensures the validation list is based on a stable
'   text list instead.
'
' INPUTS
'   Target
'     Target cell or range that should receive the TRUE/FALSE validation list.
'
'   WS_Name (optional)
'     Worksheet where the helper list should be created or maintained.
'
'   ListName (optional)
'     Workbook-level list name to use in the validation formula.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Ensures the helper list exists
'   - Applies list validation using the workbook-level name
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - DEMO_Create_BoolList
'   - DEMO_Apply_ValidationList
'
' NOTES
'   - The list name is created by DEMO_Create_BoolList, which always creates
'     BoolList; a different ListName must already exist in the workbook
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no target range was supplied
        If Target Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' ENSURE BOOL LIST
'------------------------------------------------------------------------------
    'Ensure the workbook-level TRUE/FALSE helper list exists
        DEMO_Create_BoolList WS_Name

'------------------------------------------------------------------------------
' APPLY VALIDATION
'------------------------------------------------------------------------------
    'Apply validation using the workbook-level BoolList name
        DEMO_Apply_ValidationList Target, "=" & ListName

End Sub

Public Sub DEMO_Set_WorkbookName( _
    ByVal WB As Workbook, _
    ByVal NameText As String, _
    ByVal TargetCell As Range)
'
'------------------------------------------------------------------------------
'                              SET WORKBOOK NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Creates or refreshes a workbook-level defined name pointing to one cell.
'
' WHY THIS EXISTS
'   DEMO workbooks use workbook-level names so action macros can read
'   control-panel settings without hard-coding cell addresses repeatedly.
'
' INPUTS
'   WB
'     Target workbook.
'
'   NameText
'     Workbook-level name to create or refresh.
'
'   TargetCell
'     Single target cell that the name should refer to.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Deletes the existing workbook-level name when present
'   - Recreates the name to point to the requested target cell
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - Workbook.Names: Delete and Add
'   - Range.Parent.Name and Range.Address
'
' NOTES
'   - A missing existing name is ignored: the delete runs under
'     On Error Resume Next
'   - The name refers to the sheet-qualified address of TargetCell; only
'     Nothing is checked, not the cell count
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing workbook reference
        If WB Is Nothing Then
            Err.Raise vbObjectError + 2004, _
                      "M_DEMO_BUILDER.DEMO_Set_WorkbookName", _
                      "Workbook reference cannot be Nothing."
        End If

    'Reject a blank name
        If Len(Trim$(NameText)) = 0 Then
            Err.Raise vbObjectError + 2005, _
                      "M_DEMO_BUILDER.DEMO_Set_WorkbookName", _
                      "Workbook name text cannot be blank."
        End If

    'Reject a missing target cell
        If TargetCell Is Nothing Then
            Err.Raise vbObjectError + 2006, _
                      "M_DEMO_BUILDER.DEMO_Set_WorkbookName", _
                      "Target cell cannot be Nothing."
        End If

'------------------------------------------------------------------------------
' DELETE EXISTING NAME
'------------------------------------------------------------------------------
    'Delete the existing workbook-level name when present
        On Error Resume Next
        WB.Names(NameText).Delete
        On Error GoTo 0

'------------------------------------------------------------------------------
' CREATE NAME
'------------------------------------------------------------------------------
    'Create the workbook-level name so it refers to the requested target cell
        WB.Names.Add Name:=NameText, _
                     RefersTo:="='" & TargetCell.Parent.Name & "'!" & TargetCell.Address

End Sub

Public Sub DEMO_Btn_Add( _
    ByVal WS As Worksheet, _
    ByVal ShapeName As String, _
    ByVal CaptionText As String, _
    ByVal LeftPos As Double, _
    ByVal TopPos As Double, _
    ByVal ShapeWidth As Double, _
    ByVal ShapeHeight As Double, _
    Optional ByVal ActionMacro As String = "", _
    Optional ByVal FillColor As Long = COLOR_BUTTON)
'
'------------------------------------------------------------------------------
'                               ADD DEMO BUTTON
'------------------------------------------------------------------------------
' PURPOSE
'   Creates one DEMO button as a worksheet Shape
'
' WHY THIS EXISTS
'   Shape-based buttons are portable, easy to style, and easy to assign to a
'   macro during workbook-layout generation
'
' INPUTS
'   WS
'     Worksheet where the button should be created
'
'   ShapeName
'     Internal shape name
'
'   CaptionText
'     Visible button caption
'
'   LeftPos
'     Button left position
'
'   TopPos
'     Button top position
'
'   ShapeWidth
'     Button width
'
'   ShapeHeight
'     Button height
'
'   ActionMacro (optional)
'     Macro name to assign to the shape's OnAction property. Blank, the
'     default, leaves the button without a macro.
'
'   FillColor (optional)
'     Fill color to apply to the button shape
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Creates a rounded-rectangle shape
'   - Applies button styling
'   - Uses FillColor when provided, otherwise defaults to COLOR_BUTTON
'   - Writes the requested caption
'   - Assigns the requested action macro; a blank ActionMacro leaves OnAction
'     empty, so clicking the button does nothing instead of raising "Cannot
'     run the macro" for a procedure that does not exist
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Worksheet.Shapes.AddShape (msoShapeRoundedRectangle)
'   - Shape Fill, Line, OnAction and TextFrame2
'   - ThisWorkbook.Name
'
' NOTES
'   - OnAction is qualified with the name of ThisWorkbook unless ActionMacro
'     already contains "!"
'   - The caption is white, 10 pt, bold and centered; the line color is
'     RGB(90, 90, 90)
'
' HISTORY
'   2026-11-06 - A blank ActionMacro leaves OnAction empty instead of naming
'                the undeclared DEMO_DEMOButton_NotAssigned (#129).
'   2026-04-22
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Btn_Add"

    Dim Shp As Shape    'Created button shape

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing worksheet reference
        If WS Is Nothing Then
            Err.Raise vbObjectError + 2007, _
                      "M_DEMO_BUILDER.DEMO_Btn_Add", _
                      "Worksheet reference cannot be Nothing."
        End If
    'Reject a blank shape name
        If Len(Trim$(ShapeName)) = 0 Then
            Err.Raise vbObjectError + 2008, _
                      "M_DEMO_BUILDER.DEMO_Btn_Add", _
                      "Shape name cannot be blank."
        End If
    'Reject nonpositive dimensions
        If ShapeWidth <= 0 Or ShapeHeight <= 0 Then
            Err.Raise vbObjectError + 2009, _
                      "M_DEMO_BUILDER.DEMO_Btn_Add", _
                      "Button width and height must be positive."
        End If

'------------------------------------------------------------------------------
' CREATE SHAPE
'------------------------------------------------------------------------------
    'Create a rounded-rectangle shape for the DEMO button
        Set Shp = WS.Shapes.AddShape( _
                      Type:=msoShapeRoundedRectangle, _
                      Left:=LeftPos, _
                      Top:=TopPos, _
                      Width:=ShapeWidth, _
                      Height:=ShapeHeight)

'------------------------------------------------------------------------------
' ASSIGN BASIC PROPERTIES
'------------------------------------------------------------------------------
    'Assign the internal shape name
        Shp.Name = ShapeName
    'Assign the requested action macro
        If Len(Trim$(ActionMacro)) = 0 Then
            Shp.OnAction = vbNullString
        ElseIf InStr(1, ActionMacro, "!", vbTextCompare) > 0 Then
            Shp.OnAction = ActionMacro
        Else
            Shp.OnAction = "'" & ThisWorkbook.Name & "'!" & ActionMacro
        End If

'------------------------------------------------------------------------------
' FORMAT SHAPE
'------------------------------------------------------------------------------
    'Apply button styling and visible caption text
        With Shp

        'Apply the requested fill color
            .Fill.ForeColor.RGB = FillColor
        'Apply the standard line color
            .Line.ForeColor.RGB = RGB(90, 90, 90)
        'Write the visible caption text
            .TextFrame2.TextRange.Text = CaptionText
        'Apply the standard font size
            .TextFrame2.TextRange.Font.Size = 10
        'Apply the standard font color
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        'Apply bold font styling
            .TextFrame2.TextRange.Font.Bold = msoTrue
        'Center the text vertically
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
        'Center the text horizontally
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

        End With

End Sub

Public Sub DEMO_Hide_HelperColumns( _
    ByVal WS As Worksheet, _
    ByVal ColumnAddress As String)
'
'------------------------------------------------------------------------------
'                             HIDE HELPER COLUMNS
'------------------------------------------------------------------------------
' PURPOSE
'   Hides one or more helper columns used to store internal lists or support
'   data on a DEMO sheet.
'
' WHY THIS EXISTS
'   DEMO sheets often need helper ranges for:
'     - validation lists
'     - internal lookup values
'     - support formulas
'
'   Those ranges should usually remain invisible to the end user.
'
' INPUTS
'   WS
'     Worksheet containing the helper columns.
'
'   ColumnAddress
'     Column address string such as:
'       - "AA:AZ"
'       - "XFD:XFD"
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits quietly if WS Is Nothing
'   - Exits quietly if ColumnAddress is blank
'   - Hides the requested worksheet columns
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - Worksheet.Columns(...).Hidden
'
' NOTES
'   - DEMO_Create_BoolList uses it to hide AA:AZ
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no worksheet was supplied
        If WS Is Nothing Then Exit Sub

    'Exit quietly when no column address was supplied
        If Len(Trim$(ColumnAddress)) = 0 Then Exit Sub

'------------------------------------------------------------------------------
' APPLY HIDE
'------------------------------------------------------------------------------
    'Hide the requested helper columns
        WS.Columns(ColumnAddress).Hidden = True

End Sub

Public Function DEMO_Create_Table( _
    ByVal TargetRange As Range, _
    ByVal TableName As String, _
    Optional ByVal TableStyleName As String = "TableStyleMedium2") _
    As ListObject
'
'------------------------------------------------------------------------------
'                                 CREATE TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Creates or recreates an Excel ListObject on a target range.
'
' WHY THIS EXISTS
'   DEMO workbooks often benefit from standardized result or log tables.
'   This helper centralizes table creation so callers do not need to repeat the
'   same ListObject setup logic.
'
' INPUTS
'   TargetRange
'     Source range for the table, including header row.
'
'   TableName
'     Name to assign to the ListObject.
'
'   TableStyleName (optional)
'     Built-in or custom table style name.
'
' RETURNS
'   ListObject
'     The created table object.
'
' BEHAVIOR
'   - Deletes an existing same-sheet table with the same name when present
'   - Creates a new ListObject from the supplied range
'   - Applies the requested table style
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - Worksheet.ListObjects: Add and Delete
'   - ListObject.Name and TableStyle
'
' NOTES
'   - The existing table is looked up on the worksheet of TargetRange, and a
'     missing one is ignored (On Error Resume Next)
'   - TargetRange must include the header row
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Create_Table"

    Dim Lo  As ListObject   'Created or replaced table

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing target range
        If TargetRange Is Nothing Then
            Err.Raise vbObjectError + 2010, _
                      "M_DEMO_BUILDER.DEMO_Create_Table", _
                      "Target range cannot be Nothing."
        End If

    'Reject a blank table name
        If Len(Trim$(TableName)) = 0 Then
            Err.Raise vbObjectError + 2011, _
                      "M_DEMO_BUILDER.DEMO_Create_Table", _
                      "Table name cannot be blank."
        End If

'------------------------------------------------------------------------------
' DELETE EXISTING TABLE
'------------------------------------------------------------------------------
    'Delete an existing same-sheet table with the requested name when present
        On Error Resume Next
        TargetRange.Worksheet.ListObjects(TableName).Delete
        On Error GoTo 0

'------------------------------------------------------------------------------
' CREATE TABLE
'------------------------------------------------------------------------------
    'Create the new table from the target range
        Set Lo = TargetRange.Worksheet.ListObjects.Add( _
                    SourceType:=xlSrcRange, _
                    Source:=TargetRange, _
                    XlListObjectHasHeaders:=xlYes)

    'Assign the requested table name
        Lo.Name = TableName

    'Apply the requested table style
        Lo.TableStyle = TableStyleName

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the created table
        Set DEMO_Create_Table = Lo

End Function

Public Sub DEMO_Clear_TableBody( _
    ByVal TableObject As ListObject)
'
'------------------------------------------------------------------------------
'                               CLEAR TABLE BODY
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the data body of a table while preserving its header row and
'   structure.
'
' WHY THIS EXISTS
'   DEMO result/log tables often need to be cleared between runs without
'   destroying the table definition itself.
'
' INPUTS
'   TableObject
'     Table to clear.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits quietly if TableObject Is Nothing
'   - Deletes the DataBodyRange when present
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - ListObject.DataBodyRange
'
' NOTES
'   - The body rows are deleted with DataBodyRange.Delete, not cleared
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no table was supplied
        If TableObject Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' CLEAR BODY
'------------------------------------------------------------------------------
    'Delete the table body rows when present
        If Not TableObject.DataBodyRange Is Nothing Then
            TableObject.DataBodyRange.Delete
        End If

End Sub

Public Sub DEMO_Append_TableRow( _
    ByVal TableObject As ListObject, _
    ByVal RowValues As Variant)
'
'------------------------------------------------------------------------------
'                               APPEND TABLE ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one row of values to an Excel table.
'
' WHY THIS EXISTS
'   DEMO result/log tables are easier to populate consistently through one
'   helper routine rather than repeating ListRow logic in multiple macros.
'
' INPUTS
'   TableObject
'     Table that should receive the new row.
'
'   RowValues
'     One-dimensional Variant array of values to write into the new row.
'     Typical use:
'       Array(Value1, Value2, Value3, ...)
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Exits quietly if TableObject Is Nothing
'   - Adds one new ListRow
'   - Writes the supplied array values into the new row, left to right
'   - Stops at the smaller of:
'       * the number of supplied values
'       * the number of table columns
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - ListObject.ListRows.Add and ListColumns.Count
'   - ListRow.Range
'   - Application.Min
'
' NOTES
'   - The first value written is element LBound(RowValues), whatever the base
'     of the array
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Append_TableRow"

    Dim LR          As ListRow  'Newly appended row
    Dim i           As Long     'Column/value index
    Dim LoBound     As Long     'Lower bound of the values array
    Dim HiBound     As Long     'Upper bound of the values array
    Dim MaxWrite    As Long     'Maximum number of values to write

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Exit quietly when no table was supplied
        If TableObject Is Nothing Then Exit Sub

    'Reject a non-array RowValues argument
        If Not IsArray(RowValues) Then
            Err.Raise vbObjectError + 2012, _
                      "M_DEMO_BUILDER.DEMO_Append_TableRow", _
                      "RowValues must be a one-dimensional array."
        End If

'------------------------------------------------------------------------------
' ADD ROW AND SIZE WRITE
'------------------------------------------------------------------------------
    'Add one new row to the table
        Set LR = TableObject.ListRows.Add

    'Get the bounds of the supplied values array
        LoBound = LBound(RowValues)
        HiBound = UBound(RowValues)

    'Determine how many values can be written
        MaxWrite = Application.Min(HiBound - LoBound + 1, TableObject.ListColumns.Count)

'------------------------------------------------------------------------------
' WRITE VALUES
'------------------------------------------------------------------------------
    'Write the supplied values into the new row, left to right
        For i = 1 To MaxWrite
            LR.Range.Cells(1, i).Value = RowValues(LoBound + i - 1)
        Next i

End Sub

Public Sub DEMO_Prepare_LabeledInputSection( _
    ByVal WS As Worksheet, _
    ByVal SectionHeaderRange As Range, _
    ByVal SectionTitle As String, _
    ByVal LabelRange As Range, _
    ByVal ValueRange As Range)
'
'------------------------------------------------------------------------------
'                        PREPARE LABELED INPUT SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the standard setup for a vertical label/value input section
'
' WHY THIS EXISTS
'   Many DEMO sheets share the same pattern:
'     - a section header
'     - a label column
'     - a parallel input/value column
'
'   Centralizing that setup avoids repeated formatting code in each DEMO sheet
'
' INPUTS
'   WS
'     Target worksheet
'
'   SectionHeaderRange
'     Range used for the section title band
'
'   SectionTitle
'     Text to display in the section title band
'
'   LabelRange
'     Vertical range containing the labels
'
'   ValueRange
'     Vertical range containing the user-editable values
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Writes the section header
'   - Applies standard label formatting
'   - Applies standard input-cell formatting
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - DEMO_Write_BandHeader
'   - DEMO_Format_Labels
'   - DEMO_Format_InputCell
'
' NOTES
'   - WS is accepted but not read; the ranges carry their own worksheet
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLY STANDARD SECTION FORMATTING
'------------------------------------------------------------------------------
    'Write the section header band
        DEMO_Write_BandHeader SectionHeaderRange, SectionTitle
    'Apply standard formatting to the label cells
        DEMO_Format_Labels LabelRange
    'Apply standard formatting to the input cells
        DEMO_Format_InputCell ValueRange

End Sub

Public Sub DEMO_Write_NamedInputRow( _
    ByVal WB As Workbook, _
    ByVal WS As Worksheet, _
    ByVal LabelCell As Range, _
    ByVal ValueCell As Range, _
    ByVal LabelText As String, _
    ByVal DefaultValue As Variant, _
    Optional ByVal WorkbookNameText As String = "", _
    Optional ByVal ValidationKind As DEMOInputValidationKind = DEMOInputValidationNone, _
    Optional ByVal ValidationSource As String = "", _
    Optional ByVal ValidationMin As Double = 0, _
    Optional ByVal ValidationMax As Double = 0, _
    Optional ByVal AllowDecimal As Boolean = True, _
    Optional ByVal NumberFormatText As String = "")
'
'------------------------------------------------------------------------------
'                            WRITE NAMED INPUT ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one standard label/value row for a DEMO sheet control panel
'
' WHY THIS EXISTS
'   DEMO control panels often require the same row-level actions:
'     - write a label
'     - seed a default value
'     - apply validation
'     - apply number format
'     - bind the input cell to a workbook-level name
'
'   Consolidating those actions keeps sheet-builder code compact and consistent
'
' INPUTS
'   WB
'     Target workbook for workbook-level names
'
'   WS
'     Target worksheet
'
'   LabelCell
'     Cell that will hold the label text
'
'   ValueCell
'     Cell that will hold the default/editable value
'
'   LabelText
'     Text for the label cell
'
'   DefaultValue
'     Initial value to write to the input cell
'
'   WorkbookNameText
'     Optional workbook-level name to bind to ValueCell
'
'   ValidationKind
'     Validation mode to apply
'
'   ValidationSource
'     Source string for list validation
'
'   ValidationMin
'     Numeric validation minimum
'
'   ValidationMax
'     Numeric validation maximum
'
'   AllowDecimal
'     True for decimal numeric validation, False for whole-number validation
'
'   NumberFormatText
'     Optional number format to apply to ValueCell
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Writes the label and default value
'   - Applies validation according to ValidationKind
'   - Applies an optional number format
'   - Creates or refreshes an optional workbook-level name
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - DEMO_Create_BoolList
'   - DEMO_Apply_ValidationList
'   - DEMO_Apply_NumericValidation
'   - DEMO_Set_WorkbookName
'
' NOTES
'   - WS is read only to name the worksheet that holds the Boolean helper list
'   - ValidationSource applies to the list kind; ValidationMin, ValidationMax
'     and AllowDecimal apply to the numeric kind; the Boolean kind uses the
'     =BoolList name
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE LABEL AND DEFAULT VALUE
'------------------------------------------------------------------------------
    'Write the label text
        LabelCell.Value = LabelText
    'Write the default/input value
        ValueCell.Value = DefaultValue

'------------------------------------------------------------------------------
' APPLY VALIDATION
'------------------------------------------------------------------------------
        Select Case ValidationKind
            Case DEMOInputValidationNone
            'Leave the cell without validation
            Case DEMOInputValidationList
            'Apply a standard list validation
                DEMO_Apply_ValidationList ValueCell, ValidationSource
            Case DEMOInputValidationNumeric
            'Apply numeric validation with the requested numeric policy
                DEMO_Apply_NumericValidation ValueCell, ValidationMin, ValidationMax, AllowDecimal
            Case DEMOInputValidationBoolean
            'Ensure the workbook-level boolean list exists for this sheet context
                DEMO_Create_BoolList WS.Name
            'Apply the shared boolean dropdown
                DEMO_Apply_ValidationList ValueCell, "=BoolList"
            Case Else
            'Add no validation for any other kind
        End Select

'------------------------------------------------------------------------------
' APPLY OPTIONAL NUMBER FORMAT
'------------------------------------------------------------------------------
    'Apply the requested number format when provided
        If Len(NumberFormatText) > 0 Then
            ValueCell.NumberFormat = NumberFormatText
        End If

'------------------------------------------------------------------------------
' APPLY OPTIONAL WORKBOOK NAME
'------------------------------------------------------------------------------
    'Create or refresh the workbook-level name when requested
        If Len(WorkbookNameText) > 0 Then
            DEMO_Set_WorkbookName WB, WorkbookNameText, ValueCell
        End If

End Sub

Public Sub DEMO_Btn_AddGrid( _
    ByVal WS As Worksheet, _
    ByVal AnchorCell As Range, _
    ByRef ButtonSpecs As Variant, _
    Optional ByVal ButtonsPerRow As Long = 2, _
    Optional ByVal ButtonWidth As Double = 150, _
    Optional ByVal ButtonHeight As Double = 25, _
    Optional ByVal GapX As Double = 20, _
    Optional ByVal GapY As Double = 13, _
    Optional ByVal OffsetLeft As Double = 4, _
    Optional ByVal OffsetTop As Double = 8)
'
'------------------------------------------------------------------------------
'                               ADD BUTTON GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Creates a regular grid of DEMO buttons from a compact button specification
'
' WHY THIS EXISTS
'   DEMO sheets often contain repeated button-layout code with only:
'     - button name
'     - button caption
'     - optional action macro
'     - optional fill color
'     - grid position
'
' INPUTS
'   WS
'     Worksheet receiving the button grid
'
'   AnchorCell
'     Top-left anchor cell for the grid layout
'
'   ButtonSpecs
'     Variant array where each item must itself be an array with one of these
'     supported layouts:
'
'         Array(ButtonName, ButtonCaption)
'         Array(ButtonName, ButtonCaption, ActionMacro)
'         Array(ButtonName, ButtonCaption, FillColor)
'         Array(ButtonName, ButtonCaption, ActionMacro, FillColor)
'
'     Notes
'       - When the 3rd element is text, it is treated as ActionMacro
'       - Without an ActionMacro the button has no macro
'       - When the 3rd element is numeric, it is treated as FillColor
'       - When the 4th element is present, it is treated as FillColor
'
'   ButtonsPerRow (optional)
'     Number of buttons to place on each grid row
'
'   ButtonWidth (optional)
'     Width of each button
'
'   ButtonHeight (optional)
'     Height of each button
'
'   GapX (optional)
'     Horizontal gap between buttons
'
'   GapY (optional)
'     Vertical gap between buttons
'
'   OffsetLeft (optional)
'     Left offset from AnchorCell
'
'   OffsetTop (optional)
'     Top offset from AnchorCell
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Computes a regular button grid from AnchorCell
'   - Resolves an optional action macro per button
'   - Resolves an optional fill color per button
'   - Creates each button through DEMO_Btn_Add
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - DEMO_Btn_Add
'
' NOTES
'   - ButtonSpecs is declared ByRef but only read
'   - A third element that is neither text nor numeric is ignored
'
' HISTORY
'   2026-11-06 - A button without an ActionMacro gets no macro (#129).
'   2026-04-22
'
' UPDATED
'   2026-11-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Btn_AddGrid"

    Dim i           As Long     'Loop index
    Dim GridRow     As Long     'Zero-based button row
    Dim GridCol     As Long     'Zero-based button column
    Dim ButtonLeft  As Double   'Computed button left position
    Dim ButtonTop   As Double   'Computed button top position
    Dim ActionMacro As String   'Resolved button action macro
    Dim FillColor   As Long     'Resolved button fill color
    Dim SpecItem    As Variant  'Current button specification item
    Dim SpecUBound  As Long     'Upper bound of the current spec item

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing worksheet reference
        If WS Is Nothing Then
            Err.Raise vbObjectError + 2131, _
                      "M_DEMO_BUILDER.DEMO_Btn_AddGrid", _
                      "Worksheet reference cannot be Nothing."
        End If

    'Reject a missing anchor cell
        If AnchorCell Is Nothing Then
            Err.Raise vbObjectError + 2132, _
                      "M_DEMO_BUILDER.DEMO_Btn_AddGrid", _
                      "AnchorCell cannot be Nothing."
        End If

    'Reject a non-array button specification input
        If Not IsArray(ButtonSpecs) Then
            Err.Raise vbObjectError + 2133, _
                      "M_DEMO_BUILDER.DEMO_Btn_AddGrid", _
                      "ButtonSpecs must be an array."
        End If

    'Reject an invalid buttons-per-row value
        If ButtonsPerRow < 1 Then
            Err.Raise vbObjectError + 2134, _
                      "M_DEMO_BUILDER.DEMO_Btn_AddGrid", _
                      "ButtonsPerRow must be >= 1."
        End If

    'Reject nonpositive button dimensions
        If ButtonWidth <= 0 Or ButtonHeight <= 0 Then
            Err.Raise vbObjectError + 2135, _
                      "M_DEMO_BUILDER.DEMO_Btn_AddGrid", _
                      "ButtonWidth and ButtonHeight must be positive."
        End If

'------------------------------------------------------------------------------
' BUILD BUTTON GRID
'------------------------------------------------------------------------------
    'Create each button from the supplied grid specification
        For i = LBound(ButtonSpecs) To UBound(ButtonSpecs)

        'Read the current button specification item
            SpecItem = ButtonSpecs(i)

        'Reject a non-array button specification item
            If Not IsArray(SpecItem) Then
                Err.Raise vbObjectError + 2136, _
                          "M_DEMO_BUILDER.DEMO_Btn_AddGrid", _
                          "Each ButtonSpecs item must be an array."
            End If

        'Capture the upper bound of the current specification item
            SpecUBound = UBound(SpecItem)

        'Reject a specification item with fewer than name + caption
            If SpecUBound < 1 Then
                Err.Raise vbObjectError + 2137, _
                          "M_DEMO_BUILDER.DEMO_Btn_AddGrid", _
                          "Each ButtonSpecs item must contain at least ButtonName and ButtonCaption."
            End If

        'Compute the zero-based grid row
            GridRow = (i - LBound(ButtonSpecs)) \ ButtonsPerRow

        'Compute the zero-based grid column
            GridCol = (i - LBound(ButtonSpecs)) Mod ButtonsPerRow

        'Compute the left position of the current button
            ButtonLeft = AnchorCell.Left + OffsetLeft + (GridCol * (ButtonWidth + GapX))

        'Compute the top position of the current button
            ButtonTop = AnchorCell.Top + OffsetTop + (GridRow * (ButtonHeight + GapY))

'------------------------------------------------------------------------------
' RESOLVE OPTIONAL ACTION MACRO / FILL COLOR
'------------------------------------------------------------------------------
        'Initialize the default action macro: none
            ActionMacro = vbNullString

        'Initialize the default fill color
            FillColor = COLOR_BUTTON

        'Resolve the optional 3rd element
            If SpecUBound >= 2 Then

            'Treat text as ActionMacro
                If VarType(SpecItem(2)) = vbString Then
                    ActionMacro = CStr(SpecItem(2))

                'Treat numeric values as FillColor
                ElseIf IsNumeric(SpecItem(2)) Then
                    FillColor = CLng(SpecItem(2))

                End If

            End If

        'Resolve the optional 4th element as FillColor
            If SpecUBound >= 3 Then
                If IsNumeric(SpecItem(3)) Then
                    FillColor = CLng(SpecItem(3))
                End If
            End If

'------------------------------------------------------------------------------
' CREATE CURRENT BUTTON
'------------------------------------------------------------------------------
        'Create the current button with the resolved action macro and fill color
            DEMO_Btn_Add _
                WS, _
                CStr(SpecItem(0)), _
                CStr(SpecItem(1)), _
                ButtonLeft, _
                ButtonTop, _
                ButtonWidth, _
                ButtonHeight, _
                ActionMacro, _
                FillColor

        Next i

End Sub

Public Function DEMO_Create_TableSection( _
    ByVal WS As Worksheet, _
    ByVal SectionHeaderRange As Range, _
    ByVal SectionTitle As String, _
    ByVal HeaderTopLeft As Range, _
    ByRef Headers As Variant, _
    ByVal TableName As String, _
    Optional ByVal TableStyleName As String = "TableStyleMedium6") As ListObject
'
'------------------------------------------------------------------------------
'                             CREATE TABLE SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Builds a standard titled section containing a formatted ListObject table
'
' WHY THIS EXISTS
'   Many DEMO sheets need the same output/log pattern:
'     - section header band
'     - one formatted header row
'     - one seeded blank row
'     - immediate ListObject creation
'
'   Centralizing this pattern keeps sheet builders much shorter and more
'   consistent
'
' INPUTS
'   WS
'     Target worksheet
'
'   SectionHeaderRange
'     Range used for the section title band
'
'   SectionTitle
'     Title text for the section
'
'   HeaderTopLeft
'     Top-left cell of the table header row
'
'   Headers
'     One-dimensional array of column captions
'
'   TableName
'     Name to assign to the ListObject
'
'   TableStyleName
'     Table style to apply after creation
'
' RETURNS
'   ListObject
'     The created ListObject
'
' BEHAVIOR
'   - Validates the requested worksheet, ranges, headers, and table name
'   - Writes the section header
'   - Writes the header captions across one row
'   - Formats the header row
'   - Seeds one blank data row
'   - Replaces an existing same-sheet table with the same name when present
'   - Creates the ListObject immediately
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - DEMO_Write_BandHeader
'
' NOTES
'   Headers must be supplied as a one-dimensional array with at least one
'   element
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Create_TableSection"

    Dim HeaderCount As Long         'Number of table columns
    Dim i           As Long         'Loop index
    Dim HeaderRange As Range        'Formatted header row
    Dim TableRange  As Range        'Range used to create the ListObject
    Dim Lo          As ListObject   'Created table object

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing worksheet reference
        If WS Is Nothing Then
            Err.Raise vbObjectError + 2121, _
                      "M_DEMO_BUILDER.DEMO_Create_TableSection", _
                      "Worksheet reference cannot be Nothing."
        End If

    'Reject a missing section-header range
        If SectionHeaderRange Is Nothing Then
            Err.Raise vbObjectError + 2122, _
                      "M_DEMO_BUILDER.DEMO_Create_TableSection", _
                      "SectionHeaderRange cannot be Nothing."
        End If

    'Reject a missing header top-left cell
        If HeaderTopLeft Is Nothing Then
            Err.Raise vbObjectError + 2123, _
                      "M_DEMO_BUILDER.DEMO_Create_TableSection", _
                      "HeaderTopLeft cannot be Nothing."
        End If

    'Reject a non-array header specification
        If Not IsArray(Headers) Then
            Err.Raise vbObjectError + 2124, _
                      "M_DEMO_BUILDER.DEMO_Create_TableSection", _
                      "Headers must be a one-dimensional array."
        End If

    'Reject a blank table name
        If Len(Trim$(TableName)) = 0 Then
            Err.Raise vbObjectError + 2125, _
                      "M_DEMO_BUILDER.DEMO_Create_TableSection", _
                      "TableName cannot be blank."
        End If

'------------------------------------------------------------------------------
' COUNT HEADER CAPTIONS
'------------------------------------------------------------------------------
    'Determine the number of header captions
        HeaderCount = UBound(Headers) - LBound(Headers) + 1

    'Reject an empty header specification
        If HeaderCount < 1 Then
            Err.Raise vbObjectError + 2126, _
                      "M_DEMO_BUILDER.DEMO_Create_TableSection", _
                      "Headers must contain at least one element."
        End If

'------------------------------------------------------------------------------
' WRITE SECTION HEADER
'------------------------------------------------------------------------------
    'Write the section title band
        DEMO_Write_BandHeader SectionHeaderRange, SectionTitle

'------------------------------------------------------------------------------
' WRITE TABLE HEADERS
'------------------------------------------------------------------------------
    'Write each header caption across the target row
        For i = LBound(Headers) To UBound(Headers)
            HeaderTopLeft.Offset(0, i - LBound(Headers)).Value = Headers(i)
        Next i

    'Resolve the full header row range
        Set HeaderRange = HeaderTopLeft.Resize(1, HeaderCount)

'------------------------------------------------------------------------------
' FORMAT HEADER ROW
'------------------------------------------------------------------------------
    'Apply the standard table-header formatting
        With HeaderRange
            .Interior.Color = COLOR_SUBHEADER
            .Font.Bold = True
            .Font.Color = vbWhite
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

'------------------------------------------------------------------------------
' SEED INITIAL TABLE RANGE
'------------------------------------------------------------------------------
    'Clear the first data row so the ListObject can be created immediately
        HeaderRange.Offset(1, 0).Resize(1, HeaderCount).ClearContents

    'Define the 2-row source range for table creation
        Set TableRange = HeaderRange.Resize(2, HeaderCount)

'------------------------------------------------------------------------------
' DELETE EXISTING SAME-SHEET TABLE
'------------------------------------------------------------------------------
    'Delete an existing same-sheet table with the requested name when present
        On Error Resume Next
        WS.ListObjects(TableName).Delete
        On Error GoTo 0

'------------------------------------------------------------------------------
' CREATE TABLE
'------------------------------------------------------------------------------
    'Create the ListObject from the prepared header-plus-blank-row range
        Set Lo = WS.ListObjects.Add( _
                    SourceType:=xlSrcRange, _
                    Source:=TableRange, _
                    XlListObjectHasHeaders:=xlYes)

    'Assign the requested table name
        Lo.Name = TableName

    'Apply the requested built-in table style
        Lo.TableStyle = TableStyleName

'------------------------------------------------------------------------------
' RETURN TABLE
'------------------------------------------------------------------------------
    'Return the created ListObject
        Set DEMO_Create_TableSection = Lo

End Function

'
'------------------------------------------------------------------------------
'
'                                BUTTON STATES
'
'------------------------------------------------------------------------------
'

Public Sub DEMO_Btn_ApplyState( _
    ByVal Shp As Shape, _
    ByVal StateName As String)
'
'------------------------------------------------------------------------------
'                              APPLY BUTTON STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Applies one named visual state to a shape-based button
'
' WHY THIS EXISTS
'   DEMO buttons use a small set of visual states such as:
'     - normal
'     - hover
'     - pressed
'
'   Centralizing that mapping keeps the button behavior consistent
'
' INPUTS
'   Shp
'     Target shape button
'
'   StateName
'     Requested state name
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Applies the corresponding fill/text styling for the requested state
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Shape Fill, Line and TextFrame2
'   - BTN_PRIMARY_N, BTN_PRIMARY_H, BTN_PRIMARY_P and WHITE
'
' NOTES
'   - State names are matched without regard to case or surrounding spaces; an
'     unknown name gets the normal color
'   - The state also hides the outline of the shape
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Btn_ApplyState"

    Dim Col As Long 'Resolved fill color for the requested state

'------------------------------------------------------------------------------
' RESOLVE STATE COLOR
'------------------------------------------------------------------------------
        Select Case UCase$(Trim$(StateName))
            Case "NORMAL"
                Col = BTN_PRIMARY_N

            Case "HOVER"
                Col = BTN_PRIMARY_H

            Case "PRESSED"
                Col = BTN_PRIMARY_P

            Case Else
                Col = BTN_PRIMARY_N
        End Select

'------------------------------------------------------------------------------
' APPLY STATE
'------------------------------------------------------------------------------
        With Shp
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = Col
            .Line.Visible = msoFalse
            If .TextFrame2.HasText Then
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = WHITE
            End If
        End With

End Sub

Private Function DEMO_Btn_CaptureAppearance( _
    ByVal Shp As Shape) _
    As tButtonAppearance
'
'------------------------------------------------------------------------------
'                          CAPTURE BUTTON APPEARANCE
'------------------------------------------------------------------------------
' PURPOSE
'   Captures the current visual appearance of a shape button
'
' WHY THIS EXISTS
'   Pressed/hover animations should return the button to its exact original
'   appearance rather than forcing a generic "normal" style
'
' INPUTS
'   Shp
'     Target shape button
'
' RETURNS
'   tButtonAppearance
'     Snapshot of the original visual appearance
'
' BEHAVIOR
'   Captures fill, line, text, shadow, and position attributes
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Shape Fill, Line, TextFrame2, Shadow, Top and Left
'   - tButtonAppearance
'
' NOTES
'   - Line color and weight are stored as 0 when the line is not visible; the
'     text values are captured only when the shape has text
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Btn_CaptureAppearance"

    Dim S   As tButtonAppearance    'Captured appearance snapshot

'------------------------------------------------------------------------------
' CAPTURE SHAPE APPEARANCE
'------------------------------------------------------------------------------
        With Shp
        'Capture fill appearance
            S.FillVisible = .Fill.Visible
            S.FillColor = .Fill.ForeColor.RGB

        'Capture line appearance
            S.LineVisible = .Line.Visible

            If S.LineVisible = msoTrue Then
                S.LineColor = .Line.ForeColor.RGB
                S.LineWeight = .Line.Weight
            Else
                S.LineColor = 0
                S.LineWeight = 0!
            End If

        'Capture text appearance when text is present
            If .TextFrame2.HasText Then
                S.TextColor = .TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                S.TextBold = .TextFrame2.TextRange.Font.Bold
                S.TextSize = .TextFrame2.TextRange.Font.Size
            End If

        'Capture shadow appearance
            S.ShadowVisible = .Shadow.Visible
            S.ShadowBlur = .Shadow.Blur
            S.ShadowOffsetX = .Shadow.OffsetX
            S.ShadowOffsetY = .Shadow.OffsetY

        'Capture position
            S.TopPos = .Top
            S.LeftPos = .Left
        End With

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
        DEMO_Btn_CaptureAppearance = S

End Function

Private Sub DEMO_Btn_RestoreAppearance( _
    ByVal Shp As Shape, _
    ByRef SavedState As tButtonAppearance)
'
'------------------------------------------------------------------------------
'                          RESTORE BUTTON APPEARANCE
'------------------------------------------------------------------------------
' PURPOSE
'   Restores a previously captured visual appearance to a shape button
'
' WHY THIS EXISTS
'   Temporary button animations should leave the button exactly as it was
'   before the animation started
'
' INPUTS
'   Shp
'     Target shape button
'
'   SavedState
'     Previously captured appearance snapshot
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Restores fill, line, text, shadow, and position attributes
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Shape Fill, Line, TextFrame2, Shadow, Top and Left
'   - tButtonAppearance
'
' NOTES
'   - Values are written back as captured, including the 0 line color and
'     weight stored for a hidden line
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESTORE SHAPE APPEARANCE
'------------------------------------------------------------------------------
        With Shp
        'Restore fill appearance
            .Fill.Visible = SavedState.FillVisible
            .Fill.ForeColor.RGB = SavedState.FillColor

        'Restore line appearance
            .Line.Visible = SavedState.LineVisible
            .Line.ForeColor.RGB = SavedState.LineColor
            .Line.Weight = SavedState.LineWeight

        'Restore text appearance when text is present
            If .TextFrame2.HasText Then
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = SavedState.TextColor
                .TextFrame2.TextRange.Font.Bold = SavedState.TextBold
                .TextFrame2.TextRange.Font.Size = SavedState.TextSize
            End If

        'Restore shadow appearance
            .Shadow.Visible = SavedState.ShadowVisible
            .Shadow.Blur = SavedState.ShadowBlur
            .Shadow.OffsetX = SavedState.ShadowOffsetX
            .Shadow.OffsetY = SavedState.ShadowOffsetY

        'Restore position
            .Top = SavedState.TopPos
            .Left = SavedState.LeftPos
        End With

End Sub

Public Sub DEMO_Btn_Click()

'
'------------------------------------------------------------------------------
'                          SHOW BUTTON PRESS FEEDBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Applies a short pressed-state visual feedback to the calling shape button
'
' WHY THIS EXISTS
'   DEMO actions may optionally call this routine to simulate a button press
'   before running the underlying business logic
'
'   This routine is defensive: when no valid shape caller exists, it exits
'   quietly rather than raising an error
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Resolves the calling shape from Application.Caller when available
'   - Applies a short pressed-state animation
'   - Waits briefly using Timer() delta logic
'   - Restores the original appearance exactly
'   - Exits quietly if no valid shape caller exists
'
' ERROR POLICY
'   Best effort only; exits quietly when no valid caller shape can be resolved
'
' DEPENDENCIES
'   - DEMO_Btn_CaptureAppearance
'   - DEMO_Btn_ApplyState
'   - DEMO_Btn_RestoreAppearance
'
' NOTES
'   The short visual delay uses Timer() elapsed-delta logic so it remains safe
'   across midnight rollover
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_Btn_Click"

    Dim WS          As Worksheet            'Caller worksheet
    Dim Shp         As Shape                'Clicked shape
    Dim SavedState  As tButtonAppearance    'Captured original appearance
    Dim CallerName  As Variant              'Application.Caller value
    Dim StartS      As Double               'Timer() snapshot at delay start
    Dim ElapsedS    As Double               'Elapsed seconds during visual feedback

'------------------------------------------------------------------------------
' RESOLVE CALLER
'------------------------------------------------------------------------------
    'Read the current caller context
        CallerName = Application.Caller
    'Exit quietly when not called by a shape
        If VarType(CallerName) <> vbString Then Exit Sub
    'Resolve the active sheet
        Set WS = ActiveSheet
    'Try to resolve the calling shape on the active sheet
        On Error Resume Next
        Set Shp = WS.Shapes(CStr(CallerName))
        On Error GoTo 0
    'Exit quietly when the shape cannot be resolved
        If Shp Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' CAPTURE ORIGINAL APPEARANCE
'------------------------------------------------------------------------------
    'Capture the exact original appearance before applying the pressed effect
        SavedState = DEMO_Btn_CaptureAppearance(Shp)

'------------------------------------------------------------------------------
' APPLY PRESSED STATE
'------------------------------------------------------------------------------
    'Apply the pressed color state
        DEMO_Btn_ApplyState Shp, "Pressed"
    'Apply a small pressed-position effect
        Shp.Top = Shp.Top + 1
        Shp.Left = Shp.Left + 1
    'Apply a temporary pressed shadow
        With Shp.Shadow
            .Visible = msoTrue
            .Blur = 4
            .OffsetX = 1
            .OffsetY = 1
        End With

'------------------------------------------------------------------------------
' SMALL DELAY (VISUAL FEEDBACK)
'------------------------------------------------------------------------------
    'Capture the Timer() start snapshot for the short visual delay
        StartS = Timer

    'Wait briefly while keeping Excel responsive
        Do
        'Compute elapsed seconds since the delay started
            ElapsedS = Timer - StartS
        'Correct midnight rollover when Timer() wraps at 24 hours
            If ElapsedS < 0# Then
                ElapsedS = ElapsedS + 86400#
            End If
        'Exit once the short visual delay has elapsed
            If ElapsedS >= 0.08 Then
                Exit Do
            End If
        'Yield to Excel/UI during the brief wait
            DoEvents
        Loop

'------------------------------------------------------------------------------
' RESTORE ORIGINAL APPEARANCE
'------------------------------------------------------------------------------
    'Restore the exact original appearance
        DEMO_Btn_RestoreAppearance Shp, SavedState

End Sub

Public Sub DEMO_Btn_PlayFeedback()

'
'------------------------------------------------------------------------------
'                             PLAY BUTTON FEEDBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Invoke the shared button-click feedback helper on a best-effort basis
'
' WHY THIS EXISTS
'   Several public demo actions want the same visual pressed-state feedback
'   without duplicating local error suppression around DEMO_Btn_Click
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Calls DEMO_Btn_Click with errors suppressed
'   - Restores normal error handling afterwards
'
' ERROR POLICY
'   Best effort only
'
' DEPENDENCIES
'   - DEMO_Btn_Click
'
' NOTES
'   None
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLY BEST-EFFORT FEEDBACK
'------------------------------------------------------------------------------
    'Play the shared button-click visual feedback on a best-effort basis
        On Error Resume Next
        DEMO_Btn_Click
        On Error GoTo 0

End Sub

Public Sub DEMO_SB_SetProgress( _
    ByVal CurrentStep As Long, _
    ByVal TotalSteps As Long, _
    ByVal StepText As String)
'
'------------------------------------------------------------------------------
'                             SHOW SUITE PROGRESS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes regression-suite progress to the Excel status bar
'
' WHY THIS EXISTS
'   The regression suite can take noticeable time to complete. A visible
'   progress message makes execution easier to follow
'
' INPUTS
'   CurrentStep
'     Current completed step count
'
'   TotalSteps
'     Total number of planned suite steps
'
'   StepText
'     Short description of the current step
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Sets TotalSteps to 1 when it is smaller
'   - Clamps CurrentStep between 0 and TotalSteps
'   - Formats CurrentStep / TotalSteps as a whole percentage
'   - Writes "cPerformanceManager Regression Suite | n/N | p% | StepText" to
'     the status bar
'
' ERROR POLICY
'   Raises errors normally. Out-of-range counts are clamped, not rejected.
'
' DEPENDENCIES
'   - Application.StatusBar
'
' NOTES
'   - This routine never clears the status bar
'   - The clamping changes only the local copies of the two counts
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_SB_SetProgress"

    Dim Pct         As Double   'Completion ratio
    Dim PercentText As String   'Formatted completion percentage

'------------------------------------------------------------------------------
' NORMALIZE INPUTS
'------------------------------------------------------------------------------
    'Enforce a minimum total-step count
        If TotalSteps < 1 Then
            TotalSteps = 1
        End If

    'Clamp the current step to the valid range
        If CurrentStep < 0 Then
            CurrentStep = 0
        ElseIf CurrentStep > TotalSteps Then
            CurrentStep = TotalSteps
        End If

'------------------------------------------------------------------------------
' FORMAT PROGRESS
'------------------------------------------------------------------------------
    'Compute the completion ratio
        Pct = CurrentStep / CDbl(TotalSteps)
    'Format the completion percentage
        PercentText = Format$(Pct, "0%")

'------------------------------------------------------------------------------
' WRITE STATUS BAR
'------------------------------------------------------------------------------
    'Write the suite-progress message to the Excel status bar
        Application.StatusBar = _
            "cPerformanceManager Regression Suite | " & _
            CStr(CurrentStep) & "/" & CStr(TotalSteps) & " | " & _
            PercentText & " | " & StepText

End Sub

Public Sub DEMO_CB_AddForms( _
    ByVal WS As Worksheet, _
    ByVal AnchorRange As Range, _
    Optional ByVal CheckBoxNames As Variant)
'
'------------------------------------------------------------------------------
'                            ADD FORMS CHECK BOXES
'------------------------------------------------------------------------------
' PURPOSE
'   Add one Forms check box for each cell in the supplied anchor range
'
' WHY THIS EXISTS
'   DEMO sheets may need:
'     - one checkbox in a single anchor cell
'     - one checkbox per cell across a multi-cell range
'
'   In the multi-cell case, callers may want to pass a parallel array of
'   checkbox names instead of relying on auto-generated names
'
'   This helper centralizes:
'     - input validation
'     - deterministic control replacement
'     - centered per-cell checkbox placement
'     - optional explicit naming
'     - standard checkbox initialization
'
' INPUTS
'   WS
'     Worksheet receiving the Forms check box or check boxes
'
'   AnchorRange
'     One-cell or multi-cell range whose individual cells will each receive
'     one Forms check box
'
'   CheckBoxNames
'     Optional naming input
'
'     Allowed forms:
'       - omitted
'       - String when AnchorRange contains exactly one cell
'       - one-dimensional array of names with one element per cell in
'         AnchorRange
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Validates the worksheet, anchor range, and optional naming input
'   - Creates one Forms check box per cell in AnchorRange
'   - Deletes an existing same-sheet Forms check box with the same name when present
'   - Centers each check box inside its own anchor cell
'   - Removes the visible caption
'   - Initializes each control to unchecked
'   - Clears any linked cell
'   - Ties each control to cell movement / sizing
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - DEMO_CB_ValidateNamesInput
'   - DEMO_CB_ResolveName
'   - DEMO_AddOneFormsCheckBox
'
' NOTES
'   - The names input is validated before the first control is created, so a
'     bad CheckBoxNames adds no control
'   - Controls are created one cell at a time with no rollback; a failure in
'     a later cell leaves the earlier check boxes in place
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_CB_AddForms"

    Dim CellCount       As Long     'Number of target cells
    Dim CellIndex       As Long     '1-based current target-cell index
    Dim TargetCell      As Range    'Current target cell
    Dim ResolvedName    As String   'Resolved check-box name for the current cell

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing worksheet reference
        If WS Is Nothing Then
            Err.Raise vbObjectError + 2141, _
                      "M_DEMO_BUILDER.DEMO_CB_AddForms", _
                      "Worksheet reference cannot be Nothing"
        End If

    'Reject a missing anchor range
        If AnchorRange Is Nothing Then
            Err.Raise vbObjectError + 2142, _
                      "M_DEMO_BUILDER.DEMO_CB_AddForms", _
                      "AnchorRange cannot be Nothing"
        End If

    'Reject anchor ranges that do not belong to the target worksheet
        If Not AnchorRange.Worksheet Is WS Then
            Err.Raise vbObjectError + 2143, _
                      "M_DEMO_BUILDER.DEMO_CB_AddForms", _
                      "AnchorRange must belong to the target worksheet"
        End If

    'Reject non-rectangular multi-area ranges
        If AnchorRange.Areas.Count > 1 Then
            Err.Raise vbObjectError + 2144, _
                      "M_DEMO_BUILDER.DEMO_CB_AddForms", _
                      "AnchorRange must be one contiguous rectangular range"
        End If

'------------------------------------------------------------------------------
' COUNT TARGET CELLS
'------------------------------------------------------------------------------
    'Capture the number of target cells
        CellCount = AnchorRange.Cells.Count

'------------------------------------------------------------------------------
' VALIDATE OPTIONAL NAMES INPUT
'------------------------------------------------------------------------------
    'Validate the supplied naming input against the target-cell count
        DEMO_CB_ValidateNamesInput AnchorRange, CheckBoxNames

'------------------------------------------------------------------------------
' CREATE ONE CHECK BOX PER CELL
'------------------------------------------------------------------------------
    'Create one Forms check box per target cell
        CellIndex = 0

        For Each TargetCell In AnchorRange.Cells

        'Advance the current 1-based cell index
            CellIndex = CellIndex + 1

        'Resolve the checkbox name for the current target cell
            ResolvedName = DEMO_CB_ResolveName(TargetCell, CheckBoxNames, CellIndex)

        'Create the current Forms check box
            DEMO_AddOneFormsCheckBox WS, TargetCell, ResolvedName

        Next TargetCell

End Sub

Private Sub DEMO_CB_ValidateNamesInput( _
    ByVal AnchorRange As Range, _
    ByVal CheckBoxNames As Variant)
'
'------------------------------------------------------------------------------
'                           VALIDATE CHECK-BOX NAMES
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the optional CheckBoxNames input for DEMO_CB_AddForms
'
' WHY THIS EXISTS
'   The public checkbox builder accepts several naming forms, so validation is
'   centralized here to keep the public routine smaller and easier to read
'
' INPUTS
'   AnchorRange
'     One-cell or multi-cell target range
'
'   CheckBoxNames
'     Optional name input passed by the caller
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Counts the cells of AnchorRange
'   - Returns when CheckBoxNames was omitted
'   - Single cell: accepts a String or a one-element array; any other input
'     raises vbObjectError + 2145
'   - Multi-cell: requires an array (vbObjectError + 2146) with exactly one
'     element per cell (vbObjectError + 2147)
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Range.Cells.Count
'
' NOTES
'   - Only the shape of the input is checked; DEMO_CB_ResolveName trims each
'     element and replaces a blank one with the default name
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_CB_ValidateNamesInput"

    Dim CellCount   As Long 'Number of target cells
    Dim ArrLo       As Long 'Array lower bound
    Dim ArrHi       As Long 'Array upper bound

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Capture the number of target cells
        CellCount = AnchorRange.Cells.Count

'------------------------------------------------------------------------------
' EXIT WHEN OMITTED
'------------------------------------------------------------------------------
    'Leave when the optional names input was omitted
        If IsMissing(CheckBoxNames) Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' VALIDATE SINGLE-CELL CASE
'------------------------------------------------------------------------------
    'For a single-cell anchor, allow either a scalar string or a one-element array
        If CellCount = 1 Then

        'Accept a scalar string
            If VarType(CheckBoxNames) = vbString Then
                Exit Sub
            End If

        'Accept a one-element array
            If IsArray(CheckBoxNames) Then
                ArrLo = LBound(CheckBoxNames)
                ArrHi = UBound(CheckBoxNames)

                If ArrHi - ArrLo + 1 = 1 Then
                    Exit Sub
                End If
            End If

        'Reject all other input shapes
            Err.Raise vbObjectError + 2145, _
                      "M_DEMO_BUILDER.DEMO_CB_ValidateNamesInput", _
                      "For a single-cell anchor, CheckBoxNames must be omitted, be a String, or be a one-element array"

        End If

'------------------------------------------------------------------------------
' VALIDATE MULTI-CELL CASE
'------------------------------------------------------------------------------
    'For a multi-cell anchor, require an array when names were supplied
        If Not IsArray(CheckBoxNames) Then
            Err.Raise vbObjectError + 2146, _
                      "M_DEMO_BUILDER.DEMO_CB_ValidateNamesInput", _
                      "For a multi-cell anchor, CheckBoxNames must be omitted or be a one-dimensional array"
        End If

    'Validate the supplied array size against the target-cell count
        ArrLo = LBound(CheckBoxNames)
        ArrHi = UBound(CheckBoxNames)

        If ArrHi - ArrLo + 1 <> CellCount Then
            Err.Raise vbObjectError + 2147, _
                      "M_DEMO_BUILDER.DEMO_CB_ValidateNamesInput", _
                      "When supplied for a multi-cell anchor, CheckBoxNames must contain exactly one name per target cell"
        End If

End Sub

Private Function DEMO_CB_ResolveName( _
    ByVal TargetCell As Range, _
    ByVal CheckBoxNames As Variant, _
    ByVal CellIndex As Long) _
    As String
'
'------------------------------------------------------------------------------
'                            RESOLVE CHECK-BOX NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Resolve the effective checkbox name for one target cell
'
' WHY THIS EXISTS
'   Checkbox names may be:
'     - omitted
'     - supplied as a scalar for a single-cell anchor
'     - supplied as an array for one or many cells
'
'   Centralizing that resolution keeps the main builder routine compact
'
' INPUTS
'   TargetCell
'     Current target cell
'
'   CheckBoxNames
'     Optional naming input
'
'   CellIndex
'     1-based current target-cell index
'
' RETURNS
'   String
'     Effective checkbox name
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_CB_ResolveName"

    Dim ArrLo   As Long 'Array lower bound

'------------------------------------------------------------------------------
' RESOLVE OMITTED CASE
'------------------------------------------------------------------------------
    'Auto-generate a name when the optional names input was omitted
        If IsMissing(CheckBoxNames) Then
            DEMO_CB_ResolveName = DEMO_CB_DefaultName(TargetCell)
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE SCALAR STRING CASE
'------------------------------------------------------------------------------
    'Use the supplied scalar string for the single-cell case
        If VarType(CheckBoxNames) = vbString Then
            DEMO_CB_ResolveName = Trim$(CStr(CheckBoxNames))
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE ARRAY CASE
'------------------------------------------------------------------------------
    'Use the corresponding array element for the current cell
        ArrLo = LBound(CheckBoxNames)

        DEMO_CB_ResolveName = Trim$(CStr(CheckBoxNames(ArrLo + CellIndex - 1)))

'------------------------------------------------------------------------------
' NORMALIZE BLANK NAME
'------------------------------------------------------------------------------
    'Fall back to an auto-generated name when the resolved name is blank
        If Len(DEMO_CB_ResolveName) = 0 Then
            DEMO_CB_ResolveName = DEMO_CB_DefaultName(TargetCell)
        End If

End Function

Private Function DEMO_CB_DefaultName( _
    ByVal TargetCell As Range) _
    As String
'
'------------------------------------------------------------------------------
'                         BUILD DEFAULT CHECK-BOX NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Build a deterministic default checkbox name from one target cell
'
' WHY THIS EXISTS
'   Callers should not be forced to invent checkbox names when they do not need
'   to reference the controls later
'
' INPUTS
'   TargetCell
'     Cell that will receive the Forms check box
'
' RETURNS
'   String
'     Auto-generated checkbox name
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN DEFAULT NAME
'------------------------------------------------------------------------------
    'Build a deterministic default name from the cell address
        DEMO_CB_DefaultName = "chk_" & Replace$(TargetCell.Address(False, False), "$", vbNullString)

End Function

Private Sub DEMO_AddOneFormsCheckBox( _
    ByVal WS As Worksheet, _
    ByVal AnchorCell As Range, _
    ByVal CheckBoxName As String)
'
'------------------------------------------------------------------------------
'                           ADD ONE FORMS CHECK BOX
'------------------------------------------------------------------------------
' PURPOSE
'   Create one Forms check box centered inside one anchor cell
'
' WHY THIS EXISTS
'   The multi-cell public builder delegates the actual per-cell control
'   creation to this helper
'
' INPUTS
'   WS
'     Worksheet receiving the Forms check box
'
'   AnchorCell
'     Single target cell
'
'   CheckBoxName
'     Name assigned to the created Forms check box
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Raises vbObjectError + 2148 for a blank CheckBoxName and
'     vbObjectError + 2149 when AnchorCell is not a single cell
'   - Deletes an existing same-sheet Forms check box with that name, ignoring
'     the error when there is none
'   - Computes an 18 by 16 box centered in the cell
'   - Adds the check box and sets its name, an empty caption, unchecked value,
'     no linked cell, xlMoveAndSize placement and PrintObject TRUE
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Worksheet.CheckBoxes: Add and Delete
'   - CheckBox Name, Caption, Value, LinkedCell, Placement and PrintObject
'
' NOTES
'   - The size is fixed at 18 by 16 points
'
' UPDATED
'   2026-04-19
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_AddOneFormsCheckBox"

    Dim Cb          As CheckBox 'Created Forms check box
    Dim BoxWidth    As Double   'Check-box width
    Dim BoxHeight   As Double   'Check-box height
    Dim BoxLeft     As Double   'Resolved left position
    Dim BoxTop      As Double   'Resolved top position

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a blank control name
        If Len(Trim$(CheckBoxName)) = 0 Then
            Err.Raise vbObjectError + 2148, _
                      "M_DEMO_BUILDER.DEMO_AddOneFormsCheckBox", _
                      "CheckBoxName cannot be blank"
        End If

    'Reject multi-cell anchors in the per-cell helper
        If AnchorCell.Cells.CountLarge <> 1 Then
            Err.Raise vbObjectError + 2149, _
                      "M_DEMO_BUILDER.DEMO_AddOneFormsCheckBox", _
                      "AnchorCell must be a single cell"
        End If

'------------------------------------------------------------------------------
' DELETE EXISTING SAME-NAME CHECK BOX
'------------------------------------------------------------------------------
    'Delete an existing same-sheet Forms check box with the same name when present
        On Error Resume Next
        WS.CheckBoxes(CheckBoxName).Delete
        On Error GoTo 0

'------------------------------------------------------------------------------
' COMPUTE CHECK-BOX POSITION
'------------------------------------------------------------------------------
    'Define the standard visual check-box size
        BoxWidth = 18

    'Define the standard visual check-box size
        BoxHeight = 16

    'Center the control horizontally inside the anchor cell
        BoxLeft = AnchorCell.Left + (AnchorCell.Width - BoxWidth) / 2

    'Center the control vertically inside the anchor cell
        BoxTop = AnchorCell.Top + (AnchorCell.Height - BoxHeight) / 2

'------------------------------------------------------------------------------
' CREATE CHECK BOX
'------------------------------------------------------------------------------
    'Create the Forms check box
        Set Cb = WS.CheckBoxes.Add(BoxLeft, BoxTop, BoxWidth, BoxHeight)

    'Assign the requested control name
        Cb.Name = CheckBoxName

    'Remove the visible check-box caption
        Cb.Caption = vbNullString

    'Initialize the control to unchecked
        Cb.Value = xlOff

    'Clear any linked cell
        Cb.LinkedCell = vbNullString

    'Tie the control position and size to its cells
        Cb.Placement = xlMoveAndSize

    'Ensure the control is printed with the sheet
        Cb.PrintObject = True

End Sub

Public Function DEMO_CB_TryGetState( _
    ByVal WS As Worksheet, _
    ByVal ControlName As String, _
    ByRef IsChecked As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                             READ CHECK-BOX STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the checked state of a Forms or ActiveX check box on a worksheet
'
' WHY THIS EXISTS
'   A check box on a demo sheet is either a Forms control or an ActiveX control,
'   and the two are read through different objects. This routine hides the
'   difference and reports failure through FailMsg instead of raising (#73).
'
' INPUTS
'   WS
'     Worksheet containing the control
'
'   ControlName
'     Name of the control; for an ActiveX control, the shape and the
'     OLEObject share it
'
'   IsChecked
'     Receives TRUE when the check box is checked, FALSE otherwise
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => control found and state read successfully
'   FALSE => control missing, not a check box, mixed, or unreadable
'
' BEHAVIOR
'   - Dispatches on the control type through DEMO_CB_ResolveControl, so an
'     ActiveX check box reaches the OLEObject path
'   - A mixed state (a Forms value other than xlOn or xlOff, or an ActiveX
'     Value of Null) is refused: the caller treats the box as unavailable and
'     leaves the UI element unchanged
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - DEMO_CB_ResolveControl
'   - DEMO_CB_ErrorText
'   - Shape.ControlFormat.Value for a Forms check box
'   - OLEObject.Object, Value through CallByName, for an ActiveX check box
'
' NOTES
'   - IsChecked is FALSE on entry and whenever the function returns FALSE
'
' HISTORY
'   2026-11-14 - Moved from M_EXCEL_UI_DEMO; ActiveX dispatch and mixed
'                state (#73).
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_CB_TryGetState"

    Dim Shp         As Shape        'Resolved Forms control
    Dim CheckBoxOle As OLEObject    'Resolved ActiveX control
    Dim ValueOut    As Variant      'Control value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        DEMO_CB_TryGetState = False
        IsChecked = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' READ BY CONTROL TYPE
'------------------------------------------------------------------------------
        Select Case DEMO_CB_ResolveControl(WS, ControlName, Shp, CheckBoxOle, FailMsg)
            Case DEMO_CB_KIND_FORMS
                ValueOut = Shp.ControlFormat.Value
            Case DEMO_CB_KIND_ACTIVEX
                ValueOut = CallByName(CheckBoxOle.Object, "Value", VbGet)
            Case Else
                GoTo SafeExit
        End Select

'------------------------------------------------------------------------------
' MAP THE VALUE
'------------------------------------------------------------------------------
    'Refuse a mixed state; accept only a definite checked or unchecked value
        If IsNull(ValueOut) Then
            FailMsg = DEMO_CB_MIXED_STATE
        ElseIf Shp Is Nothing Then
            IsChecked = CBool(ValueOut)
            DEMO_CB_TryGetState = True
        ElseIf ValueOut = xlOn Then
            IsChecked = True
            DEMO_CB_TryGetState = True
        ElseIf ValueOut = xlOff Then
            IsChecked = False
            DEMO_CB_TryGetState = True
        Else
            FailMsg = DEMO_CB_MIXED_STATE
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        FailMsg = DEMO_CB_ErrorText(Err.Number, Err.Source, Err.Description)
        IsChecked = False
        DEMO_CB_TryGetState = False

End Function

Public Function DEMO_CB_TrySetState( _
    ByVal WS As Worksheet, _
    ByVal ControlName As String, _
    ByVal IsChecked As Boolean, _
    ByRef FailMsg As String) _
    As Boolean
'
'------------------------------------------------------------------------------
'                            WRITE CHECK-BOX STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the checked state of a Forms or ActiveX check box on a worksheet
'
' WHY THIS EXISTS
'   A check box on a demo sheet is either a Forms control or an ActiveX control,
'   and the two are written through different objects. This routine hides the
'   difference and reports failure through FailMsg instead of raising (#73).
'
' INPUTS
'   WS
'     Worksheet containing the control
'
'   ControlName
'     Name of the control; for an ActiveX control, the shape and the
'     OLEObject share it
'
'   IsChecked
'     Requested state
'
'   FailMsg
'     Receives a diagnostic reason when the function returns FALSE
'
' RETURNS
'   TRUE  => control found and updated successfully
'   FALSE => control missing, not a check box, or not writable
'
' BEHAVIOR
'   - Dispatches on the control type through DEMO_CB_ResolveControl
'   - Writes xlOn / xlOff to a Forms check box and True / False to an
'     ActiveX check box; either write replaces a mixed state
'
' ERROR POLICY
'   - Does NOT raise to callers
'   - Returns FALSE and populates FailMsg on failure
'
' DEPENDENCIES
'   - DEMO_CB_ResolveControl
'   - DEMO_CB_ErrorText
'   - Shape.ControlFormat.Value for a Forms check box
'   - OLEObject.Object, Value through CallByName, for an ActiveX check box
'
' NOTES
'   - When the resolver refuses the control, the result stays FALSE and FailMsg
'     holds the reason it set
'
' HISTORY
'   2026-11-14 - Moved from M_EXCEL_UI_DEMO; ActiveX dispatch (#73).
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_DEMO_BUILDER.DEMO_CB_TrySetState"

    Dim Shp         As Shape        'Resolved Forms control
    Dim CheckBoxOle As OLEObject    'Resolved ActiveX control

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

        DEMO_CB_TrySetState = False
        FailMsg = vbNullString

'------------------------------------------------------------------------------
' WRITE BY CONTROL TYPE
'------------------------------------------------------------------------------
        Select Case DEMO_CB_ResolveControl(WS, ControlName, Shp, CheckBoxOle, FailMsg)
            Case DEMO_CB_KIND_FORMS
                If IsChecked Then
                    Shp.ControlFormat.Value = xlOn
                Else
                    Shp.ControlFormat.Value = xlOff
                End If
                DEMO_CB_TrySetState = True
            Case DEMO_CB_KIND_ACTIVEX
                CallByName CheckBoxOle.Object, "Value", VbLet, IsChecked
                DEMO_CB_TrySetState = True
            Case Else
            'Leave the result False for any other control kind
        End Select

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        FailMsg = DEMO_CB_ErrorText(Err.Number, Err.Source, Err.Description)
        DEMO_CB_TrySetState = False

End Function

Private Function DEMO_CB_ResolveControl( _
    ByVal WS As Worksheet, _
    ByVal ControlName As String, _
    ByRef Shp As Shape, _
    ByRef CheckBoxOle As OLEObject, _
    ByRef FailMsg As String) _
    As Long
'
'------------------------------------------------------------------------------
'                          RESOLVE CHECK-BOX CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Finds a check box by name and reports whether it is a Forms or an ActiveX
'   check box
'
' WHY THIS EXISTS
'   An ActiveX control is also a worksheet shape with the same name. Testing
'   the shape for a Forms control first and refusing anything else would
'   reject every ActiveX check box before its OLEObject is reached, so the
'   shape type decides the path.
'
' INPUTS
'   WS
'     Worksheet containing the control
'
'   ControlName
'     Name of the control; for an ActiveX control, the shape and the OLEObject
'     share it
'
'   Shp
'     Receives the shape of a Forms check box; Nothing for any other result
'
'   CheckBoxOle
'     Receives the OLEObject of an ActiveX check box; Nothing for any other
'     result
'
'   FailMsg
'     Receives the reason when the result is DEMO_CB_KIND_NONE
'
' RETURNS
'   Long
'     DEMO_CB_KIND_FORMS with Shp set, DEMO_CB_KIND_ACTIVEX with CheckBoxOle
'     set, or DEMO_CB_KIND_NONE with FailMsg set
'
' BEHAVIOR
'   - msoFormControl shape: accepted when its FormControlType is xlCheckBox
'   - msoOLEControlObject shape: resolved through WS.OLEObjects and accepted
'     when its progID names a CheckBox
'   - any other shape: refused
'   - no shape: WS.OLEObjects is still tried, then "check box not found"
'
' ERROR POLICY
'   Raises only when Excel fails while reading a control it found; the
'   callers turn that into FailMsg
'
' DEPENDENCIES
'   - Worksheet.Shapes and OLEObjects
'   - Shape.Type, FormControlType and OLEObject.progID
'   - DEMO_CB_KIND_NONE, DEMO_CB_KIND_FORMS and DEMO_CB_KIND_ACTIVEX
'
' NOTES
'   - A missing name is expected, so both lookups run under
'     On Error Resume Next and test the result for Nothing
'   - A control is an ActiveX check box when its progID contains "CheckBox",
'     compared without regard to case
'
' HISTORY
'   2026-11-14 - Added for #73.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        DEMO_CB_ResolveControl = DEMO_CB_KIND_NONE
        Set Shp = Nothing
        Set CheckBoxOle = Nothing

        If WS Is Nothing Then
            FailMsg = "worksheet reference is Nothing"
            Exit Function
        End If

        If Len(Trim$(ControlName)) = 0 Then
            FailMsg = "control name is blank"
            Exit Function
        End If

'------------------------------------------------------------------------------
' LOOK UP THE SHAPE
'------------------------------------------------------------------------------
    'A missing name is an expected outcome, not an error
        On Error Resume Next
        Set Shp = WS.Shapes(ControlName)
        On Error GoTo 0

'------------------------------------------------------------------------------
' DISPATCH BY SHAPE TYPE
'------------------------------------------------------------------------------
        If Not Shp Is Nothing Then
            Select Case Shp.Type
                Case msoFormControl
                    If Shp.FormControlType <> xlCheckBox Then
                        FailMsg = "Forms control exists but is not a CheckBox"
                        Set Shp = Nothing
                        Exit Function
                    End If
                    DEMO_CB_ResolveControl = DEMO_CB_KIND_FORMS
                    Exit Function

                Case msoOLEControlObject
                    Set Shp = Nothing

                Case Else
                    FailMsg = "shape exists but is not a Forms or ActiveX control"
                    Set Shp = Nothing
                    Exit Function
            End Select
        End If

'------------------------------------------------------------------------------
' RESOLVE THE ACTIVEX CONTROL
'------------------------------------------------------------------------------
        On Error Resume Next
        Set CheckBoxOle = WS.OLEObjects(ControlName)
        On Error GoTo 0

        If CheckBoxOle Is Nothing Then
            FailMsg = "check box not found"
            Exit Function
        End If

        If InStr(1, CheckBoxOle.progID, "CheckBox", vbTextCompare) = 0 Then
            FailMsg = "ActiveX control exists but is not a CheckBox"
            Set CheckBoxOle = Nothing
            Exit Function
        End If

        DEMO_CB_ResolveControl = DEMO_CB_KIND_ACTIVEX

End Function

Private Function DEMO_CB_ErrorText( _
    ByVal Number As Long, _
    ByVal Source As String, _
    ByVal Description As String) _
    As String
'
'------------------------------------------------------------------------------
'                         FORMAT CHECK-BOX ERROR TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Formats a runtime error for a check-box FailMsg
'
' HISTORY
'   2026-11-14 - Added for #73.
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

        DEMO_CB_ErrorText = CStr(Number) & ": " & Description & _
            IIf(Len(Source) > 0, " | Source: " & Source, vbNullString)

End Function
