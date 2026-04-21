Attribute VB_Name = "M_DEMO_BUILDER"
'==============================================================================
' MODULE: M_DEMO_BUILDER
'------------------------------------------------------------------------------
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
'
'   Centralizing that logic in one module keeps DEMO-sheet construction:
'     - more consistent
'     - easier to maintain
'     - easier to reuse across projects
'
' DESIGN
'   - This module is workbook-oriented and defaults to ThisWorkbook for build
'     operations unless a workbook is explicitly passed by the caller.
'   - Formatting helpers are kept generic enough to be reusable outside one
'     specific DEMO sheet.
'   - Button creation supports an optional action macro so the helper is not
'     hard-wired to one single placeholder routine.
'   - Fast-mode helpers centralize temporary Application-state changes used
'     during sheet-building operations.
'
' UPDATED
'   2026-04-19
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
    Option Explicit         'Force explicit declaration of all variables
    Option Private Module   'Hide support routines from the Macro dialog

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    Public Const COLOR_TITLE        As Long = 0           'RGB(0, 0, 0)
    Public Const COLOR_SUBTITLE     As Long = 6568980
    Public Const COLOR_HEADER       As Long = 6568980
    Public Const COLOR_SUBHEADER    As Long = 13145700
    Public Const COLOR_INPUT        As Long = 13172735
    Public Const COLOR_LABEL_BACK   As Long = 15787740
    Public Const COLOR_LABEL_FONT   As Long = 0
    'Public Const COLOR_BUTTON       As Long = &HFF660A
    Public Const COLOR_BUTTON       As Long = 13145700
    Public Const COLOR_TABLE        As Long = 16774635
    Public Const COLOR_HELP         As Long = 6697728
    'Buttons color tokens
    Public Const BTN_PRIMARY_N      As Long = &HFF660A  ' #0A66FF (BGR in VBA)
    Public Const BTN_PRIMARY_H      As Long = &HB34700  ' #0047B3
    Public Const BTN_PRIMARY_P      As Long = &H853000  ' #003085
    Public Const WHITE              As Long = &HFFFFFF


'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const DEFAULT_BOOL_WS       As String = "DEMO"
    Private Const DEFAULT_BOOLLIST_NAME As String = "BoolList"

'------------------------------------------------------------------------------
' PRIVATE VARIABLES
'------------------------------------------------------------------------------
    Private mResetWorkbookName As String
    Private mResetSheetName    As String
    Private mResetShapeName    As String

'------------------------------------------------------------------------------
' PUBLIC TYPES
'------------------------------------------------------------------------------
    Public Type tDEMOFastModeState
        ScreenUpdating  As Boolean       'Saved Application.ScreenUpdating
        EnableEvents    As Boolean       'Saved Application.EnableEvents
        DisplayAlerts   As Boolean       'Saved Application.DisplayAlerts
        Calculation     As XlCalculation 'Saved Application.Calculation
    End Type

'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
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
    



Public Sub DEMO_Build_DemoTemplate( _
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
'==============================================================================
'                           BUILD DEMO TEMPLATE
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
'   None
'
' BEHAVIOR
'   - Uses the requested workbook or defaults to ThisWorkbook
'   - Gets or creates the requested worksheet
'   - Resets the sheet before rebuilding it
'   - Derives the separator column from ContentColumns
'   - Derives the first hidden column from the separator column
'   - Writes title/subtitle bands from column B to the last content column
'   - Applies configurable row/column sizing and view normalization
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - DEMO_GetOrCreateSheet
'   - DEMO_Reset_Sheet
'   - DEMO_Write_BandHeader
'   - DEMO_ColumnLetter
'
' NOTES
'   - Window-level settings can only be applied safely through the active window
'   - SeparatorColumns and HideColumnsFrom are retained in the signature only
'     for backward compatibility
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WB                          As Workbook       'Target workbook
    Dim WS                          As Worksheet      'Target worksheet

    Dim MaxRow                      As Long           'Last worksheet row
    Dim MaxCol                      As Long           'Last worksheet column

    Dim ContentParts                As Variant        'Split ContentColumns parts
    Dim ContentText                 As String         'Trimmed ContentColumns text
    Dim ContentFirstColIndex        As Long           'Resolved first content-column index
    Dim ContentLastColIndex         As Long           'Resolved last content-column index

    Dim EffectiveSeparatorColIndex  As Long           'Derived separator-column index
    Dim EffectiveHideFromColIndex   As Long           'Derived first hidden-column index
    Dim EffectiveSeparatorColumns   As String         'Derived separator-column address
    Dim EffectiveHideColumnsFrom    As String         'Derived first hidden-column text

    Dim LastVisibleRow              As Long           'Last visible worksheet row
    Dim LastVisibleCol              As Long           'Last visible worksheet column
    Dim ScrollAreaAddress           As String         'Resolved worksheet ScrollArea text

    Dim TitleBandRange              As Range          'Resolved title-band range
    Dim SubTitleBandRange           As Range          'Resolved subtitle-band range

    Dim TmpColIndex                 As Long           'Temporary column index for normalization

    Dim SavedErrNum                 As Long           'Captured error number
    Dim SavedErrSrc                 As String         'Captured error source
    Dim SavedErrDesc                As String         'Captured error description

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject an empty worksheet name
        If Len(Trim$(WS_Name)) = 0 Then
            Err.Raise vbObjectError + 2000, _
                      "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                      "Worksheet name cannot be blank."
        End If

    'Reject invalid freeze-row requests when pane freezing is enabled
        If IsFrozenPane Then
            If FreezeAtRow < 2 Then
                Err.Raise vbObjectError + 2013, _
                          "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                          "FreezeAtRow must be >= 2 when IsFrozenPane = True."
            End If
        End If

    'Reject invalid body-row ranges
        If BodyRowFrom < 1 Or BodyRowTo < BodyRowFrom Then
            Err.Raise vbObjectError + 2014, _
                      "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                      "Invalid body-row range."
        End If

    'Reject invalid width settings
        If LeftMarginColumnWidth <= 0# Or _
           ContentColumnWidth <= 0# Or _
           SeparatorColumnWidth <= 0# Then
            Err.Raise vbObjectError + 2015, _
                      "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                      "Column widths must be positive."
        End If

    'Reject invalid row-height settings
        If TitleRowHeight <= 0# Or _
           SubTitleRowHeight <= 0# Or _
           BodyRowHeight <= 0# Then
            Err.Raise vbObjectError + 2016, _
                      "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                      "Row heights must be positive."
        End If

    'Reject invalid zoom values when explicitly provided
        If ZoomPercent <> 0 Then
            If ZoomPercent < 10 Or ZoomPercent > 400 Then
                Err.Raise vbObjectError + 2017, _
                          "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                          "ZoomPercent must be 0 or between 10 and 400."
            End If
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
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
                      "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                      "TargetWorkbook must be a Workbook object when supplied."
        End If

'------------------------------------------------------------------------------
' RESET SHEET
'------------------------------------------------------------------------------
    'Get or create the target worksheet
        Set WS = DEMO_GetOrCreateSheet(WB, WS_Name)

    'Reset the worksheet to a clean reusable state before rebuilding it
        DEMO_Reset_Sheet WS

'------------------------------------------------------------------------------
' INITIALIZE WORKSHEET LIMITS
'------------------------------------------------------------------------------
    'Capture the last available worksheet row
        MaxRow = WS.Rows.Count

    'Capture the last available worksheet column
        MaxCol = WS.Columns.Count

'------------------------------------------------------------------------------
' RESOLVE CONTENT COLUMN BLOCK
'------------------------------------------------------------------------------
    'Resolve the trimmed content-column specification
        ContentText = Trim$(ContentColumns)

    'Reject a blank content-column specification
        If Len(ContentText) = 0 Then
            Err.Raise vbObjectError + 2021, _
                      "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                      "ContentColumns cannot be blank."
        End If

    'Split the content-column specification
        ContentParts = Split(ContentText, ":")

    'Resolve the content-column range when a single column was supplied
        If UBound(ContentParts) = 0 Then
            ContentFirstColIndex = WS.Range(Trim$(ContentParts(0)) & "1").Column
            ContentLastColIndex = ContentFirstColIndex

    'Resolve the content-column range when a standard two-part range was supplied
        ElseIf UBound(ContentParts) = 1 Then
            ContentFirstColIndex = WS.Range(Trim$(ContentParts(0)) & "1").Column
            ContentLastColIndex = WS.Range(Trim$(ContentParts(1)) & "1").Column

    'Reject invalid multi-part content-column specifications
        Else
            Err.Raise vbObjectError + 2022, _
                      "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                      "ContentColumns must resolve to one column or one contiguous column range."
        End If

    'Normalize reversed content-column specifications
        If ContentLastColIndex < ContentFirstColIndex Then
            TmpColIndex = ContentFirstColIndex
            ContentFirstColIndex = ContentLastColIndex
            ContentLastColIndex = TmpColIndex
        End If

    'Reject content-column blocks that do not leave room for separator + hidden block
        If ContentLastColIndex > MaxCol - 2 Then
            Err.Raise vbObjectError + 2023, _
                      "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                      "ContentColumns must leave room for a separator column and at least one hidden trailing column."
        End If

    'Reject content-column blocks that end before column B
        If ContentLastColIndex < 2 Then
            Err.Raise vbObjectError + 2024, _
                      "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                      "ContentColumns must extend at least through column B."
        End If

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
' VALIDATE ROW-HIDE REQUEST
'------------------------------------------------------------------------------
    'Reject invalid row-hide requests
        If HideRowsFrom < 1 Or HideRowsFrom > MaxRow Then
            Err.Raise vbObjectError + 2019, _
                      "M_DEMO_BUILDER.DEMO_Build_DemoTemplate", _
                      "HideRowsFrom must be between 1 and the worksheet last row."
        End If

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
        If Len(Trim$(LeftMarginColumns)) > 0 Then
            WS.Columns(LeftMarginColumns).ColumnWidth = LeftMarginColumnWidth
        End If

    'Apply the configured main content-block widths
        WS.Columns(ContentColumns).ColumnWidth = ContentColumnWidth

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

            'Apply formula-bar visibility
                Application.DisplayFormulaBar = ShowFormulaBar
        End If

CleanExit:
'------------------------------------------------------------------------------
' RE-RAISE ERROR AFTER CLEANUP
'------------------------------------------------------------------------------
    'Re-raise the captured error after cleanup when needed
        If SavedErrNum <> 0 Then
            Err.Raise SavedErrNum, SavedErrSrc, SavedErrDesc
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error before cleanup
        SavedErrNum = Err.Number
        SavedErrSrc = Err.Source
        SavedErrDesc = Err.Description

    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub


Private Function DEMO_ColumnLetter( _
    ByVal ColIndex As Long) _
    As String
'
'==============================================================================
'                             COLUMN LETTER
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
' ERROR POLICY
'   Raises when ColIndex is less than 1
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim N                   As Long          'Working column index
    Dim R                   As Long          'Current base-26 remainder
    Dim S                   As String        'Accumulated column text

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
Public Sub DEMO_Begin_FastMode( _
    ByRef StateOut As tDEMOFastModeState)
'
'==============================================================================
'                              BEGIN FAST MODE
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
'     Output structure that receives the saved Application state.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Captures current values for:
'       * ScreenUpdating
'       * EnableEvents
'       * DisplayAlerts
'       * Calculation
'   - Applies reduced-noise builder settings
'
' ERROR POLICY
'   Raises errors normally.
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' CAPTURE CURRENT STATE
'------------------------------------------------------------------------------
    'Capture the current Excel Application state
        StateOut.ScreenUpdating = Application.ScreenUpdating
        StateOut.EnableEvents = Application.EnableEvents
        StateOut.DisplayAlerts = Application.DisplayAlerts
        StateOut.Calculation = Application.Calculation

'------------------------------------------------------------------------------
' APPLY FAST MODE
'------------------------------------------------------------------------------
    'Reduce UI noise while building or rebuilding workbook structures
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
End Sub

Public Sub DEMO_End_FastMode( _
    ByRef StateIn As tDEMOFastModeState)
'
'==============================================================================
'                               END FAST MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Restores a previously saved Excel Application state.
'
' WHY THIS EXISTS
'   This is the cleanup companion to DEMO_Begin_FastMode.
'
' INPUTS
'   StateIn
'     Previously captured Application-state snapshot.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Restores:
'     - ScreenUpdating
'     - EnableEvents
'     - DisplayAlerts
'     - Calculation
'
' ERROR POLICY
'   Raises errors normally.
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' RESTORE SAVED STATE
'------------------------------------------------------------------------------
    'Restore the previously captured Excel Application state
        Application.ScreenUpdating = StateIn.ScreenUpdating
        Application.EnableEvents = StateIn.EnableEvents
        Application.DisplayAlerts = StateIn.DisplayAlerts
        Application.Calculation = StateIn.Calculation
End Sub

Public Sub DEMO_Set_RangeBorder( _
    ByVal Target As Range, _
    Optional ByVal BorderColor As Long = 0, _
    Optional ByVal BorderWeight As XlBorderWeight = xlThin, _
    Optional ByVal IncludeInside As Boolean = True, _
    Optional ByVal InsideColor As Long = vbWhite)
'
'==============================================================================
'                             SET RANGE BORDER
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
'   None.
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
'==============================================================================

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

Public Function DEMO_GetOrCreateSheet( _
    ByVal WB As Workbook, _
    ByVal SheetName As String) _
    As Worksheet
'
'==============================================================================
'                           GET OR CREATE SHEET
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
'   - Searches the workbook for the requested sheet name
'   - Creates a new worksheet if the name is not found
'   - Renames the new sheet to the requested name
'
' ERROR POLICY
'   Raises errors normally.
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS As Worksheet     'Worksheet iterator / result

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing workbook reference
        If WB Is Nothing Then
            Err.Raise vbObjectError + 2001, _
                      "M_DEMO_BUILDER.DEMO_GetOrCreateSheet", _
                      "Workbook reference cannot be Nothing."
        End If

    'Reject a blank worksheet name
        If Len(Trim$(SheetName)) = 0 Then
            Err.Raise vbObjectError + 2002, _
                      "M_DEMO_BUILDER.DEMO_GetOrCreateSheet", _
                      "Worksheet name cannot be blank."
        End If

'------------------------------------------------------------------------------
' SEARCH EXISTING SHEETS
'------------------------------------------------------------------------------
    'Search for an existing worksheet with the requested name
        For Each WS In WB.Worksheets
            If StrComp(WS.Name, SheetName, vbTextCompare) = 0 Then
                Set DEMO_GetOrCreateSheet = WS
                Exit Function
            End If
        Next WS

'------------------------------------------------------------------------------
' CREATE SHEET
'------------------------------------------------------------------------------
    'Create a new worksheet because the requested one does not yet exist
        Set WS = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))

    'Assign the requested name to the new worksheet
        WS.Name = SheetName

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the existing or newly created worksheet
        Set DEMO_GetOrCreateSheet = WS
End Function

Public Sub DEMO_Reset_Sheet( _
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
'==============================================================================
'                                RESET SHEET
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
'   ProtectPassword (optional)
'     Password used to temporarily unprotect and optionally re-protect the sheet
'
'   ReProtectAtEnd (optional)
'     TRUE  => re-protect the sheet at the end when it was successfully unprotected
'     FALSE => leave the sheet unprotected after reset
'
' RETURNS
'   None
'
' ERROR POLICY
'   Raises errors normally except for explicitly marked best-effort cleanup
'   steps guarded with On Error Resume Next
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet   'Resolved worksheet to reset
    Dim i                   As Long        'Reverse loop index
    Dim WasProtected        As Boolean     'TRUE when the sheet was protected on entry
    Dim SavedErrNumber      As Long        'Captured error number
    Dim SavedErrSource      As String      'Captured error source
    Dim SavedErrDescription As String      'Captured error description

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
                              "M_DEMO_BUILDER.DEMO_Reset_Sheet", _
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
                                      "M_DEMO_BUILDER.DEMO_Reset_Sheet", _
                                      "ActiveSheet is not a worksheet."
                        End If

                    'Resolve the active worksheet
                        Set WS = ActiveSheet

                ElseIf TypeOf WS_In Is Worksheet Then

                    'Use the supplied worksheet
                        Set WS = WS_In

                Else
                    Err.Raise vbObjectError + 2103, _
                              "M_DEMO_BUILDER.DEMO_Reset_Sheet", _
                              "WS_In must be a Worksheet when supplied."
                End If

        Else
            Err.Raise vbObjectError + 2104, _
                      "M_DEMO_BUILDER.DEMO_Reset_Sheet", _
                      "WS_In must be omitted or be a Worksheet object."
        End If

'------------------------------------------------------------------------------
' CAPTURE PROTECTION STATE
'------------------------------------------------------------------------------
    'Capture whether the sheet is protected on entry
        WasProtected = WS.ProtectContents Or WS.ProtectDrawingObjects Or WS.ProtectScenarios

'------------------------------------------------------------------------------
' OPTIONAL UNPROTECT
'------------------------------------------------------------------------------
    'Temporarily unprotect the sheet when protection is active
        If WasProtected Then
            If Len(ProtectPassword) > 0 Then
                WS.Unprotect Password:=ProtectPassword
            Else
                Err.Raise vbObjectError + 2100, _
                          "M_DEMO_BUILDER.DEMO_Reset_Sheet", _
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

            'Remove any manual page breaks
                WS.ResetAllPageBreaks

            'Hide page-break indicators on the sheet
                WS.DisplayPageBreaks = False

            On Error GoTo CleanFail
        End If

CleanExit:
'------------------------------------------------------------------------------
' OPTIONAL RE-PROTECT
'------------------------------------------------------------------------------
    'Re-protect the sheet when requested and when it was protected on entry
        If WasProtected And ReProtectAtEnd Then
            WS.Protect Password:=ProtectPassword
        End If

'------------------------------------------------------------------------------
' RE-RAISE ERROR AFTER CLEANUP
'------------------------------------------------------------------------------
    'Re-raise the captured error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description

    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub
Public Sub DEMO_Format_Labels( _
    ByVal TargetRange As Range)
'
'==============================================================================
'                               FORMAT LABELS
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
'   None.
'
' BEHAVIOR
'   - Exits quietly if TargetRange Is Nothing
'   - Applies fill, alignment, and bold formatting
'   - Applies borders
'
' ERROR POLICY
'   Raises errors normally.
'
' UPDATED
'   2026-04-19
'==============================================================================

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
'==============================================================================
'                           FORMAT INPUT CELL
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
'   None.
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
' UPDATED
'   2026-04-19
'==============================================================================

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
'==============================================================================
'                           FORMAT OUTPUT CELL
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
'   None.
'
' BEHAVIOR
'   - Exits quietly if TargetRange Is Nothing
'   - Applies a neutral output style
'   - Applies borders
'
' ERROR POLICY
'   Raises errors normally.
'
' UPDATED
'   2026-04-19
'==============================================================================

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
'==============================================================================
'                            WRITE BAND HEADER
'------------------------------------------------------------------------------
' PURPOSE
'   Writes and formats one left-aligned header band without merged cells.
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
' RETURNS
'   None.
'
' BEHAVIOR
'   - Exits quietly if TargetRange Is Nothing
'   - Ensures the range is not merged
'   - Clears prior contents from the full target range
'   - Writes the requested text into the first cell only
'   - Applies fill / font / alignment formatting
'   - Keeps the band left aligned
'   - Optionally applies a border
'
' ERROR POLICY
'   Raises errors normally.
'
' UPDATED
'   2026-04-19
'==============================================================================

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
'==============================================================================
'                         APPLY VALIDATION LIST
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
'   None.
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
' UPDATED
'   2026-04-19
'==============================================================================

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
'==============================================================================
'                         APPLY NUMERIC VALIDATION
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
'   None.
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
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
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
'==============================================================================
'                           CREATE BOOL LIST
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
'   None.
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
'   - DEMO_GetOrCreateSheet
'   - DEMO_Hide_HelperColumns
'
' NOTES
'   - This creates a text list, not a Boolean list
'   - That is intentional so the dropdown remains locale-independent
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WB      As Workbook      'Target workbook
    Dim WS      As Worksheet     'Target worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Use the workbook that contains this module
        Set WB = ThisWorkbook

    'Get or create the target worksheet
        Set WS = DEMO_GetOrCreateSheet(WB, WS_Name)

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
'==============================================================================
'                          APPLY BOOL VALIDATION
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
'   None.
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
' UPDATED
'   2026-04-19
'==============================================================================

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
'==============================================================================
'                            SET WORKBOOK NAME
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
'   None.
'
' BEHAVIOR
'   - Deletes the existing workbook-level name when present
'   - Recreates the name to point to the requested target cell
'
' ERROR POLICY
'   Raises errors normally.
'
' UPDATED
'   2026-04-19
'==============================================================================

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

Public Sub DEMO_Add_DemoButton( _
    ByVal WS As Worksheet, _
    ByVal ShapeName As String, _
    ByVal CaptionText As String, _
    ByVal LeftPos As Double, _
    ByVal TopPos As Double, _
    ByVal ShapeWidth As Double, _
    ByVal ShapeHeight As Double, _
    Optional ByVal ActionMacro As String = "DEMO_DEMOButton_NotAssigned")
'
'==============================================================================
'                             ADD DEMO BUTTON
'------------------------------------------------------------------------------
' PURPOSE
'   Creates one DEMO button as a worksheet Shape.
'
' WHY THIS EXISTS
'   Shape-based buttons are portable, easy to style, and easy to assign to a
'   placeholder or real macro during workbook-layout generation.
'
' INPUTS
'   WS
'     Worksheet where the button should be created.
'
'   ShapeName
'     Internal shape name.
'
'   CaptionText
'     Visible button caption.
'
'   LeftPos
'     Button left position.
'
'   TopPos
'     Button top position.
'
'   ShapeWidth
'     Button width.
'
'   ShapeHeight
'     Button height.
'
'   ActionMacro (optional)
'     Macro name to assign to the shape's OnAction property.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Creates a rounded-rectangle shape
'   - Applies button styling
'   - Writes the requested caption
'   - Assigns the requested action macro
'
' ERROR POLICY
'   Raises errors normally.
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Shp As Shape   'Created button shape

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing worksheet reference
        If WS Is Nothing Then
            Err.Raise vbObjectError + 2007, _
                      "M_DEMO_BUILDER.DEMO_Add_DemoButton", _
                      "Worksheet reference cannot be Nothing."
        End If
    'Reject a blank shape name
        If Len(Trim$(ShapeName)) = 0 Then
            Err.Raise vbObjectError + 2008, _
                      "M_DEMO_BUILDER.DEMO_Add_DemoButton", _
                      "Shape name cannot be blank."
        End If
    'Reject nonpositive dimensions
        If ShapeWidth <= 0 Or ShapeHeight <= 0 Then
            Err.Raise vbObjectError + 2009, _
                      "M_DEMO_BUILDER.DEMO_Add_DemoButton", _
                      "Button width and height must be positive."
        End If

'------------------------------------------------------------------------------
' CREATE SHAPE
'------------------------------------------------------------------------------
    'Create a rounded-rectangle shape for the DEMO button
        Set Shp = WS.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
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
            Shp.OnAction = "'" & ThisWorkbook.Name & "'!DEMO_DEMOButton_NotAssigned"
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
            .Fill.ForeColor.RGB = COLOR_BUTTON
            .Line.ForeColor.RGB = RGB(90, 90, 90)
            .TextFrame2.TextRange.Text = CaptionText
            .TextFrame2.TextRange.Font.Size = 10
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextFrame2.TextRange.Font.Bold = msoTrue
            .TextFrame2.VerticalAnchor = msoAnchorMiddle
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
End Sub

Public Sub DEMO_Hide_HelperColumns( _
    ByVal WS As Worksheet, _
    ByVal ColumnAddress As String)
'
'==============================================================================
'                           HIDE HELPER COLUMNS
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
'   None.
'
' BEHAVIOR
'   - Exits quietly if WS Is Nothing
'   - Exits quietly if ColumnAddress is blank
'   - Hides the requested worksheet columns
'
' ERROR POLICY
'   Raises errors normally.
'
' UPDATED
'   2026-04-19
'==============================================================================

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
'==============================================================================
'                               CREATE TABLE
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
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lo As ListObject   'Created or replaced table

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
'==============================================================================
'                              CLEAR TABLE BODY
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
'   None.
'
' BEHAVIOR
'   - Exits quietly if TableObject Is Nothing
'   - Deletes the DataBodyRange when present
'
' ERROR POLICY
'   Raises errors normally.
'
' UPDATED
'   2026-04-19
'==============================================================================

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
'==============================================================================
'                             APPEND TABLE ROW
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
'   None.
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
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LR          As ListRow   'Newly appended row
    Dim i           As Long      'Column/value index
    Dim LoBound     As Long      'Lower bound of the values array
    Dim HiBound     As Long      'Upper bound of the values array
    Dim MaxWrite    As Long      'Maximum number of values to write

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
' INITIALIZE
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
'==============================================================================
'                     PREPARE LABELED INPUT SECTION
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
'   None
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
' UPDATED
'   2026-04-19
'==============================================================================
    
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
'==============================================================================
'                         WRITE NAMED INPUT ROW
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
'   None
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
' UPDATED
'   2026-04-19
'==============================================================================
    
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


Public Sub DEMO_Add_ButtonGrid( _
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
'==============================================================================
'                           ADD BUTTON GRID
'------------------------------------------------------------------------------
' PURPOSE
'   Creates a regular grid of DEMO buttons from a compact button specification
'
' WHY THIS EXISTS
'   DEMO sheets often contain repeated button-layout code with only:
'     - button name
'     - button caption
'     - optional action macro
'     - grid position
'
' INPUTS
'   ButtonSpecs
'     Variant array where each item is either:
'         Array(ButtonName, ButtonCaption)
'     or:
'         Array(ButtonName, ButtonCaption, ActionMacro)
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i               As Long      'Loop index
    Dim GridRow         As Long      'Zero-based button row
    Dim GridCol         As Long      'Zero-based button column
    Dim ButtonLeft      As Double    'Computed button left position
    Dim ButtonTop       As Double    'Computed button top position
    Dim ActionMacro     As String    'Optional button action macro

'------------------------------------------------------------------------------
' BUILD BUTTON GRID
'------------------------------------------------------------------------------
    For i = LBound(ButtonSpecs) To UBound(ButtonSpecs)

        GridRow = (i - LBound(ButtonSpecs)) \ ButtonsPerRow
        GridCol = (i - LBound(ButtonSpecs)) Mod ButtonsPerRow

        ButtonLeft = AnchorCell.Left + OffsetLeft + (GridCol * (ButtonWidth + GapX))
        ButtonTop = AnchorCell.Top + OffsetTop + (GridRow * (ButtonHeight + GapY))

        'Resolve the optional action macro when supplied
            If UBound(ButtonSpecs(i)) >= 2 Then
                ActionMacro = CStr(ButtonSpecs(i)(2))
            Else
                ActionMacro = "DEMO_DEMOButton_NotAssigned"
            End If

        'Create the current button
            DEMO_Add_DemoButton _
                WS, _
                CStr(ButtonSpecs(i)(0)), _
                CStr(ButtonSpecs(i)(1)), _
                ButtonLeft, _
                ButtonTop, _
                ButtonWidth, _
                ButtonHeight, _
                ActionMacro

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
'==============================================================================
'                         CREATE TABLE SECTION
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
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeaderCount         As Long          'Number of table columns
    Dim i                   As Long          'Loop index
    Dim HeaderRange         As Range         'Formatted header row
    Dim TableRange          As Range         'Range used to create the ListObject
    Dim Lo                  As ListObject    'Created table object

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
' INITIALIZE
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
'                               BUTTON STATES
'
'------------------------------------------------------------------------------
'




Public Sub Btn_ApplyState( _
    ByVal Shp As Shape, _
    ByVal StateName As String)
'
'==============================================================================
'                               APPLY BUTTON STATE
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
'   None
'
' BEHAVIOR
'   Applies the corresponding fill/text styling for the requested state
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Col As Long   'Resolved fill color for the requested state

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


Private Function Btn_CaptureAppearance( _
    ByVal Shp As Shape) _
    As tButtonAppearance
'
'==============================================================================
'                         CAPTURE BUTTON APPEARANCE
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
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S As tButtonAppearance   'Captured appearance snapshot

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
    Btn_CaptureAppearance = S

End Function


Private Sub Btn_RestoreAppearance( _
    ByVal Shp As Shape, _
    ByRef SavedState As tButtonAppearance)
'
'==============================================================================
'                         RESTORE BUTTON APPEARANCE
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
'   None
'
' BEHAVIOR
'   Restores fill, line, text, shadow, and position attributes
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-19
'==============================================================================

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



Public Sub Btn_Click()
'
'==============================================================================
'                                BUTTON CLICK
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
'   None
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
'   - Btn_CaptureAppearance
'   - Btn_ApplyState
'   - Btn_RestoreAppearance
'
' NOTES
'   The short visual delay uses Timer() elapsed-delta logic so it remains safe
'   across midnight rollover
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet           'Caller worksheet
    Dim Shp                 As Shape               'Clicked shape
    Dim SavedState          As tButtonAppearance   'Captured original appearance
    Dim CallerName          As Variant             'Application.Caller value
    Dim StartS              As Double              'Timer() snapshot at delay start
    Dim ElapsedS            As Double              'Elapsed seconds during visual feedback

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
        SavedState = Btn_CaptureAppearance(Shp)

'------------------------------------------------------------------------------
' APPLY PRESSED STATE
'------------------------------------------------------------------------------
    'Apply the pressed color state
        Btn_ApplyState Shp, "Pressed"
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
        Btn_RestoreAppearance Shp, SavedState

End Sub

Public Sub DEMO_PlayButtonFeedback()

'
'==============================================================================
'                       DEMO_PlayButtonFeedback
'------------------------------------------------------------------------------
' PURPOSE
'   Invoke the shared button-click feedback helper on a best-effort basis
'
' WHY THIS EXISTS
'   Several public demo actions want the same visual pressed-state feedback
'   without duplicating local error suppression around Btn_Click
'
' RETURNS
'   None
'
' ERROR POLICY
'   Best effort only
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' APPLY BEST-EFFORT FEEDBACK
'------------------------------------------------------------------------------
    'Play the shared button-click visual feedback on a best-effort basis
        On Error Resume Next
        Btn_Click
        On Error GoTo 0

End Sub


Public Sub DEMO_SB_SetProgress( _
    ByVal CurrentStep As Long, _
    ByVal TotalSteps As Long, _
    ByVal StepText As String)
'
'==============================================================================
'                            SUITE SET PROGRESS
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
'   None
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Pct                 As Double       'Completion ratio
    Dim PercentText         As String       'Formatted completion percentage

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

Public Sub DEMO_AddFormsCheckBox( _
    ByVal WS As Worksheet, _
    ByVal AnchorRange As Range, _
    Optional ByVal CheckBoxNames As Variant)

'
'==============================================================================
'                           DEMO_AddFormsCheckBox
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
'   None
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
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CellCount           As Long          'Number of target cells
    Dim CellIndex           As Long          '1-based current target-cell index
    Dim TargetCell          As Range         'Current target cell
    Dim ResolvedName        As String        'Resolved check-box name for the current cell

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing worksheet reference
        If WS Is Nothing Then
            Err.Raise vbObjectError + 2141, _
                      "M_DEMO_BUILDER.DEMO_AddFormsCheckBox", _
                      "Worksheet reference cannot be Nothing"
        End If

    'Reject a missing anchor range
        If AnchorRange Is Nothing Then
            Err.Raise vbObjectError + 2142, _
                      "M_DEMO_BUILDER.DEMO_AddFormsCheckBox", _
                      "AnchorRange cannot be Nothing"
        End If

    'Reject anchor ranges that do not belong to the target worksheet
        If Not AnchorRange.Worksheet Is WS Then
            Err.Raise vbObjectError + 2143, _
                      "M_DEMO_BUILDER.DEMO_AddFormsCheckBox", _
                      "AnchorRange must belong to the target worksheet"
        End If

    'Reject non-rectangular multi-area ranges
        If AnchorRange.Areas.Count > 1 Then
            Err.Raise vbObjectError + 2144, _
                      "M_DEMO_BUILDER.DEMO_AddFormsCheckBox", _
                      "AnchorRange must be one contiguous rectangular range"
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Capture the number of target cells
        CellCount = AnchorRange.Cells.Count

'------------------------------------------------------------------------------
' VALIDATE OPTIONAL NAMES INPUT
'------------------------------------------------------------------------------
    'Validate the supplied naming input against the target-cell count
        DEMO_ValidateCheckBoxNamesInput AnchorRange, CheckBoxNames

'------------------------------------------------------------------------------
' CREATE ONE CHECK BOX PER CELL
'------------------------------------------------------------------------------
    'Create one Forms check box per target cell
        CellIndex = 0

        For Each TargetCell In AnchorRange.Cells

            'Advance the current 1-based cell index
                CellIndex = CellIndex + 1

            'Resolve the checkbox name for the current target cell
                ResolvedName = DEMO_ResolveCheckBoxName(TargetCell, CheckBoxNames, CellIndex)

            'Create the current Forms check box
                DEMO_AddOneFormsCheckBox WS, TargetCell, ResolvedName

        Next TargetCell

End Sub


Private Sub DEMO_ValidateCheckBoxNamesInput( _
    ByVal AnchorRange As Range, _
    ByVal CheckBoxNames As Variant)

'
'==============================================================================
'                     DEMO_ValidateCheckBoxNamesInput
'------------------------------------------------------------------------------
' PURPOSE
'   Validate the optional CheckBoxNames input for DEMO_AddFormsCheckBox
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
'   None
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CellCount           As Long      'Number of target cells
    Dim ArrLo               As Long      'Array lower bound
    Dim ArrHi               As Long      'Array upper bound

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
                          "M_DEMO_BUILDER.DEMO_ValidateCheckBoxNamesInput", _
                          "For a single-cell anchor, CheckBoxNames must be omitted, be a String, or be a one-element array"

        End If

'------------------------------------------------------------------------------
' VALIDATE MULTI-CELL CASE
'------------------------------------------------------------------------------
    'For a multi-cell anchor, require an array when names were supplied
        If Not IsArray(CheckBoxNames) Then
            Err.Raise vbObjectError + 2146, _
                      "M_DEMO_BUILDER.DEMO_ValidateCheckBoxNamesInput", _
                      "For a multi-cell anchor, CheckBoxNames must be omitted or be a one-dimensional array"
        End If

    'Validate the supplied array size against the target-cell count
        ArrLo = LBound(CheckBoxNames)
        ArrHi = UBound(CheckBoxNames)

        If ArrHi - ArrLo + 1 <> CellCount Then
            Err.Raise vbObjectError + 2147, _
                      "M_DEMO_BUILDER.DEMO_ValidateCheckBoxNamesInput", _
                      "When supplied for a multi-cell anchor, CheckBoxNames must contain exactly one name per target cell"
        End If

End Sub


Private Function DEMO_ResolveCheckBoxName( _
    ByVal TargetCell As Range, _
    ByVal CheckBoxNames As Variant, _
    ByVal CellIndex As Long) _
    As String

'
'==============================================================================
'                       DEMO_ResolveCheckBoxName
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
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ArrLo               As Long      'Array lower bound

'------------------------------------------------------------------------------
' RESOLVE OMITTED CASE
'------------------------------------------------------------------------------
    'Auto-generate a name when the optional names input was omitted
        If IsMissing(CheckBoxNames) Then
            DEMO_ResolveCheckBoxName = DEMO_DefaultCheckBoxName(TargetCell)
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE SCALAR STRING CASE
'------------------------------------------------------------------------------
    'Use the supplied scalar string for the single-cell case
        If VarType(CheckBoxNames) = vbString Then
            DEMO_ResolveCheckBoxName = Trim$(CStr(CheckBoxNames))
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE ARRAY CASE
'------------------------------------------------------------------------------
    'Use the corresponding array element for the current cell
        ArrLo = LBound(CheckBoxNames)

        DEMO_ResolveCheckBoxName = Trim$(CStr(CheckBoxNames(ArrLo + CellIndex - 1)))

'------------------------------------------------------------------------------
' NORMALIZE BLANK NAME
'------------------------------------------------------------------------------
    'Fall back to an auto-generated name when the resolved name is blank
        If Len(DEMO_ResolveCheckBoxName) = 0 Then
            DEMO_ResolveCheckBoxName = DEMO_DefaultCheckBoxName(TargetCell)
        End If

End Function


Private Function DEMO_DefaultCheckBoxName( _
    ByVal TargetCell As Range) _
    As String

'
'==============================================================================
'                        DEMO_DefaultCheckBoxName
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
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN DEFAULT NAME
'------------------------------------------------------------------------------
    'Build a deterministic default name from the cell address
        DEMO_DefaultCheckBoxName = "chk_" & Replace$(TargetCell.Address(False, False), "$", vbNullString)

End Function


Private Sub DEMO_AddOneFormsCheckBox( _
    ByVal WS As Worksheet, _
    ByVal AnchorCell As Range, _
    ByVal CheckBoxName As String)

'
'==============================================================================
'                        DEMO_AddOneFormsCheckBox
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
'   None
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-19
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cb                  As CheckBox    'Created Forms check box
    Dim BoxWidth            As Double      'Check-box width
    Dim BoxHeight           As Double      'Check-box height
    Dim BoxLeft             As Double      'Resolved left position
    Dim BoxTop              As Double      'Resolved top position

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
' INITIALIZE
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

