' ============================================================================
'  Suspect Report -- collapsible tree grouping macro
' ============================================================================
'
'  Excel's native row grouping (the +/- collapse buttons) can only be
'  created from inside Excel itself -- the SharePoint/Graph REST API that
'  suspect.py uses to write this sheet's data has no equivalent action.
'  So suspect.py writes the tree flat (every row visible) plus a hidden
'  "level" marker in column A (blank/1/2/3), and this macro reads that
'  column to build the real grouping the moment anyone opens the sheet.
'
'  It only regroups when the data has actually changed since the last
'  time it ran -- see the marker-cell check in ApplyTreeGrouping below.
'  Re-grouping already-grouped, unchanged rows is a pointless no-op that
'  can also visibly re-collapse a tree someone had manually expanded, so
'  this is a hard requirement, not just an optimization.
'
'  Grouping is applied (the +/- buttons exist at every level) but nothing
'  starts hidden -- the tree is fully expanded by default, and it's up to
'  whoever's looking at it to manually collapse whatever they don't need.
'  Suspect-level rows (the leaves) alternate light green / white per item,
'  matching suspect.py's original look.
'
'  ---------------------------------------------------------------------
'  ONE-TIME SETUP (do this once, then never again):
'  ---------------------------------------------------------------------
'  1. Open the workbook in Excel Desktop.
'  2. Press Alt+F11 to open the VBA editor.
'  3. In the Project Explorer (left side), find "Suspect Report" under
'     Microsoft Excel Objects -- double-click it to open its code module,
'     and paste in the "SUSPECT REPORT SHEET MODULE" section below.
'  4. Double-click "ThisWorkbook" (also under Microsoft Excel Objects) and
'     paste in the "THISWORKBOOK MODULE" section below.
'  5. Right-click any Excel Objects item -> Insert -> Module, to create a
'     new standard module (this is "Module3" below -- if Excel names
'     yours something else, use that name instead everywhere this file
'     says Module3). Paste in the "STANDARD MODULE" section below.
'  6. Save the workbook as macro-enabled (.xlsm -- it already is one) and
'     make sure macros are enabled when it's opened.
'
'  That's it. From then on, every time suspect.py refreshes the data,
'  opening or clicking into the "Suspect Report" tab will automatically
'  (re)build the collapsible tree.
' ============================================================================


' ============================================================================
'  STANDARD MODULE  (Module3 -- rename below too if yours is named differently)
' ============================================================================
Option Explicit

Private Const SHEET_NAME As String = "Suspect Report"
Private Const LEVEL_COL As String = "A"
Private Const DATA_START_ROW As Long = 4

Private Const BUILD_MARKER_CELL As String = "O1"   ' Python writes a fresh value here every run
Private Const GROUPED_MARKER_CELL As String = "O2"  ' This macro writes here once it has grouped that version
Private Const LAST_TREE_ROW_CELL As String = "O3"   ' Python writes the last tree row here (summary block starts after it)

Public Sub ApplyTreeGrouping()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim buildMarker As String
    buildMarker = CStr(ws.Range(BUILD_MARKER_CELL).Value)
    If buildMarker = "" Then Exit Sub  ' suspect.py hasn't written any data yet

    ' Already grouped this exact data version -- do nothing. Re-running
    ' the grouping on unchanged data would also stomp on anyone who had
    ' manually expanded/collapsed parts of the tree since.
    If buildMarker = CStr(ws.Range(GROUPED_MARKER_CELL).Value) Then Exit Sub

    Dim lastTreeRow As Long
    If Not IsNumeric(ws.Range(LAST_TREE_ROW_CELL).Value) Then Exit Sub
    lastTreeRow = CLng(ws.Range(LAST_TREE_ROW_CELL).Value)
    If lastTreeRow < DATA_START_ROW Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanFail

    ' Reset any previous grouping first, so a shrunk or reshaped tree
    ' from the new data doesn't inherit stale group boundaries.
    Dim clearLastRow As Long
    clearLastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If clearLastRow >= DATA_START_ROW Then
        ws.Rows(DATA_START_ROW & ":" & clearLastRow).ClearOutline
    End If

    ws.Outline.SummaryRow = xlSummaryAbove

    ' Suspect-row (level 3) striping alternates light green (#E2EFDA) /
    ' white, reset at the start of each item's suspect block -- matches
    ' suspect.py's original openpyxl design before grouping moved here.
    Dim stripeColor As Long
    stripeColor = RGB(226, 239, 218)  ' #E2EFDA

    Dim r As Long, lvl As Variant, prevLvl As Variant
    Dim stripeOn As Boolean
    prevLvl = Empty
    For r = DATA_START_ROW To lastTreeRow
        lvl = ws.Cells(r, LEVEL_COL).Value

        ' Outline level only -- rows are left visible below, so the
        ' +/- buttons exist for anyone who wants to manually collapse,
        ' but nothing starts hidden.
        If lvl = 1 Or lvl = 2 Or lvl = 3 Then
            ws.Rows(r).OutlineLevel = CLng(lvl) + 1   ' Excel outline levels are 1-based (1 = no grouping)
        Else
            ws.Rows(r).OutlineLevel = 1
        End If
        ws.Rows(r).Hidden = False

        If lvl = 3 Then
            If prevLvl <> 3 Then stripeOn = False  ' first suspect row under an item: no stripe
            If stripeOn Then
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Interior.Color = stripeColor
            Else
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Interior.ColorIndex = xlNone
            End If
            stripeOn = Not stripeOn
        End If

        prevLvl = lvl
    Next r

    ws.Range(GROUPED_MARKER_CELL).Value = buildMarker

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


' ============================================================================
'  SUSPECT REPORT SHEET MODULE
'  (paste into the code-behind for the "Suspect Report" sheet itself --
'  NOT a standard module. Double-click "Suspect Report" in the VBA
'  editor's Project Explorer to get to it.)
' ============================================================================
Option Explicit

Private Sub Worksheet_Activate()
    Module3.ApplyTreeGrouping
End Sub


' ============================================================================
'  THISWORKBOOK MODULE
'  (covers the case where the file opens with the Suspect Report tab
'  already active, since Worksheet_Activate only fires on a change of
'  tab, not on initial load of the front tab.)
'
'  IF ThisWorkbook ALREADY HAS a Workbook_Open sub (very likely, if this
'  workbook does anything else on open): do NOT paste a second one --
'  VBA will reject it as an "Ambiguous name". Instead, just add the
'  3-line If block below into the EXISTING Workbook_Open sub, anywhere
'  before its End Sub:
'
'      If ActiveSheet.Name = "Suspect Report" Then
'          Module3.ApplyTreeGrouping
'      End If
'
'  Only paste the full Private Sub below if ThisWorkbook has no
'  Workbook_Open sub at all yet.
' ============================================================================
Option Explicit

Private Sub Workbook_Open()
    If ActiveSheet.Name = "Suspect Report" Then
        Module3.ApplyTreeGrouping
    End If
End Sub
