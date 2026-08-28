Option Explicit

Private Const SHEET_NAME As String = "Suspect Report"
Private Const LEVEL_COL As String = "A"

' Rows 1-62 are reserved for a separate table this macro does not own
' and must never touch. suspect.py's own content starts at row 63: the
' Module Summary first (fixed at SUMMARY_TITLE_ROW), then a blank-row
' gap, then the tree at a fixed TREE_TITLE_ROW (not computed from the
' summary's size, so a future table above can rely on it staying put).
Private Const RESERVED_TOP_ROWS As Long = 62
Private Const SUMMARY_TITLE_ROW As Long = RESERVED_TOP_ROWS + 1   ' 63
Private Const TREE_TITLE_ROW As Long = 71
Private Const TREE_SUBTITLE_ROW As Long = TREE_TITLE_ROW + 1      ' 72
Private Const TREE_HEADER_ROW As Long = TREE_TITLE_ROW + 2        ' 73
Private Const DATA_START_ROW As Long = TREE_HEADER_ROW + 1        ' 74

Private Const BUILD_MARKER_CELL As String = "O71"   ' Python writes a fresh value here every run
Private Const GROUPED_MARKER_CELL As String = "O72"  ' This macro writes here once it has grouped that version
Private Const LAST_TREE_ROW_CELL As String = "O73"   ' Python writes the last tree row here

Public Sub ApplyTreeGrouping()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim buildMarker As String
    buildMarker = CStr(ws.Range(BUILD_MARKER_CELL).Value)
    If buildMarker = "" Then Exit Sub  ' suspect.py hasn't written any data yet

    Dim lastTreeRow As Long
    If Not IsNumeric(ws.Range(LAST_TREE_ROW_CELL).Value) Then Exit Sub
    lastTreeRow = CLng(ws.Range(LAST_TREE_ROW_CELL).Value)
    If lastTreeRow < DATA_START_ROW Then Exit Sub

    ' Same data version as the last grouping: skip only if every row's
    ' outline level still matches column A. Someone who manually cleared
    ' or ungrouped the outline gets repaired here; manual expand/collapse
    ' alone never triggers a rebuild (it doesn't change outline levels),
    ' so we won't stomp on how anyone left the tree folded.
    If buildMarker = CStr(ws.Range(GROUPED_MARKER_CELL).Value) Then
        ' Widths can be repaired on their own -- a full regroup just for
        ' a resized column would stomp on the user's fold state.
        If Not ColumnWidthsAreCorrect(ws) Then ApplyColumnWidths ws
        If TreeIsConsistent(ws, lastTreeRow) Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False  ' merges below would otherwise prompt
    On Error GoTo CleanFail

    ' Reset all previous grouping, all the way down. suspect.py's range
    ' clear wipes values but NOT row outline levels, so when a new report
    ' is shorter than the last one, rows below the new tree would keep
    ' stale groups (blank rows with orphaned +/- buttons) if we only
    ' cleared through the new data's extent.
    ws.Rows(DATA_START_ROW & ":" & ws.Rows.Count).ClearOutline

    ' Links are re-created below as =HYPERLINK() formulas (see
    ' AddJamaHyperlink), so re-running just overwrites them. This clears
    ' any real Hyperlink objects a previous macro version may have left on
    ' column B, so the two link styles never coexist.
    ws.Range(ws.Cells(DATA_START_ROW, 2), ws.Cells(lastTreeRow, 2)).Hyperlinks.Delete

    ws.Outline.SummaryRow = xlSummaryAbove

    FormatBanner ws

    ' Bulk defaults for the whole tree body (this is the suspect-row
    ' look); the per-row pass below only overrides what differs by level.
    ' Underline is reset here so a row that stopped being a link loses the
    ' link underline; linked rows re-add it in AddJamaHyperlink.
    With ws.Range(ws.Cells(DATA_START_ROW, 2), ws.Cells(lastTreeRow, 4))
        .Interior.ColorIndex = xlNone
        .Font.Bold = False
        .Font.Size = 10
        .Font.Color = vbBlack
        .Font.Underline = xlUnderlineStyleNone
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(DATA_START_ROW, 2), ws.Cells(lastTreeRow, 2)) _
        .HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(DATA_START_ROW, 3), ws.Cells(lastTreeRow, 4)) _
        .HorizontalAlignment = xlCenter

    ' One pass over the tree: outline level plus per-level look, both
    ' driven by column A. Suspect rows (level 3) stripe alternating
    ' light green (#E2EFDA) / no fill, restarting at each item's block.
    ' Row visibility is set in one shot after the loop via ShowLevels.
    '
    ' A level of "2x" (instead of a plain 2) means the row is still an
    ' item row, but one appended from the report's "Unlinked Downstream"
    ' section (no downstream link at all, not just a suspect one) rather
    ' than a normal item -- its background gets replaced with peach
    ' (#F8CBAD) so it stands out from ordinary item rows. See ParseLevel.
    Dim r As Long, rawLvl As Variant, baseLvl As Long, prevBaseLvl As Long
    Dim isMissing As Boolean
    Dim stripeOn As Boolean
    Dim rowBand As Range
    prevBaseLvl = -1
    For r = DATA_START_ROW To lastTreeRow
        rawLvl = ws.Cells(r, LEVEL_COL).Value
        ParseLevel rawLvl, baseLvl, isMissing
        Set rowBand = ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))

        If baseLvl = 1 Then      ' module
            ws.Rows(r).OutlineLevel = 2
            rowBand.Interior.Color = RGB(25, 124, 10)   ' #197c0a
            rowBand.Font.Bold = True
            rowBand.Font.Color = vbWhite
        ElseIf baseLvl = 2 Then  ' item (or missing item, appended at the end)
            ws.Rows(r).OutlineLevel = 3
            rowBand.Interior.Color = RGB(138, 193, 218)  ' #8ac1da
            rowBand.Font.Bold = True
            AddJamaHyperlink ws, r
        ElseIf baseLvl = 3 Then  ' suspect
            ws.Rows(r).OutlineLevel = 4
            If prevBaseLvl <> 3 Then stripeOn = False  ' first suspect row under an item: no stripe
            If stripeOn Then rowBand.Interior.Color = RGB(255, 255, 255) ' #FFFFFF
            stripeOn = Not stripeOn
            AddJamaHyperlink ws, r
        Else                 ' category (no level number in column A)
            ws.Rows(r).OutlineLevel = 1
            rowBand.Interior.Color = RGB(56, 88, 153)    ' #385899
            rowBand.Font.Bold = True
            rowBand.Font.Color = vbWhite
            rowBand.Font.Size = 11
        End If

        If isMissing Then
            rowBand.Interior.Color = RGB(248, 203, 173)   ' #F8CBAD -- flags a missing/unlinked row
        End If

        prevBaseLvl = baseLvl
    Next r

    FormatSummary ws

    ' Collapse to Category level only -- Module/Item/Suspect rows start
    ' hidden behind the outline's +/- buttons; Category rows (outline
    ' level 1) are the only ones visible until the user expands further.
    ws.Outline.ShowLevels RowLevels:=1

    ApplyColumnWidths ws

    ws.Range(GROUPED_MARKER_CELL).Value = buildMarker

CleanFail:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ---- Banner, header row, and summary block -----------------------------
' Layout positions mirror suspect.py's row plan: rows 1-62 are reserved
' for a separate table (not built here, left untouched). The Module
' Summary comes first, at a fixed SUMMARY_TITLE_ROW: title, blank,
' summary header, one row per module, Total row. Then a blank-row gap,
' then the tree at a fixed TREE_TITLE_ROW: title, subtitle, column
' headers, then the tree body from DATA_START_ROW through whatever row
' the LAST_TREE_ROW_CELL marker names.

Private Sub FormatBanner(ws As Worksheet)
    ' Title banner across B:D on TREE_TITLE_ROW.
    With ws.Range("B" & TREE_TITLE_ROW & ":D" & TREE_TITLE_ROW)
        .Merge
        .Interior.Color = RGB(56, 88, 153)  ' #385899
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = vbWhite
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(TREE_TITLE_ROW).RowHeight = 30

    ' Subtitle across B:D on TREE_SUBTITLE_ROW.
    With ws.Range("B" & TREE_SUBTITLE_ROW & ":D" & TREE_SUBTITLE_ROW)
        .Merge
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = RGB(102, 102, 102)     ' #666666
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(TREE_SUBTITLE_ROW).RowHeight = 18

    ' Column header row.
    With ws.Range("B" & TREE_HEADER_ROW & ":D" & TREE_HEADER_ROW)
        .Interior.Color = RGB(255, 192, 0)   ' #FFC000
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B" & TREE_HEADER_ROW).HorizontalAlignment = xlLeft
    ws.Range("C" & TREE_HEADER_ROW & ":D" & TREE_HEADER_ROW).HorizontalAlignment = xlCenter
End Sub

Private Sub FormatSummary(ws As Worksheet)
    ' Module Summary block above the tree, at a fixed row (see layout
    ' comment above) -- suspect.py refuses to write if this block would
    ' ever grow large enough to reach TREE_TITLE_ROW, so it's safe to
    ' anchor purely off SUMMARY_TITLE_ROW here.
    Dim titleRow As Long, headerRow As Long, totalRow As Long
    titleRow = SUMMARY_TITLE_ROW
    headerRow = titleRow + 2
    If Trim(CStr(ws.Cells(headerRow, 2).Value)) = "" Then Exit Sub  ' summary block not written

    ' Scan down from the header while column B keeps having a value --
    ' the block's rows are always contiguous (no blank rows in between),
    ' so the last non-blank one is the Total row.
    totalRow = headerRow
    Do While Trim(CStr(ws.Cells(totalRow + 1, 2).Value)) <> ""
        totalRow = totalRow + 1
    Loop
    If totalRow <= headerRow Then Exit Sub  ' header written but no module rows

    With ws.Range(ws.Cells(titleRow, 2), ws.Cells(titleRow, 6))
        .Merge
        .Interior.Color = RGB(16, 124, 65)   ' #107C41
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = vbWhite
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(headerRow, 2), ws.Cells(headerRow, 6))
        .Interior.Color = RGB(112, 173, 71)  ' #70AD47
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Per-module data rows (may be absent if the summary had no modules).
    If totalRow > headerRow + 1 Then
        With ws.Range(ws.Cells(headerRow + 1, 2), ws.Cells(totalRow - 1, 6))
            .Interior.ColorIndex = xlNone
            .Font.Bold = False
            .Font.Size = 11
            .Font.Color = vbBlack
            .VerticalAlignment = xlCenter
        End With
        ws.Range(ws.Cells(headerRow + 1, 2), ws.Cells(totalRow - 1, 2)) _
            .HorizontalAlignment = xlLeft
        ws.Range(ws.Cells(headerRow + 1, 3), ws.Cells(totalRow - 1, 6)) _
            .HorizontalAlignment = xlCenter
    End If

    With ws.Range(ws.Cells(totalRow, 2), ws.Cells(totalRow, 6))
        .Interior.Color = RGB(169, 208, 142) ' #A9D08E
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = vbBlack
        .VerticalAlignment = xlCenter
    End With
    ws.Cells(totalRow, 2).HorizontalAlignment = xlLeft
    ws.Range(ws.Cells(totalRow, 3), ws.Cells(totalRow, 6)).HorizontalAlignment = xlCenter
End Sub

' ---- JAMA hyperlinks ---------------------------------------------------
' Every item (level 2) and suspect (level 3) row links its column-B label
' to the matching JAMA item. The link is built from the numeric ID that
' suspect.py wrote in column C; only that number varies between links.
' projectId is fixed at 47. Applied locally (no Graph API calls), so
' hundreds of links add no throttling risk to the Python build.
'
' The URL pieces are literal strings inside this one sub (not module-level
' Consts) so the sub is fully self-contained -- nothing outside it has to
' be pasted correctly for this to compile and run.

Private Sub AddJamaHyperlink(ws As Worksheet, r As Long)
    ' Link column B of row r to its JAMA item, then apply the link look
    ' (blue underline, Calibri 10) while preserving whatever background
    ' fill the caller already applied to this row (category/module/item
    ' color or suspect stripe) -- the fill must stay visible behind the
    ' link text, not get cleared out by the formula/format changes here.
    '
    ' A =HYPERLINK() formula is used rather than Hyperlinks.Add because the
    ' JAMA URL has a '#' fragment followed by a '?projectId=47' query, and
    ' Hyperlinks.Add splits on '#' and can mangle the query. HYPERLINK
    ' passes the URL string through verbatim.
    Dim numericId As Variant
    numericId = ws.Cells(r, 3).Value           ' column C = numeric ID
    If Not IsNumeric(numericId) Then Exit Sub   ' no id -> nothing to link

    Dim cell As Range
    Set cell = ws.Cells(r, 2)

    ' Snapshot the fill so it can be restored after the formula write.
    Dim hadFill As Boolean, savedFillColor As Long
    hadFill = (cell.Interior.ColorIndex <> xlNone)
    If hadFill Then savedFillColor = cell.Interior.Color

    Dim url As String, label As String
    url = "https://diality-prod.jamacloud.com/perspective.req#/items/" & CStr(CLng(numericId)) & "?projectId=47"
    ' Keep the existing '[id] Name' label (with its leading indent) as the
    ' link text. Double any quotes so the formula string stays valid.
    label = Replace(CStr(cell.Value), """", """""")
    cell.Formula = "=HYPERLINK(""" & url & """,""" & label & """)"

    If hadFill Then
        cell.Interior.Color = savedFillColor
    Else
        cell.Interior.ColorIndex = xlNone
    End If

    cell.Font.Name = "Calibri"
    cell.Font.Size = 10
    cell.Font.Color = RGB(5, 99, 193)          ' #0563C1 standard link blue
    cell.Font.Underline = xlUnderlineStyleSingle
End Sub

' ---- Column widths -----------------------------------------------------
' This module owns the column-width spec (suspect.py does not touch
' widths). Targets are in pixels; VBA's ColumnWidth is in characters and
' Column.Width (read-only) is in points, so SetColumnPixelWidth converges
' on each pixel target through the points ratio (1 px = 0.75 pt at
' Excel's 96-DPI baseline).

Private Function WidthColumns() As Variant
    WidthColumns = Array("A", "B", "C", "D", "E", "F")
End Function

Private Function WidthPixels() As Variant
    ' Same order as WidthColumns. Column A only carries the outline-level
    ' numbers this module reads, so it is collapsed to zero width
    ' (hidden -- VBA still reads hidden cells fine).
    WidthPixels = Array(0, 550, 110, 135, 135, 150)
End Function

Private Sub ApplyColumnWidths(ws As Worksheet)
    Dim cols As Variant, px As Variant, i As Long
    cols = WidthColumns()
    px = WidthPixels()
    For i = LBound(cols) To UBound(cols)
        SetColumnPixelWidth ws, CStr(cols(i)), CLng(px(i))
    Next i
End Sub

Private Sub SetColumnPixelWidth(ws As Worksheet, colLetter As String, targetPx As Long)
    Dim col As Range
    Set col = ws.Columns(colLetter)

    If targetPx <= 0 Then
        col.ColumnWidth = 0
        Exit Sub
    End If

    ' A hidden (zero-width) column has no measurable points-per-character
    ' ratio, so seed it with Excel's default width first.
    If col.ColumnWidth = 0 Then col.ColumnWidth = 8.43

    ' Points-per-character is affine (character width plus fixed cell
    ' padding), so a single ratio adjustment isn't exact -- iterate; two
    ' passes normally land within half a pixel.
    Dim targetPt As Double, pass As Long
    targetPt = targetPx * 0.75
    For pass = 1 To 3
        If Abs(col.Width - targetPt) <= 0.375 Then Exit For
        col.ColumnWidth = col.ColumnWidth * targetPt / col.Width
    Next pass
End Sub

Private Function ColumnWidthsAreCorrect(ws As Worksheet) As Boolean
    ' One-pixel tolerance (0.75 pt): Excel snaps widths to whole pixels.
    Dim cols As Variant, px As Variant, i As Long
    cols = WidthColumns()
    px = WidthPixels()
    For i = LBound(cols) To UBound(cols)
        If Abs(ws.Columns(CStr(cols(i))).Width - px(i) * 0.75) > 0.75 Then Exit Function
    Next i
    ColumnWidthsAreCorrect = True
End Function

Private Function TreeIsConsistent(ws As Worksheet, lastTreeRow As Long) As Boolean
    ' True when every tree row's outline level matches its column-A level.
    ' Only structure is checked -- row visibility (expanded/collapsed) is
    ' the user's choice and deliberately ignored.
    Dim r As Long, baseLvl As Long, isMissing As Boolean, expected As Long
    For r = DATA_START_ROW To lastTreeRow
        ParseLevel ws.Cells(r, LEVEL_COL).Value, baseLvl, isMissing
        If baseLvl = 1 Or baseLvl = 2 Or baseLvl = 3 Then
            expected = baseLvl + 1
        Else
            expected = 1
        End If
        If ws.Rows(r).OutlineLevel <> expected Then Exit Function
    Next r
    TreeIsConsistent = True
End Function

' ---- Level parsing ------------------------------------------------------
' Column A holds a plain 1/2/3 for a normal module/item/suspect row, or
' "1x"/"2x"/"3x" for the same level when suspect.py flagged that row as
' missing (an item with no downstream link at all, not just a suspect
' one) -- see suspect.py's _leveled(). Blank (category rows) or anything
' else parses to baseLvl = 0.
Private Sub ParseLevel(ByVal rawValue As Variant, ByRef baseLvl As Long, ByRef isMissing As Boolean)
    Dim s As String
    s = Trim(CStr(rawValue))
    isMissing = False
    baseLvl = 0
    If Len(s) = 0 Then Exit Sub
    If Right(s, 1) = "x" Or Right(s, 1) = "X" Then
        isMissing = True
        s = Left(s, Len(s) - 1)
    End If
    If IsNumeric(s) Then baseLvl = CLng(s)
End Sub


