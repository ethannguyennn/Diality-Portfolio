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

    ws.Outline.SummaryRow = xlSummaryAbove

    FormatBanner ws

    ' Bulk defaults for the whole tree body (this is the suspect-row
    ' look); the per-row pass below only overrides what differs by level.
    With ws.Range(ws.Cells(DATA_START_ROW, 2), ws.Cells(lastTreeRow, 4))
        .Interior.ColorIndex = xlNone
        .Font.Bold = False
        .Font.Size = 10
        .Font.Color = vbBlack
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
    Dim r As Long, lvl As Variant, prevLvl As Variant
    Dim stripeOn As Boolean
    Dim rowBand As Range
    prevLvl = Empty
    For r = DATA_START_ROW To lastTreeRow
        lvl = ws.Cells(r, LEVEL_COL).Value
        Set rowBand = ws.Range(ws.Cells(r, 2), ws.Cells(r, 4))

        If lvl = 1 Then      ' module
            ws.Rows(r).OutlineLevel = 2
            rowBand.Interior.Color = RGB(112, 173, 71)   ' #70AD47
            rowBand.Font.Bold = True
            rowBand.Font.Color = vbWhite
        ElseIf lvl = 2 Then  ' item
            ws.Rows(r).OutlineLevel = 3
            rowBand.Interior.Color = RGB(169, 208, 142)  ' #A9D08E
            rowBand.Font.Bold = True
        ElseIf lvl = 3 Then  ' suspect
            ws.Rows(r).OutlineLevel = 4
            If prevLvl <> 3 Then stripeOn = False  ' first suspect row under an item: no stripe
            If stripeOn Then rowBand.Interior.Color = RGB(226, 239, 218)  ' #E2EFDA
            stripeOn = Not stripeOn
        Else                 ' category (no level number in column A)
            ws.Rows(r).OutlineLevel = 1
            rowBand.Interior.Color = RGB(16, 124, 65)    ' #107C41
            rowBand.Font.Bold = True
            rowBand.Font.Color = vbWhite
            rowBand.Font.Size = 11
        End If

        prevLvl = lvl
    Next r

    FormatSummary ws, lastTreeRow

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
' Layout positions mirror suspect.py's row plan: title row 1, subtitle
' row 2, column headers row 3, tree from row 4 through O3, then a blank
' row, the "Module Summary" title, a blank row, the summary header, one
' row per module, and the Total row (the last used row in column B).

Private Sub FormatBanner(ws As Worksheet)
    ' Title banner across B1:D1.
    With ws.Range("B1:D1")
        .Merge
        .Interior.Color = RGB(16, 124, 65)   ' #107C41
        .Font.Bold = True
        .Font.Size = 16
        .Font.Color = vbWhite
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 30

    ' Subtitle across B2:D2.
    With ws.Range("B2:D2")
        .Merge
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = RGB(102, 102, 102)     ' #666666
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(2).RowHeight = 18

    ' Column header row.
    With ws.Range("B3:D3")
        .Interior.Color = RGB(255, 192, 0)   ' #FFC000
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B3").HorizontalAlignment = xlLeft
    ws.Range("C3:D3").HorizontalAlignment = xlCenter
End Sub

Private Sub FormatSummary(ws As Worksheet, lastTreeRow As Long)
    ' Module Summary block below the tree (see layout comment above).
    Dim titleRow As Long, headerRow As Long, totalRow As Long
    titleRow = lastTreeRow + 2
    headerRow = titleRow + 2
    totalRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If totalRow <= headerRow Then Exit Sub  ' summary block not written

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
    WidthPixels = Array(0, 215, 110, 135, 135, 150)
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
    Dim r As Long, lvl As Variant, expected As Long
    For r = DATA_START_ROW To lastTreeRow
        lvl = ws.Cells(r, LEVEL_COL).Value
        If lvl = 1 Or lvl = 2 Or lvl = 3 Then
            expected = CLng(lvl) + 1
        Else
            expected = 1
        End If
        If ws.Rows(r).OutlineLevel <> expected Then Exit Function
    Next r
    TreeIsConsistent = True
End Function
