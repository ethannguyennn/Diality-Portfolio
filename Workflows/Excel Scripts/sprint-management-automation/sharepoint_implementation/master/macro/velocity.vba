
Option Explicit
 
' -- CONFIG ---------------------------------------------------------
Private Const DATA_SHEET As String = "Sprint Template"
Private Const TOC_SHEET As String = "Velocity"
 
' 1-based column numbers in DATA_SHEET (A=1, D=4, E=5, F=6, I=9)
Private Const C_SPRINT As Long = 1     ' col A
Private Const C_ASSIGNEE As Long = 4   ' col D
Private Const C_TYPE As Long = 5       ' col E
Private Const C_SCOPE As Long = 6      ' col F
Private Const C_STATUS As Long = 9     ' col I
 
Private mRunning As Boolean            ' re-entrancy guard
 
' ----------------------------------------------------------------
'  ENTRY POINT — called from Worksheet_Activate on the Velocity tab
' ----------------------------------------------------------------
Public Sub RunSprintVelocityUpdate()
 
    If mRunning Then Exit Sub
    mRunning = True
 
    On Error GoTo 0
 
    Dim wasScreenUpdating As Boolean, wasEnableEvents As Boolean
    Dim wasCalc As XlCalculation
    wasScreenUpdating = Application.ScreenUpdating
    wasEnableEvents = Application.EnableEvents
    wasCalc = Application.Calculation
 
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait
 
    Dim wsData As Worksheet
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0
    If wsData Is Nothing Then
        MsgBox "Sheet """ & DATA_SHEET & """ not found.", vbExclamation
        GoTo CleanExit
    End If
 
    If Not IsNumeric(wsData.Range("B2").Value) Then
        MsgBox "B2 does not contain a valid sprint number.", vbExclamation
        GoTo CleanExit
    End If
    Dim sprint As Long
    sprint = CLng(wsData.Range("B2").Value)
 
    Dim usedRng As Range
    Set usedRng = wsData.UsedRange
    Dim lastRow As Long, lastCol As Long
    lastRow = usedRng.Row + usedRng.Rows.Count - 1
    lastCol = usedRng.Column + usedRng.Columns.Count - 1
    If lastCol < 9 Then lastCol = 9
 
    If lastRow < 4 Then
        MsgBox "No data rows found in """ & DATA_SHEET & """.", vbExclamation
        GoTo CleanExit
    End If
 
    Dim rawData As Variant
    rawData = wsData.Range(wsData.Cells(4, 1), wsData.Cells(lastRow, lastCol)).Value
    If Not IsArray(rawData) Then
        Dim single1(1 To 1, 1 To 1) As Variant
        single1(1, 1) = rawData
        rawData = single1
    End If
 
    ' ---- filter rows for this sprint ----
    Dim filteredIdx() As Long
    Dim fCount As Long
    fCount = 0
    ReDim filteredIdx(1 To UBound(rawData, 1))
    Dim fi As Long
    For fi = 1 To UBound(rawData, 1)
        If IsNumeric(rawData(fi, C_SPRINT)) Then
            If CLng(rawData(fi, C_SPRINT)) = sprint Then
                fCount = fCount + 1
                filteredIdx(fCount) = fi
            End If
        End If
    Next fi
 
    If fCount = 0 Then
        MsgBox "No rows found for Sprint " & sprint & ".", vbExclamation
        GoTo CleanExit
    End If
    ReDim Preserve filteredIdx(1 To fCount)
 
    Dim sheetName As String
    sheetName = BuildVelocitySheet(sprint, rawData, filteredIdx)
 
    UpdateTOC sprint, sheetName, TOC_SHEET
    UpdateVelocityChart rawData, TOC_SHEET, sprint
    
    ThisWorkbook.Worksheets(TOC_SHEET).Activate

 
CleanExit:
    Application.ScreenUpdating = wasScreenUpdating
    Application.EnableEvents = wasEnableEvents
    Application.Calculation = wasCalc
    Application.Cursor = xlDefault
    mRunning = False
    Exit Sub
 
CleanFail:
    MsgBox "Sprint Velocity update failed: " & Err.Description, vbCritical
    Resume CleanExit
End Sub
 
' ----------------------------------------------------------------
'  BuildVelocitySheet — builds/replaces "Sprint N Velocity", hides it,
'  returns its name
' ----------------------------------------------------------------
Private Function BuildVelocitySheet(ByVal sprint As Long, rawData As Variant, filteredIdx() As Long) As String
 
    Dim fCount As Long
    fCount = UBound(filteredIdx) - LBound(filteredIdx) + 1
 
    Dim statuses As Variant: statuses = StatusesList()
    Dim typesArr As Variant: typesArr = TypesList()
    Dim scopesArr As Variant: scopesArr = ScopesList()
 
    Dim statusColorMap As Object: Set statusColorMap = StatusColorsMap()
    Dim scopeColorMap As Object: Set scopeColorMap = ScopeColorsMap()
    Dim typeColorMap As Object: Set typeColorMap = TypeColorsMap()
 
    ' ---- shared loop vars ----
    Dim idx As Long, r As Long, c As Long, p As Long
    Dim nm As String, stv As String, tv As String, scv As String
    Dim si As Long, ti As Long, sci As Long, sti As Long
    Dim tOrd As Long, scOrd As Long, tName As String, scName As String
    Dim rw As Long, s1HRow As Long, s1LastCol As Long, totalRow As Long
    Dim rowTotal As Long, cnt As Long, grandTotal As Long, peopleCount As Long
    Dim colTotals() As Long
    Dim s1ChartRange As Range, s1Chart As ChartObject, ser As Series
    Dim s2Top As Long, s2HRow As Long, typeOrder As Variant
    Dim writtenTypeNames() As String, typeRowCount As Long
    Dim s2ChartRange As Range, s2Chart As ChartObject, s2pIdx As Long
    Dim s3Top As Long, s3HRow As Long, scopeOrder As Variant
    Dim writtenScopeNames() As String, scopeRowCount As Long
    Dim s3ChartRange As Range, s3Chart As ChartObject, s3pIdx As Long
    Dim s4Top As Long, s4HRow As Long
    Dim writtenStatusNames() As String, s4Count As Long
    Dim s4ChartRange As Range, s4Chart As ChartObject, s4pIdx As Long
    Dim sheetName As String, existing As Worksheet, vel As Worksheet
    Dim people() As String
    Dim inner As Object
 
    ' ---------- aggregate: status by assignee ----------
    Dim aMap As Object: Set aMap = CreateObject("Scripting.Dictionary")
    For idx = 1 To fCount
        r = filteredIdx(idx)
        nm = Trim(CStr(rawData(r, C_ASSIGNEE)))
        If nm = "" Then nm = "Unassigned"
        stv = NormalizeStatus(Trim(CStr(rawData(r, C_STATUS))), statuses)
        If Not aMap.Exists(nm) Then
            Set inner = CreateObject("Scripting.Dictionary")
            For si = LBound(statuses) To UBound(statuses)
                inner(statuses(si)) = 0
            Next si
            aMap.Add nm, inner
        End If
        aMap(nm)(stv) = aMap(nm)(stv) + 1
    Next idx
 
    If aMap.Count > 0 Then
        people = DictKeysToStringArray(aMap)
        SortStringArray people
    Else
        ReDim people(0 To -1)
    End If
 
    ' ---------- aggregate: by process type ----------
    Dim typeMap As Object: Set typeMap = CreateObject("Scripting.Dictionary")
    For ti = LBound(typesArr) To UBound(typesArr)
        typeMap(typesArr(ti)) = 0
    Next ti
    For idx = 1 To fCount
        r = filteredIdx(idx)
        tv = Trim(CStr(rawData(r, C_TYPE)))
        If typeMap.Exists(tv) Then
            typeMap(tv) = typeMap(tv) + 1
        Else
            If typeMap.Exists("Other") Then
                typeMap("Other") = typeMap("Other") + 1
            Else
                typeMap.Add "Other", 1
            End If
        End If
    Next idx
 
    ' ---------- aggregate: scope changes ----------
    Dim scopeMap As Object: Set scopeMap = CreateObject("Scripting.Dictionary")
    For sci = LBound(scopesArr) To UBound(scopesArr)
        scopeMap(scopesArr(sci)) = 0
    Next sci
    For idx = 1 To fCount
        r = filteredIdx(idx)
        scv = Trim(CStr(rawData(r, C_SCOPE)))
        If scopeMap.Exists(scv) Then
            scopeMap(scv) = scopeMap(scv) + 1
        Else
            If scopeMap.Exists("Other") Then
                scopeMap("Other") = scopeMap("Other") + 1
            Else
                scopeMap.Add "Other", 1
            End If
        End If
    Next idx
 
    ' ---------- aggregate: by status ----------
    Dim statusMap As Object: Set statusMap = CreateObject("Scripting.Dictionary")
    For sti = LBound(statuses) To UBound(statuses)
        statusMap(statuses(sti)) = 0
    Next sti
    For idx = 1 To fCount
        r = filteredIdx(idx)
        stv = NormalizeStatus(Trim(CStr(rawData(r, C_STATUS))), statuses)
        statusMap(stv) = statusMap(stv) + 1
    Next idx
 
    ' ---------- create / replace the velocity sheet ----------
    sheetName = "Sprint " & sprint & " Velocity"
 
    Application.DisplayAlerts = False
    On Error Resume Next
    Set existing = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If Not existing Is Nothing Then existing.Delete
    Application.DisplayAlerts = True
 
    Set vel = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    vel.Name = sheetName
 
    rw = 1
 
    ' ===== Section 1: Status by Assignee =====
    vel.Cells(rw, 1).Value = "Sprint " & sprint & "  |  Status by Assignee"
    With vel.Cells(rw, 1).Font
        .Bold = True
        .Size = 13
        .Color = HexToLong("#1F3864")
    End With
    rw = rw + 1
 
    s1HRow = rw
    vel.Cells(rw, 1).Value = "Assignee"
    For c = LBound(statuses) To UBound(statuses)
        vel.Cells(rw, 2 + c).Value = statuses(c)
    Next c
    s1LastCol = 2 + (UBound(statuses) - LBound(statuses) + 1)   ' "Total" column
    vel.Cells(rw, s1LastCol).Value = "Total"
 
    With vel.Range(vel.Cells(rw, 1), vel.Cells(rw, s1LastCol))
        .Interior.Color = HexToLong("#2F5496")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With
    For c = LBound(statuses) To UBound(statuses)
        If statusColorMap.Exists(statuses(c)) Then
            With vel.Cells(rw, 2 + c)
                .Interior.Color = statusColorMap(statuses(c))
                .Font.Color = HexToLong("#1F3864")
            End With
        End If
    Next c
    rw = rw + 1
 
    ReDim colTotals(LBound(statuses) To UBound(statuses))
    grandTotal = 0
    peopleCount = 0
    If Not (LBound(people) > UBound(people)) Then
        peopleCount = UBound(people) - LBound(people) + 1
        For p = LBound(people) To UBound(people)
            vel.Cells(rw, 1).Value = people(p)
            rowTotal = 0
            For c = LBound(statuses) To UBound(statuses)
                cnt = aMap(people(p))(statuses(c))
                vel.Cells(rw, 2 + c).Value = cnt
                rowTotal = rowTotal + cnt
                colTotals(c) = colTotals(c) + cnt
            Next c
            vel.Cells(rw, s1LastCol).Value = rowTotal
            grandTotal = grandTotal + rowTotal
            rw = rw + 1
        Next p
    End If
 
    vel.Cells(rw, 1).Value = "TOTAL"
    For c = LBound(statuses) To UBound(statuses)
        vel.Cells(rw, 2 + c).Value = colTotals(c)
    Next c
    vel.Cells(rw, s1LastCol).Value = grandTotal
    With vel.Range(vel.Cells(rw, 1), vel.Cells(rw, s1LastCol))
        .Font.Bold = True
        .Interior.Color = HexToLong("#BDD7EE")
    End With
    totalRow = rw
    rw = rw + 1
 
    If peopleCount > 0 Then
        Set s1ChartRange = vel.Range(vel.Cells(s1HRow, 1), vel.Cells(s1HRow + peopleCount, s1LastCol - 1))
        Set s1Chart = vel.ChartObjects.Add( _
            Left:=vel.Cells(s1HRow, s1LastCol + 1).Left + 15, _
            Top:=vel.Cells(s1HRow, 1).Top, _
            Width:=580, _
            Height:=Application.WorksheetFunction.Max(300, peopleCount * 22 + 80))
        With s1Chart.Chart
            .ChartType = xlBarStacked
            .SetSourceData Source:=s1ChartRange, PlotBy:=xlColumns
            .hasTitle = True
            .ChartTitle.Text = "Sprint " & sprint & " — Status by Assignee"
            For Each ser In .SeriesCollection
                If statusColorMap.Exists(ser.Name) Then
                    ser.Interior.Color = statusColorMap(ser.Name)
                End If
            Next ser
        End With
    End If
    rw = rw + 2
 
    ' ===== Section 2: Count by Process Type =====
    s2Top = rw
    vel.Cells(rw, 1).Value = "Count by Process Type"
    With vel.Cells(rw, 1).Font
        .Bold = True: .Size = 13: .Color = HexToLong("#1F3864")
    End With
    rw = rw + 1
 
    vel.Cells(rw, 1).Value = "Process Type"
    vel.Cells(rw, 2).Value = "Count"
    With vel.Range(vel.Cells(rw, 1), vel.Cells(rw, 2))
        .Interior.Color = HexToLong("#375623")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With
    s2HRow = rw
    rw = rw + 1
 
    typeOrder = ArrayAppend(typesArr, "Other")
    ReDim writtenTypeNames(1 To (UBound(typeOrder) - LBound(typeOrder) + 1))
    typeRowCount = 0
    For tOrd = LBound(typeOrder) To UBound(typeOrder)
        tName = CStr(typeOrder(tOrd))
        If typeMap.Exists(tName) Then
            If typeMap(tName) > 0 Then
                typeRowCount = typeRowCount + 1
                vel.Cells(rw, 1).Value = tName
                vel.Cells(rw, 2).Value = typeMap(tName)
                writtenTypeNames(typeRowCount) = tName
                rw = rw + 1
            End If
        End If
    Next tOrd
 
    If typeRowCount > 0 Then
        Set s2ChartRange = vel.Range(vel.Cells(s2HRow, 1), vel.Cells(s2HRow + typeRowCount, 2))
        Set s2Chart = vel.ChartObjects.Add( _
            Left:=vel.Cells(s2Top, 4).Left + 15, Top:=vel.Cells(s2Top, 1).Top, _
            Width:=380, Height:=280)
        With s2Chart.Chart
            .ChartType = xlPie
            .SetSourceData Source:=s2ChartRange, PlotBy:=xlColumns
            .hasTitle = True
            .ChartTitle.Text = "Sprint " & sprint & " — Process Type"
            With .SeriesCollection(1)
                .HasDataLabels = True
                .DataLabels.ShowPercentage = True
                .DataLabels.ShowValue = False
                .DataLabels.ShowCategoryName = True
                For s2pIdx = 1 To typeRowCount
                    If typeColorMap.Exists(writtenTypeNames(s2pIdx)) Then
                        .Points(s2pIdx).Interior.Color = typeColorMap(writtenTypeNames(s2pIdx))
                    End If
                Next s2pIdx
            End With
        End With
    End If
    rw = rw + 2
 
    ' ===== Section 3: Scope Changes =====
    s3Top = rw
    vel.Cells(rw, 1).Value = "Scope Changes"
    With vel.Cells(rw, 1).Font
        .Bold = True: .Size = 13: .Color = HexToLong("#1F3864")
    End With
    rw = rw + 1
 
    vel.Cells(rw, 1).Value = "Scope Change"
    vel.Cells(rw, 2).Value = "Count"
    With vel.Range(vel.Cells(rw, 1), vel.Cells(rw, 2))
        .Interior.Color = HexToLong("#C55A11")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With
    s3HRow = rw
    rw = rw + 1
 
    scopeOrder = ArrayAppend(scopesArr, "Other")
    ReDim writtenScopeNames(1 To (UBound(scopeOrder) - LBound(scopeOrder) + 1))
    scopeRowCount = 0
    For scOrd = LBound(scopeOrder) To UBound(scopeOrder)
        scName = CStr(scopeOrder(scOrd))
        If scopeMap.Exists(scName) Then
            If scopeMap(scName) > 0 Then
                scopeRowCount = scopeRowCount + 1
                vel.Cells(rw, 1).Value = scName
                vel.Cells(rw, 2).Value = scopeMap(scName)
                writtenScopeNames(scopeRowCount) = scName
                rw = rw + 1
            End If
        End If
    Next scOrd
 
    If scopeRowCount > 0 Then
        Set s3ChartRange = vel.Range(vel.Cells(s3HRow, 1), vel.Cells(s3HRow + scopeRowCount, 2))
        Set s3Chart = vel.ChartObjects.Add( _
            Left:=vel.Cells(s3Top, 4).Left + 15, Top:=vel.Cells(s3Top, 1).Top, _
            Width:=380, Height:=260)
        With s3Chart.Chart
            .ChartType = xlPie
            .SetSourceData Source:=s3ChartRange, PlotBy:=xlColumns
            .hasTitle = True
            .ChartTitle.Text = "Sprint " & sprint & " — Scope Changes"
            With .SeriesCollection(1)
                .HasDataLabels = True
                .DataLabels.ShowPercentage = True
                .DataLabels.ShowValue = False
                .DataLabels.ShowCategoryName = True
                For s3pIdx = 1 To scopeRowCount
                    If scopeColorMap.Exists(writtenScopeNames(s3pIdx)) Then
                        .Points(s3pIdx).Interior.Color = scopeColorMap(writtenScopeNames(s3pIdx))
                    End If
                Next s3pIdx
            End With
        End With
    End If
    rw = rw + 2
 
    ' ===== Section 4: Count by Status =====
    s4Top = rw
    vel.Cells(rw, 1).Value = "Count by Status"
    With vel.Cells(rw, 1).Font
        .Bold = True: .Size = 13: .Color = HexToLong("#1F3864")
    End With
    rw = rw + 1
 
    vel.Cells(rw, 1).Value = "Status"
    vel.Cells(rw, 2).Value = "Count"
    With vel.Range(vel.Cells(rw, 1), vel.Cells(rw, 2))
        .Interior.Color = HexToLong("#7030A0")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With
    s4HRow = rw
    rw = rw + 1
 
    ReDim writtenStatusNames(1 To (UBound(statuses) - LBound(statuses) + 1))
    s4Count = 0
    For c = LBound(statuses) To UBound(statuses)
        s4Count = s4Count + 1
        vel.Cells(rw, 1).Value = statuses(c)
        vel.Cells(rw, 2).Value = statusMap(statuses(c))
        writtenStatusNames(s4Count) = CStr(statuses(c))
        rw = rw + 1
    Next c
 
    Set s4ChartRange = vel.Range(vel.Cells(s4HRow, 1), vel.Cells(s4HRow + s4Count, 2))
    Set s4Chart = vel.ChartObjects.Add( _
        Left:=vel.Cells(s4Top, 4).Left + 15, Top:=vel.Cells(s4Top, 1).Top, _
        Width:=400, Height:=260)
    With s4Chart.Chart
        .ChartType = xlPie
        .SetSourceData Source:=s4ChartRange, PlotBy:=xlColumns
        .hasTitle = True
        .ChartTitle.Text = "Sprint " & sprint & " — Ticket Status"
        With .SeriesCollection(1)
            .HasDataLabels = True
            .DataLabels.ShowPercentage = True
            .DataLabels.ShowValue = False
            .DataLabels.ShowCategoryName = True
            For s4pIdx = 1 To s4Count
                If statusColorMap.Exists(writtenStatusNames(s4pIdx)) Then
                    .Points(s4pIdx).Interior.Color = statusColorMap(writtenStatusNames(s4pIdx))
                End If
            Next s4pIdx
        End With
    End With
 
    ' ---- autofit + hide ----
    vel.UsedRange.Columns.AutoFit
    vel.Visible = xlSheetHidden
 
    BuildVelocitySheet = sheetName
End Function
 
' ----------------------------------------------------------------
'  UpdateTOC — writes/refreshes one row in the Velocity TOC
' ----------------------------------------------------------------
Private Sub UpdateTOC(ByVal sprint As Long, ByVal sheetName As String, ByVal tocName As String)
 
    Dim toc As Worksheet
    On Error Resume Next
    Set toc = ThisWorkbook.Worksheets(tocName)
    On Error GoTo 0
    If toc Is Nothing Then
        Debug.Print "Warning: TOC sheet """ & tocName & """ not found — skipping TOC update."
        Exit Sub
    End If
 
    Dim todayDt As Date: todayDt = Now
    Dim dateStr As String
    dateStr = Month(todayDt) & "/" & Day(todayDt) & "/" & Year(todayDt)
 
    Dim rawHour As Long: rawHour = Hour(todayDt)
    Dim ampm As String: ampm = IIf(rawHour >= 12, "PM", "AM")
    Dim hour12 As Long: hour12 = rawHour Mod 12
    If hour12 = 0 Then hour12 = 12
    Dim minutesStr As String: minutesStr = Format(Minute(todayDt), "00")
    Dim timeStr As String: timeStr = hour12 & ":" & minutesStr & " " & ampm
 
    Dim a1Val As String
    a1Val = Trim(CStr(toc.Range("A1").Value))
    If a1Val = "" Then
        With toc.Range("A1")
            .Value = "Sprint Velocity Log"
            .Font.Bold = True
            .Font.Size = 16
            .Font.Color = HexToLong("#1F3864")
        End With
    End If
 
    With toc.Range("A2:C2")
        .Value = Array("Sprint", "Date Generated", "Time Generated")
        .Interior.Color = HexToLong("#2F5496")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With
 
    Dim scanValues As Variant
    scanValues = toc.Range(toc.Cells(3, 1), toc.Cells(202, 1)).Value
 
    Dim displayName As String: displayName = "Sprint " & sprint
    Dim foundRow As Long: foundRow = -1
    Dim nextEmptyRow As Long: nextEmptyRow = 3
    Dim i As Long, cellVal As String
 
    For i = 1 To 200
        cellVal = Trim(CStr(scanValues(i, 1)))
        If cellVal = "" Then
            nextEmptyRow = 2 + i
            Exit For
        End If
        If cellVal = displayName Then foundRow = 2 + i
        nextEmptyRow = 2 + i + 1
    Next i
 
    Dim writeRow As Long
    writeRow = IIf(foundRow >= 1, foundRow, nextEmptyRow)
 
    Dim nameCell As Range
    Set nameCell = toc.Cells(writeRow, 1)
    nameCell.Value = displayName
 
    If nameCell.Hyperlinks.Count > 0 Then nameCell.Hyperlinks.Delete
    toc.Hyperlinks.Add Anchor:=nameCell, Address:="", _
                        SubAddress:="'" & sheetName & "'!A1", _
                        ScreenTip:="Open " & sheetName, _
                        TextToDisplay:=displayName
 
    nameCell.Font.Color = HexToLong("#0563C1")
    nameCell.Font.Underline = xlUnderlineStyleSingle
 
    toc.Cells(writeRow, 2).Value = dateStr
    toc.Cells(writeRow, 3).Value = timeStr
 
    toc.UsedRange.Columns.AutoFit
End Sub
 
' ----------------------------------------------------------------
'  UpdateVelocityChart — Committed vs Done + Completion % combo chart
'  Data table at F2:I(n+2); chart placed below it. Rebuilt each run.
' ----------------------------------------------------------------
Private Sub UpdateVelocityChart(rawData As Variant, ByVal tocName As String, ByVal currentSprint As Long)
 
    Const CHART_TITLE As String = "Sprint Committed vs Done with Completion %"
 
    Dim toc As Worksheet
    On Error Resume Next
    Set toc = ThisWorkbook.Worksheets(tocName)
    On Error GoTo 0
    If toc Is Nothing Then Exit Sub
 
    Dim sprintNums As Object: Set sprintNums = CreateObject("Scripting.Dictionary")
    Dim sprintDone As Object: Set sprintDone = CreateObject("Scripting.Dictionary")
 
    Dim i As Long, sNum As Long
    Dim scopeVal As String, statusVal As String
    For i = LBound(rawData, 1) To UBound(rawData, 1)
        If IsNumeric(rawData(i, C_SPRINT)) Then
            sNum = CLng(rawData(i, C_SPRINT))
            If sNum <> 0 And sNum <= currentSprint Then
                scopeVal = Trim(CStr(rawData(i, C_SCOPE)))
                statusVal = Trim(CStr(rawData(i, C_STATUS)))
                If scopeVal <> "Move out of Sprint" Then
                    If Not sprintNums.Exists(sNum) Then
                        sprintNums.Add sNum, 0
                        sprintDone.Add sNum, 0
                    End If
                    sprintNums(sNum) = sprintNums(sNum) + 1
                    If statusVal = "Done" Or IsDoneAliasStatus(statusVal) Then sprintDone(sNum) = sprintDone(sNum) + 1
                End If
            End If
        End If
    Next i
 
    If sprintNums.Count = 0 Then Exit Sub
 
    Dim allSprints() As Long
    ReDim allSprints(0 To sprintNums.Count - 1)
    Dim k As Long, keyV As Variant
    k = 0
    For Each keyV In sprintNums.Keys
        allSprints(k) = CLng(keyV)
        k = k + 1
    Next keyV
    SortLongArray allSprints
 
    Dim totalCount As Long: totalCount = UBound(allSprints) - LBound(allSprints) + 1
    Dim startIdx As Long: startIdx = 0
    If totalCount > 10 Then startIdx = totalCount - 10
    Dim sortedSprints() As Long
    ReDim sortedSprints(0 To totalCount - 1 - startIdx)
    Dim jj As Long
    For jj = startIdx To totalCount - 1
        sortedSprints(jj - startIdx) = allSprints(jj)
    Next jj
 
    Dim n As Long: n = UBound(sortedSprints) - LBound(sortedSprints) + 1
 
    Dim co As ChartObject, hasTitle As Boolean
    For Each co In toc.ChartObjects
        hasTitle = False
        On Error Resume Next
        hasTitle = co.Chart.hasTitle
        On Error GoTo 0
        If hasTitle Then
            If co.Chart.ChartTitle.Text = CHART_TITLE Then co.Delete
        End If
    Next co
 
    toc.Range(toc.Cells(2, 6), toc.Cells(13, 9)).Clear
 
    With toc.Range(toc.Cells(2, 6), toc.Cells(2, 9))
        .Value = Array("Sprint", "Committed", "Done", "Completion %")
        .Interior.Color = HexToLong("#2F5496")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With
 
    Dim r2 As Long, sp As Long, committed As Long, doneCt As Long, pct As Double
    For r2 = 1 To n
        sp = sortedSprints(r2 - 1)
        committed = sprintNums(sp)
        doneCt = sprintDone(sp)
        pct = IIf(committed > 0, doneCt / committed, 0)
        toc.Cells(2 + r2, 6).Value = CStr(sp)
        toc.Cells(2 + r2, 7).Value = committed
        toc.Cells(2 + r2, 8).Value = doneCt
        toc.Cells(2 + r2, 9).Value = pct
    Next r2
 
    toc.Range(toc.Cells(3, 9), toc.Cells(2 + n, 9)).NumberFormat = "0%"
    toc.Range(toc.Cells(2, 6), toc.Cells(2 + n, 9)).Columns.AutoFit
 
    Dim seriesDataRange As Range, sprintLabelRange As Range, anchorCell As Range
    Set seriesDataRange = toc.Range(toc.Cells(2, 7), toc.Cells(2 + n, 9))
    Set sprintLabelRange = toc.Range(toc.Cells(3, 6), toc.Cells(2 + n, 6))
    Set anchorCell = toc.Cells(2 + n + 2, 6)
 
    Dim chartObj As ChartObject
    Set chartObj = toc.ChartObjects.Add(Left:=anchorCell.Left, Top:=anchorCell.Top, Width:=500, Height:=300)
 
    Dim serCommitted As Series, serDone As Series, serPct As Series
    With chartObj.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=seriesDataRange, PlotBy:=xlColumns
        .hasTitle = True
        .ChartTitle.Text = CHART_TITLE
 
        Set serCommitted = .SeriesCollection(1)
        Set serDone = .SeriesCollection(2)
        Set serPct = .SeriesCollection(3)
 
        serCommitted.XValues = sprintLabelRange
        serDone.XValues = sprintLabelRange
        serPct.XValues = sprintLabelRange
 
        serCommitted.Interior.Color = HexToLong("#4472C4")
        serDone.Interior.Color = HexToLong("#ED7D31")
 
        serPct.ChartType = xlLineMarkers
        serPct.AxisGroup = xlSecondary
        serPct.Border.Color = HexToLong("#70AD47")
        serPct.MarkerStyle = xlMarkerStyleCircle
 
        serCommitted.HasDataLabels = True
        serDone.HasDataLabels = True
        serPct.HasDataLabels = True
        serPct.DataLabels.NumberFormat = "0%"
 
        With .Axes(xlValue, xlPrimary)
            .hasTitle = True
            .AxisTitle.Text = "Work Items"
            .TickLabels.NumberFormat = "0"
            .MinimumScale = 0
        End With
 
        On Error Resume Next
        With .Axes(xlValue, xlSecondary)
            .hasTitle = True
            .AxisTitle.Text = "Completion %"
            .TickLabels.NumberFormat = "0%"
            .MaximumScale = 1
            .MinimumScale = 0
        End With
        On Error GoTo 0
    End With
End Sub
 
' ----------------------------------------------------------------
'  Small helpers
' ----------------------------------------------------------------
Private Function StatusesList() As Variant
    StatusesList = Array("Blocked", "Done", "In-progress", "Not started", "On-hold", "N/A")
End Function
 
Private Function TypesList() As Variant
    TypesList = Array("Big-V", "Little-V", "Bug", "Dialin ticket", "Support")
End Function
 
Private Function ScopesList() As Variant
    ScopesList = Array("Scoped", "Added", "Move out of Sprint")
End Function
 
Private Function StatusColorsMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("Done") = HexToLong("#4FB06D")
    d("Blocked") = HexToLong("#BF2C34")
    d("In-progress") = HexToLong("#F5C26B")
    d("Not started") = HexToLong("#CBD6E2")
    d("On-hold") = HexToLong("#F07857")
    Set StatusColorsMap = d
End Function
 
Private Function ScopeColorsMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("Scoped") = HexToLong("#4FB06D")
    d("Added") = HexToLong("#F07857")
    d("Move out of Sprint") = HexToLong("#BF2C34")
    Set ScopeColorsMap = d
End Function
 
Private Function TypeColorsMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("Big-V") = HexToLong("#CBD6E2")
    d("Little-V") = HexToLong("#BE398D")
    d("Bug") = HexToLong("#BF2C34")
    d("Dialin ticket") = HexToLong("#F07857")
    d("Support") = HexToLong("#4FB06D")
    Set TypeColorsMap = d
End Function
 
' Bug/Dialin statuses where the fix work is effectively finished (just
' waiting on test sign-off or the next release cutover) -- counted as
' Done here to match SUMMARY_DONE_ALIASES in update_sprint.py's Sprint
' Summary "Done" column formula. Keep this list in sync with that one.
Private Function DoneAliasStatuses() As Variant
    DoneAliasStatuses = Array("Ready for Fix to be Tested", "Testing Fix in Progress", _
                              "Ready for Updated Release", "Needs Final Fix Details")
End Function

Private Function IsDoneAliasStatus(ByVal raw As String) As Boolean
    Dim aliases As Variant: aliases = DoneAliasStatuses()
    Dim i As Long
    For i = LBound(aliases) To UBound(aliases)
        If raw = aliases(i) Then
            IsDoneAliasStatus = True
            Exit Function
        End If
    Next i
    IsDoneAliasStatus = False
End Function

Private Function NormalizeStatus(ByVal raw As String, statuses As Variant) As String
    Dim i As Long
    For i = LBound(statuses) To UBound(statuses)
        If statuses(i) = raw Then
            NormalizeStatus = raw
            Exit Function
        End If
    Next i
    If IsDoneAliasStatus(raw) Then
        NormalizeStatus = "Done"
        Exit Function
    End If
    NormalizeStatus = "In-progress"
End Function
 
Private Function ArrayAppend(arr As Variant, extra As String) As Variant
    Dim n As Long: n = UBound(arr) - LBound(arr) + 1
    Dim outArr() As Variant
    ReDim outArr(0 To n)
    Dim i As Long
    For i = 0 To n - 1
        outArr(i) = arr(LBound(arr) + i)
    Next i
    outArr(n) = extra
    ArrayAppend = outArr
End Function
 
Private Function DictKeysToStringArray(d As Object) As String()
    Dim keysArr As Variant
    keysArr = d.Keys
    Dim outArr() As String
    ReDim outArr(0 To d.Count - 1)
    Dim i As Long
    For i = 0 To d.Count - 1
        outArr(i) = CStr(keysArr(i))
    Next i
    DictKeysToStringArray = outArr
End Function
 
Private Sub SortStringArray(arr() As String)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(j), arr(i), vbTextCompare) < 0 Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub
 
Private Sub SortLongArray(arr() As Long)
    Dim i As Long, j As Long, tmp As Long
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub
 
Private Function HexToLong(ByVal hexStr As String) As Long
    Dim h As String
    h = Replace(hexStr, "#", "")
    Dim rC As Long, gC As Long, bC As Long
    rC = CLng("&H" & Mid(h, 1, 2))
    gC = CLng("&H" & Mid(h, 3, 2))
    bC = CLng("&H" & Mid(h, 5, 2))
    HexToLong = RGB(rC, gC, bC)
End Function



