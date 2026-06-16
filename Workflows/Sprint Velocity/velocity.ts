function main(workbook: ExcelScript.Workbook): void {

    // ══════════════════════════════════════════════════════════════
    //  CONFIG
    // ══════════════════════════════════════════════════════════════
    const DATA_SHEET = "Sprint Template";
    const TOC_SHEET = "Velocity";          // ← NEW

    const C_SPRINT = 0;
    const C_ASSIGNEE = 3;
    const C_TYPE = 4;
    const C_SCOPE = 5;
    const C_STATUS = 8;

    const STATUSES = ["Blocked", "Done", "In-progress", "Not started", "On-hold", "N/A"];
    const TYPES = ["Big-V", "Little-V", "Bug", "Dialin ticket", "Support"];
    const SCOPES = ["Scoped", "Added", "Move out of Sprint"];

    const STATUS_COLORS: { [s: string]: string } = {
        "Done": "#4FB06D",
        "Blocked": "#BF2C34",
        "In-progress": "#F5c26B",
        "Not started": "#CBD6E2",
        "On-hold": "#F07857"
    };
    const SCOPE_COLORS: { [s: string]: string } = {
        "Scoped": "#4FB06D",
        "Added": "#F07857",
        "Move out of Sprint": "#BF2C34"
    };
    const TYPE_COLORS: { [s: string]: string } = {
        "Big-V": "#CBD6E2",
        "Little-V": "#BE398D",
        "Bug": "#BF2C34",
        "Dialin ticket": "#F07857",
        "Support": "#4FB06D"
    };

    const ROW_PT = 15;
    const COL_PT = 72;

    // ══════════════════════════════════════════════════════════════
    //  LOAD & FILTER DATA
    // ══════════════════════════════════════════════════════════════
    const dataSheet = workbook.getWorksheet(DATA_SHEET);
    if (!dataSheet) throw new Error(`Sheet "${DATA_SHEET}" not found.`);

    const sprint = Number(dataSheet.getRange("B2").getValue());
    if (isNaN(sprint)) throw new Error("B2 does not contain a valid sprint number.");

    const used = dataSheet.getUsedRange();
    const numRows = used.getRowCount();
    const numCols = Math.max(9, used.getColumnCount());

    const rawData = dataSheet
        .getRangeByIndexes(3, 0, numRows - 3, numCols)
        .getValues();

    const rows = rawData.filter(row => Number(row[C_SPRINT]) === sprint);
    if (!rows.length) throw new Error(`No rows found for Sprint ${sprint}.`);

    function normalizeStatus(raw: string): string {
        return STATUSES.includes(raw) ? raw : "In-progress";
    }

    // ══════════════════════════════════════════════════════════════
    //  AGGREGATE
    // ══════════════════════════════════════════════════════════════

    const aMap: { [name: string]: { [status: string]: number } } = {};
    rows.forEach(row => {
        const name = String(row[C_ASSIGNEE]).trim() || "Unassigned";
        const status = normalizeStatus(String(row[C_STATUS]).trim());
        if (!aMap[name]) { aMap[name] = {}; STATUSES.forEach(s => (aMap[name][s] = 0)); }
        if (aMap[name][status] !== undefined) aMap[name][status]++;
    });
    const people = Object.keys(aMap).sort();

    const typeMap: { [k: string]: number } = {};
    TYPES.forEach(t => (typeMap[t] = 0));
    rows.forEach(row => {
        const t = String(row[C_TYPE]).trim();
        typeMap[t] !== undefined
            ? typeMap[t]++
            : (typeMap["Other"] = (typeMap["Other"] || 0) + 1);
    });
    const typeRows = [...TYPES, "Other"]
        .filter(t => (typeMap[t] || 0) > 0)
        .map(t => [t, typeMap[t]] as [string, number]);

    const scopeMap: { [k: string]: number } = {};
    SCOPES.forEach(s => (scopeMap[s] = 0));
    rows.forEach(row => {
        const s = String(row[C_SCOPE]).trim();
        scopeMap[s] !== undefined
            ? scopeMap[s]++
            : (scopeMap["Other"] = (scopeMap["Other"] || 0) + 1);
    });
    const scopeRows = [...SCOPES, "Other"]
        .filter(s => (scopeMap[s] || 0) > 0)
        .map(s => [s, scopeMap[s]] as [string, number]);

    const statusMap: { [k: string]: number } = {};
    STATUSES.forEach(s => (statusMap[s] = 0));
    rows.forEach(row => {
        const s = normalizeStatus(String(row[C_STATUS]).trim());
        if (statusMap[s] !== undefined) statusMap[s]++;
    });
    const statusRows = STATUSES.map(s => [s, statusMap[s]] as [string, number]);

    // ══════════════════════════════════════════════════════════════
    //  BUILD VELOCITY SHEET
    // ══════════════════════════════════════════════════════════════
    const sheetName = `Sprint ${sprint} Velocity`;
    const existing = workbook.getWorksheet(sheetName);
    if (existing) existing.delete();
    const vel = workbook.addWorksheet(sheetName);

    function sectionTitle(row: number, text: string): void {
        const c = vel.getRangeByIndexes(row, 0, 1, 1);
        c.setValue(text);
        c.getFormat().getFont().setBold(true);
        c.getFormat().getFont().setSize(13);
        c.getFormat().getFont().setColor("#1F3864");
    }

    function tableHeader(row: number, labels: string[], bgColor: string): void {
        const c = vel.getRangeByIndexes(row, 0, 1, labels.length);
        c.setValues([labels]);
        c.getFormat().getFill().setColor(bgColor);
        c.getFormat().getFont().setBold(true);
        c.getFormat().getFont().setColor("#FFFFFF");
    }

    function tableData(startRow: number, vals: (string | number)[][]): void {
        if (!vals.length) return;
        vel.getRangeByIndexes(startRow, 0, vals.length, vals[0].length)
            .setValues(vals as (string | number | boolean)[][]);
    }

    // ─── Section 1: Status by Assignee ───────────────────────────
    let r = 0;
    sectionTitle(r++, `Sprint ${sprint}  |  Status by Assignee`);

    const s1Heads = ["Assignee", ...STATUSES, "Total"];
    tableHeader(r, s1Heads, "#2F5496");
    const s1HRow = r++;

    STATUSES.forEach((s, i) => {
        const color = STATUS_COLORS[s];
        if (color) {
            const cell = vel.getRangeByIndexes(s1HRow, i + 1, 1, 1);
            cell.getFormat().getFill().setColor(color);
            cell.getFormat().getFont().setColor("#1F3864");
        }
    });

    const s1Body: (string | number)[][] = people.map(p => {
        const cnts = STATUSES.map(s => aMap[p][s] || 0);
        return [p, ...cnts, cnts.reduce((a, b) => a + b, 0)];
    });
    tableData(r, s1Body);
    r += s1Body.length;

    const colTotals = STATUSES.map((_s, i) =>
        s1Body.reduce((sum, row) => sum + (row[i + 1] as number), 0)
    );
    const grandTotal = colTotals.reduce((a, b) => a + b, 0);
    const gRow = vel.getRangeByIndexes(r, 0, 1, s1Heads.length);
    gRow.setValues([["TOTAL", ...colTotals, grandTotal]]);
    gRow.getFormat().getFont().setBold(true);
    gRow.getFormat().getFill().setColor("#BDD7EE");
    r++;

    const s1CRange = vel.getRangeByIndexes(s1HRow, 0, people.length + 1, STATUSES.length + 1);
    const s1Chart = vel.addChart(ExcelScript.ChartType.barStacked, s1CRange, ExcelScript.ChartSeriesBy.columns);
    s1Chart.getTitle().setText(`Sprint ${sprint} — Status by Assignee`);
    s1Chart.setTop(s1HRow * ROW_PT);
    s1Chart.setLeft(s1Heads.length * COL_PT + 15);
    s1Chart.setWidth(580);
    s1Chart.setHeight(Math.max(300, people.length * 22 + 80));
    // @ts-ignore
    s1Chart.getSeries().forEach(series => {
        const color = STATUS_COLORS[series.getName()];
        if (color) series.getFormat().getFill().setSolidColor(color);
    });
    r += 2;

    // ─── Section 2: Count by Process Type ────────────────────────
    const s2Top = r;
    sectionTitle(r++, "Count by Process Type");
    tableHeader(r, ["Process Type", "Count"], "#375623");
    const s2HRow = r++;
    tableData(r, typeRows);
    r += typeRows.length;

    if (typeRows.length > 0) {
        const s2CRange = vel.getRangeByIndexes(s2HRow, 0, typeRows.length + 1, 2);
        const s2Chart = vel.addChart(ExcelScript.ChartType.pie, s2CRange, ExcelScript.ChartSeriesBy.columns);
        s2Chart.getTitle().setText(`Sprint ${sprint} — Process Type`);
        s2Chart.setTop(s2Top * ROW_PT);
        s2Chart.setLeft(3 * COL_PT + 15);
        s2Chart.setWidth(380);
        s2Chart.setHeight(280);
        s2Chart.getDataLabels().setShowPercentage(true);
        s2Chart.getDataLabels().setShowValue(false);
        s2Chart.getDataLabels().setShowCategoryName(true);
        // @ts-ignore
        const s2Slices = s2Chart.getSeries()[0].getPoints();
        typeRows.forEach(([name], i) => {
            const color = TYPE_COLORS[name];
            if (color) s2Slices[i].getFormat().getFill().setSolidColor(color);
        });
    }
    r += 2;

    // ─── Section 3: Scope Changes ─────────────────────────────────
    const s3Top = r;
    sectionTitle(r++, "Scope Changes");
    tableHeader(r, ["Scope Change", "Count"], "#C55A11");
    const s3HRow = r++;
    tableData(r, scopeRows);
    r += scopeRows.length;

    if (scopeRows.length > 0) {
        const s3CRange = vel.getRangeByIndexes(s3HRow, 0, scopeRows.length + 1, 2);
        const s3Chart = vel.addChart(ExcelScript.ChartType.pie, s3CRange, ExcelScript.ChartSeriesBy.columns);
        s3Chart.getTitle().setText(`Sprint ${sprint} — Scope Changes`);
        s3Chart.setTop(s3Top * ROW_PT);
        s3Chart.setLeft(3 * COL_PT + 15);
        s3Chart.setWidth(380);
        s3Chart.setHeight(260);
        s3Chart.getDataLabels().setShowPercentage(true);
        s3Chart.getDataLabels().setShowValue(false);
        s3Chart.getDataLabels().setShowCategoryName(true);
        // @ts-ignore
        const s3Slices = s3Chart.getSeries()[0].getPoints();
        scopeRows.forEach(([name], i) => {
            const color = SCOPE_COLORS[name];
            if (color) s3Slices[i].getFormat().getFill().setSolidColor(color);
        });
    }
    r += 2;

    // ─── Section 4: Count by Status ──────────────────────────────
    const s4Top = r;
    sectionTitle(r++, "Count by Status");
    tableHeader(r, ["Status", "Count"], "#7030A0");
    const s4HRow = r++;
    tableData(r, statusRows);
    r += statusRows.length;

    const s4CRange = vel.getRangeByIndexes(s4HRow, 0, statusRows.length + 1, 2);
    const s4Chart = vel.addChart(ExcelScript.ChartType.pie, s4CRange, ExcelScript.ChartSeriesBy.columns);
    s4Chart.getTitle().setText(`Sprint ${sprint} — Ticket Status`);
    s4Chart.setTop(s4Top * ROW_PT);
    s4Chart.setLeft(3 * COL_PT + 15);
    s4Chart.setWidth(400);
    s4Chart.setHeight(260);
    s4Chart.getDataLabels().setShowPercentage(true);
    s4Chart.getDataLabels().setShowValue(false);
    s4Chart.getDataLabels().setShowCategoryName(true);
    // @ts-ignore
    const s4Slices = s4Chart.getSeries()[0].getPoints();
    statusRows.forEach(([name], i) => {
        const color = STATUS_COLORS[name];
        if (color) s4Slices[i].getFormat().getFill().setSolidColor(color);
    });

    // ─── Autofit ──────────────────────────────────────────────────
    vel.getUsedRange().getFormat().autofitColumns();

    // ══════════════════════════════════════════════════════════════
    //  HIDE VELOCITY SHEET IMMEDIATELY  ← NEW
    //  Must come after all content is written — hiding first would
    //  not cause errors but this order is safer and more explicit.
    // ══════════════════════════════════════════════════════════════
    vel.setVisibility(ExcelScript.SheetVisibility.hidden);

    // ══════════════════════════════════════════════════════════════
    //  UPDATE TOC  ← NEW
    // ══════════════════════════════════════════════════════════════
    updateTOC(workbook, sprint, sheetName, TOC_SHEET);
    updateVelocityChart(workbook, rawData, TOC_SHEET);
}

// ════════════════════════════════════════════════════════════════════
//  updateVelocityChart
//  Writes a "Committed vs Done with Completion %" combo chart to the
//  Velocity TOC sheet.  Data table starts at E2; chart is placed
//  directly below it.  Rebuilds from scratch on every run.
// ════════════════════════════════════════════════════════════════════
function updateVelocityChart(
    workbook: ExcelScript.Workbook,
    rawData: (string | number | boolean)[][],
    tocName: string
): void {
    const C_SPRINT = 0;
    const C_SCOPE  = 5;
    const C_STATUS = 8;
    const CHART_TITLE = "Sprint Committed vs Done with Completion %";

    const toc = workbook.getWorksheet(tocName);
    if (!toc) return;

    // ── Aggregate committed / done counts by sprint ───────────────
    const sprintData: { [num: number]: { committed: number; done: number } } = {};
    for (const row of rawData) {
        const sprintNum = Number(row[C_SPRINT]);
        if (!sprintNum || isNaN(sprintNum)) continue;
        const scope  = String(row[C_SCOPE]).trim();
        const status = String(row[C_STATUS]).trim();
        if (scope === "Move out of Sprint") continue;
        if (!sprintData[sprintNum]) sprintData[sprintNum] = { committed: 0, done: 0 };
        sprintData[sprintNum].committed++;
        if (status === "Done") sprintData[sprintNum].done++;
    }

    const sortedSprints = Object.keys(sprintData)
        .map(k => Number(k))
        .sort((a, b) => a - b)
        .slice(-10);

    if (sortedSprints.length === 0) return;

    // ── Remove any existing chart with this title ─────────────────
    // @ts-ignore
    toc.getCharts().forEach(c => {
        try { if (c.getTitle().getText() === CHART_TITLE) c.delete(); }
        catch (_e) { /* chart has no title */ }
    });

    // ── Clear old data table area: E2 down 12 rows (header + 10 + buffer) ──
    toc.getRangeByIndexes(1, 4, 12, 4).clear(ExcelScript.ClearApplyTo.all);

    // ── Write header row at E2 (row index 1, col index 4) ────────
    const hdr = toc.getRangeByIndexes(1, 4, 1, 4);
    hdr.setValues([["Sprint", "Committed", "Done", "Completion %"]]);
    hdr.getFormat().getFill().setColor("#2F5496");
    hdr.getFormat().getFont().setBold(true);
    hdr.getFormat().getFont().setColor("#FFFFFF");

    // ── Write data rows ───────────────────────────────────────────
    // Sprint stored as string so Excel treats it as a category axis, not a series.
    const tableRows: (string | number)[][] = sortedSprints.map(s => {
        const d = sprintData[s];
        const pct = d.committed > 0 ? d.done / d.committed : 0;
        return [String(s), d.committed, d.done, pct];
    });
    toc.getRangeByIndexes(2, 4, tableRows.length, 4)
        .setValues(tableRows as (string | number | boolean)[][]);

    // Format Completion % column (H) as percentage
    toc.getRangeByIndexes(2, 7, tableRows.length, 1).setNumberFormatLocal("0%");

    // Autofit E–H columns to their content
    toc.getRangeByIndexes(1, 4, tableRows.length + 1, 4).getFormat().autofitColumns();

    // ── Create chart anchored below the data table ────────────────
    // Chart range is F–H only (Committed, Done, Completion %) so the
    // Sprint column (E) cannot be misread as a data series by Excel.
    // Sprint numbers are wired in as category labels via setXAxisValues.
    const seriesDataRange  = toc.getRangeByIndexes(1, 5, tableRows.length + 1, 3);
    const sprintLabelRange = toc.getRangeByIndexes(2, 4, tableRows.length, 1);
    const anchorCell       = toc.getRangeByIndexes(2 + tableRows.length + 1, 4, 1, 1);

    const chart = toc.addChart(
        ExcelScript.ChartType.columnClustered,
        seriesDataRange,
        ExcelScript.ChartSeriesBy.columns
    );
    chart.getTitle().setText(CHART_TITLE);
    chart.setTop(anchorCell.getTop());
    chart.setLeft(anchorCell.getLeft());
    chart.setWidth(500);
    chart.setHeight(300);

    // ── Configure series ──────────────────────────────────────────
    // With Sprint excluded, series are: [0]=Committed, [1]=Done, [2]=Completion %
    // @ts-ignore
    const allSeries: ExcelScript.ChartSeries[] = chart.getSeries();

    // Bind sprint numbers as the x-axis category labels for every series
    allSeries[0].setXAxisValues(sprintLabelRange);
    allSeries[1].setXAxisValues(sprintLabelRange);
    allSeries[2].setXAxisValues(sprintLabelRange);

    const seriesCommitted = allSeries[0];
    const seriesDone      = allSeries[1];
    const seriesPct       = allSeries[2];

    seriesCommitted.getFormat().getFill().setSolidColor("#4472C4"); // blue bars
    seriesDone.getFormat().getFill().setSolidColor("#ED7D31");      // orange bars

    // Convert Completion % to a line with circle markers on the secondary axis
    seriesPct.setChartType(ExcelScript.ChartType.lineMarkers);
    // @ts-ignore
    seriesPct.setAxisGroup(ExcelScript.ChartAxisGroup.secondary);
    seriesPct.getFormat().getLine().setColor("#70AD47");            // green line
    seriesPct.setMarkerStyle(ExcelScript.ChartMarkerStyle.circle);

    // ── Data labels ───────────────────────────────────────────────
    seriesCommitted.getDataLabels().setShowValue(true);
    seriesDone.getDataLabels().setShowValue(true);
    seriesPct.getDataLabels().setShowValue(true);
    seriesPct.getDataLabels().setNumberFormat("0%");

    // ── Y-axis titles and formats ─────────────────────────────────
    const primaryAxis = chart.getAxes().getValueAxis();
    primaryAxis.getTitle().setText("Work Items");
    primaryAxis.getTitle().setVisible(true);
    primaryAxis.setNumberFormat("0");   // whole numbers, not percentages
    primaryAxis.setMinimum(0);

    // Secondary axis: getItem(value, secondary) is the correct ExcelScript API.
    try {
        // @ts-ignore
        const secAxis: ExcelScript.ChartAxis = chart.getAxes().getItem(ExcelScript.ChartAxisType.value, ExcelScript.ChartAxisGroup.secondary);
        secAxis.getTitle().setText("Completion %");
        secAxis.getTitle().setVisible(true);
        secAxis.setNumberFormat("0%");
        secAxis.setMaximum(1);
        secAxis.setMinimum(0);
    } catch (_e) { /* secondary axis not yet available at script time */ }
}

// ════════════════════════════════════════════════════════════════════
//  updateTOC
//  Writes or refreshes one row in the Velocity TOC for this sprint.
//  • First run ever: initializes title + column headers
//  • Same sprint re-run: updates the date, refreshes hyperlink
//  • New sprint: appends a new row at the bottom
// ════════════════════════════════════════════════════════════════════
function updateTOC(
    workbook: ExcelScript.Workbook,
    sprint: number,
    sheetName: string,
    tocName: string
): void {

    const toc = workbook.getWorksheet(tocName);
    if (!toc) {
        console.log(`Warning: TOC sheet "${tocName}" not found — skipping TOC update.`);
        return;
    }

    // M/D/YYYY  e.g.  6/15/2026
    const today = new Date();
    const dateStr = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;

    // ── One-time setup: write title + headers if A1 is empty ────
    const a1Val = String(toc.getRange("A1").getValue()).trim();
    if (a1Val === "") {
        const title = toc.getRange("A1");
        title.setValue("Sprint Velocity Log");
        title.getFormat().getFont().setBold(true);
        title.getFormat().getFont().setSize(16);
        title.getFormat().getFont().setColor("#1F3864");

        const hdr = toc.getRange("A2:B2");
        hdr.setValues([["Sprint", "Date Generated"]]);
        hdr.getFormat().getFill().setColor("#2F5496");
        hdr.getFormat().getFont().setBold(true);
        hdr.getFormat().getFont().setColor("#FFFFFF");
    }

    // ── Batch-read column A rows 3–202 to find existing entry ───
    //    (one API call instead of looping per cell)
    const scanValues = toc.getRangeByIndexes(2, 0, 200, 1).getValues();
    const displayName = `Sprint ${sprint}`;   // e.g. "Sprint 5"

    let foundRow = -1;
    let nextEmptyRow = 2;   // 0-indexed; row index 2 = sheet row 3

    for (let i = 0; i < scanValues.length; i++) {
        const cellVal = String(scanValues[i][0]).trim();
        if (cellVal === "") {
            nextEmptyRow = 2 + i;
            break;
        }
        if (cellVal === displayName) {
            foundRow = 2 + i;
        }
        nextEmptyRow = 2 + i + 1;
    }

    const writeRow = foundRow >= 0 ? foundRow : nextEmptyRow;

    // ── Sprint name cell with clickable hyperlink ────────────────
    const nameCell = toc.getRangeByIndexes(writeRow, 0, 1, 1);
    nameCell.setValue(displayName);
    nameCell.setHyperlink({
        address: "",
        documentReference: `'${sheetName}'!A1`,   // e.g. 'Sprint 5 Velocity'!A1
        screenTip: `Open ${sheetName}`,
        textToDisplay: displayName
    });
    nameCell.getFormat().getFont().setColor("#0563C1");
    nameCell.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.single);

    // ── Date cell ────────────────────────────────────────────────
    toc.getRangeByIndexes(writeRow, 1, 1, 1).setValue(dateStr);

    // ── Autofit ──────────────────────────────────────────────────
    toc.getUsedRange().getFormat().autofitColumns();
}
