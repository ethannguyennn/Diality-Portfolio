
function main(workbook: ExcelScript.Workbook): void {

    // ══════════════════════════════════════════════════════════════
    //  CONFIG — update DATA_SHEET if the tab name differs
    // ══════════════════════════════════════════════════════════════
    const DATA_SHEET = "Sprint Template";

    // 0-based column indices  (A=0  B=1  C=2 … H=7)
    const C_SPRINT = 0; // A — Sprint #
    const C_ASSIGNEE = 2; // C — Current Assignee  (B = Original, ignored)
    const C_TYPE = 3; // D — Process Type
    const C_SCOPE = 4; // E — Scope Changes
    const C_STATUS = 7; // H — Ticket Status

    // Exact string values from the reference sheet
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

    const ROW_PT = 15;  // approx Excel row height in points (used for chart placement)
    const COL_PT = 72;  // approx Excel column width in points

    // ══════════════════════════════════════════════════════════════
    //  LOAD & FILTER DATA
    // ══════════════════════════════════════════════════════════════
    const dataSheet = workbook.getWorksheet(DATA_SHEET);
    if (!dataSheet) throw new Error(`Sheet "${DATA_SHEET}" not found.`);

    // B2 holds the sprint number currently being managed
    const sprint = Number(dataSheet.getRange("B2").getValue());
    if (isNaN(sprint)) throw new Error("B2 does not contain a valid sprint number.");

    const used = dataSheet.getUsedRange();
    const numRows = used.getRowCount();
    const numCols = Math.max(8, used.getColumnCount());

    // Sheet layout: Row 1 = title area | Row 2 = metadata (B2 = sprint#)
    //               Row 3 = column headers | Row 4+ = data  →  start at index 3
    const rawData = dataSheet
        .getRangeByIndexes(3, 0, numRows - 3, numCols)
        .getValues();

    const rows = rawData.filter(row => Number(row[C_SPRINT]) === sprint);
    if (!rows.length) throw new Error(`No rows found for Sprint ${sprint}.`);

    // ══════════════════════════════════════════════════════════════
    //  AGGREGATE
    // ══════════════════════════════════════════════════════════════

    // 1 ─ Status by Assignee
    const aMap: { [name: string]: { [status: string]: number } } = {};
    rows.forEach(row => {
        const name = String(row[C_ASSIGNEE]).trim() || "Unassigned";
        const status = String(row[C_STATUS]).trim();
        if (!aMap[name]) { aMap[name] = {}; STATUSES.forEach(s => (aMap[name][s] = 0)); }
        if (aMap[name][status] !== undefined) aMap[name][status]++;
    });
    const people = Object.keys(aMap).sort();

    // 2 ─ Count by Process Type  (skip zero-count entries so pie has no empty slices)
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

    // 3 ─ Scope Changes
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

    // 4 ─ Count by Status
    const statusMap: { [k: string]: number } = {};
    STATUSES.forEach(s => (statusMap[s] = 0));
    rows.forEach(row => {
        const s = String(row[C_STATUS]).trim();
        if (statusMap[s] !== undefined) statusMap[s]++;
    });
    const statusRows = STATUSES.map(s => [s, statusMap[s]] as [string, number]);

    // ══════════════════════════════════════════════════════════════
    //  BUILD VELOCITY SHEET  (deleted & recreated each run)
    // ══════════════════════════════════════════════════════════════
    const sheetName = `Sprint ${sprint} Velocity`;
    const existing = workbook.getWorksheet(sheetName);
    if (existing) existing.delete();
    const vel = workbook.addWorksheet(sheetName);

    // ─── helpers ─────────────────────────────────────────────────
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

    // ─────────────────────────────────────────────────────────────
    //  SECTION 1 — Status by Assignee
    //  Table: Assignee | Blocked | Done | In-progress | Not started | On-hold | N/A | Total
    //  Chart: Stacked horizontal bar (to the right of the table)
    // ─────────────────────────────────────────────────────────────
    let r = 0;

    sectionTitle(r++, `Sprint ${sprint}  |  Status by Assignee`);

    const s1Heads = ["Assignee", ...STATUSES, "Total"];
    tableHeader(r, s1Heads, "#2F5496");
    const s1HRow = r++;

    // Color each status column header individually
    STATUSES.forEach((s, i) => {
        const color = STATUS_COLORS[s];
        if (color) {
            const cell = vel.getRangeByIndexes(s1HRow, i + 1, 1, 1);
            cell.getFormat().getFill().setColor(color);
            cell.getFormat().getFont().setColor("#1F3864");
        }
    });

    // Build body rows:  [name, blocked, done, in-progress, not started, on-hold, N/A, total]
    const s1Body: (string | number)[][] = people.map(p => {
        const cnts = STATUSES.map(s => aMap[p][s] || 0);
        return [p, ...cnts, cnts.reduce((a, b) => a + b, 0)];
    });
    tableData(r, s1Body);
    r += s1Body.length;

    // Grand-total row at the bottom
    const colTotals = STATUSES.map((_s, i) =>
        s1Body.reduce((sum, bodyRow) => sum + (bodyRow[i + 1] as number), 0)
    );
    const grandTotal = colTotals.reduce((a, b) => a + b, 0);
    const gRow = vel.getRangeByIndexes(r, 0, 1, s1Heads.length);
    gRow.setValues([["TOTAL", ...colTotals, grandTotal]]);
    gRow.getFormat().getFont().setBold(true);
    gRow.getFormat().getFill().setColor("#BDD7EE");
    r++;

    // Stacked bar chart — excludes the "Total" column, placed to the right of the table
    const s1CRange = vel.getRangeByIndexes(s1HRow, 0, people.length + 1, STATUSES.length + 1);
    const s1Chart = vel.addChart(
        ExcelScript.ChartType.barStacked,
        s1CRange,
        ExcelScript.ChartSeriesBy.columns
    );
    s1Chart.getTitle().setText(`Sprint ${sprint} — Status by Assignee`);
    s1Chart.setTop(s1HRow * ROW_PT);
    s1Chart.setLeft(s1Heads.length * COL_PT + 15);
    s1Chart.setWidth(580);
    s1Chart.setHeight(Math.max(300, people.length * 22 + 80));

    // Apply status colors to each chart series
    // @ts-ignore — ExcelScript types unavailable in local TS environment
    s1Chart.getSeries().forEach(series => {
        const color = STATUS_COLORS[series.getName()];
        if (color) series.getFormat().getFill().setSolidColor(color);
    });

    r += 2; // gap

    // ─────────────────────────────────────────────────────────────
    //  SECTION 2 — Count by Process Type
    //  Table: Process Type | Count
    //  Chart: Pie (to the right of the table)
    // ─────────────────────────────────────────────────────────────
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

        // Color each pie slice by process type name
        // @ts-ignore — ExcelScript types unavailable in local TS environment
        const s2Slices = s2Chart.getSeries()[0].getPoints();
        typeRows.forEach(([name], i) => {
            const color = TYPE_COLORS[name];
            if (color) s2Slices[i].getFormat().getFill().setSolidColor(color);
        });
    }

    r += 2;

    // ─────────────────────────────────────────────────────────────
    //  SECTION 3 — Scope Changes
    //  Table: Scope Change | Count
    //  Chart: Pie (to the right of the table)
    // ─────────────────────────────────────────────────────────────
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

        // Color each pie slice by scope name
        // @ts-ignore — ExcelScript types unavailable in local TS environment
        const s3Slices = s3Chart.getSeries()[0].getPoints();
        scopeRows.forEach(([name], i) => {
            const color = SCOPE_COLORS[name];
            if (color) s3Slices[i].getFormat().getFill().setSolidColor(color);
        });
    }

    r += 2;

    // ─────────────────────────────────────────────────────────────
    //  SECTION 4 — Count by Status
    //  Table: Status | Count
    //  Chart: Pie
    // ─────────────────────────────────────────────────────────────
    const s4Top = r;
    sectionTitle(r++, "Count by Status");
    tableHeader(r, ["Status", "Count"], "#7030A0");
    const s4HRow = r++;
    tableData(r, statusRows);
    r += statusRows.length;

    const s4CRange = vel.getRangeByIndexes(s4HRow, 0, statusRows.length + 1, 2);
    const s4Chart = vel.addChart(
        ExcelScript.ChartType.pie,
        s4CRange,
        ExcelScript.ChartSeriesBy.columns
    );
    s4Chart.getTitle().setText(`Sprint ${sprint} — Ticket Status`);
    s4Chart.setTop(s4Top * ROW_PT);
    s4Chart.setLeft(3 * COL_PT + 15);
    s4Chart.setWidth(400);
    s4Chart.setHeight(260);
    s4Chart.getDataLabels().setShowPercentage(true);
    s4Chart.getDataLabels().setShowValue(false);
    s4Chart.getDataLabels().setShowCategoryName(true);

    // Color each pie slice using STATUS_COLORS
    // @ts-ignore — ExcelScript types unavailable in local TS environment
    const s4Slices = s4Chart.getSeries()[0].getPoints();
    statusRows.forEach(([name], i) => {
        const color = STATUS_COLORS[name];
        if (color) s4Slices[i].getFormat().getFill().setSolidColor(color);
    });

    // ─────────────────────────────────────────────────────────────
    //  AUTOFIT COLUMNS
    // ─────────────────────────────────────────────────────────────
    vel.getUsedRange().getFormat().autofitColumns();
}
