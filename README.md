# VBA Matrix Scheduler

> High-performance 2D Schedule & Attendance Matrix Generator for Microsoft Excel VBA.  
> Powered by the In-Memory Columnar DataFrame (`cDataFrame`) Architecture.

---

## 1. Overview

Traditional Excel VBA macros for generating cross-tabulation schedule matrices typically suffer from two major drawbacks:
1. **Slow Cell-by-Cell Worksheet Access**: Looping through rows and columns using `Cells(r, c).Value = ...` creates massive COM interop overhead, taking 10–60+ seconds on realistic enterprise rosters.
2. **Brittle Lookups & Merges**: Using worksheet formulas (`VLOOKUP`, `INDEX/MATCH`) or nested $O(N \times M)$ loops makes handling empty spacer rows, team groupings, and multi-shift conflicts tedious and error-prone.

**`VBA Matrix Scheduler`** solves these challenges by combining:
- **In-Memory Columnar Engine (`cDataFrame` & `cColumn`)**: Performs all sorting, grouping, matrix transformations, and index lookups entirely in RAM.
- **Single COM Read/Write**: Dumps the entire 2D matrix back to the worksheet in a single instantaneous `.Value2` block write (<0.1s).
- **$O(1)$ Hash Matrix Indexing**: Uses `SetIndex` and `Loc(Person, Day)` to map transactions directly to matrix intersections at constant time.
- **Automated Conflict Resolution**: Supports `"CONCAT"` (e.g., combining morning `WFH` and evening `OT` into `"WFH / OT"`), `"SUM"` (for work hours), and `"OVERWRITE"`.
- **Decoupled Styling Engine (`cTableStyler`)**: Beautiful executive theme styling, group banners, weekend/weekday headers, and status color tags.

---

## 2. Architecture & Data Flow

```mermaid
flowchart TD
    subgraph Input ["Source Excel Tables"]
        T1["tbEmployees (Master)<br/>Name, Team, Role"]
        T2["tbEvents (Transactions)<br/>Person, Day, Status"]
    end

    subgraph Memory ["In-Memory Processing (cDataFrame)"]
        D1["dfSource: Sort by Team & Name in RAM"]
        D2["dfMatrix: Construct Name + Day 1..N columns<br/>Insert Team Section Banners & Spacer Rows"]
        D3["SetIndex('Name') -> O(1) Hash Map Index"]
        D4["Iterate tbEvents -> Loc(Person, Day) = Status<br/>Apply Conflict Policy: CONCAT / SUM / OVERWRITE"]
    end

    subgraph Output ["Single COM Presentation"]
        W1["Sheet 'ScheduleMatrix' (B5:AF100)<br/>Single Block Write via .Value2"]
        W2["cTableStyler Presentation Layer<br/>Navy Header, Team Banners, Status Color Tags"]
    end

    T1 --> D1
    D1 --> D2
    D2 --> D3
    T2 --> D4
    D3 --> D4
    D4 --> W1
    W1 --> W2
```

---

## 3. Project Structure

```text
vba-matrix-scheduler/
├── README.md                        # Project documentation & usage guide
├── MatrixScheduler.xlsm             # 30-day schedule & attendance demo workbook
├── MonthlyTaskCompletionMatrix.xlsm # Unified 12-month task completion matrix (daily detail + weekly grid)
├── .gitignore                       # Git ignore rules for Office & temp files
├── src/
│   └── vba/
│       ├── cColumn.cls              # 1D Variant vector store with geometric buffer expansion
│       ├── cDataFrame.cls           # In-memory columnar DataFrame (BuildCalendarMatrix, TryParseExcelDate, Loc, Sort, Merge)
│       ├── cDataValidator.cls       # Schema contract & data quality assurance engine
│       ├── cReportBuilder.cls       # Executive report composer (AddSectionHeader, AddQuickJumpBar, AddActionButton)
│       ├── cTableStyler.cls         # Fluent styling engine (ApplyWeeklyCalendarGrid, SpotlightColumn, DimColumns)
│       ├── mMatrixScheduler.bas     # Matrix generation and event mapping workflow
│       └── mMonthlyTaskMatrix.bas   # Unified 12-month task completion matrix orchestrator
└── scripts/
    └── build_workbook.py            # Automated Python script to build & verify MatrixScheduler.xlsm
```

---

## 4. Quick Start (Using Interactive Demo Workbook)

1. Open `MatrixScheduler.xlsm` in Microsoft Excel.
2. Enable macros when prompted by Excel.
3. On sheet **`ScheduleMatrix`**, you will find 4 action buttons:
   - **`1. Generate Matrix`**: Reads employee records from `tbEmployees`, groups by team, and generates a 30-day schedule grid with styled team headers.
   - **`2. Map Schedule Data`**: Reads transaction records from `tbEvents` and writes status tags into cell intersections with automated color highlighting (`Leave` = Coral, `WFH` = Sky Blue, `OT` = Purple, `Work` = Mint Green).
   - **`Run All (1 + 2)`**: Executes the full end-to-end pipeline in one click.
   - **`Clear Matrix`**: Clears the output matrix range.

### 4.1 Monthly Task Completion Matrix (`MonthlyTaskCompletionMatrix.xlsm`)

Open **`MonthlyTaskCompletionMatrix.xlsm`** for the unified single-sheet monthly task completion report:

- **Single-Sheet Unified View (`TaskCompletionMatrix`)**: All 12 months (`2026-01` through `2026-12`) are stacked vertically in a single aligned 35-column grid (`Assignee`, `Department`, `D01`..`D31`, `Month Total`, `Daily Avg`).
- **Weekly Grid Separation**: Each month displays `WEEK 1`..`W5` + `SUMMARY` group banners and thick outer week borders generated by `cTableStyler.ApplyWeeklyCalendarGrid`.
- **Daily Live Auto-Update**: `ThisWorkbook` hooks (`Workbook_Open`, `Workbook_SheetChange`, and `Workbook_SheetActivate`) automatically recalculate task completion counts from `tbTasks` whenever the workbook opens or task data changes, and spotlight today's date column (`SpotlightColumn`).
- **Same-Sheet Quick-Jump Bar**: Clickable `Jan '26`..`Dec '26` navigation pills (`cReportBuilder.AddQuickJumpBar`) scroll directly to any month block on the sheet.

---

## 5. Integrating into Existing Workbooks

To integrate this engine into any existing Excel project:

1. Press `ALT + F11` to open the VBA Editor.
2. Right-click your VBA project and choose **Import File...** (`Ctrl + M`).
3. Import the 4 files from `src/vba/`:
   - `cColumn.cls`
   - `cDataFrame.cls`
   - `cTableStyler.cls`
   - `mMatrixScheduler.bas`
4. Call the core procedures directly in your code:

```vb
Sub RunMySchedulePipeline()
    Dim wsEmp As Worksheet, wsMatrix As Worksheet, wsEvents As Worksheet
    Dim tbEmp As ListObject, tbEvents As ListObject
    
    Set wsEmp = ThisWorkbook.Sheets("Employees")
    Set wsEvents = ThisWorkbook.Sheets("Events")
    Set wsMatrix = ThisWorkbook.Sheets("ScheduleMatrix")
    
    Set tbEmp = wsEmp.ListObjects("tbEmployees")
    Set tbEvents = wsEvents.ListObjects("tbEvents")
    
    ' 1. Generate 30-day matrix starting at cell B5
    GenerateMatrixFromTable SourceTable:=tbEmp, _
                            StartCell:=wsMatrix.Range("B5"), _
                            NumDays:=30, _
                            GroupCol:="Team", _
                            NameCol:="Name", _
                            SpacerRows:=True
                            
    ' 2. Map events to cell intersections via RAM
    MapEventsToMatrix EventsTable:=tbEvents, _
                      StartCell:=wsMatrix.Range("B5"), _
                      PersonCol:="Person", _
                      DateCol:="Day", _
                      StatusCol:="Status", _
                      ConflictMode:="CONCAT"
                      
    ' 3. Apply status color tags
    ApplyStatusFormatting wsMatrix.Range("B5")
End Sub
```

---

## 6. License & Credits

- **Author**: Technodevmoo AI Engineering
- **Architecture**: Powered by `cDataFrame` Column-Oriented Table Store
- **License**: MIT


---

## 9. Developer Guidelines & VBA Class Gotchas (Best Practices)

When developing or calling methods in `cDataFrame` and related classes, keep in mind these key VBA language subtleties:

1. **Invoke Object-Returning Methods with `Call` or `Set` (Avoid Error 449):**
   Classes with default members (e.g., `Item(RowIndex, ColumnKey)` in `cDataFrame`) cause VBA to evaluate the default property if called as a bare statement with parentheses (`df.LoadFrom3DArray(...)`). Always use `Call df.Method(...)` or `Set df = df.Method(...)` to discard or assign the returned object cleanly without triggering `Error 449 (Argument not optional)`.
2. **Optional Object Parameters Must Omit Default Values:**
   VBA syntax strictly forbids default values on object parameters (`Optional ByVal TargetWorkbook As Object = Nothing` causes a compile syntax error). Use `Optional ByVal TargetWorkbook As Object` and perform `If TargetWorkbook Is Nothing Then` inside the procedure.
3. **Keep Class Modules Fully Decoupled:**
   Never reference standard test or runner modules (e.g. `LogPrint` in `mTestRunner.bas`) inside class modules. Standalone classes must compile cleanly in isolation.
4. **Prevent Headless Excel COM Hangs During Automated Tests:**
   VBE pops a modal dialog (`#32770`) on any compile/syntax error, halting headless Excel automation until the watchdog timeout fires. Always run a pre-execution compile check or employ an asynchronous Win32 dialog monitor (`WM_COMMAND` with `PostMessage`) during automated test runs.
5. **Enforce Strict CRLF (`\r\n`) Line Endings:**
   The VBE import mechanism on Windows requires CRLF line endings. Corrupted line breaks (`\n` or `\r\r\n`) can break the `VERSION 1.0 CLASS` header, causing class files to be imported as plain standard modules.
