import win32com.client
import os
import sys
import datetime
import random

project_dir = os.path.dirname(os.path.abspath(__file__))
vba_dir = os.path.join(project_dir, "vba")
xlsm_path = os.path.join(project_dir, "MonthlyTaskCompletionMatrix.xlsm")
xlsb_path = os.path.join(project_dir, "MonthlyTaskCompletionMatrix.xlsb")
sync_dir = r"C:\Users\User\Documents\Projects\vba-matrix-scheduler"

print(f"Project directory: {project_dir}", flush=True)
print(f"Target XLSM: {xlsm_path}", flush=True)

# Clean up existing files
for p in [xlsm_path, xlsb_path]:
    if os.path.exists(p):
        try:
            os.remove(p)
            print(f"Removed old {os.path.basename(p)}", flush=True)
        except Exception as e:
            print(f"Warning: could not remove {p}: {e}", flush=True)

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    print("1. Creating blank workbook...", flush=True)
    wb = excel.Workbooks.Add()
    wb.VBProject.Name = "MonthlyTaskMatrixProject"

    ws_control = wb.Sheets(1)
    ws_control.Name = "ControlPanel"
    ws_emp = wb.Sheets.Add(After=ws_control)
    ws_emp.Name = "Employees"
    ws_tasks = wb.Sheets.Add(After=ws_emp)
    ws_tasks.Name = "RawTasks"

    # --------------------------------------------------------------------------
    # 2. Populate Employees Master
    # --------------------------------------------------------------------------
    print("2. Populating Employees master table...", flush=True)
    employees_data = [
        ["EmployeeID", "Name", "Department", "Role", "DailyTarget"],
        ["EMP-001", "Somchai Prasert", "Engineering", "Tech Lead", 2],
        ["EMP-002", "Wichai Srisuk", "Engineering", "Senior Backend Dev", 2],
        ["EMP-003", "Anan Chaiyaphum", "Engineering", "Frontend Specialist", 2],
        ["EMP-004", "Kanya Wongsuwan", "Engineering", "QA Automation Lead", 3],
        ["EMP-005", "Nattaporn Saelim", "Design", "Product Design Lead", 2],
        ["EMP-006", "Piti Rattanaporn", "Design", "UI/UX Designer", 2],
        ["EMP-007", "Siriporn Meesuk", "Design", "Design Systems Engineer", 2],
        ["EMP-008", "Chaiwat Kongkaew", "Product", "Principal Product Manager", 1],
        ["EMP-009", "Thanya Somboon", "Product", "Agile Coach / Scrum Master", 2],
        ["EMP-010", "Worawat Boonmee", "Product", "Senior Data Analyst", 2],
        ["EMP-011", "Maneerat Thongdee", "Operations", "Operations Manager", 2],
        ["EMP-012", "Narongrit Kaewdee", "Operations", "DevOps & SRE Specialist", 2],
    ]

    for r_idx, row in enumerate(employees_data, 1):
        for c_idx, val in enumerate(row, 1):
            ws_emp.Cells(r_idx + 1, c_idx + 1).Value = val

    emp_range = ws_emp.Range(ws_emp.Cells(2, 2), ws_emp.Cells(len(employees_data) + 1, len(employees_data[0]) + 1))
    tb_emp = ws_emp.ListObjects.Add(1, emp_range, None, 1)
    tb_emp.Name = "tbEmployees"
    tb_emp.TableStyle = "TableStyleMedium2"
    ws_emp.Columns("B:F").AutoFit()

    # --------------------------------------------------------------------------
    # 3. Populate Realistic 2026 Raw Tasks (2,300+ rows)
    # --------------------------------------------------------------------------
    print("3. Generating 2026 raw tasks data (~2,300 tasks)...", flush=True)
    random.seed(42)

    task_titles = {
        "Engineering": [
            ("API Endpoint Refactoring", "Feature", 3),
            ("Database Query Index Optimization", "DevOps", 5),
            ("Resolve Redis Cache Invalidation Bug", "Bugfix", 2),
            ("Implement OAuth2 JWT Middleware", "Security", 3),
            ("Add Retry Circuit Breaker", "Resilience", 5),
            ("Kubernetes HPA Tuning", "DevOps", 3),
            ("Docker Multi-stage Build Optimization", "DevOps", 2),
            ("Automated E2E Cypress Suite Run", "QA", 2),
            ("Kafka Event Schema Migration", "Architecture", 5),
            ("Fix Memory Leak in WebSocket Gateway", "Bugfix", 3),
            ("Refactor Customer Billing Service", "Feature", 8),
            ("CI/CD Pipeline Security Scan Integration", "DevOps", 3),
        ],
        "Design": [
            ("Design System Color Token Migration", "Design", 2),
            ("User Journey Mapping for Checkout Flow", "Research", 5),
            ("High-Fidelity Wireframes for Mobile App", "Design", 3),
            ("Accessibility WCAG 2.1 AA Audit", "Audit", 3),
            ("Component Variant Expansion in Figma", "Design", 2),
            ("Usability Testing Session & Synthesis", "Research", 5),
            ("Design System Typography Harmonization", "Design", 2),
            ("Dark Mode Palette Specification", "Design", 3),
        ],
        "Product": [
            ("Q2 Feature Prioritization Matrix", "Planning", 3),
            ("Customer Churn Root-Cause Analysis", "Analytics", 5),
            ("Sprint 14 Backlog Refinement & Sizing", "Planning", 2),
            ("User Feedback Sentiment Clustering", "Research", 3),
            ("Quarterly OKR Tracking Dashboard Update", "Reporting", 2),
            ("Competitive Benchmarking Report", "Strategy", 5),
            ("A/B Test Cohort Significance Review", "Analytics", 3),
            ("Stakeholder Roadmap Alignment Brief", "Strategy", 3),
        ],
        "Operations": [
            ("SSL/TLS Wildcard Certificate Renewal", "Security", 2),
            ("Incident Post-Mortem & Remediation Action", "Ops", 3),
            ("PostgreSQL Replica Failover Drill", "Resilience", 5),
            ("Production VPC Peering Security Review", "Security", 3),
            ("Monthly Cloud Cost Allocation Audit", "FinOps", 3),
            ("Backup Restoration Integrity Test", "Ops", 3),
            ("Grafana Dashboard Alert Threshold Tuning", "Monitoring", 2),
            ("Runbook Automation via Ansible", "Automation", 5),
        ]
    }

    raw_headers = ["TaskID", "TaskDate", "Assignee", "Department", "TaskTitle", "Category", "Status", "StoryPoints"]
    task_rows = [raw_headers]

    task_id_counter = 1001
    start_date = datetime.date(2026, 1, 1)
    end_date = datetime.date(2026, 12, 31)
    total_days = (end_date - start_date).days + 1

    dept_members = {}
    for emp in employees_data[1:]:
        dept = emp[2]
        dept_members.setdefault(dept, []).append(emp[1])

    today_date = datetime.date.today()

    for day_offset in range(total_days):
        current_date = start_date + datetime.timedelta(days=day_offset)
        is_weekend = (current_date.weekday() >= 5)

        if is_weekend:
            daily_task_count = random.randint(0, 2)
        else:
            daily_task_count = random.randint(6, 11)

        # Ensure today always has visible tasks even on quiet days
        if current_date == today_date and daily_task_count < 6:
            daily_task_count = 7

        for _ in range(daily_task_count):
            dept = random.choice(list(dept_members.keys()))
            assignee = random.choice(dept_members[dept])
            task_choice = random.choice(task_titles[dept])
            title, category, points = task_choice

            status_roll = random.random()
            if current_date <= today_date:
                if status_roll < 0.82:
                    status = "Completed"
                elif status_roll < 0.92:
                    status = "In Progress"
                elif status_roll < 0.97:
                    status = "Pending"
                else:
                    status = "Blocked"
            else:
                if status_roll < 0.70:
                    status = "Pending"
                elif status_roll < 0.94:
                    status = "In Progress"
                else:
                    status = "Blocked"

            task_rows.append([
                f"TSK-{task_id_counter}",
                current_date.strftime("%Y-%m-%d"),
                assignee,
                dept,
                title,
                category,
                status,
                points
            ])
            task_id_counter += 1

    print(f"   Generated {len(task_rows) - 1} task records across 365 days.", flush=True)

    print("   Writing batch array to RawTasks worksheet...", flush=True)
    num_rows = len(task_rows)
    num_cols = len(raw_headers)
    target_range = ws_tasks.Range(ws_tasks.Cells(1, 1), ws_tasks.Cells(num_rows, num_cols))
    target_range.Value = task_rows

    tb_tasks = ws_tasks.ListObjects.Add(1, target_range, None, 1)
    tb_tasks.Name = "tbTasks"
    tb_tasks.TableStyle = "TableStyleLight9"
    ws_tasks.Columns("A:H").AutoFit()

    # --------------------------------------------------------------------------
    # 4. Format ControlPanel Worksheet
    # --------------------------------------------------------------------------
    print("4. Designing ControlPanel dashboard...", flush=True)
    ws_control.Activate()
    ws_control.Tab.Color = 0x27AE60

    ws_control.Range("B2:J2").Merge()
    ws_control.Range("B2").Value = "Executive Monthly Task Completion Matrix"
    ws_control.Range("B2").Font.Size = 18
    ws_control.Range("B2").Font.Bold = True
    ws_control.Range("B2").Font.Color = 0xFFFFFF
    ws_control.Range("B2").Interior.Color = 0x502A1A
    ws_control.Range("B2").HorizontalAlignment = -4108
    ws_control.Range("B2").RowHeight = 36

    ws_control.Range("B3:J3").Merge()
    ws_control.Range("B3").Value = "High-Performance In-Memory Matrix Generation via cDataFrame & cReportBuilder"
    ws_control.Range("B3").Font.Size = 11
    ws_control.Range("B3").Font.Italic = True
    ws_control.Range("B3").Font.Color = 0x555555
    ws_control.Range("B3").Interior.Color = 0xF0F4F8
    ws_control.Range("B3").HorizontalAlignment = -4108
    ws_control.Range("B3").RowHeight = 22

    btn1 = ws_control.Buttons().Add(100, 85, 240, 36)
    btn1.Text = "Refresh Master Matrix (All 12 Months)"
    btn1.OnAction = "Btn_RefreshMatrix"

    btn2 = ws_control.Buttons().Add(360, 85, 200, 36)
    btn2.Text = "Jump to Current Month"
    btn2.OnAction = "Btn_JumpToCurrentMonth"

    btn3 = ws_control.Buttons().Add(580, 85, 170, 36)
    btn3.Text = "Clear Matrix Sheet"
    btn3.OnAction = "Btn_ClearSheets"

    info_rows = [
        ["Key Features", "Implementation Architecture & Benefits"],
        ["Unified Single-Sheet Matrix", "All 12 monthly tables (Jan-Dec) stacked vertically on one TaskCompletionMatrix sheet with aligned D01..D31 columns."],
        ["Daily Live Auto-Update", "Workbook_Open auto-refreshes every day on open; RawTasks edits auto-refresh on switching back to TaskCompletionMatrix."],
        ["Today Column Spotlight", "Dynamically highlights the active month and today's calendar column (Date) plus live YTD / Today / MTD KPI cards."],
        ["Weekly Grid Separation", "Distinct vertical borders (xlMedium) after Days 7, 14, 21, 28, and month end, plus alternating week fills and WEEK 1-5 banners."],
        ["Same-Sheet Quick Jump", "Clickable month pills (Jan..Dec) at the top of TaskCompletionMatrix scroll straight to any month block in 1 click."],
        ["In-Memory Tally Engine", "cDataFrame loads 2,300+ tasks in RAM and rebuilds all 12 months cleanly in ~2 seconds via cReportBuilder."]
    ]

    for r_idx, (k, v) in enumerate(info_rows, 1):
        ws_control.Cells(r_idx + 8, 2).Value = k
        ws_control.Cells(r_idx + 8, 3).Value = v

    info_range = ws_control.Range(ws_control.Cells(9, 2), ws_control.Cells(len(info_rows) + 8, 3))
    tb_info = ws_control.ListObjects.Add(1, info_range, None, 1)
    tb_info.Name = "tbFeatures"
    tb_info.TableStyle = "TableStyleMedium14"
    ws_control.Columns("B").ColumnWidth = 26
    ws_control.Columns("C").ColumnWidth = 85

    # --------------------------------------------------------------------------
    # 5. Import VBA Modules
    # --------------------------------------------------------------------------
    print("5. Importing VBA modules...", flush=True)
    modules = [
        "cColumn.cls",
        "cDataFrame.cls",
        "cTableStyler.cls",
        "cDataValidator.cls",
        "cReportBuilder.cls",
        "cFormulaEngine.cls",
        "mMonthlyTaskMatrix.bas"
    ]

    for mod_name in modules:
        mod_file = os.path.join(vba_dir, mod_name)
        if not os.path.exists(mod_file):
            raise FileNotFoundError(f"Missing VBA module: {mod_file}")
        wb.VBProject.VBComponents.Import(mod_file)
        print(f"   Imported {mod_name}", flush=True)

    # --------------------------------------------------------------------------
    # 6. Save Base Workbook First
    # --------------------------------------------------------------------------
    print(f"6. Saving initial workbook to {xlsm_path}...", flush=True)
    wb.SaveAs(xlsm_path, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled

    # --------------------------------------------------------------------------
    # 7. Execute Unified Single-Sheet Generation (All 12 Months Stacked)
    # --------------------------------------------------------------------------
    print("7. Running GenerateCombinedMonthlyMatrix via Excel COM...", flush=True)
    macro_name = f"'{os.path.basename(xlsm_path)}'!mMonthlyTaskMatrix.GenerateCombinedMonthlyMatrix"
    excel.Run(macro_name, True, False)
    print("   Unified 12-month TaskCompletionMatrix sheet generated successfully!", flush=True)

    # --------------------------------------------------------------------------
    # 7b. Inject Auto-Update Event Handlers into ThisWorkbook (survives sheet rebuilds)
    # --------------------------------------------------------------------------
    print("7b. Installing daily auto-refresh VBA event handlers in ThisWorkbook...", flush=True)
    ws_matrix = wb.Sheets("TaskCompletionMatrix")

    wb_code = "\r\n".join([
        "Option Explicit",
        "Private Sub Workbook_Open()",
        "    On Error Resume Next",
        "    mMonthlyTaskMatrix.AutoRefreshIfNeeded True",
        "End Sub",
        "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)",
        "    On Error Resume Next",
        "    If Sh.Name = \"RawTasks\" Or Sh.Name = \"Employees\" Then",
        "        mMonthlyTaskMatrix.MarkMatrixDirty",
        "    End If",
        "End Sub",
        "Private Sub Workbook_SheetActivate(ByVal Sh As Object)",
        "    On Error Resume Next",
        "    If Sh.Name = \"TaskCompletionMatrix\" Then",
        "        mMonthlyTaskMatrix.AutoRefreshIfNeeded False",
        "    End If",
        "End Sub"
    ])
    wb_cm = wb.VBProject.VBComponents("ThisWorkbook").CodeModule
    if wb_cm.CountOfLines > 0:
        wb_cm.DeleteLines(1, wb_cm.CountOfLines)
    wb_cm.AddFromString(wb_code)

    # Set opening view to TaskCompletionMatrix top
    ws_matrix.Activate()
    ws_matrix.Range("B2").Select()

    # Save final generated state
    print(f"8. Saving finalized workbook...", flush=True)
    wb.Save()

    print(f"9. Saving binary format {xlsb_path}...", flush=True)
    wb.SaveAs(xlsb_path, 50)  # 50 = xlExcel12

    wb.Close(False)
    print("Workbook created and closed cleanly.", flush=True)

    # --------------------------------------------------------------------------
    # 8. Copy to Standalone Repository
    # --------------------------------------------------------------------------
    if os.path.exists(sync_dir):
        print(f"10. Syncing to standalone repo: {sync_dir}...", flush=True)
        import shutil
        target_xlsm = os.path.join(sync_dir, "MonthlyTaskCompletionMatrix.xlsm")
        shutil.copy2(xlsm_path, target_xlsm)
        target_vba = os.path.join(sync_dir, "src", "vba", "mMonthlyTaskMatrix.bas")
        shutil.copy2(os.path.join(vba_dir, "mMonthlyTaskMatrix.bas"), target_vba)
        print(f"    Synced {target_xlsm}", flush=True)
        print(f"    Synced {target_vba}", flush=True)

    print("ALL STEPS COMPLETED WITH 100% SUCCESS!", flush=True)

except Exception as ex:
    print(f"ERROR: {ex}", file=sys.stderr, flush=True)
    import traceback
    traceback.print_exc()
    sys.exit(1)
finally:
    excel.Quit()
