import win32com.client, os, sys, zipfile

project_dir = r"C:\Users\User\Documents\Projects\vba-matrix-scheduler"
vba_dir = os.path.join(project_dir, "src", "vba")
xlsm_path = os.path.join(project_dir, "MatrixScheduler.xlsm")
xlsb_path = os.path.join(project_dir, "MatrixScheduler.xlsb")
zip_path = os.path.join(project_dir, "MatrixScheduler-Package.zip")

for p in [xlsm_path, xlsb_path, zip_path]:
    if os.path.exists(p):
        try:
            os.remove(p)
        except Exception:
            pass

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    print("1. Creating Workbook...", flush=True)
    wb = excel.Workbooks.Add()
    
    ws_matrix = wb.Sheets(1)
    ws_matrix.Name = "ScheduleMatrix"
    ws_emp = wb.Sheets.Add(After=ws_matrix)
    ws_emp.Name = "Employees"
    ws_events = wb.Sheets.Add(After=ws_emp)
    ws_events.Name = "Events"
    
    emp_data = [
        ["EmployeeID", "Name", "Team", "Role"],
        ["EMP-001", "Somchai Prasert", "Engineering", "Tech Lead"],
        ["EMP-002", "Wichai Srisuk", "Engineering", "Backend Dev"],
        ["EMP-003", "Anan Chaiyaphum", "Engineering", "Frontend Dev"],
        ["EMP-004", "Kanya Wongsuwan", "Engineering", "QA Engineer"],
        ["EMP-005", "Nattaporn Saelim", "Design", "Product Designer"],
        ["EMP-006", "Piti Rattanaporn", "Design", "UI/UX Designer"],
        ["EMP-007", "Siriporn Meesuk", "Design", "Graphic Designer"],
        ["EMP-008", "Chaiwat Kongkaew", "Product", "Product Manager"],
        ["EMP-009", "Thanya Somboon", "Product", "Scrum Master"],
        ["EMP-010", "Worawat Boonmee", "Product", "Data Analyst"],
        ["EMP-011", "Maneerat Thongdee", "Operations", "Ops Manager"],
        ["EMP-012", "Narongrit Kaewdee", "Operations", "Support Specialist"]
    ]
    for r_idx, row in enumerate(emp_data, 1):
        for c_idx, val in enumerate(row, 1):
            ws_emp.Cells(r_idx + 3, c_idx + 1).Value = val
    emp_range = ws_emp.Range(ws_emp.Cells(4, 2), ws_emp.Cells(len(emp_data) + 3, len(emp_data[0]) + 1))
    ws_emp.ListObjects.Add(1, emp_range, None, 1).Name = "tbEmployees"
    ws_emp.Columns("B:E").AutoFit()

    events_data = [
        ["EventID", "Person", "Day", "Status", "Note"],
        ["EV-101", "Somchai Prasert", 1, "Work", "Sprint Planning"],
        ["EV-102", "Somchai Prasert", 8, "WFH", "Remote Focus Day"],
        ["EV-103", "Somchai Prasert", 15, "OT", "Release Deployment"],
        ["EV-104", "Somchai Prasert", 22, "Leave", "Annual Leave"],
        ["EV-105", "Wichai Srisuk", 2, "Work", "API Refactor"],
        ["EV-106", "Wichai Srisuk", 5, "OFF", "Scheduled Rest Day"],
        ["EV-107", "Wichai Srisuk", 12, "WFH", "Bug Fixing"],
        ["EV-108", "Wichai Srisuk", 18, "OT", "DB Migration"],
        ["EV-109", "Anan Chaiyaphum", 3, "Work", "UI Redesign"],
        ["EV-110", "Anan Chaiyaphum", 10, "WFH", "Component Library"],
        ["EV-111", "Anan Chaiyaphum", 17, "Leave", "Sick Leave"],
        ["EV-112", "Kanya Wongsuwan", 4, "Work", "Regression Test"],
        ["EV-113", "Kanya Wongsuwan", 6, "OFF", "Rest Day"],
        ["EV-114", "Kanya Wongsuwan", 11, "OT", "Test Automation"],
        ["EV-115", "Kanya Wongsuwan", 24, "WFH", "Test Planning"],
        ["EV-116", "Nattaporn Saelim", 1, "Work", "Design Sprint"],
        ["EV-117", "Nattaporn Saelim", 9, "Leave", "Personal Leave"],
        ["EV-118", "Nattaporn Saelim", 16, "WFH", "Wireframes"],
        ["EV-119", "Piti Rattanaporn", 7, "WFH", "Figma Components"],
        ["EV-120", "Piti Rattanaporn", 14, "Work", "User Interview"],
        ["EV-121", "Siriporn Meesuk", 10, "Leave", "Personal Leave"],
        ["EV-122", "Siriporn Meesuk", 11, "WFH", "Asset Export"],
        ["EV-123", "Chaiwat Kongkaew", 1, "Work", "Roadmap Review"],
        ["EV-124", "Chaiwat Kongkaew", 9, "WFH", "Strategy Deck"],
        ["EV-125", "Chaiwat Kongkaew", 25, "Leave", "Vacation"],
        ["EV-126", "Thanya Somboon", 5, "Work", "Sprint Retrospective"],
        ["EV-127", "Thanya Somboon", 19, "WFH", "Backlog Grooming"],
        ["EV-128", "Worawat Boonmee", 8, "Work", "ETL Validation"],
        ["EV-129", "Worawat Boonmee", 16, "WFH", "Report Generation"],
        ["EV-130", "Maneerat Thongdee", 1, "Work", "Ops Sync"],
        ["EV-131", "Maneerat Thongdee", 13, "Leave", "Personal Leave"],
        ["EV-132", "Narongrit Kaewdee", 15, "OT", "Production Support"],
        ["EV-133", "Narongrit Kaewdee", 28, "WFH", "Weekend Standby"]
    ]
    for r_idx, row in enumerate(events_data, 1):
        for c_idx, val in enumerate(row, 1):
            ws_events.Cells(r_idx, c_idx).Value = val
    events_range = ws_events.Range(ws_events.Cells(1, 1), ws_events.Cells(len(events_data), len(events_data[0])))
    ws_events.ListObjects.Add(1, events_range, None, 1).Name = "tbEvents"
    ws_events.Columns("A:E").AutoFit()

    ws_matrix.Cells(2, 2).Value = "Team Schedule Matrix Dashboard"
    ws_matrix.Cells(2, 2).Font.Size = 16
    ws_matrix.Cells(2, 2).Font.Bold = True
    ws_matrix.Cells(2, 2).Font.Color = 0x523A1F

    ws_matrix.Cells(3, 2).Value = "Dynamic In-Memory Matrix Engine powered by cDataFrame | Click buttons below to generate and map:"
    ws_matrix.Cells(3, 2).Font.Size = 10
    ws_matrix.Cells(3, 2).Font.Italic = True
    ws_matrix.Cells(3, 2).Font.Color = 0x707070

    btn1 = ws_matrix.Buttons().Add(120, 60, 160, 32)
    btn1.Text = "1. Generate Matrix"
    btn1.OnAction = "Btn_GenerateMatrix"
    
    btn2 = ws_matrix.Buttons().Add(290, 60, 160, 32)
    btn2.Text = "2. Map Schedule Data"
    btn2.OnAction = "Btn_MapEvents"
    
    btn3 = ws_matrix.Buttons().Add(460, 60, 140, 32)
    btn3.Text = "Run All (1 + 2)"
    btn3.OnAction = "Btn_RunAll"

    btn4 = ws_matrix.Buttons().Add(610, 60, 110, 32)
    btn4.Text = "Clear Matrix"
    btn4.OnAction = "Btn_ClearMatrix"

    print("2. Importing VBA modules...", flush=True)
    for mod_file in ["cColumn.cls", "cDataFrame.cls", "cTableStyler.cls", "mMatrixScheduler.bas"]:
        mod_path = os.path.join(vba_dir, mod_file)
        wb.VBProject.VBComponents.Import(mod_path)
        print(f"   Imported {mod_file}", flush=True)

    ws_matrix.Activate()
    
    print("3. Saving MatrixScheduler.xlsm...", flush=True)
    wb.SaveAs(xlsm_path, 52)
    print("4. Saving MatrixScheduler.xlsb (Binary format)...", flush=True)
    wb.SaveAs(xlsb_path, 50)
    wb.Close(False)
    
    print("5. Packaging ZIP...", flush=True)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.write(xlsm_path, 'MatrixScheduler.xlsm')
        z.write(xlsb_path, 'MatrixScheduler.xlsb')
        z.write(os.path.join(project_dir, 'README.md'), 'README.md')
    print("SUCCESS! All artifacts packaged cleanly.", flush=True)
finally:
    excel.Quit()
