import win32com.client
import os
import sys

def build_workbook():
    project_dir = r"C:\Users\User\Documents\Projects\vba-matrix-scheduler"
    vba_dir = os.path.join(project_dir, "src", "vba")
    xlsm_path = os.path.join(project_dir, "MatrixScheduler.xlsm")
    
    if os.path.exists(xlsm_path):
        try:
            os.remove(xlsm_path)
        except Exception:
            pass

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Add()
        
        # 1. Setup Sheets
        ws_matrix = wb.Sheets(1)
        ws_matrix.Name = "ScheduleMatrix"
        
        ws_emp = wb.Sheets.Add(After=ws_matrix)
        ws_emp.Name = "Employees"
        
        ws_events = wb.Sheets.Add(After=ws_emp)
        ws_events.Name = "Events"
        
        # --- Sheet Employees Data ---
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
                ws_emp.Cells(r_idx, c_idx).Value = val
                
        # Create Excel Table tbEmployees
        emp_range = ws_emp.Range(ws_emp.Cells(1, 1), ws_emp.Cells(len(emp_data), len(emp_data[0])))
        ws_emp.ListObjects.Add(1, emp_range, None, 1).Name = "tbEmployees"
        ws_emp.Columns("A:D").AutoFit()
        
        # --- Sheet Events Data ---
        events_data = [
            ["EventID", "Person", "Day", "Status", "Note"],
            ["EV-101", "Somchai Prasert", 1, "Work", "Office Kickoff"],
            ["EV-102", "Somchai Prasert", 5, "WFH", "Sprint Planning"],
            ["EV-103", "Somchai Prasert", 12, "Leave", "Annual Leave"],
            ["EV-104", "Somchai Prasert", 15, "OT", "Release Night"],
            ["EV-105", "Somchai Prasert", 15, "WFH", "After-hours support"], # Duplicate Day -> Tests CONCAT
            ["EV-106", "Wichai Srisuk", 1, "Work", "Core API Dev"],
            ["EV-107", "Wichai Srisuk", 2, "WFH", "Refactoring"],
            ["EV-108", "Wichai Srisuk", 8, "Leave", "Personal errand"],
            ["EV-109", "Wichai Srisuk", 15, "OT", "Release Night"],
            ["EV-110", "Anan Chaiyaphum", 3, "WFH", "UI Migration"],
            ["EV-111", "Anan Chaiyaphum", 4, "WFH", "Component testing"],
            ["EV-112", "Anan Chaiyaphum", 20, "Leave", "Sick Leave"],
            ["EV-113", "Kanya Wongsuwan", 5, "Work", "Regression Pass"],
            ["EV-114", "Kanya Wongsuwan", 15, "OT", "Release Verification"],
            ["EV-115", "Kanya Wongsuwan", 22, "WFH", "Test Automation"],
            ["EV-116", "Nattaporn Saelim", 2, "Work", "Design System"],
            ["EV-117", "Nattaporn Saelim", 6, "Leave", "Family trip"],
            ["EV-118", "Nattaporn Saelim", 18, "WFH", "Prototype Review"],
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
                
        # Create Excel Table tbEvents
        events_range = ws_events.Range(ws_events.Cells(1, 1), ws_events.Cells(len(events_data), len(events_data[0])))
        ws_events.ListObjects.Add(1, events_range, None, 1).Name = "tbEvents"
        ws_events.Columns("A:E").AutoFit()

        # --- Sheet ScheduleMatrix Layout & Buttons ---
        ws_matrix.Cells(2, 2).Value = "Team Schedule Matrix Dashboard"
        ws_matrix.Cells(2, 2).Font.Size = 16
        ws_matrix.Cells(2, 2).Font.Bold = True
        ws_matrix.Cells(2, 2).Font.Color = 0x523A1F # RGB(31, 58, 82)
        
        ws_matrix.Cells(3, 2).Value = "Dynamic In-Memory Matrix Engine powered by cDataFrame | Click buttons below to generate and map:"
        ws_matrix.Cells(3, 2).Font.Size = 10
        ws_matrix.Cells(3, 2).Font.Italic = True
        ws_matrix.Cells(3, 2).Font.Color = 0x707070
        
        # Add Action Buttons
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

        # Import VBA Modules
        vba_modules = ["cColumn.cls", "cDataFrame.cls", "cTableStyler.cls", "mMatrixScheduler.bas"]
        for mod_file in vba_modules:
            mod_path = os.path.join(vba_dir, mod_file)
            wb.VBProject.VBComponents.Import(mod_path)
            print(f"Imported: {mod_file}")

        # Run Btn_ClearMatrix and Generate Initial Matrix via code
        ws_matrix.Activate()
        
        # Save Workbook
        wb.SaveAs(xlsm_path, 52) # 52 = xlOpenXMLWorkbookMacroEnabled
        print(f"Workbook successfully saved to: {xlsm_path}")
        
        # Test Live Generation
        excel.Run("Btn_ClearMatrix")
        ws_emp_tb = ws_emp.ListObjects("tbEmployees")
        # Direct call to test silent execution
        excel.Run("GenerateMatrixFromTable", ws_emp_tb, ws_matrix.Range("B5"), 30, "Team", "Name", True)
        ws_events_tb = ws_events.ListObjects("tbEvents")
        excel.Run("MapEventsToMatrix", ws_events_tb, ws_matrix.Range("B5"), "Person", "Day", "Status", "CONCAT")
        excel.Run("ApplyStatusFormatting", ws_matrix.Range("B5"))
        print("Live matrix generation test PASSED!")
        
        wb.Save()
        wb.Close(True)
    except Exception as e:
        print(f"Build failed: {e}")
        try:
            wb.Close(False)
        except Exception:
            pass
        raise
    finally:
        excel.Quit()

if __name__ == "__main__":
    build_workbook()
