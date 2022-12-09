Attribute VB_Name = "test_speed"
Sub sheet_speed_test()
    Dim ws_n As Integer: ws_n = Worksheets.Count
    Dim v_array As Variant
    ReDim v_array(1 To ws_n, 1 To 2)
    Dim irow As Integer: irow = 1
    
    'opton
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    
    'timesheets
    For Each ws In Worksheets
        v_array(irow, 1) = ws.Name
        
        ws.Activate
        t0 = Now
        ActiveSheet.Calculate
        v_array(irow, 2) = Now - t0
        irow = irow + 1
    Next ws
    
    'optoff
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    
    'return results
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("sheet_speed_test_results").Delete
    Application.DisplayAlerts = True
    Sheets.Add(Before:=Sheets(1)).Name = "sheet_speed_test_results"
    
    Range("A1") = "WorksheeetName"
    Range("B1") = "CalculationTime"
    Range("A2:B" & ws_n + 1) = v_array
    Cells.EntireColumn.AutoFit
End Sub
