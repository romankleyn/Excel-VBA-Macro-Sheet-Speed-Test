# Excel-VBA-Macro-Sheet-Speed-Test
Find the slowest worksheets in an Excel workbook where VBA is activated, copy the code and run. It will find where the bottle neck is if related to specific excel sheet

Sheets turns off all auto calculations, gets names, activates sheets, runs calculations on them as is using Calculate Sheet(Shift + F9), times each sheet calculation time. Turns on all auto calculations, and returns results in new sheet named "sheet_speed_test_results".
