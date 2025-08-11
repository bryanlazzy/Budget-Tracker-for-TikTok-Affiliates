Attribute VB_Name = "Module1"
Option Explicit

Const COLOR_PROFIT As Long = 5296274 ' Green
Const COLOR_LOSS As Long = 255 ' Red
Const COLOR_NEUTRAL As Long = 16777215 ' White


Sub ApplyColorCoding(ws As Worksheet, startRow As Long, endRow As Long, profitColumn As String)
    Dim i As Long
    Dim profitValue As Double
    
    For i = startRow To endRow
        If ws.Cells(i, 1).Value <> "" Then
            profitValue = ws.Range(profitColumn & i).Value
            
            If profitValue > 0 Then
                
                ws.Range(profitColumn & i).Interior.Color = COLOR_PROFIT
                ws.Range(profitColumn & i).Font.Color = RGB(0, 0, 0)
            ElseIf profitValue < 0 Then
                
                ws.Range(profitColumn & i).Interior.Color = COLOR_LOSS
                ws.Range(profitColumn & i).Font.Color = RGB(255, 255, 255)
            Else
                
                ws.Range(profitColumn & i).Interior.Color = COLOR_NEUTRAL
                ws.Range(profitColumn & i).Font.Color = RGB(0, 0, 0)
            End If
        End If
    Next i
End Sub


Sub CreateMacroButtons()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnRange As Range
    
    
    Set ws = ThisWorkbook.Sheets("Daily Tracking")
    ws.Select
    
    
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    
    Set btnRange = ws.Range("J3:L4")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    With btn
        .Caption = "Add Daily Entry"
        .OnAction = "AddDailyEntry"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    
    Set btnRange = ws.Range("J6:L7")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    With btn
        .Caption = "Calculate Monthly"
        .OnAction = "CalculateMonthlyTotals"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    
    Set btnRange = ws.Range("J9:L10")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    With btn
        .Caption = "Refresh Colors"
        .OnAction = "RefreshColorCoding"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    
    Set btnRange = ws.Range("J12:L13")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    With btn
        .Caption = "Monthly Summary"
        .OnAction = "GoToMonthlySheet"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    
    Set btnRange = ws.Range("J15:L16")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    With btn
        .Caption = "Yearly Summary"
        .OnAction = "GoToYearlySheet"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    
    Set ws = ThisWorkbook.Sheets("Monthly Summary")
    
    
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    
    Set btnRange = ws.Range("I3:K4")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    With btn
        .Caption = "Back to Daily"
        .OnAction = "GoToDailySheet"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    
    Set btnRange = ws.Range("I6:K7")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    With btn
        .Caption = "Yearly Summary"
        .OnAction = "GoToYearlySheet"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    
    Set ws = ThisWorkbook.Sheets("Yearly Summary")
    
    
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    
    Set btnRange = ws.Range("J3:L4")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    With btn
        .Caption = "Back to Daily"
        .OnAction = "GoToDailySheet"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    
    Set btnRange = ws.Range("J6:L7")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    With btn
        .Caption = "Monthly Summary"
        .OnAction = "GoToMonthlySheet"
        .Font.Bold = True
        .Font.Size = 10
    End With
End Sub


Sub GoToDailySheet()
    ThisWorkbook.Sheets("Daily Tracking").Select
End Sub

Sub GoToMonthlySheet()
    ThisWorkbook.Sheets("Monthly Summary").Select
End Sub

Sub GoToYearlySheet()
    ThisWorkbook.Sheets("Yearly Summary").Select
End Sub


Sub SetupNavigation()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Daily Tracking")
    
    ws.Select
    ws.Range("A2").Value = "Ready to track your TikTok affiliate earnings!"
    ws.Range("A2").Font.Italic = True
    
    
    Call CreateMacroButtons
End Sub


Sub CreateDailyTracking()
    Dim ws As Worksheet
    
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Daily Tracking")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Daily Tracking"
    End If
    
    ws.Select
    
    
    With ws
        .Range("A1").Value = "TikTok Affiliate Daily Budget Tracker"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").HorizontalAlignment = xlCenter
        
        
        .Range("A3").Value = "Date"
        .Range("B3").Value = "Affiliate Income"
        .Range("C3").Value = "Other Income"
        .Range("D3").Value = "Total Income"
        .Range("E3").Value = "Expenses"
        .Range("F3").Value = "Net Profit/Loss"
        .Range("G3").Value = "Status"
        .Range("H3").Value = "Notes"
        
        
        .Range("A3:H3").Font.Bold = True
        .Range("A3:H3").Interior.Color = RGB(200, 200, 200)
        .Range("A3:H3").Borders.LineStyle = xlContinuous
        
        
        .Columns("A").ColumnWidth = 12
        .Columns("B:G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 20
        
        
        .Range("D4").Formula = "=B4+C4"
        .Range("F4").Formula = "=D4-E4"
        .Range("G4").Formula = "=IF(F4>0,""Profit"",IF(F4<0,""Loss"",""Break Even""))"
    End With
End Sub


Sub CreateMonthlySummary()
    Dim ws As Worksheet
    
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Monthly Summary")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Monthly Summary"
    End If
    
    ws.Select
    
    With ws
        .Range("A1").Value = "TikTok Affiliate Monthly Summary"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1:G1").Merge
        .Range("A1").HorizontalAlignment = xlCenter
        
        
        .Range("A3").Value = "Month/Year"
        .Range("B3").Value = "Total Income"
        .Range("C3").Value = "Total Expenses"
        .Range("D3").Value = "Net Profit/Loss"
        .Range("E3").Value = "Status"
        .Range("F3").Value = "Profit Margin %"
        .Range("G3").Value = "Days Active"
        
        
        .Range("A3:G3").Font.Bold = True
        .Range("A3:G3").Interior.Color = RGB(200, 200, 200)
        .Range("A3:G3").Borders.LineStyle = xlContinuous
        
        
        .Columns("A:G").ColumnWidth = 15
        
        
        .Range("D4").Formula = "=B4-C4"
        .Range("E4").Formula = "=IF(D4>0,""Profit"",IF(D4<0,""Loss"",""Break Even""))"
        .Range("F4").Formula = "=IF(B4<>0,ROUND(D4/B4*100,2),0)"
    End With
End Sub


Sub CreateYearlySummary()
    Dim ws As Worksheet
    
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Yearly Summary")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Yearly Summary"
    End If
    
    ws.Select
    
    With ws
        .Range("A1").Value = "TikTok Affiliate Yearly Summary"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1:H1").Merge
        .Range("A1").HorizontalAlignment = xlCenter
        
        
        .Range("A3").Value = "Year"
        .Range("B3").Value = "Total Income"
        .Range("C3").Value = "Total Expenses"
        .Range("D3").Value = "Net Profit/Loss"
        .Range("E3").Value = "Status"
        .Range("F3").Value = "Profit Margin %"
        .Range("G3").Value = "Avg Monthly Earnings"
        .Range("H3").Value = "Growth Rate %"
        
        
        .Range("A3:H3").Font.Bold = True
        .Range("A3:H3").Interior.Color = RGB(200, 200, 200)
        .Range("A3:H3").Borders.LineStyle = xlContinuous
        
        
        .Columns("A:H").ColumnWidth = 15
        
        
        .Range("D4").Formula = "=B4-C4"
        .Range("E4").Formula = "=IF(D4>0,""Profit"",IF(D4<0,""Loss"",""Break Even""))"
        .Range("F4").Formula = "=IF(B4<>0,ROUND(D4/B4*100,2),0)"
        .Range("G4").Formula = "=IF(B4<>0,ROUND(B4/12,2),0)"
    End With
End Sub


Sub InitializeBudgetTracker()
    Application.ScreenUpdating = False
    
    
    Cells.Clear
    
    Call CreateDailyTracking
    
    Call CreateMonthlySummary
    
    Call CreateYearlySummary
    
    Call SetupNavigation
    
    Application.ScreenUpdating = True
    
    MsgBox "TikTok Affiliate Budget Tracker initialized successfully!", vbInformation, "Setup Complete"
End Sub


Sub AddDailyEntry()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim newDate As Date
    Dim affiliateIncome As Double
    Dim otherIncome As Double
    Dim expenses As Double
    Dim notes As String
    
    Set ws = ThisWorkbook.Sheets("Daily Tracking")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    
    newDate = InputBox("Enter date (MM/DD/YYYY):", "Date Entry", Date)
    affiliateIncome = Val(InputBox("Enter affiliate income:", "Affiliate Income", 0))
    otherIncome = Val(InputBox("Enter other income:", "Other Income", 0))
    expenses = Val(InputBox("Enter expenses:", "Expenses", 0))
    notes = InputBox("Enter notes (optional):", "Notes", "")
    
    
    With ws
        .Cells(lastRow, 1).Value = newDate
        .Cells(lastRow, 2).Value = affiliateIncome
        .Cells(lastRow, 3).Value = otherIncome
        .Cells(lastRow, 4).Formula = "=B" & lastRow & "+C" & lastRow
        .Cells(lastRow, 5).Value = expenses
        .Cells(lastRow, 6).Formula = "=D" & lastRow & "-E" & lastRow
        .Cells(lastRow, 7).Formula = "=IF(F" & lastRow & ">0,""Profit"",IF(F" & lastRow & "<0,""Loss"",""Break Even""))"
        .Cells(lastRow, 8).Value = notes
    End With
    
    
    Call ApplyColorCoding(ws, lastRow, lastRow, "F")
    
    MsgBox "Daily entry added successfully!", vbInformation, "Entry Added"
End Sub


Sub RefreshColorCoding()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    On Error Resume Next
    
    
    Set ws = ThisWorkbook.Sheets("Daily Tracking")
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 4 Then
            Call ApplyColorCoding(ws, 4, lastRow, "F")
        End If
    End If
    
    
    Set ws = ThisWorkbook.Sheets("Monthly Summary")
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 4 Then
            Call ApplyColorCoding(ws, 4, lastRow, "D")
        End If
    End If
    
    
    Set ws = ThisWorkbook.Sheets("Yearly Summary")
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 4 Then
            Call ApplyColorCoding(ws, 4, lastRow, "D")
        End If
    End If
    
    On Error GoTo 0
    
    MsgBox "Color coding refreshed for all sheets!", vbInformation, "Refresh Complete"
End Sub


Sub CalculateMonthlyTotals()
    Dim dailyWs As Worksheet
    Dim monthlyWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentMonth As String
    Dim monthlyIncome As Double
    Dim monthlyExpenses As Double
    Dim monthlyRow As Long
    
    Set dailyWs = ThisWorkbook.Sheets("Daily Tracking")
    Set monthlyWs = ThisWorkbook.Sheets("Monthly Summary")
    
    lastRow = dailyWs.Cells(dailyWs.Rows.Count, 1).End(xlUp).Row
    monthlyRow = 4
    
    
    monthlyWs.Range("A4:G100").Clear
    
    If lastRow < 4 Then
        MsgBox "No daily data found to calculate monthly totals.", vbInformation, "No Data"
        Exit Sub
    End If
    
    For i = 4 To lastRow
        If dailyWs.Cells(i, 1).Value <> "" Then
            currentMonth = Format(dailyWs.Cells(i, 1).Value, "mmm yyyy")
            monthlyIncome = monthlyIncome + dailyWs.Cells(i, 4).Value
            monthlyExpenses = monthlyExpenses + dailyWs.Cells(i, 5).Value
            
            
            If i = lastRow Or Format(dailyWs.Cells(i + 1, 1).Value, "mmm yyyy") <> currentMonth Then
                
                With monthlyWs
                    .Cells(monthlyRow, 1).Value = currentMonth
                    .Cells(monthlyRow, 2).Value = monthlyIncome
                    .Cells(monthlyRow, 3).Value = monthlyExpenses
                    .Cells(monthlyRow, 4).Formula = "=B" & monthlyRow & "-C" & monthlyRow
                    .Cells(monthlyRow, 5).Formula = "=IF(D" & monthlyRow & ">0,""Profit"",IF(D" & monthlyRow & "<0,""Loss"",""Break Even""))"
                    .Cells(monthlyRow, 6).Formula = "=IF(B" & monthlyRow & "<>0,ROUND(D" & monthlyRow & "/B" & monthlyRow & "*100,2),0)"
                End With
                
                monthlyRow = monthlyRow + 1
                monthlyIncome = 0
                monthlyExpenses = 0
            End If
        End If
    Next i
    
    
    If monthlyRow > 4 Then
        Call ApplyColorCoding(monthlyWs, 4, monthlyRow - 1, "D")
    End If
    
    MsgBox "Monthly totals calculated successfully!", vbInformation, "Calculation Complete"
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    
    If Not Intersect(Target, Range("B:F")) Is Nothing Then
        Call RefreshColorCoding
    End If
End Sub
