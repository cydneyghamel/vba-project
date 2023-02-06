Attribute VB_Name = "Module1"
Sub StockSummary():

    ' Assumptions:
    ' Each sheet has unique set of tickers
    ' Dates are in ascending order for each ticker
    
    ' Define variables
    Dim startRow As Integer
    Dim endRow As Double
    
    ' Loop through sheets
    For Each ws In ThisWorkbook.Worksheets
        startRow = 2
        endRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
 
        ' Set column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Declare variables
        Dim summaryRow As Integer
        Dim r As Double
        Dim ticker As String
        Dim xopen As Double
        Dim xclose As Double
        Dim vol As Double
        Dim gi As Double
        Dim giTick As String
        Dim gd As Double
        Dim gdTick As String
        Dim gVol As Double
        Dim gVolTick As String
        
        ' Assign row for summary output
        summaryRow = 2
        
        ' Loop through row, where r = current row #
        For r = startRow To endRow
        
            ' If the value in the current row's 1st colum does not equal the saved ticker value...
            If (ws.Cells(r, 1).Value <> ticker) Then
                ticker = ws.Cells(r, 1).Value
                xopen = ws.Cells(r, 3).Value
                vol = ws.Cells(r, 7).Value
            End If
            
            ' If the value in the next row's 1st column does not equal the saved ticker value...
            If (ws.Cells(r + 1, 1).Value <> ticker) Then
                Dim change As Double
                Dim pChange As Double
                
                ' Update the variables
                xclose = ws.Cells(r, 6).Value
                vol = vol + ws.Cells(r, 7).Value
                change = xclose - xopen
                pChange = change / xopen
                
                ' Insert data into spreadsheet/set value of cells
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = change
                ws.Cells(summaryRow, 11).Value = pChange
                ws.Cells(summaryRow, 12).Value = vol
                
                ' Format colors for yearly change
                If (change < 0) Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                ElseIf (change > 0) Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                End If
                
                ' If pChange > Greatest % Increase, set values
                If (pChange > gi) Then
                    gi = pChange
                    giTick = ticker
                End If
                
                ' If pChange < Greatest % Decrease, set values
                If (pChange < gd) Then
                    gd = pChange
                    gdTick = ticker
                End If
                
                ' If pChange > gVol, set values
                If (pChange > gVol) Then
                    gVol = pChange
                    gVolTick = ticker
                End If
                
                ' Updates summaryRow
                summaryRow = summaryRow + 1
            
            Else
                ' Update vol
                vol = vol + ws.Cells(r, 7).Value
            End If
            
        Next r
         
         ' Set headers for additional outputs (Greatest % Increase, Greatest % Decrease, Greatest Total Volume
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Designate locations for outputs
        ws.Cells(2, 16).Value = giTick
        ws.Cells(2, 17).Value = gi
        ws.Cells(3, 16).Value = gdTick
        ws.Cells(3, 17).Value = gd
        ws.Cells(4, 16).Value = gVolTick
        ws.Cells(4, 17).Value = gVol 'Q
        
        ' Format columns to percentages
        ws.Columns(11).NumberFormat = "0.00%"
        ws.Columns(17).NumberFormat = "0.00%"
        
    Next ws

    
End Sub

        
