VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub RunMacroOnAllSheets()
    
    ' Variable for worksheet
    Dim ws As Worksheet
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Sheets
    
        ' Set the active sheet
        ws.Activate
        
        ' Run the stock performance calculation macro for the current sheet
        CalculationsWorksheet
        
    Next ws
    
End Sub
Sub CalculationsWorksheet()

    ' Variables for datatypes
    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    Dim percentageChange As Double
    Dim i As Long
    
    ' Variables for greatest values
    Dim greatestIncrease As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecrease As Double
    Dim greatestDecreaseTicker As String
    Dim greatestVolume As Double
    Dim greatestVolumeTicker As String
    
    ' Initialize greatest values
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in the data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    ' Add headers for the summary table
    ws.Cells(1, 9).Value = "Ticker"                 ' Column I
    ws.Cells(1, 10).Value = "Quarterly Change"      ' Column J
    ws.Cells(1, 11).Value = "Percentage Change"     ' Column K
    ws.Cells(1, 12).Value = "Total Stock Volume"    ' Column L
    
    summaryRow = 2 ' Start row for the summary table
    startRow = 2 ' Start of the first ticker block
    
    ' Loop through each row of data
    For i = 2 To lastRow
    
        ' Check if the ticker changes (or it's the last row)
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow Then
        
            ' Capture the ticker
            ticker = ws.Cells(startRow, 1).Value
            
            ' Determine start and end rows for the ticker
            endRow = i
            
            ' Calculate opening price, closing price, and total volume
            openingPrice = ws.Cells(startRow, 3).Value
            closingPrice = ws.Cells(endRow, 4).Value
            totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))
            
            ' Calculate percentage change
            percentageChange = 0
            If openingPrice <> 0 Then
                percentageChange = ((closingPrice - openingPrice) / openingPrice)
            End If
            
            ' Output results to the summary table
            ws.Cells(summaryRow, 9).Value = ticker
            ws.Cells(summaryRow, 10).Value = closingPrice - openingPrice
            ws.Cells(summaryRow, 11).Value = percentageChange
            ws.Cells(summaryRow, 12).Value = totalVolume
            
            ' Apply conditional formatting to Column J (Quarterly Change)
            Dim cell As Range
            Set cell = ws.Cells(summaryRow, 10) ' Column J
                If cell.Value > 0 Then
                cell.Interior.ColorIndex = 4 ' Light green
                ElseIf cell.Value < 0 Then
                cell.Interior.ColorIndex = 3 ' Light red
                Else
                cell.Interior.ColorIndex = xlNone ' No fill
                End If
            
            ' Track greatest percentage increase
            If percentageChange > greatestIncrease Then
                greatestIncrease = percentageChange
                greatestIncreaseTicker = ticker
            End If
            ' Track greatest percentage decrease
            If percentageChange < greatestDecrease Then
                greatestDecrease = percentageChange
                greatestDecreaseTicker = ticker
            End If
            ' Track greatest total volume
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
            ' Move to the next summary row
            summaryRow = summaryRow + 1
            ' Reset start row for the next ticker
            startRow = i + 1
        End If
        
    Next i
    
    ' Output the greatest values below the summary table
    ws.Cells(1, 14).Value = "Greatest % Increase"
    ws.Cells(1, 15).Value = greatestIncreaseTicker
    ws.Cells(1, 16).Value = greatestIncrease
    ws.Cells(2, 14).Value = "Greatest % Decrease"
    ws.Cells(2, 15).Value = greatestDecreaseTicker
    ws.Cells(2, 16).Value = greatestDecrease
    ws.Cells(3, 14).Value = "Greatest Total Volume"
    ws.Cells(3, 15).Value = greatestVolumeTicker
    ws.Cells(3, 16).Value = greatestVolume
    
End Sub
