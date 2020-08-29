# VBA-Challenge
Repo for Akriti Suri's VBA Homework 8.29.20

Sub GetTicker()

Dim WS_Count As Integer
Dim I As Integer

WS_Count = ActiveWorkbook.Worksheets.Count
For I = 1 To WS_Count


 Dim Sheet1 As Worksheet
 worksheetCount = ThisWorkbook.Worksheets.Count
 
 
 Dim Ticker As String
 Ticker = ThisWorkbook.Worksheets(1).Cells(2, 1).Value
 
Next I

 'Find the column count and start the output on the following one
 Dim ColumnCount As Integer
 Set Sheet1 = ThisWorkbook.Worksheets(1)
 
 
 ColumnCount = Sheet1.UsedRange.Columns.Count
 
 
 Dim RowCount As Long
 RowCount = Sheet1.UsedRange.Rows.Count


 'Create 4 new columns for each output in the same worksheet
 Sheet1.Cells(1, ColumnCount + 1) = "Ticker Symbol"
 Sheet1.Cells(1, ColumnCount + 2) = "Yearly Change"
 Sheet1.Cells(1, ColumnCount + 3) = "Percent Change"
 Sheet1.Cells(1, ColumnCount + 4) = "Total Stock Volume"
 
 'Start Loop and repeat following steps for amount of rows in the worksheet
    
    For Counter = 2 To RowCount
        'Read the value from ticker input column and setting it into output column
            Sheet1.Cells(Counter, ColumnCount + 1) = Sheet1.Cells(Counter, 1)
            
        'Read the value of open price column and closing price column and subtract the 2 values and set it to the 2nd output column
            Sheet1.Cells(Counter, ColumnCount + 2) = Sheet1.Cells(Counter, 6) - Sheet1.Cells(Counter, 3)
            
        'Read the value of 2nd output column and divide by open price column and multiply by 100 and set it into the 3rd output column
            Sheet1.Cells(Counter, ColumnCount + 3) = (Sheet1.Cells(Counter, ColumnCount + 2) / Sheet1.Cells(Counter, 3)) * 100
            
        'Read the value from volume column and set it into the 4th output column
            Sheet1.Cells(Counter, ColumnCount + 4) = Sheet1.Cells(Counter, 7)
            
        'If the value in Yearly Change Column is <0 then color that cell green and if it is >0 then color that cell red
  
            If Sheet1.Cells(Counter, ColumnCount + 2) > 0 Then
                Sheet1.Cells(Counter, ColumnCount + 2).Interior.ColorIndex = 4
            End If
        
            If Sheet1.Cells(Counter, ColumnCount + 2) < 0 Then
                Sheet1.Cells(Counter, ColumnCount + 2).Interior.ColorIndex = 3
            End If
        
        'Create 3 new columns for each output in the same worksheet
            Sheet1.Cells(2, ColumnCount + 6) = "Greatest % Increase"
            Sheet1.Cells(3, ColumnCount + 6) = "Greatest % Decrease"
            Sheet1.Cells(4, ColumnCount + 6) = "Greatest Total Volume"
            Sheet1.Cells(1, ColumnCount + 7) = "Ticker ID"
            Sheet1.Cells(1, ColumnCount + 8) = "Value"
        
    Next Counter
    
    'Read the highest value from the Percent Change column and set it into the 7th output column
    Dim rng As Range
    
    Set rng = Sheet1.Cells(Counter, ColumnCount + 3)
    Sheet1.Cells(2, ColumnCount + 8) = WorksheetFunction.Max(rng)
    Sheet1.Cells(3, ColumnCount + 8) = WorksheetFunction.Min(rng)
    
    Dim rng2 As Range
    Set rng2 = Sheet1.Cells(Counter, ColumnCount + 4)
    Sheet1.Cells(4, ColumnCount + 8) = WorksheetFunction.Max(rng2)
    

    
End Sub

