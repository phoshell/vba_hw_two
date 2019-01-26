Attribute VB_Name = "Module1"
Sub ThisAssignmentWasHard()

     ' Declare variables
    Dim ticker As String
    Dim total_volume_stock As Double
    Dim worksheet_summary As Integer
    Dim i As Long
    Dim LastRow As Long
    Dim SheetYear As Worksheet
    
    ' Loop though all sheets
    For Each SheetYear In Worksheets
    SheetYear.Activate
    
    ' identify space for the answers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"
    
    ' Sets values
    total_volume_stock = 0
    worksheet_summary = 2
    
    ' Find last row to loop
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' For Loop starts here
    For i = 2 To LastRow
        
            ' grouping tickers
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = Cells(i, 1).Value
                
                ' getting the sums for total stock volume
                total_volume_stock = total_volume_stock + Cells(i, 7).Value
                
                ' showing the total stock and ticker output
                Range("I" & worksheet_summary).Value = ticker
                
                Range("J" & worksheet_summary).Value = total_volume_stock
        
                ' add 1 to summary and reset total stock volume
                worksheet_summary = worksheet_summary + 1
                total_volume_stock = 0
                
                ' if not true
                Else
                total_volume_stock = total_volume_stock + Cells(i, 7).Value
                
                End If
        Next i
    Next SheetYear
End Sub

