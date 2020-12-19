Attribute VB_Name = "Module1"
Sub alphabetTesting()

    Dim index, LastRow As Long
    Dim counter As Integer
    
    counter = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Cells(1, 9) = "Ticker"

    For index = 3 To LastRow
    
        If (Cells(index, 1).Value <> Cells(index - 1, 1)) Then
            Cells(counter, 9).Value = Cells(index - 1, 1)
            counter = counter + 1
        End If
        
    Next index
    
End Sub

Sub stockDiff()
    Dim openStock As Double
    Dim closeStock As Double
    Dim index, LastRow As Integer
    Dim priceChange As Double
    
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    
    LastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    For index = 2 To LastRow
        openStock = WorksheetFunction.VLookup(Cells(index, 9), Range("A:C"), 3, 0)
        closeStock = WorksheetFunction.VLookup(Cells(index, 9), Range("A:F"), 6)
        
        priceChange = closeStock - openStock
        Cells(index, 10) = priceChange
        
        If (openStock = 0) Then
            Cells(index, 11) = 0
        ElseIf (openStock > 0) Then
            Cells(index, 11) = priceChange / openStock
        End If
    Next index
    
    For index = 2 To LastRow
        If (Cells(index, 10) > 0) Then
                Cells(index, 10).Interior.ColorIndex = 4
            ElseIf (Cells(index, 10) < 0) Then
                Cells(index, 10).Interior.ColorIndex = 3
        End If
    Next index
    
End Sub

Sub totalStock()
    Dim tStock, index, LastRow As Long
    Dim counter As Integer
    
    Cells(1, 12) = "Total Stock Volume"
    
    counter = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For index = 2 To LastRow
        tStock = tStock + Cells(index, 7)
        
        If (Cells(index, 1) <> Cells(index + 1, 1)) Then
            Cells(counter, 12) = tStock
            counter = counter + 1
            tStock = 0
        End If
        
    Next index
    
End Sub

Sub bonus()

    Cells(2, 14).Value = "Greatest Percent Increase"
    Cells(3, 14).Value = "Greatest Percent Decrease"
    Cells(4, 14).Value = "Greatest Total Stock"
    
    
    Cells(2, 15).Value = WorksheetFunction.Max(Range("K:K"))
    Cells(3, 15).Value = WorksheetFunction.Min(Range("K:K"))
    Cells(4, 15).Value = WorksheetFunction.Max(Range("L:L"))

    
End Sub




