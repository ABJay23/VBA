Attribute VB_Name = "Module2"
Sub SheetLoop()

  Dim ws As Worksheet
  For Each ws In Worksheets
  ws.Activate

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "yearly change"
    ws.Range("K1").Value = "percent change"
    ws.Range("L1").Value = "total stock volume"
    
    'state variables
    Dim totalvolume As Double
    Dim startrow As Long
    Dim table2row As Integer
    
   
    
    
    
    table2row = 0
    totalvolume = 0
    startrow = 2
   
   'Loop through rows
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
 
        For i = 2 To RowCount
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            startopenprice = ws.Cells(startrow, 3).Value
            endcloseprice = ws.Cells(i, 6).Value
            yearlychange = endcloseprice - startopenprice
            percentchange = yearlychange / startopenprice
            startrow = i + 1
            
            ws.Range("I" & 2 + table2row).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + table2row).Value = yearlychange
            ws.Range("K" & 2 + table2row).Value = percentchange
            ws.Range("L" & 2 + table2row).Value = totalvolume
            totalvolume = 0
            yearlychange = 0
        
            table2row = table2row + 1
            
        If (yearlychange > 0) Then
            'fill column with green
            ws.Range("J" & 2 + table2row).Interior.ColorIndex = 4
            
        ElseIf (yearlychange <= 0) Then
            'fill column with red
            ws.Range("J" & 2 + table2row).Interior.ColorIndex = 3
            
        End If
            
            
            
      




            End If
        Next i
    Next ws
End Sub
