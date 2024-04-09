Attribute VB_Name = "AnalysisStocks"
Sub StockValues()
    For Each ws In Worksheets
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'The Dims
        Dim tickername As String
        Dim tickervolume As Double
        Dim summaryticker As Integer
        Dim openprice As Double
        Dim closeprice As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        
 'Set Values
 tickervolume = 0
summaryticker = 2
    openprice = ws.Cells(2, 3).Value
        
  'Label Columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
     ws.Cells(1, 11).Value = "Percent Change"
     ws.Cells(1, 12).Value = "Total Stock Volume"

'starting loop
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'Set the ticker names
          tickername = ws.Cells(i, 1).Value
          tickervolume = tickervolume + ws.Cells(i, 7).Value
          ws.Range("I" & summaryticker).Value = tickername
          ws.Range("L" & summaryticker).Value = tickervolume

 'Closing Price
    closeprice = ws.Cells(i, 6).Value

'Yearly Change
    yearlychange = (closeprice - openprice)
    ws.Range("J" & summaryticker).Value = yearlychange

'Percentage Change
    If openprice = 0 Then
    percentchange = 0
                
Else
    percentchange = yearlychange / openprice
                
 End If
 
 'Yearly Change Per Ticker
    ws.Range("K" & summaryticker).Value = percentchange
    ws.Range("K" & summaryticker).NumberFormat = "0.00%"
'Row Count
    summaryticker = summaryticker + 1
 'Volume
    tickervolume = 0
'Opening Price
openprice = ws.Cells(i + 1, 3)

Else
              
'Add Vol of Trade
    tickervolume = tickervolume + ws.Cells(i, 7).Value
End If
 
 Next i

'Conditional formatting Color for Yearly Change
    lastrowsummarytable = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
'Color code yearly change
For i = 2 To lastrowsummarytable
    If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 43
 
 Else
    ws.Cells(i, 10).Interior.ColorIndex = 3
            
End If
 
 Next i

'Labels for Return Columns
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
            
'Percent Change and Total Value

For i = 2 To lastrowsummarytable
'Maximum percent change

 If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowsummarytable)) Then
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
    ws.Cells(2, 17).NumberFormat = "0.00%"

'Min Percent Change
 ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowsummarytable)) Then
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
    ws.Cells(3, 17).NumberFormat = "0.00%"
            
'Max Volume
ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowsummarytable)) Then
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    
'Ending the loop

 End If
        Next i
  Next ws
        End Sub




