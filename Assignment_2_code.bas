Attribute VB_Name = "Module111"
Sub fun():

For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
WorksheetName = ws.Name

Dim ticker As String

Dim tickerVol As Double
tickerVol = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim openingPrice As Double
openingPrice = ws.Cells(2, 3)

Dim closingPrice As Double
Dim quarterlyChange As Double
Dim percentChange As Double

Dim ticker2 As String
Dim greatest_increase As Double
Dim greatest_decrease As Double
Dim greatest_volume As Double


greatest_increase = 0
greatest_decrease = 0
greatest_volume = 0



ws.Cells(1, "I") = "Ticker"
ws.Cells(1, "J") = "Quarterly Change"
ws.Cells(1, "K") = "Percent Change"
ws.Cells(1, "L") = "Total Stock Volume"

ws.Cells(1, "P") = "Ticker"
ws.Cells(1, "Q") = "Value"


For r = 2 To lastrow

    If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
    
        ticker = ws.Cells(r, 1).Value
        tickerVol = tickerVol + ws.Cells(r + 7).Value
        ws.Range("I" & Summary_Table_Row).Value = ticker
  
  
        ws.Range("L" & Summary_Table_Row).Value = tickerVol
        
        closePrice = ws.Cells(r, 6).Value
        quarterlyChange = (closePrice - openingPrice)
        
        ws.Range("J" & Summary_Table_Row).Value = quarterlyChange
        
     If (openingPrice = 0) Then
        
            percentChange = 0
            
        Else
        
            percentChange = quarterlyChange / openingPrice
        
End If

        ws.Range("K" & Summary_Table_Row).Value = percentChange
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
    Summary_Table_Row = Summary_Table_Row + 1
    
    tickerVol = 0
    
    openingPrice = ws.Cells(r + 1, 3)
    
Else

tickerVol = tickerVol + ws.Cells(r, 7).Value

End If


If ws.Cells(r, 11).Value > ws.Cells(2, 17).Value Then
ws.Cells(2, 17).Value = greatest_increase
ws.Cells(2, 17).Value = ws.Cells(r, 11).Value


End If

r = r + 1


If ws.Cells(r, 11).Value < ws.Cells(3, 17).Value Then
ws.Cells(3, 17) = greatest_decrease
ws.Cells(2, 17).Value = ws.Cells(r, 11).Value

End If

r = r + 1


If ws.Cells(r, 12).Value > ws.Cells(4, 17).Value Then
ws.Cells(4, 17).Value = greatest_volume
ws.Cells(4, 17).Value = ws.Cells(r, 12).Value

End If

r = r + 1



If ws.Cells(r, 10).Value > 0 Then
        ws.Cells(r, 10).Interior.ColorIndex = 4
    
ElseIf ws.Cells(r, 10).Value < 0 Then
    ws.Cells(r, 10).Interior.ColorIndex = 3
    
End If

Next r
Next ws


End Sub
