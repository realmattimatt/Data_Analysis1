Option Explicit

Sub get_ticker()
    
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Quarterly_change As Double
    Dim Ticker_vol As Double
    Dim Percent_change
    Dim Open_price As Double
    Dim Close_price As Double
    Dim Lastrow As Long
    Dim i As Long
    Dim Summary_row As Long
    Dim range As range
    Dim Value As Double
    Dim Greatest_increase As Double
    Dim Greatest_decrease As Double
    Dim Greatest_vol As Double
    Dim Increase_ticker As String
    Dim Decrease_ticker As String
    Dim Vol_ticker As String
    
    
    For Each ws In Worksheets
    
        ws.range("J1").Value = ("Ticker")
        ws.range("K1").Value = ("Quarterly Change")
        ws.range("L1").Value = ("Percent Change")
        ws.range("M1").Value = ("Total Stock Volume")
        ws.range("P1").Value = ("Ticker")
        ws.range("Q1").Value = ("Value")
        ws.range("O2").Value = ("Greatest % Increase")
        ws.range("O3").Value = ("Greatest % Decrease")
        ws.range("O4").Value = ("Greatest Total Volume")
        ws.Columns("J:M").AutoFit
        
       
       
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Summary_row = 2
        Ticker_vol = 0
        Greatest_increase = 0
        Greatest_decrease = 0
        Greatest_vol = 0
        ws.range("Q2:Q3").NumberFormat = "0.00%"
        ws.range("Q4").NumberFormat = "#,##0.00"
        
        
        For i = 2 To Lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                Ticker = ws.Cells(i, 1).Value
                Open_price = ws.Cells(i - (i - 2), 3).Value
                Close_price = ws.Cells(i, 6).Value
                Quarterly_change = Close_price - Open_price
                Ticker_vol = Ticker_vol + ws.Cells(i, 7).Value
                If Open_price <> 0 Then
                    Percent_change = (Quarterly_change / Open_price)
                Else
                    Percent_change = 0
                End If
                
                
                ws.Cells(Summary_row, 10).Value = Ticker
                ws.Cells(Summary_row, 11).Value = Quarterly_change
                ws.Cells(Summary_row, 12).Value = Percent_change
                ws.Cells(Summary_row, 13).Value = Ticker_vol
                
                
                With ws.Cells(Summary_row, 11).FormatConditions.Delete
                ws.Cells(Summary_row, 11).FormatConditions.Add Type:=xlCellValue, _
                    Operator:=xlGreater, Formula1:="=0"
                ws.Cells(Summary_row, 11).FormatConditions(1).Interior.Color = RGB(144, 238, 144)
                
                ws.Cells(Summary_row, 11).FormatConditions.Add Type:=xlCellValue, _
                    Operator:=xlLess, Formula1:="=0"
                ws.Cells(Summary_row, 11).FormatConditions(2).Interior.Color = RGB(255, 99, 71)
                
                End With
                
                If Percent_change > Greatest_increase Then
                    Greatest_increase = Percent_change
                    Increase_ticker = Ticker
                End If
                
                If Percent_change < Greatest_increase Then
                    Greatest_decrease = Percent_change
                    Decrease_ticker = Ticker
                End If
                
                If Ticker_vol > Greatest_vol Then
                    Greatest_vol = Ticker_vol
                    Vol_ticker = Ticker
                End If
            
                Summary_row = Summary_row + 1
                Ticker_vol = 0
            End If
            
        Next i
        ws.range("P2").Value = Increase_ticker
        ws.range("Q2").Value = Greatest_increase
        ws.range("Q2").NumberFormat = "0.00%"
        ws.range("P3").Value = Decrease_ticker
        ws.range("Q3").Value = Greatest_decrease
        ws.range("Q3").NumberFormat = "0.00%"
        ws.range("P4").Value = Vol_ticker
        ws.range("Q4").Value = Greatest_vol
        ws.Columns("O:Q").AutoFit
        
    Next ws
End Sub
