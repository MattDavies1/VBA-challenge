Attribute VB_Name = "Module1"
' Sub to apply subroutine to all sheets in a workbook
' Run this subroutine to apply the below summarizer to the entirety of the workbook
Sub run_on_sheets():
    Dim num_ws As Integer
    num_ws = ActiveWorkbook.Worksheets.Count
    
    For i = 1 To num_ws
        Worksheets(i).Activate
        Call summarizer
     Next i

End Sub

' Sub which answers VBA Challenge

Sub summarizer():
        'Define variables for all loops
        Dim length As Long
        Dim count_symbol As Long
        Dim count_vol As LongLong
        Dim price_open As Double
        Dim price_close As Double
        Dim count_price As Double
        
                
        ' Make Data Table Labels
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change in Price"
        Range("L1").Value = "Yearly Pct. Change in Price"
        Range("M1").Value = "Total Stock Volume"
        
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        
        
    ' Find length of the sheet
        length = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
        
    ' Generate list of all Symbols
        
        count_symbol = 1
        
        For i = 2 To length
            If Cells(i, 1).Value <> Cells(count_symbol, 10).Value Then
                Cells(count_symbol + 1, 10).Value = Cells(i, 1).Value
                count_symbol = count_symbol + 1
            End If
        Next i
        
    ' Volume Summing Funtion
        
        count_vol = 2
        
        For i = 2 To length
            If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                total_vol = total_vol + Cells(i, 7).Value
            Else
                total_vol = total_vol + Cells(i, 7).Value
                Cells(count_vol, 13).Value = total_vol
                total_vol = 0
                count_vol = count_vol + 1
            End If
        Next i
        
    ' Function to generate delta price and % price change over year
        
        count_price = 2
        price_open = 0
        price_close = 0
        is_first = 1
        
        For i = 2 To length
            If (Cells(i, 1).Value = Cells(i + 1, 1).Value) And (is_first = 1) Then
                price_open = Cells(i, 3).Value
                is_first = 0
            ElseIf (Cells(i, 1).Value = Cells(i + 1, 1).Value) And (is_first = 0) Then
            ElseIf (Cells(i, 1).Value <> Cells(i + 1, 1).Value) And (price_open <> 0) Then
                price_close = Cells(i, 6).Value
                Cells(count_price, 11).Value = price_close - price_open
                Cells(count_price, 12).Value = FormatPercent((price_close - price_open) / price_open, 2)
                count_price = count_price + 1
                is_first = 1
            Else
                count_price = count_price + 1
                is_first = 1
            End If
        Next i
        
    ' Format Yearly Change Cell colours
        For i = 2 To Range("K" & Rows.Count).End(xlUp).Row
            If Cells(i, "K").Value < 0 Then
                Cells(i, "K").Interior.Color = vbRed
            ElseIf Cells(i, "K").Value > 0 Then
                Cells(i, "K").Interior.Color = vbGreen
            End If
        Next i
        
    ' Find Greatest Increase Decrease and Volume
        ' Set bins for comparison
        greatest_pct = Array("ticker", 0)
        least_pct = Array("ticker", 0)
        greatest_vol = Array("ticker", 0)
        
        For i = 2 To Range("J" & Rows.Count).End(xlUp).Row
            If Cells(i, "L").Value > greatest_pct(1) Then
                greatest_pct(0) = Cells(i, "J").Value
                greatest_pct(1) = Cells(i, "L").Value
            End If
            If Cells(i, "L").Value < least_pct(1) Then
                least_pct(0) = Cells(i, "J").Value
                least_pct(1) = Cells(i, "L").Value
            End If
            If Cells(i, "M").Value > greatest_vol(1) Then
                greatest_vol(0) = Cells(i, "J").Value
                greatest_vol(1) = Cells(i, "M").Value
            End If
        Next i
        
        ' Place Notable Values on sheet
        Range("P2").Value = greatest_pct(0)
        Range("Q2").Value = FormatPercent(greatest_pct(1))
        Range("P3").Value = least_pct(0)
        Range("Q3").Value = FormatPercent(least_pct(1))
        Range("P4").Value = greatest_vol(0)
        Range("Q4").Value = greatest_vol(1)
        
        'Fit all column widths to data
        ActiveSheet.Columns("J:O").AutoFit
    
End Sub

