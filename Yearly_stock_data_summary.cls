VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_data()

    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets

        'Add Column Names
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"


        'Define variables
    
    
        Dim lastrow As Long
        Dim ticker As String
        Dim year_open As Double
        Dim year_close As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_stock_volume As Double
        Dim summary_line As Integer
    
       'Set last row value
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        summary_line = 2
    
        'Grab ticker and associated values from first column
        For i = 1 To lastrow
            If ws.Range("A" & i + 1).Value <> ws.Range("A" & i).Value Then
        
            'Grab ticker from first row of new tickers and add it to summary list
            ticker = ws.Range("A" & i + 1).Value
            ws.Range("I" & summary_line).Value = ticker
        
            'Grab year open value and store it
            year_open = ws.Range("C" & i + 1).Value
        
            'Start adding volumes
            total_stock_volume = ws.Range("G" & i + 1).Value
            
        
        
            'Grab value for year close and include the last stock volume values
            ElseIf ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value And ws.Range("A" & i + 1).Value <> ws.Range("A" & i + 2).Value Then
            year_close = ws.Range("F" & i + 1).Value
            yearly_change = year_close - year_open
            percent_change = yearly_change / year_open
            
            total_stock_volume = total_stock_volume + Range("G" & i + 1)
            
                'Add values to summary lines and color condition for yearly change
                ws.Range("J" & summary_line).Value = yearly_change
                ws.Range("K" & summary_line).Value = percent_change
                ws.Range("L" & summary_line).Value = total_stock_volume
            
                'Format cells
                ws.Range("K" & summary_line).NumberFormat = "0.00%"
            
                'Format color condition
                If yearly_change < 0 Then
                ws.Range("J" & summary_line).Interior.ColorIndex = 3
                Else
                ws.Range("J" & summary_line).Interior.ColorIndex = 4
                End If
            
            'Once all values are pasted in summary, start new summary line
            summary_line = summary_line + 1
            
            'Grab values for all other volumes in ticker series. Logic is when first and second row are equal
            Else
            total_stock_volume = total_stock_volume + ws.Range("G" & i + 1).Value
            
            End If
            
        Next i
            
            
        'Add Greatest Inc & Dec Summary Table
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        Dim greatest_inc As Double
        Dim greatest_dec As Double
        Dim greatest_vol As Double
        Dim ticker1 As String
        Dim ticker2 As String
        Dim ticker3 As String
        
        'Set values to make comparison dynamic between current highest and next line
        greatest_inc = ws.Range("K2").Value
        greatest_dec = ws.Range("K2").Value
        greatest_vol = ws.Range("L2").Value
        
        For i = 2 To summary_line
            If ws.Range("K" & i).Value > greatest_inc Then
            greatest_inc = ws.Range("K" & i).Value
            ticker1 = ws.Range("I" & i).Value
            
            End If
            
            If ws.Range("K" & i).Value < greatest_dec Then
            greatest_dec = ws.Range("K" & i).Value
            ticker2 = ws.Range("I" & i).Value
            
            End If
            
            If ws.Range("L" & i).Value > greatest_vol Then
            greatest_vol = ws.Range("L" & i).Value
            ticker3 = ws.Range("I" & i).Value
        
            End If
            
        Next i
        'Display results
        ws.Range("O2").Value = ticker1
        ws.Range("O3").Value = ticker2
        ws.Range("O4").Value = ticker3
        
        ws.Range("P2").Value = greatest_inc
        ws.Range("P3").Value = greatest_dec
        ws.Range("P4").Value = greatest_vol
        
        'Format cells
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
        ws.Range("P4").NumberFormat = "0.00E+00"
    
        'Format Column widths
        ws.Columns("A:P").AutoFit
        
    Next ws

End Sub
