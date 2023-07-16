Attribute VB_Name = "Module1"
Option Explicit

Sub StockDataChallenge()
    
    Const FIRST_DATA_ROW As Integer = 2
    
    Dim LastRow As Long
    Dim Ticker As String
    Dim TotalRow As Long
    Dim InputRow As Long
    Dim YearlyChange As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PercentChange As Double
    Dim StockVol As Variant
    
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Volume"
    
    
    
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    'prepare for first stock
    TotalRow = FIRST_DATA_ROW
    OpenPrice = Cells(FIRST_DATA_ROW, 3).Value
    ClosePrice = Cells(FIRST_DATA_ROW, 6).Value
    StockVol = 0
    For InputRow = FIRST_DATA_ROW To LastRow
        Ticker = Cells(InputRow, 1).Value
        StockVol = StockVol + Cells(InputRow, 7).Value
        If Cells(InputRow + 1, 1).Value <> Ticker Then
            'input
            ClosePrice = Cells(InputRow, 6).Value
            'calculations
            YearlyChange = Round(ClosePrice - OpenPrice, 2)
            PercentChange = ((ClosePrice - OpenPrice) / OpenPrice)
            'output
            Cells(TotalRow, 10).Value = Ticker
            Cells(TotalRow, 11).Value = YearlyChange
            Cells(TotalRow, 12).Value = PercentChange
            Cells(TotalRow, 12).Value = Format(PercentChange, "percent")
            Cells(TotalRow, 13).Value = StockVol
            
            

            'prepare for next stock
            TotalRow = TotalRow + 1
            StockVol = 0
            
            
          
        End If

    Next InputRow
    MsgBox ("done")
End Sub
