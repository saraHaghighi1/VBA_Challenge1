Attribute VB_Name = "Module1"
Sub FillRowForTicker()
    Dim Ticker_counts As Long
    Dim tickers As Variant
    Dim data As Variant
    Dim i As Long
    
    ' Get unique tickers
    GetUniqueTickers
    
    ' Get data into array
    data = Range("A2:G" & Cells(Rows.Count, "A").End(xlUp).Row).Value
    
    Ticker_counts = Cells(Rows.Count, "I").End(xlUp).Row - 1 ' Calculate ticker count
    
    ' Loop through tickers
    For i = 2 To Ticker_counts + 1
        Dim ticker As String
        ticker = Cells(i, 9).Value ' Get ticker from column I
        
         'Calculate and fill the row for the ticker
        Cells(i, 10).Value = CalculateYearlyChange(ticker, data)
        Cells(i, 11).Value = CalculatePercentageChange(ticker, data)
        Cells(i, 11).NumberFormat = "00.0%"
        Cells(i, 12).Value = CalculateTotalStock(ticker, data)
    Next i

    ApplyConditionalFormatting
    
    ' Set headers
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage change"
    Range("L1").Value = "Total Stock Volume"
    Range("Q3").Value = "Greatest % Increase"
    Range("Q4").Value = "Greatest % Decrease"
    Range("Q5").Value = "Total Stock Volume"
    Range("R2").Value = "Ticker"
    Range("S2").Value = "Value"
    
  
'Greatest % Increase
 Range("I:L").Sort key1:=Range("K1"), Order1:=xlDescending, Header:=xlYes
 Range("S3").Value = Range("K2").Value
 Range("R3").Value = Range("I2").Value
 Range("S3").NumberFormat = "0.00%"

'Greatest % Decrease
 Range("I:L").Sort key1:=Range("K1"), Order1:=xlAscending, Header:=xlYes
 Range("S4").Value = Range("K2").Value
 Range("R4").Value = Range("I2").Value
 Range("S4").NumberFormat = "0.00%"
'Greatest Total Volume
Range("I:L").Sort key1:=Range("L1"), Order1:=xlDescending, Header:=xlYes
Range("S5").Value = Range("L2").Value
Range("R5").Value = Range("I2").Value
End Sub

Sub GetUniqueTickers()
    ' Get unique tickers using Advanced Filter
    Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I1"), Unique:=True
End Sub

Function CalculateTotalStock(ticker As String, data As Variant) As LongLong
    ' Calculate Total Stock Volume for ticker
    Dim TotalStock As LongLong
    TotalStock = 0
    
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
        If data(i, 1) = ticker Then
            TotalStock = TotalStock + data(i, 7)
        End If
    Next i
    
    CalculateTotalStock = TotalStock
End Function

Function CalculateYearlyChange(ticker As String, data As Variant) As Double
    ' Calculate Yearly Change for ticker
    Dim openPrice As Double
    Dim closePrice As Double
    
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
        If data(i, 1) = ticker Then
            openPrice = data(i, 3) ' Open price
            Exit For
        End If
    Next i
    
    For i = UBound(data, 1) To LBound(data, 1) Step -1
        If data(i, 1) = ticker Then
            closePrice = data(i, 6) ' Close price
            Exit For
        End If
    Next i
    
    CalculateYearlyChange = closePrice - openPrice
End Function

Function CalculatePercentageChange(ticker As String, data As Variant) As Double
    ' Calculate Percentage Change for ticker
    Dim openPrice As Double
    Dim closePrice As Double
    
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
        If data(i, 1) = ticker Then
            openPrice = data(i, 3) ' Open price
            Exit For
        End If
    Next i
    
    For i = UBound(data, 1) To LBound(data, 1) Step -1
        If data(i, 1) = ticker Then
            closePrice = data(i, 6) ' Close price
            Exit For
        End If
    Next i
    
    If openPrice <> 0 Then
        CalculatePercentageChange = (closePrice - openPrice) / openPrice
    Else
        CalculatePercentageChange = 0
    End If
End Function

Sub ApplyConditionalFormatting()
    ' Apply conditional formatting for Percentage change column (Column J)
    Dim rng As Range
    Set rng = Range("J:J")
    
    ' Apply conditional formatting for positive change (green)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0) ' Green
    End With
    
    ' Apply conditional formatting for negative change (red)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0) ' Red
    End With
    
    ' Apply formatting
    rng.FormatConditions(1).StopIfTrue = False
    rng.FormatConditions(2).StopIfTrue = False
End Sub

