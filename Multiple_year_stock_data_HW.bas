Sub stockAnalysisWkbk()
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        stockAnalysisSheet ws
    Next
        
End Sub
Sub stockAnalysisSheet(ws As Worksheet)
    'defs
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim typeCount As Integer
    Dim yearlyChange As Double
    Dim stockVolume As Double
    Dim GPI, GPD, GTV As Double
    Dim GPITick, GPDTick, GTVTick As String
    '---------------------------
    
    'inits
    ticker = ""
    typeCount = 0
    stockVolume = 0
    GPI = -1.7E+308 'min double
    GPD = 1.7E+308 'max double
    GTV = 0
    '---------------------------
    
    'headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    '-----------------------------
    
    'Read first data row
    ticker = ws.Cells(2, 1).Value
    openingPrice = ws.Cells(2, 3).Value
    stockVolume = stockVolume + ws.Cells(2, 7).Value / 1000000 'prevent overflow
    '------------------------------
    
    'Following rows
    For r = 3 To ws.Cells(Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 1).Value <> ticker Then
            'ticker, openingPrice, stockVolume already exist

            'new summary row index
            typeCount = typeCount + 1
            
            'Read Previous Row
            closingPrice = ws.Cells(r - 1, 6).Value
            
            'Calculate
            yearlyChange = closingPrice - openingPrice
            
            'Fill Summary Table
            ws.Cells(typeCount + 1, 9).Value = ticker
            ws.Cells(typeCount + 1, 10).Value = yearlyChange
            If openingPrice > 0 Then
                ws.Cells(typeCount + 1, 11).Value = yearlyChange / openingPrice
            Else
                ws.Cells(typeCount + 1, 11).Value = 0
            End If
            ws.Cells(typeCount + 1, 12).Value = stockVolume * 1000000 'prevent overflow
            
            'Record edge cases
            If openingPrice > 0 Then
                If yearlyChange / openingPrice > GPI Then
                    GPI = yearlyChange / openingPrice
                    GPITick = ticker
                End If
            
            
                If yearlyChange / openingPrice < GPD Then
                    GPD = yearlyChange / openingPrice
                    GPDTick = ticker
                End If
            End If
            
            If stockVolume * 1000000 > GTV Then
                GTV = stockVolume * 1000000
                GTVTick = ticker
            End If
            
            'Color Backgrounds
            If yearlyChange > 0 Then
                ws.Cells(typeCount + 1, 10).Interior.ColorIndex = 4
            ElseIf yearlyChange < 0 Then
                ws.Cells(typeCount + 1, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(typeCount + 1, 10).Interior.ColorIndex = 2
            End If
            ws.Cells(typeCount + 1, 11).NumberFormat = "0.00%"
            
            
            'Read row
            ticker = ws.Cells(r, 1).Value
            openingPrice = ws.Cells(r, 3).Value
            stockVolume = 0
        End If
        
        'Accumulate every row
        stockVolume = stockVolume + ws.Cells(r, 7).Value / 1000000
                  
    Next r
    '-----------------------------------
    
    'Fill edge case Table
    ws.Cells(2, 16).Value = GPITick
    ws.Cells(3, 16).Value = GPDTick
    ws.Cells(4, 16).Value = GTVTick
    
    ws.Cells(2, 17).Value = GPI
    ws.Cells(3, 17).Value = GPD
    ws.Cells(4, 17).Value = GTV
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    '-----------------------------------
    
End Sub

