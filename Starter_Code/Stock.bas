Attribute VB_Name = "Module1"
Sub StockData()

    Dim ws As Worksheet
    Dim ticker As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim total As Double
    Dim openPrice As Double
    Dim closedPrice As Double
    Dim summaryTable As Double
    Dim LastRow As Long
    Dim LastRowTwo As Long
   
    Application.ScreenUpdating = False

    For Each ws In Worksheets

        ' Set headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
       
        ' Initialize variables
        summaryTable = 2
        openPrice = ws.Cells(2, 3).Value
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Set headers for greatest values
        ws.Cells(2, 15).Value = "Greatest Increase"
        ws.Cells(3, 15).Value = "Greatest Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        ' Main loop to process each row
        For i = 2 To LastRow
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                openPrice = ws.Cells(i + 1, 3).Value
            End If

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closedPrice = ws.Cells(i, 6).Value
                yearlyChange = closedPrice - openPrice
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice
                Else
                    percentChange = 0
                End If
                total = total + ws.Cells(i, 7).Value

                ' Store values in the summary table
                ws.Cells(summaryTable, 9).Value = ticker
                ws.Cells(summaryTable, 10).Value = yearlyChange
                ws.Cells(summaryTable, 11).Value = percentChange
                ws.Cells(summaryTable, 12).Value = total

                summaryTable = summaryTable + 1
                total = 0
            Else
                total = total + ws.Cells(i, 7).Value
            End If
        Next i

        ' Apply conditional formatting
        LastRowTwo = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For i = 2 To LastRowTwo
            If ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.Color = vbRed
            ElseIf ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.Color = vbGreen
            Else
                ws.Cells(i, 10).Interior.ColorIndex = xlNone ' Keep it white
            End If
        Next i

        ' Calculate greatest values
        With WorksheetFunction
            Dim greatestIncrease As Double
            Dim greatestDecrease As Double
            Dim greatestTotal As Double
            Dim greatIncTick As Long
            Dim greatDecTick As Long
            Dim greatTotTick As Long

            greatestIncrease = .Max(ws.Range("K:K"))
            greatIncTick = .Match(greatestIncrease, ws.Range("K:K"), 0)
            ws.Cells(2, 17).Value = ws.Cells(greatIncTick, 9).Value
            ws.Cells(2, 16).Value = greatestIncrease

            greatestDecrease = .Min(ws.Range("K:K"))
            greatDecTick = .Match(greatestDecrease, ws.Range("K:K"), 0)
            ws.Cells(3, 17).Value = ws.Cells(greatDecTick, 9).Value
            ws.Cells(3, 16).Value = greatestDecrease

            greatestTotal = .Max(ws.Range("L:L"))
            greatTotTick = .Match(greatestTotal, ws.Range("L:L"), 0)
            ws.Cells(4, 17).Value = ws.Cells(greatTotTick, 9).Value
            ws.Cells(4, 16).Value = greatestTotal
        End With

        ' Auto-fit columns
        ws.Columns("A:Q").AutoFit

    Next ws

    Application.ScreenUpdating = True

End Sub

