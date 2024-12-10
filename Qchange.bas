Attribute VB_Name = "Module1"

Sub QChange()

    Dim ws As Worksheet
    Dim i As Long
    Dim LastRow As Long
    Dim EarliestOpenPrice As Double
    Dim LastTickerUnique As Long
    Dim UniqueTicker As String
    Dim total As Double
    Dim j As Long
    Dim change As Double
    Dim start As Long
    Dim percentChange As Double
    Dim OutputRow As Long
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolume As Double
    Dim MaxIncreaseTicker As String
    Dim MaxDecreaseTicker As String
    Dim MaxVolumeTicker As String

    ' Loop through each sheet
    For Each ws In Worksheets

        ' Create new headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ' Set initial values
        j = 0
        total = 0
        change = 0
        start = 2
        MaxIncrease = -99999
        MaxDecrease = 99999
        MaxVolume = 0

        ' Find the last row in column A
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through rows
        For i = 2 To LastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                total = total + ws.Cells(i, 7).Value

                ' Find the first non-zero starting value
                If ws.Cells(start, 3).Value = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If

                ' Calculate change
                change = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
                If ws.Cells(start, 3).Value <> 0 Then
                    percentChange = change / ws.Cells(start, 3).Value
                Else
                    percentChange = 0
                End If

                ' Update Max Increase, Decrease, and Volume
                If percentChange > MaxIncrease Then
                    MaxIncrease = percentChange
                    MaxIncreaseTicker = ws.Cells(i, 1).Value
                End If
                If percentChange < MaxDecrease Then
                    MaxDecrease = percentChange
                    MaxDecreaseTicker = ws.Cells(i, 1).Value
                End If
                If total > MaxVolume Then
                    MaxVolume = total
                    MaxVolumeTicker = ws.Cells(i, 1).Value
                End If

                ' Print the results
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = change
                ws.Range("K" & 2 + j).Value = percentChange
                ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                ws.Range("L" & 2 + j).Value = total

                ' Start of the next stock ticker
                start = i + 1
                j = j + 1

            Else
                total = total + ws.Cells(i, 7).Value
            End If

        Next i

        ' Apply Conditional Formatting for Column J (Quarterly Change)
        With ws.Range("J2:J" & j + 1).FormatConditions
            .Delete ' Remove previous formatting
            ' Positive (Green)
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            .Item(.Count).Interior.Color = RGB(0, 255, 0)
            ' Negative (Red)
            .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            .Item(.Count).Interior.Color = RGB(255, 0, 0)
        End With

        ' Output the greatest values
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("Q2").Value = MaxIncreaseTicker
        ws.Range("Q3").Value = MaxDecreaseTicker
        ws.Range("Q4").Value = MaxVolumeTicker
        ws.Range("R2").Value = MaxIncrease
        ws.Range("R3").Value = MaxDecrease
        ws.Range("R4").Value = MaxVolume
        ws.Range("R2:R3").NumberFormat = "0.00%"

    Next ws

    MsgBox "Process completed.", vbInformation

End Sub

