VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysisAllSheets()
    Dim ws As Worksheet
    Dim totalVolume As Double
    Dim rowNum As Long
    Dim change As Double
    Dim resultRow As Integer
    Dim startRow As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim maxIncreaseRow As Long
    Dim maxDecreaseRow As Long
    Dim maxVolumeRow As Long

    For Each ws In ThisWorkbook.Worksheets
        ws.Activate

        With ws.Range("I1:L1")
            .Cells(1, 1).Value = "Ticker"
            .Cells(1, 2).Value = "Quarterly Change"
            .Cells(1, 3).Value = "Percent Change"
            .Cells(1, 4).Value = "Total Stock Volume"
        End With
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        resultRow = 0
        totalVolume = 0
        startRow = 2

        rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        For rowNum = 2 To rowCount
     
            If ws.Cells(rowNum + 1, 1).Value <> ws.Cells(rowNum, 1).Value Then
                totalVolume = totalVolume + ws.Cells(rowNum, 7).Value

                If totalVolume = 0 Then
                    ws.Range("I" & 2 + resultRow).Value = ws.Cells(rowNum, 1).Value
                    ws.Range("J" & 2 + resultRow).Value = 0
                    ws.Range("K" & 2 + resultRow).Value = "%" & 0
                    ws.Range("L" & 2 + resultRow).Value = 0

                Else
                    If ws.Cells(startRow, 3).Value = 0 Then
                        For findValueRow = startRow To rowNum
                            If ws.Cells(findValueRow, 3).Value <> 0 Then
                                startRow = findValueRow
                                Exit For
                            End If
                        Next findValueRow
                    End If

                    change = ws.Cells(rowNum, 6).Value - ws.Cells(startRow, 3).Value
                    percentChange = change / ws.Cells(startRow, 3).Value

                    startRow = rowNum + 1

                    ws.Range("I" & 2 + resultRow).Value = ws.Cells(rowNum, 1).Value
                    ws.Range("J" & 2 + resultRow).Value = change
                    ws.Range("J" & 2 + resultRow).NumberFormat = "0.00"
                    ws.Range("K" & 2 + resultRow).Value = percentChange
                    ws.Range("K" & 2 + resultRow).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + resultRow).Value = totalVolume

                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + resultRow).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + resultRow).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + resultRow).Interior.ColorIndex = 0
                    End Select
                End If

                totalVolume = 0
                change = 0
                resultRow = resultRow + 1

            Else
                totalVolume = totalVolume + ws.Cells(rowNum, 7).Value

            End If

        Next rowNum

        ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & resultRow + 1)) * 100
        ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & resultRow + 1)) * 100
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & resultRow + 1))

        maxIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & resultRow + 1)), ws.Range("K2:K" & resultRow + 1), 0)
        maxDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & resultRow + 1)), ws.Range("K2:K" & resultRow + 1), 0)
        maxVolumeRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & resultRow + 1)), ws.Range("L2:L" & resultRow + 1), 0)

        ws.Range("P2").Value = ws.Cells(maxIncreaseRow + 1, 9).Value
        ws.Range("P3").Value = ws.Cells(maxDecreaseRow + 1, 9).Value
        ws.Range("P4").Value = ws.Cells(maxVolumeRow + 1, 9).Value

        ws.Columns("L").AutoFit
    Next ws

End Sub

