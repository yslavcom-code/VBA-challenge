Attribute VB_Name = "Module1"
Sub ProcessSheets()
    Dim sheet As Integer
    
    For sheet = 1 To 4
        Dim sheetName As String
        sheetName = "Q" & sheet
    
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets(sheetName)
        Call ProcessSheet(ws)
    Next sheet
    
End Sub

Sub ProcessSheet(ws As Worksheet)

    ws.Cells(1, 8).Value = "Ticker"
    ws.Cells(1, 9).Value = "Quaterly Change"
    ws.Cells(1, 10).Value = "Percent Change"
    ws.Cells(1, 11).Value = "Total Stock Volume"

    Dim ticker As String
    Dim newTicker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As LongLong
    Dim volumeNow As LongLong
    Dim start As Boolean
    start = True
    
    Dim filledRow As Long
       
    ticker = ""
    newTicker = ""
    filledRow = 2
    
    volume = 0
    
    Dim lastRow  As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim Row As Long
    For Row = 2 To lastRow
        newTicker = ws.Cells(Row, 1).Value
        volumeNow = ws.Cells(Row, 7).Value
        If ticker <> newTicker Then
            If start = True Then
                start = False
                volume = volumeNow
                openPrice = ws.Cells(Row, 3).Value
                ticker = newTicker
            Else
                ' previous ticker
                Call calcData(ws, filledRow, closePrice, openPrice, volume, ticker)
                filledRow = filledRow + 1
                
                volume = volumeNow
                openPrice = ws.Cells(Row, 3).Value
                ticker = newTicker
            End If
        Else
            volume = volume + volumeNow
            closePrice = ws.Cells(Row, 6).Value
        End If
    Next Row
    ' process the last ticker
    Call calcData(ws, filledRow, closePrice, openPrice, volume, ticker)

End Sub


Private Sub calcData(ws As Worksheet, filledRow As Long, closePrice As Double, openPrice As Double, volume As LongLong, ticker As String)
    Dim change As Double
    change = closePrice - openPrice
    ws.Cells(filledRow, 9).Value = change
    If change >= 0 Then
        ws.Cells(filledRow, 9).Interior.Color = RGB(0, 255, 0)
    Else
        ws.Cells(filledRow, 9).Interior.Color = RGB(255, 0, 0)
    End If
    ws.Cells(filledRow, 10).Value = change / openPrice
    ws.Cells(filledRow, 10).NumberFormat = "0.00%"
    ws.Cells(filledRow, 11).Value = volume
    ws.Cells(filledRow, 8).Value = ticker

End Sub

