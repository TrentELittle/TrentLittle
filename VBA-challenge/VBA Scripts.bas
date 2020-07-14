Attribute VB_Name = "Module1"

Sub YearlyChanges()

Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Headings
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        'Variables
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim Ticker As String
        Dim PercentChange As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim x As Long
        
        'Set Initial Open Price
        OpenPrice = Cells(2, Column + 2).Value
        
        
        'Loop through all stocks to find values
        For x = 2 To LastRow
            If Cells(x + 1, Column).Value <> Cells(x, Column).Value Then
                'Ticker Value
                Ticker = Cells(x, Column).Value
                Cells(Row, Column + 8).Value = Ticker
                ClosePrice = Cells(x, Column + 5).Value
                'Yearly Change Value
                YearlyChange = ClosePrice - OpenPrice
                Cells(Row, Column + 9).Value = YearlyChange
                'Percent Change Value
                If (OpenPrice = 0 And ClosePrice = 0) Then
                    PercentChange = 0
                ElseIf (OpenPrice = 0 And ClosePrice <> 0) Then
                    PercentChange = 1
                Else
                    PercentChange = YearlyChange / OpenPrice
                    Cells(Row, Column + 10).Value = PercentChange
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                'Volume Value
                Volume = Volume + Cells(x, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                
                'Re-Setting Values for the next row in the loop
                Row = Row + 1
                OpenPrice = Cells(x + 1, Column + 2)
                Volume = 0
            Else
                Volume = Volume + Cells(x, Column + 6).Value
            End If
        Next x
        
        'Setting Color to show if Yearly Change was Positive, Negative, or Zero
        YearlyChangeLastRow = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row
        For y = 2 To YearlyChangeLastRow
            If (Cells(y, Column + 9).Value > 0 Or Cells(y, Column + 9).Value = 0) Then
                Cells(y, Column + 9).Interior.ColorIndex = 4
            ElseIf Cells(y, Column + 9).Value < 0 Then
                Cells(y, Column + 9).Interior.ColorIndex = 3
            ElseIf Cells(y, Column + 9).Value = 0 Then
                Cells(y, Column + 9).Interior.ColorIndex = 6
            End If
        Next y
    Next ws
End Sub


