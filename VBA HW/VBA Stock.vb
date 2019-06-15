VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'VBA HOMEWORK : VEENA KOTTOOR
    Sub StockData()
    
        ' Go through all the worksheets
    Dim WS As Worksheet
        For Each WS In ActiveWorkbook.Worksheets
        WS.Activate
        
            ' Determine the Last Row
            LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    

            ' Add new headings
            
            Cells(1, 9).Value = "Ticker"
       
            Cells(1, 10).Value = "Total Stock Volume"
            
            'Create Variables to hold Value
            
            Dim OpenPrice As Double
            Dim ClosePrice As Double
            Dim TickerName As String
            Dim Volume As Double
            Volume = 0
            Dim Row As Double
            Row = 2
            Dim Column_Number As Integer
            Column_Number = 1
            Dim i As Long
            
            'Set Initial Open Price
            OpenPrice = Cells(2, Column_Number + 2).Value
            
             ' Loop through all ticker symbol
            
            For i = 2 To LastRow
                If Cells(i + 1, Column_Number).Value <> Cells(i, Column_Number).Value Then
                    ' Set Ticker
                    TickerName = Cells(i, Column_Number).Value
                    Cells(Row, Column_Number + 8).Value = TickerName
                    ' Set ClosePrice
                    ClosePrice = Cells(i, Column_Number + 5).Value
                    ' Add TotalVolume
                    Volume = Volume + Cells(i, Column_Number + 6).Value
                    Cells(Row, Column_Number + 9).Value = Volume
                    ' Increment rows
                    Row = Row + 1
                    ' Reset the OpenPrice
                    OpenPrice = Cells(i + 1, Column_Number + 2)
                    ' Reset the VolumeTotal
                    Volume = 0
                Else
                    Volume = Volume + Cells(i, Column_Number + 6).Value
                End If
            Next i
            
        Next WS
            
    End Sub





