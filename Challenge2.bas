Attribute VB_Name = "Module1"
Sub Challenge()
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        Dim ticker As String
        Dim openPrice As Double
        Dim closePrice As Double
        Dim volume As Double
        
        Dim percentChange As Double
        
        Dim gpI As Double
        Dim gpD As Double
        Dim gtV As Double
        
        Dim gpI_ticker As String
        Dim gpD_ticker As String
        Dim gtV_ticker As String
        
        Dim lastRow As Long
        Dim r As Integer
    
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        r = 2
        
        gpI = 0
        gpD = 0
        gtV = 0
        
        For i = 2 To lastRow
            
            ticker = ws.Cells(i, 1).Value
            
            If volume = 0 Then
                openPrice = ws.Cells(i, 3).Value
            End If
            
            volume = volume + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> ticker Then
                closePrice = ws.Cells(i, 6)
                percentChange = (closePrice - openPrice) / openPrice
                
                ws.Cells(r, 9).Value = ticker
                ws.Cells(r, 10).Value = closePrice - openPrice
                ws.Cells(r, 11).Value = percentChange
                ws.Cells(r, 12).Value = volume
                
                If percentChange > gpI Then
                    gpI = percentChange
                    gpI_ticker = ticker
                ElseIf percentChange < gpD Then
                    gpD = percentChange
                    gpD_ticker = ticker
                End If
                
                If volume > gtV Then
                    gtV = volume
                    gtV_ticker = ticker
                End If
                
                volume = 0
                r = r + 1
                
            End If
            
        Next i
        
        ws.Cells(2, 16).Value = gpI_ticker
        ws.Cells(2, 17).Value = gpI
        
        ws.Cells(3, 16).Value = gpD_ticker
        ws.Cells(3, 17).Value = gpD
        
        ws.Cells(4, 16).Value = gtV_ticker
        ws.Cells(4, 17).Value = gtV
        
    Next ws
    
End Sub






