Attribute VB_Name = "Module1"
Sub Ticker()
    
    Dim ws As Worksheet
    Dim LastRowA As Long
    Dim LastRowF As Long
    Dim ColA() As Variant
    Dim ColF() As Variant
    Dim outputRow As Long
    Dim Result As Double
    Dim PerResult As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim UniqueDict As Object
    Set UniqueDict = CreateObject("Scripting.Dictionary")
    
    For Each ws In ThisWorkbook.Sheets
        
        LastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        Set UniqueRange = ws.Range("A1:A" & LastRowA)
        UniqueRange.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ws.Range("I1"), Unique:=True
                           
        LastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        LastRowF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
        UniqueDict.RemoveAll

        ColA = ws.Range("A2:A" & LastRowA).Value
        ColF = ws.Range("F2:F" & LastRowF).Value

    For i = 1 To UBound(ColA)
        UniqueDict(ColA(i, 1)) = 1
        
        Next i

        
         outputRow = 2
                
    For Each UniqueValue In UniqueDict.Keys
            
            Result = 0
            PerResult = 0
            OpenPrice = 0
            ClosePrice = 0

           
    For i = 1 To UBound(ColA)
        If ColA(i, 1) = UniqueValue Then
        If OpenPrice = 0 Then
        OpenPrice = ws.Cells(i + 1, "C").Value
    
    End If
        
        ClosePrice = ColF(i, 1)
                
    End If
            
        Next i

           
    If OpenPrice <> 0 Then
        Result = ClosePrice - OpenPrice
        PerResult = Result / OpenPrice
            
    End If

            
        ws.Cells(outputRow, "J").Value = Result
        ws.Cells(outputRow, "K").Value = PerResult
        ws.Cells(outputRow, "K").NumberFormat = "0.00%"
        outputRow = outputRow + 1
        
    Next UniqueValue
        
    Next ws
        
    Call Format
        
    End Sub
        
    Sub Format()
        
    Dim ws As Worksheet
    Dim LastRowK As Long
    Dim LastRowJ As Long
        
    For Each ws In ThisWorkbook.Sheets
        
    LastRowK = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    LastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        
    Set ColourRange = ws.Range("K2:K" & LastRowK)
    Set ColourRange2 = ws.Range("J2:J" & LastRowJ)
        
    ColourRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    ColourRange.FormatConditions(1).Interior.Color = RGB(255, 0, 0) ' Red
    ColourRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
    ColourRange.FormatConditions(2).Interior.Color = RGB(255, 255, 0) ' Yellow
    ColourRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
    ColourRange.FormatConditions(3).Interior.Color = RGB(0, 255, 0) ' Green
        
    ColourRange2.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
    ColourRange2.FormatConditions(1).Interior.Color = RGB(255, 0, 0) ' Red
    ColourRange2.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
    ColourRange2.FormatConditions(2).Interior.Color = RGB(255, 255, 0) ' Yellow
    ColourRange2.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
    ColourRange2.FormatConditions(3).Interior.Color = RGB(0, 255, 0) ' Green
        
    Next ws
         
    Call TotalVolume
        
    End Sub
        
    Sub TotalVolume()
    
    Dim ws As Worksheet
    Dim LastRowA As Long
    Dim LastRowI As Long
    Dim ColA() As Variant
    Dim ColI() As Variant
    Dim TotalVolume As Double
    Dim TickerDict As Object
    Set TickerDict = CreateObject("Scripting.Dictionary")
    
    For Each ws In ThisWorkbook.Sheets
        LastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        LastRowI = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        TickerDict.RemoveAll
                
        ColA = ws.Range("A2:A" & LastRowA).Value
        ColI = ws.Range("I2:I" & LastRowI).Value

        
    For i = 1 To UBound(ColA)
        Dim Ticker As String
        Ticker = ColA(i, 1)
        If Not TickerDict.Exists(Ticker) Then
            TickerDict(Ticker) = 0
            
    End If
    
    Next i
         
    For i = 1 To UBound(ColI)
        Ticker = ColI(i, 1)
        TotalVolume = 0
           
    For j = 1 To UBound(ColA)
        If ColA(j, 1) = Ticker Then
        TotalVolume = TotalVolume + ws.Cells(j + 1, 7).Value
    
    End If
    
    Next j
            
    TickerDict(Ticker) = TotalVolume
        
    Next i

    For i = 1 To UBound(ColI)
        
        Ticker = ColI(i, 1)
        ws.Cells(i + 1, 12).Value = TickerDict(Ticker)
        
    Next i
    
    Next ws
    
    Call Bonus
    
End Sub
    
  Sub Bonus()
  
  Dim ws As Worksheet
  Dim MaxValL As Double
  Dim MinValK As Double
  Dim MaxValK2 As Double
  Dim CorrespondingValueL As String
  Dim CorrespondingValueK As String
  Dim CorrespondingValueK2 As String
  Dim CellL As Range, CellK As Range, CellK2 As Range
  
  For Each ws In ThisWorkbook.Sheets
             
        MaxValL = 0
        MinValK = 0
        MaxValK2 = 0
        CorrespondingValueL = ""
        CorrespondingValueK = ""
        CorrespondingValueK2 = ""
        
        For Each CellL In ws.Columns("L").Cells
            If IsNumeric(CellL.Value) Then
            If CellL.Value > MaxValL Then
            MaxValL = CellL.Value
            CorrespondingValueL = ws.Cells(CellL.Row, "I").Value
        
        End If
        End If
        Next CellL
        
        For Each CellK In ws.Columns("K").Cells
            If IsNumeric(CellK.Value) Then
            If CellK.Value < MinValK Or MinValK = 0 Then
                MinValK = CellK.Value
                CorrespondingValueK = ws.Cells(CellK.Row, "I").Value
        End If
        End If
        Next CellK
        
        For Each CellK2 In ws.Columns("K").Cells
            If IsNumeric(CellK2.Value) Then
            If CellK2.Value > MaxValK2 Then
                MaxValK2 = CellK2.Value
                CorrespondingValueK2 = ws.Cells(CellK2.Row, "I").Value
        End If
        End If
        Next CellK2
        
             
            
        ws.Cells(4, "P").Value = MaxValL
        ws.Cells(4, "O").Value = CorrespondingValueL
        ws.Cells(3, "P").Value = MinValK
        ws.Cells(3, "O").Value = CorrespondingValueK
        ws.Cells(2, "P").Value = MaxValK2
        ws.Cells(2, "O").Value = CorrespondingValueK2
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(1, 14).Value = "Measure"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Range("P4").NumberFormat = "0"
        ws.Range("P3").NumberFormat = "0.00%"
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells.EntireColumn.AutoFit
        
                  
        
        
    Next ws

End Sub

