Sub Ticker()

For Each ws In Worksheets

Dim Ticker_Name As String

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim Ticker_Total As Double
Ticker_Total = 0

Dim Open_Number As Double
Open_Number = 0

Dim Close_Number As Double
Close_Number = 0

Dim Yearly_Change As Double

Dim Percentage_Chage As Double
Percentage_Change = 0

Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total_Volume As Double

WorksheetName = ws.Name

' Set titles for needed inputs

ws.Cells(1, 9).Value = "Ticker Name"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


' Lastrow statement

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If Open_Number = 0 Then
    
    Open_Number = ws.Cells(i, 3).Value
    
End If
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Ticker Name

    Ticker_Name = ws.Cells(i, 1).Value
    
    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    
' Yearly Change
    
    Close_Number = ws.Cells(i, 6).Value
    
    Yearly_Change = Close_Number - Open_Number
    
    ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
    
    
' Ticker Total with non matching cells
    
    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
    ws.Range("L" & Summary_Table_Row).Value = Ticker_Total

' Percentage Change
    
    Percentage_Change = (Yearly_Change / Open_Number) * 100
    
    ws.Range("K" & Summary_Table_Row).Value = "%" & Percentage_Change
    
    
' Reset Numbers and Add 1 to Summary Table Row
    
    Ticker_Total = 0
    
    Open_Number = 0
    
    Summary_Table_Row = Summary_Table_Row + 1
    
Else

' Ticker Total for all other cells

    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
    If ws.Range("J" & Summary_Table_Row).Value < 0 Then

    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    
    Else
    
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
End If

    

End If

Next i


lastrow2 = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Greatest_Total_Volume = ws.Cells(2, 12).Value
    Greatest_Increase = ws.Cells(2, 11).Value
    Greatest_Decrease = ws.Cells(2, 11).Value
        
For i = 2 To lastrow2

If Cells(i, 12).Value > Greatest_Total_Volume Then
Greatest_Total_Volume = Cells(i, 12).Value
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value

Else

Greatest_Total_Volume = Greatest_Total_Volume

End If

If Cells(i, 11).Value > Greatest_Increase Then
Greatest_Increase = Cells(i, 11).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value

Else

Greatest_Total_Volume = Greatest_Total_Volume

End If

If ws.Cells(i, 11).Value < Greatest_Decrease Then
Greatest_Decrease = ws.Cells(i, 11).Value
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

Else

Greatest_Decrease = Greatest_Decrease

End If

ws.Cells(2, 17).Value = Format(Greatest_Increase, "Percent")
ws.Cells(3, 17).Value = Format(Greatest_Decrease, "Percent")
ws.Cells(4, 17).Value = Format(Greatest_Total_Volume, "Scientific")

Next i

    Worksheets(WorksheetName).Columns("A:Z").AutoFit
    
Next ws


End Sub
