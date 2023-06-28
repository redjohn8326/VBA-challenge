Attribute VB_Name = "Module1"
Sub Year_Total()

    ' Worksheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets

    ' Get the number of rows
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Dim LastRow2 As Long
    
    ' Set an initial variable for the yearly change
    Dim Yearly_Change As Double
    Dim Yearly_Open As Double
    Dim Yearly_Close As Double
    Yearly_Open = Range("C2").Value
    
    ' Set an initial variable for the percent change
    Dim Percent_Change As Double
    
    ' Set an initial variable for holding the ticker
    Dim ticker As String

  ' Set an initial variable for holding the volumte total per ticker
    Dim Volume_Total As Double
    Volume_Total = 0

    ' Keep track of the location for ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Insert headers and values we are looking for
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest total volume"

    ' Loop through all rows
    For i = 2 To LastRow
   
    ' Check if the following ticker is different than the current one
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ' Set the ticker
        ticker = ws.Cells(i, 1).Value

        ' Add to the Volume Total
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value
        
        ' Set Yearly Change
        Yearly_Close = ws.Cells(i, 6).Value
        Yearly_Change = Yearly_Close - Yearly_Open

        ' Set Percentage
        Percent_Change = Yearly_Change / Yearly_Open

        ' Print the ticker in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = ticker
        
        ' Print the Yearly Change to the Summary Table
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

        ' Print the Percent Change to the Summary Table
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
               
        ' Print the Volume Total to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = Volume_Total

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the Volume Total, Yearly Change and Percent Change
        Volume_Total = 0
        Yearly_Open = ws.Cells(i + 1, 3).Value
    
        ' Check if the following ticker is the same than the current one
    Else
        ' Add to the Brand Total
        Volume_Total = Volume_Total + ws.Cells(i, 7).Value
    End If
    
    Next i
    
    ' Colour formatting
    LastRow2 = ws.Cells(Rows.Count, "K").End(xlUp).Row
    For j = 2 To LastRow
    
        If ws.Cells(j, 11).Value > 0 Then
             ws.Cells(j, 11).Interior.ColorIndex = 4
        ElseIf ws.Cells(j, 11).Value < 0 Then
            ws.Cells(j, 11).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 11).Interior.ColorIndex = 8
    
        End If
    Next j

' Greatest increase, decrease and total volume
    Dim lookupRange As Range
    Set lookupRange = ws.Range("K" & LastRow)
    Dim maxTicker As Variant
    Dim minTicker As Variant
    Dim VolTicket As Variant
    
    Great_Increase = WorksheetFunction.Max(ws.Range("K1:K" & LastRow))
    Great_Decrease = WorksheetFunction.Min(ws.Range("K1:K" & LastRow))
    Great_TotVolume = WorksheetFunction.Max(ws.Range("L1:L" & LastRow))
    
    maxTicker = Application.XLookup(Great_Increase, ws.Range("K1:K" & LastRow), ws.Range("I1:I" & LastRow))
    minTicker = Application.XLookup(Great_Decrease, ws.Range("K1:K" & LastRow), ws.Range("I1:I" & LastRow))
    VolTicker = Application.XLookup(Great_TotVolume, ws.Range("L1:L" & LastRow), ws.Range("I1:I" & LastRow))
    ws.Range("P2").Value = maxTicker
    ws.Range("P3").Value = minTicker
    ws.Range("P4").Value = VolTicker
    ws.Range("Q2").Value = Great_Increase
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").Value = Great_Decrease
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").Value = Great_TotVolume


    Next ws
End Sub


