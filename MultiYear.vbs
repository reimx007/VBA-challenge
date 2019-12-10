Attribute VB_Name = "Module11"
Sub stocks()
Dim ws As Worksheet

For Each ws In Worksheets

  ' Set initial variables for the ticker,
  ' volume, summary table, opening value, closing value and last row
  Dim ticker As String
  Dim volume, opening, closing, yearChange, perChange As Double
  Dim lastRow, rowCount As Long
  Dim Summary_Table_Row As Integer
  
  'Challenge variables
  Dim lastSumRow As Long
  Dim maxChange, minChange, maxVol As Double
  Dim maxTick, minTick, tickVol As String
  Dim rng, rngVol, rngSummary As Range
  
  'Set initial variable values
  volume = 0
  lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  Summary_Table_Row = 2
  'rowCount is used to determine how many rows are within the same ticker
  rowCount = 0
  
  ' Add headers to summary table
  ws.Range("J1:M1").Font.Bold = True
  ws.Range("J1").Value = "Ticker"
  ws.Range("K1").Value = "Yearly Change"
  ws.Range("L1").Value = "Percent Change"
  ws.Range("M1").Value = "Total Stock Volume"
  

  ' Loop through all stocks
  For i = 2 To lastRow
   
    ' Check if we are still reading the same ticker value, if it changes to new ticker:
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker name, volume, closing amount, row count. Calulate changes.
      ticker = ws.Cells(i, 1).Value
      closing = ws.Cells(i, 6).Value
      volume = volume + ws.Cells(i, 7).Value
      rowCount = rowCount + 1
      opening = ws.Cells(i - rowCount + 1, 3).Value
      yearChange = closing - opening
      'perChange = opening / yearChange
        If (opening = 0 Or closing = 0 Or yearChange = 0) Then
        perChange = 0
        Else
        perChange = opening / yearChange
        End If
        
        ' Print the ticker, opening, volume in the Summary Table if positive
        If yearChange >= 0 Then
      ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
      ws.Range("K" & Summary_Table_Row).Borders.LineStyle = xlContinuous
      ws.Range("K" & Summary_Table_Row).Value = yearChange
      ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
      ws.Range("L" & Summary_Table_Row).Borders.LineStyle = xlContinuous
      ws.Range("L" & Summary_Table_Row).Value = perChange
      ws.Range("J" & Summary_Table_Row).Value = ticker
      ws.Range("J" & Summary_Table_Row).Borders.LineStyle = xlContinuous
      ws.Range("M" & Summary_Table_Row).Value = volume
      ws.Range("M" & Summary_Table_Row).Borders.LineStyle = xlContinuous
      Else
      ' Print the ticker, opening, volume in the Summary Table if negative
      ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
      ws.Range("K" & Summary_Table_Row).Borders.LineStyle = xlContinuous
      ws.Range("K" & Summary_Table_Row).Value = yearChange
      ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
      ws.Range("L" & Summary_Table_Row).Borders.LineStyle = xlContinuous
      ws.Range("L" & Summary_Table_Row).Value = perChange
      ws.Range("J" & Summary_Table_Row).Value = ticker
      ws.Range("J" & Summary_Table_Row).Borders.LineStyle = xlContinuous
      ws.Range("M" & Summary_Table_Row).Value = volume
      ws.Range("M" & Summary_Table_Row).Borders.LineStyle = xlContinuous
      End If

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the volume and row count
      rowCount = 0
      volume = 0
    ' If the cell immediately following a row is the ticker...
    Else

      ' Add to volume and row count. Get closing value.
      volume = volume + ws.Cells(i, 7).Value
      rowCount = rowCount + 1
      closing = ws.Cells(i, 6).Value
    End If
     
  Next i
'Challenge things
' Add headers to NEW summary table
  ws.Range("O1:Q1").Font.Bold = True
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O2").Font.Bold = True
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O3").Font.Bold = True
  ws.Range("O4").Value = "Greatest Total Volume"
  ws.Range("O4").Font.Bold = True
'Get last row of summary table
lastSumRow = ws.Cells(Rows.Count, 12).End(xlUp).Row

Set rng = ws.Range("L2:L" & lastSumRow)
Set rngVol = ws.Range("M2:M" & lastSumRow)
maxChange = Application.WorksheetFunction.Max(rng)
minChange = Application.WorksheetFunction.Min(rng)
maxVol = Application.WorksheetFunction.Max(rngVol)
ws.Range("Q2").Value = maxChange
ws.Range("Q3").Value = minChange
ws.Range("Q4").Value = maxVol

maxTick = Application.VLookup(maxChange, ws.Range("J2:M" & lastSumRow), 1, False)
ws.Range("P2").Value = maxTick
minTick = Application.VLookup(minChange, ws.Range("J2:M" & lastSumRow), 1, False)
ws.Range("P3").Value = minTick
'tickVol = Application.VLookup(maxVol, ws.Range("J2:M" & lastSumRow), 1, False)
'ws.Range("P4").Value = tickVol

'MsgBox ws.Name
Next

End Sub
