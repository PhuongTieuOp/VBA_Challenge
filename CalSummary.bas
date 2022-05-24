Attribute VB_Name = "CalSummary"
' Calculate yearly changed, percentage and total stock volume exchanged
Sub CalculateSummary()

  ' Declare variables
  Dim vSummaryTableRow As Integer
  Dim vLastRow As Long
  Dim vYrOpenPrice As Double
  Dim vYrChangePrice As Double
  Dim vPctChange As Double
  Dim vTotStkVolume As Double
  
  ' Initialise variables
  vSummaryTableRow = 2
  
  ' Setup summary table header
  Range("I1") = "Ticker"
  Range("J1") = "Yearly Changed"
  Range("K1") = "Percentage Changed"
  Range("L1") = "Total Stock Volume"
  vYrOpenPrice = Range("C2")
  
  vLastRow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all Ticker's data rows
  For i = 2 To vLastRow

    ' Check if we are still within the same ticker name, if it is not...
    If Cells(i + 1, 1) <> Cells(i, 1) Then

      ' Calculate and assign value to variables
      vYrChangePrice = Cells(i, 6) - vYrOpenPrice
      If vYrOpenPrice <> 0 Then
        vPctChange = vYrChangePrice / vYrOpenPrice
      End If
      vTotStkVolume = vTotStkVolume + Cells(i, 7)

      ' Print the Ticker Name and its relavant details to the Summary Table
      Range("I" & vSummaryTableRow) = Cells(i, 1)
      Range("J" & vSummaryTableRow) = vYrChangePrice
      Range("K" & vSummaryTableRow) = vPctChange
      Range("L" & vSummaryTableRow) = vTotStkVolume

      ' Change color to Red = 3 if value is negative, if positive color Green = 4
      If vYrChangePrice < 0 Then
        Range("J" & vSummaryTableRow).Interior.ColorIndex = 3
      Else
        Range("J" & vSummaryTableRow).Interior.ColorIndex = 4
      End If

      ' After printing out, increment the summary table row count, and reset variables for next row
      vSummaryTableRow = vSummaryTableRow + 1
      vYrOpenPrice = Cells(i + 1, 3)
      vYrChangePrice = 0
      vPctChange = 0
      vTotStkVolume = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add value to Total variable
      vTotStkVolume = vTotStkVolume + Cells(i, 7)
      
    End If
  Next i

End Sub



