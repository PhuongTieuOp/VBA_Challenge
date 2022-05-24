Attribute VB_Name = "CalGreatestValue"

' Calculate greatest % increase, decrease, and total stock volume exchanged
'
Sub CalculateGreatestValue()

  Dim vMaxPctIncrease As Double
  Dim vMaxPctTickerName As String
  Dim vMinPctIncrease As Double
  Dim vMinPctTickerName As String
  Dim vMaxTotStkVolume As Double
  Dim vCurrPctChange As Double
  Dim vCurrTickerName As String
  Dim vCurrTotStkVolumne As Double
  
  ' Setup greatest value table header
  Range("O1") = "Ticker"
  Range("P1") = "Value"
  Range("N2") = "Greatest % Increase"
  Range("N3") = "Greatest % Decrease"
  Range("N4") = "Greatest Total Stock Volume"
     
  ' Assign value for the first row read
  vMaxPctIncrease = Cells(2, 11)
  vMinPctIncrease = Cells(2, 11)
  vMaxPctTickerName = Cells(2, 9)
  vMinPctTickerName = Cells(2, 9)
  vTickerName = Cells(2, 9)
  vTotStkVolume = Cells(2, 12)
  
  Range("K1").Select
  Range(Selection, Selection.End(xlDown)).Select
  vLastRow = Cells(Rows.Count, 1).End(xlUp).Row
 
 ' Loop thru Summary Table to find Greatest % Increase

  For i = 2 To vLastRow
  
    vCurrPctChange = Cells(i, 11)
    vCurrTotStkVolume = Cells(i, 12)
    vCurrTickerName = Cells(i, 9)
    ' Find Greatest Percentage Increase
    If vCurrPctChange > vMaxPctIncrease Then
         vMaxPctTickerName = vCurrTickerName
         vMaxPctIncrease = vCurrPctChange
    End If
    
    ' Find Greatest Percentage Decrease
    If vCurrPctChange < vMinPctIncrease Then
         vMinPctTickerName = vCurrTickerName
         vMinPctIncrease = vCurrPctChange
    End If
    
    ' Find Greatest Total Volume
    If vCurrTotStkVolume > vTotStkVolume Then
         vTickerName = vCurrTickerName
         vTotStkVolume = vCurrTotStkVolume
    End If

  Next i
  
  ' Print the Ticker Name and its relavant details to the greater value table
  Range("O2") = vMaxPctTickerName
  Range("P2") = vMaxPctIncrease
  Range("O3") = vMinPctTickerName
  Range("P3") = vMinPctIncrease
  Range("O4") = vTickerName
  Range("P4") = vTotStkVolume
  Range("I1").Select
  
End Sub

