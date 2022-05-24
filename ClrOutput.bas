Attribute VB_Name = "ClrOutput"
' Clear the calculated output for a worksheet
'
Sub ClearOutput()
    
    ' Clear summary table contents
    Range("I1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ' Clear summary table fill color
    Range("I1:L1").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .PatternColorIndex = xlNoFill
    End With
    
    ' Clear greatest value table
    Range("N1:P6").Select
    Selection.ClearContents
    
    Range("I1").Select
End Sub

