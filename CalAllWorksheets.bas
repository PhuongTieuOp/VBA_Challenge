Attribute VB_Name = "CalAllWorksheets"
' Calculate summary and greatest value tables
'
Sub CalculateAllWorksheets()
    Dim vWS As Integer
    
    vWS = Application.Worksheets.Count
    
    For i = 1 To vWS
        Worksheets(i).Activate
        CalculateSummary
        CalculateGreatestValue
    Next i
    
    Worksheets(1).Activate
    Worksheets(1).Range("I1").Select
    
End Sub
