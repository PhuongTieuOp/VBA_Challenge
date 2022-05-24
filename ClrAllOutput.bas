Attribute VB_Name = "ClrAllOutput"
' Clear the calculated output for all worksheets
'
Sub ClearAllOutput()
    Dim vWS As Integer
    
    vWS = Application.Worksheets.Count
    
    For i = 1 To vWS
        Worksheets(i).Activate
        ClearOutput
    Next i
    
    Worksheets(1).Activate
    Worksheets(1).Range("I1").Select
    
End Sub

