Sub SetColor()
    Worksheets("Current Sheet").Activate
    Dim irow
    Dim iResult
	
    For irow = 1 To 3000
      ' Range("A" & Chr(34) & irow & Chr(34)).Activate
       iResult = Range("A" & irow).Row Mod 2
            If iResult <> 0 Or irow = 1 Then
            Sheet1.Rows(irow).Interior.ColorIndex=3
        Else
           Sheet1.Rows(irow).Interior.ColorIndex=0        
        End If
      
    Next irow
End Sub

Sub Reset_DefaultColor()
    Dim irow
    For irow = 1 To 3000
        Sheet1.Rows(irow).Interior.ColorIndex = 0
    Next irow
End Sub
