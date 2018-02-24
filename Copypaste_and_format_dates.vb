Sub copyFormatDates()
    Dim xvar As Variant
    Worksheets("formated_data").Range("A2:G43").Clear
    Worksheets("formated_data").Range("A2:G43").Value = Worksheets("Exported").Range("A2:G43").Value
    Worksheets("formated_data").Range("I2:K43").Value = Worksheets("Exported").Range("I2:K43").Value
    
    For i = 2 To 43
    
    xvar = CStr(Worksheets("Exported").Range("H" & i).Value)
    xvar1 = CStr(Worksheets("Exported").Range("L" & i).Value)
    
    If Len(xvar) > 1 & xvar <> "" Then
        Worksheets("formated_data").Range("H" & i).Value = Format(Mid(xvar, 1, 2) & "/" & Mid(xvar, 3, 2) & "/" & Mid(xvar, 5, 4), "MM-dd-yyyy")
    Else
        Worksheets("formated_data").Range("H" & i).Value = ""
    End If
    
    If xvar1 <> "0" Then
        Worksheets("formated_data").Range("L" & i).Value = Format(Mid(xvar1, 1, 2) & "/" & Mid(xvar1, 3, 2) & "/" & Mid(xvar1, 5, 4), "MM-dd-yyyy")
    Else
        Worksheets("formated_data").Range("L" & i).Value = ""
    End If
    Next i
      
      
   Worksheets("formated_data").Activate
    
End Sub

