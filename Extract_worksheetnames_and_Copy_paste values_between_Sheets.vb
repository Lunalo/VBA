
'On  This Workbook extract sheets into combobox and call the form
Sub CallForm()
Dim sht As Worksheet
frmPix.cboPickNames.Clear
Call frmPix.enableOK
For Each sht In ActiveWorkbook.Worksheets
    If Mid(sht.Name, 1, 1) = "S" Then
       frmPix.cboPickNames.AddItem sht.Name
    
    End If
Next sht

frmPix.Show
   
End Sub


' All this is done on the form
Private Sub cboPickNames_Change()
    Call enableOK
End Sub

Private Sub cmdExit_Click()
    frmPix.Hide
End Sub

'Copy  pasting
Private Sub cmdOK_Click()
    Worksheets("Malawi Pixel Analysis_30_12").Activate
    Range("A4:G50000").Clear
    Worksheets(Me.cboPickNames.Text).Range("A2:G50000").Copy
    Worksheets("Malawi Pixel Analysis_30_12").Range("A4").PasteSpecial
     
   'Worksheets("Malawi Pixel Analysis_30_12").Range("A4:G1000000").Value = Worksheets(Me.cboPickNames.Text).Range("A2:G100000").Value
    Me.Hide
End Sub

Sub enableOK()
    If Me.cboPickNames.Text = "" Then
        Me.cmdOK.Enabled = False
    Else
       Me.cmdOK.Enabled = True
    End If
End Sub
