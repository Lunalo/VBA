	'On  Sheet data call

	Sub SetDataEntryForm()
		frmDataEntry.txtYear.Value = Year(Now())
		frmDataEntry.Show
		frmDataEntry.txtWptID.SetFocus
	End Sub

	Private Sub cmdConfirmEntry_Click()
		Dim iNumNonEmpty, iRowNumCurr, iColNumCurr, iCol As Variant
		'Call this for purpose of initialising
		Call setHhLim
		Call setJerLimit
		Call SetLimOldStock
		Call setLimWtpid
		
		'Make DATA sheet to be active
		Worksheets("DATA").Activate
			
			With ActiveSheet
			  Range("H1").Select
			  iColNumCurr = Range("H1").Column
			  iRowNumCurr = Range("H1").Row
			  
			  iNumNonEmpty = Range("H1:H30000").Cells.SpecialCells(xlCellTypeConstants).Count
			  
			  Cells(iRowNumCurr + iNumNonEmpty, iColNumCurr).Activate
		   
			  Cells(ActiveCell.Row, iColNumCurr) = Me.txtWptID.Value
			  Cells(ActiveCell.Row, iColNumCurr + 1) = Me.txtDay.Value & "/" & Me.txtMonth.Value & "/" & Me.txtYear.Value
			  Cells(ActiveCell.Row, iColNumCurr + 2) = Me.txtJerDeliv.Value
			  Cells(ActiveCell.Row, iColNumCurr + 3) = Me.txtOldStock.Value
			  Cells(ActiveCell.Row, iColNumCurr + 5) = Me.txtHh.Value
			
				'Change color if geographical information of waterpoint is missing in the database
				 For iCol = 1 To 7
				   If Cells(ActiveCell.Row, iColNumCurr) <> "" Then
					If Cells(ActiveCell.Row, iColNumCurr - iCol) = "" Then
					   Cells(ActiveCell.Row, iColNumCurr - iCol).Interior.Color = RGB(255, 0, 0)
					Else
						 Cells(ActiveCell.Row, iColNumCurr - iCol).Interior.Color = RGB(0, 0, 0)
					End If
						
				  Else
					Cells(ActiveCell.Row, iColNumCurr - iCol).Interior.Color = RGB(0, 0, 0)
				  End If
				Next iCol
			End With
		 Call ResetForm
		
	End Sub

	'Household number must have a length of 3 or less
	Sub setHhLim()
		Dim iVal As Variant
		iVal = Me.txtHh.Value
		If Len(iVal) > 3 Then
			MsgBox ("The value must have a length of atleast 3 Characters")
			Me.txtHh.Value = ""
			Me.txtHh.SetFocus
			Me.cmdConfirmEntry.Enabled = False
		End If
	End Sub

	'Jericans limit must within 1 and 9
	Sub setJerLimit()
		Dim iVal As Variant
		iVal = Me.txtJerDeliv.Value
		If Len(iVal) > 1 Then
			MsgBox ("The value must have a length of 1")
			Me.txtJerDeliv.Value = ""
			Me.txtJerDeliv.SetFocus
			Me.cmdConfirmEntry.Enabled = False
		End If
	End Sub

	'This is condition for enabling data entry confirmation
	Sub enableEnter()
		If Me.txtJerDeliv.Value = "" Or txtWptID.Value = "" Or txtYear.Value = "" Or txtDay.Value = "" Or txtOldStock.Value = "" Or Me.txtMonth.Value = "" Or Me.txtHh.Value = "" Or Len(txtYear.Value) < 4 Then
			Me.cmdConfirmEntry.Enabled = False
		Else
			Me.cmdConfirmEntry.Enabled = True
		End If
	End Sub

	' Month value must be non empty, numeric and within a range of 1 and 12
	Sub monthCheck()
		Dim monthval As Integer
		If Me.txtMonth.Value <> "" Then
			If IsNumeric(txtMonth.Value) Then
				monthval = CInt(Me.txtMonth.Value)
				If monthval < 1 Or monthval > 12 Then
					MsgBox ("The value entered should be integer between 1-12")
					Me.cmdConfirmEntry.Enabled = False
					 Me.txtMonth.Value = ""
				End If
		   Else
				Me.cmdConfirmEntry.Enabled = False
				Me.txtMonth.Value = ""
				Me.txtMonth.SetFocus
			End If
		End If
	End Sub

	'Waterpoint ID must be non emptyand numeric
	Sub checkWtptID()
		If Me.txtWptID.Value <> "" Then
			If Not IsNumeric(Me.txtWptID.Value) Then
					MsgBox ("The value should be Integer")
				   Me.txtWptID.Value = ""
			  
			End If
		 End If
	End Sub

	'Year Value must be 4 digit and current/or las year, numeric and non empty
	Sub yearCheck()
		If Me.txtYear.Value <> "" Then
			If IsNumeric(txtYear.Value) Then
					If Len(txtYear.Value) = 4 Then
						If (CInt(txtYear.Value) <> (Year(Now()) - 1)) And (CInt(txtYear.Value) <> Year(Now())) Then
							MsgBox ("The value entered should be integer of either " & Year(Now()) - 1 & " or " & Year(Now()))
							Me.cmdConfirmEntry.Enabled = False
							Me.txtYear.Value = ""
											
						End If
					 ElseIf (Len(txtYear.Value) < 4 Or Len(txtYear.Value) > 4) Then
						MsgBox ("The Year Value must have length of 4 to continue")
						Me.txtYear.Value = ""
					 End If
			Else
				 Me.cmdConfirmEntry.Enabled = False
				Me.txtMonth.Value = Year(Now())
		   End If
		 End If
	End Sub

	'Number of Households must be numeric and non empty
	Sub checkHhNum()
		If Me.txtHh.Value <> "" Then
			If Not IsNumeric(Me.txtHh.Value) Then
					MsgBox ("The value should be Integer")
				   Me.txtHh.Value = ""
				 
			  End If
		 End If
	End Sub

	'OldStock value must be non empty and numeric
	Sub checkOldStock()
		If Me.txtOldStock.Value <> "" Then
			If Not IsNumeric(Me.txtOldStock.Value) Then
					MsgBox ("The value should be Integer")
				   Me.txtOldStock.Value = ""
						   
			End If
		 End If
	End Sub

	'Jericans delivered must be indicated in numeric form
	Sub checkJerDeliv()
		If Me.txtJerDeliv.Value <> "" Then
			If Not IsNumeric(Me.txtJerDeliv.Value) Then
					MsgBox ("The value should be Integer")
				   Me.txtJerDeliv.Value = ""
					   
			End If
		 End If
		
	End Sub

	'Old stock value must be one digit
	Sub SetLimOldStock()
		Dim iVal As Variant
		iVal = Me.txtOldStock.Value
		If Len(iVal) > 1 Then
			MsgBox ("The value must have a length of 1")
			Me.txtOldStock.Value = ""
			Me.txtOldStock.SetFocus
		End If
	End Sub

	'Length of waterpoint ID must either be 8 or 7
	Sub setLimWtpid()
		Dim iWaterpoint As Variant
		iWaterpoint = Me.txtWptID.Value
		If iWaterpoint <> "" Then
			If Len(iWaterpoint) > 8 Or Len(iWaterpoint) < 7 Then
				MsgBox ("Waterpoint ID must contain 7 or 8 characters")
				Me.txtWptID.Value = ""
			End If
		End If
	End Sub

	'First value for waterpoint ID must be 8 to continue
	Sub checkFirstVal()
		Dim strVal As String
		strVal = Me.txtWptID.Value
		If Mid(strVal, 1, 1) <> 8 And Me.txtWptID.Value <> "" Then
			MsgBox ("The first value of waterpoint ID must be 8")
			Me.txtWptID.Value = ""
		 End If
	End Sub

	'Provide rules for leap years and non leap years. Also set ending values for each month
	Sub checkMonthYear()
		Dim iYear, iMonth, iDay
		iYear = Me.txtYear.Value
		iMonth = Me.txtMonth.Value
		iDay = Me.txtDay.Value
		  If txtYear.Value <> "" & txtMonth.Value <> "" And txtDay.Value <> "" Then
				If iMonth = 2 Then
					If (iYear Mod 400) = 0 Or ((iYear Mod 4) = 0 And (iYear Mod 100) = 0) Then
						If iDay > 29 And iDay >= 1 Then
							MsgBox ("The maximum day value should be 29 for leap years")
							txtDay.Value = ""
						End If
					Else
						 If iDay > 28 And iDay >= 1 Then
							MsgBox ("The maximum day value should be 28 for non leap years")
							txtDay.Value = ""
						 End If
					End If
			
				ElseIf (iMonth = 4 Or iMonth = 6 Or iMonth = 9 Or iMonth = 11) Then
					 If iDay > 30 And iDay >= 1 Then
						MsgBox ("Last day of this month is 30")
						txtDay.Value = ""
					End If
				Else
					 If iDay > 31 And iDay >= 1 Then
						MsgBox ("Last day of this month is 31")
						txtDay.Value = ""
					End If
				
				End If
			End If
	 End Sub
			
	'Check whether day value is numeric and control contains value
	Sub checkNumericDay()
		If Not Me.txtDay.Value = "" Then
			If Not IsNumeric(Me.txtDay.Value) Then
				MsgBox ("Day value must be numeric to continue")
				Me.txtDay.Value = ""
			End If
		End If
	End Sub

	'Call Respective methods/ Subroutines in the control events
		
	Private Sub UserForm_Initialize()
		 Call enableEnter
	End Sub

	Sub ResetForm()
		Me.txtHh.Value = ""
		Me.txtOldStock.Value = ""
		Me.txtJerDeliv.Value = ""
		Me.txtWptID.Value = ""
		Me.txtDay.Value = ""
		Me.txtMonth.Value = ""
		Me.txtYear.Value = Year(Now())
		Me.txtWptID.SetFocus
	End Sub

	Private Sub txtMonth_AfterUpdate()
		Call checkMonthYear
	End Sub

	Private Sub txtDay_AfterUpdate()
		Call checkMonthYear
	End Sub

	Private Sub txtDay_Change()
		Call enableEnter
		checkNumericDay
	End Sub

	Private Sub txtHh_AfterUpdate()
		Call setHhLim
	End Sub

	Private Sub cmdReset_Click()
	   Call ResetForm
	End Sub
		
	Private Sub txtWptID_Change()
		Call enableEnter
		Call checkWtptID
		Call checkFirstVal
	End Sub

	Private Sub txtYear_AfterUpdate()
		Call yearCheck
		Call enableEnter
		Call checkMonthYear
	End Sub

	Private Sub txtOldStock_Change()
		Call enableEnter
		Call checkOldStock
	End Sub

	Private Sub txtWptID_AfterUpdate()
		  Call setLimWtpid
	End Sub
		
	 Private Sub txtMonth_Change()
		Call monthCheck
		Call enableEnter
		If Not Me.txtMonth.Value = "" Then
			If Not IsNumeric(Me.txtMonth.Value) Then
				MsgBox ("Month value must be numeric to continue")
				Me.txtMonth.Value = ""
			End If
		End If
	End Sub

	Private Sub txtHh_Change()
		Call enableEnter
		Call checkHhNum
	End Sub

	Private Sub txtJerDeliv_AfterUpdate()
		Call setJerLimit
	End Sub

	Private Sub txtOldStock_AfterUpdate()
	  Call SetLimOldStock
	End Sub

	Private Sub txtJerDeliv_Change()
		Call enableEnter
		Call checkJerDeliv
	End Sub
	Private Sub cmdCancel_Click()
		frmDataEntry.Hide
	End Sub

	Private Sub txtYear_Change()
		If Not Me.txtYear.Value = "" Then
			If Not IsNumeric(Me.txtYear.Value) Then
				MsgBox ("Year value must be numeric to continue")
				Me.txtYear.Value = ""
			End If
		End If
	End Sub
