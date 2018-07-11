<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function Z_MinRate()
	Z_MinRate = 0
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT MinWage FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		Z_MinRate = rsRate("MinWage")
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function
Function Z_InHouseRate()
	Z_InHouseRate = 0
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT inHouse FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		Z_InHouseRate = rsRate("inHouse")
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function


DIM tmpIntr(), tmpTown(), tmpIntrName(), tmpLang(), tmpClass(), tmpBill(), tmpAhrs(), tmpApp(), tmpInst(), tmpDept(), tmpAmt(), tmpFac(), tmpMonthYr(), tmpCtr(), tmpMonthYr2(), tmpMonthYr3()
DIM tmpMonthYr4(), tmpHrs(), tmpHHrs(), tmpMile(), tmpToll(), arrTS(), arrAuthor(), arrPage(), tmpTrain(), tmpIHTrain(), tmpbhrs(), arrBody(), tmpHrs2(), tmpHrs3(), tmpHrs4() , tmpHrs5(), tmpZip()
DIM tmpHrsHP(), tmpHrsHP2()
DIM tmpTrngIntr(), tmpTrng_Hrs()
Server.ScriptTimeout = 360000

%>
<%
tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
tmpdate = replace(date, "/", "") 
tmpTime = replace(FormatDateTime(time, 3), ":", "")
ctr = 13


RepCSV =  "Mileage" & tmpdate & "-" & tmpTime & ".csv" 
	tmpMonthYear = MonthName(Month(tmpReport(1))) & " - " & Year(tmpReport(1))
	strMSG = "Mileage report for the month of " & tmpMonthYear
	strHead = "<td class='tblgrn'>File Number</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Miles</td>" & vbCrlf & _
		"<td class='tblgrn'>Miles Amount</td>" & vbCrlf & _
		"<td class='tblgrn'>Receipts Amount</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf 
	CSVHead = "File Number,Last Name,First Name,Miles,Miles Amount,Receipts Amount,Total"
	strFY = Year(tmpReport(1))
	strMo =  Month(tmpReport(1))
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT FileNum, actmil, overmile, appDate, i.[index] as myIntrIndex, Toll, mileageproc, m.[mileageRate] " & _
			"FROM [Request_T] AS r " & _
			"INNER JOIN [Interpreter_T] AS i ON r.[IntrID] = i.[Index] " & _
			"INNER JOIN [MileageRate_T] AS m ON r.[index] = r.[index] " & _
			"WHERE r.[instID] <> 479 " & _
			"AND Month(appDate) = " & strMo & _
			"AND Year(appDate) = " & strFY & " "
	If Z_CZero(tmpReport(4)) > 0 Then
		sqlRep = sqlRep & "AND IntrID = " & tmpReport(4) & " "
		strMSG = strMSG & " for " & GetIntr(tmpReport(4)) & "."
	End If
	sqlRep = sqlRep & "AND LbconfirmToll = 1 AND mileageproc IS NULL ORDER BY [last name], [first name], appDate"

	rsRep.Open sqlRep, g_strCONN, 1, 3
	strPySrc = "383000"
	If Len(strFY) = 4 Then strFY = Right(strFY, 2)
	strFY = Z_CLng(strFY)
	If strMo > 6 Then strFY = strFY + 1
	fltMyRate = 0
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			'IntrName = rsRep("Last Name") & ", " & rsRep("First Name")
			strMile = Z_Czero(rsRep("actmil"))
			IntrID = rsRep("myIntrIndex")
			strTol = Z_Czero(rsRep("Toll"))
			fltMyRate = rsRep("mileageRate")
			lngIDx = SearchArraysHours(IntrID, tmpIntr)
			If lngIdx < 0 Then
				ReDim Preserve tmpIntr(x)
				ReDim Preserve tmpMile(x)
				ReDim Preserve tmpToll(x)
				
				tmpIntr(x) = IntrID
				tmpMile(x) = strMile
				tmpToll(x) = strTol
				x = x + 1
			Else	
				tmpMile(lngIdx) = tmpMile(lngIdx) + strMile
				tmpToll(lngIdx) = tmpToll(lngIdx) + strTol
			End If

			rsRep("mileageproc") = Date ' TODO: put this back in effect!
			rsRep.Update
			rsRep.MoveNext
		Loop
		y = 0
		Do Until y = x
			ReDim Preserve arrBody(y)
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			myMile = Z_Czero(tmpMile(y))
			myToll = Z_Czero(tmpToll(y))
	
			mileTOT = MileageAmt(myMile) + myToll
			mileAmt = MileageAmt(myMile)
			arrBody(y) = "<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center' style='width: 80%;'>" & vbCrLf & _
				"<tr><td align='center' colspan='14'>ASCENTRIA CARE ALLIANCE</td></tr>" & vbCrLf & _
				"<tr><td align='center' colspan='14'>Monthly Staff Expense Report</td></tr>" & vbCrLf & _
				"<tr><td colspan='14'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr><td align='right' colspan='2'>Name:</td><td align='left'><nobr><b><u>" & GetIntr(tmpIntr(y)) & "</u></b></td></tr>" & vbCrLf & _
				"<tr><td align='right' colspan='2'>Address:</td><td align='left'><nobr><b><u>" & Z_IntrAdr(tmpIntr(y)) & "</u></b></td></tr>" & vbCrLf & _
				"<tr><td align='right' colspan='2'>City:</td><td align='left'><nobr><b><u>" & Z_IntrCity(tmpIntr(y)) & "</u></b></td></tr>" & vbCrLf & _
				"<tr><td align='right' colspan='2'>Zip:</td><td align='left'><nobr><b><u>" & Z_IntrZip(tmpIntr(y)) & "</u></b></td></tr>" & vbCrLf & _
				"<tr><td align='right' colspan='2'>Month:</td><td align='left'><nobr><b><u>" & Month(tmpReport(1)) & "/" & Year(tmpReport(1)) & "</u></b></td></tr>" & vbCrLf & _
				"<tr><td colspan='6'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr><td align='center' colspan='14'>" & vbCrLf & _
					"<table bgColor='white' border='0' style='width: 80%;' cellSpacing='0' cellPadding='0' align='center'>" & vbCrLf & _
						"<tr><td>Non Client Miles:</td><td align='left'>" & Z_FormatNumber(myMile, 2) & "</td><td align='center'>&nbsp;&nbsp;X&nbsp;&nbsp;</td><td> " & Z_FormatNumber(MileageRate(), 2) & "</td><td align='left'>$" & Z_FormatNumber(mileAmt, 2) & "</td></tr>" & vbCrLf & _
						"<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td align='left'>Parking/Tolls:</td><td align='left'>$" & Z_FormatNumber(myToll, 2) & "</td></tr>" & vbCrLf & _
						"<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td align='left'>Fares:</td><td align='left'>$0.00</td></tr>" & vbCrLf & _
						"<tr><td>&nbsp;</td><td align='left' colspan='3'>Total Travel Expense <b>(Feeds to Travel line below)</b>:</td><td align='left'>$" & Z_FormatNumber(mileTOT, 2) & "</td></tr>" & vbCrLf & _
					"</table>" & vbCrLf & _					
				"</td></tr>" & vbCrLf & _	
				"<tr><td colspan='14'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr><td colspan='14'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Entity</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Region</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Site</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Department</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Service</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid; width: 200px;'><b>G/L</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid; width: 300px;'><b>Fiscal Yr</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Payer Source</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Person Code</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Position Code</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>FAS 117</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>TBD2</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'><b>Amount</b></td>" & vbCrLf & _	
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Travel - Mileage</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>65010</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strFY & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strPySrc & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _	
				"</tr>" & vbCrLf & _
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Travel - Meals</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>65020</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strFY & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strPySrc & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Travel - Lodging</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>65000</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strFY & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strPySrc & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Travel - Other</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>65020</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strFY & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strPySrc & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Postage &amp; Delivery</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>63190</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strFY & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strPySrc & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Telephone</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>63210</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strFY & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strPySrc & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Seminars &amp; Workshops</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>63220</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strFY & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strPySrc & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Program Supplies</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>60350</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strFY & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strPySrc & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>Office Supplies</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>14</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH02</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>NH04</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>880</td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>8100</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>63170</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strFY & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>" & strPySrc & "</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>99999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>1</td>" & vbCrLf & _		
					"<td align='center' style='border: 1px solid;'>999999</td>" & vbCrLf & _
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _
				"</tr>" & vbCrLf & _	
				"<tr>" & vbCrLf & _	
					"<td align='right' style='border: 1px solid;' colspan='13'><b>Total</b></td>" & vbCrLf & _	
					"<td align='center' style='border: 1px solid;'>&nbsp;</td>" & vbCrLf & _	
				"</tr>" & vbCrLf & _	
				"<tr><td align='right' colspan='2'><nobr><b>Payment is hearby requested by:</b></td></tr>" & vbCrLf & _
				"<tr><td colspan='6'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr>" & vbCrLf & _	
					"<td align='right' colspan='2'>Signature:</td>" & vbCrLf & _
					"<td align='left' colspan='4'><nobr><b><u><i>ELECTRONICALLY APPROVED</i></u></b>&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf & _
					"Date:<u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></td>" & vbCrLf & _
				"</tr>" & vbCrLf & _
				"<tr><td colspan='6'>&nbsp;</td></tr>" & vbCrLf & _
				"<tr>" & vbCrLf & _	
					"<td align='right' colspan='2'>Approved By:</td>" & vbCrLf & _
					"<td align='left' colspan='4'><nobr><b><u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></b></u></td>" & vbCrLf & _
				"</tr>" & vbCrLf & _
				"<tr><td colspan='6'>&nbsp;</td></tr>" & vbCrLf & _
				"</table><div style=""page-break-after: always;""><br></div>" & vbCrLf
				
				
		'	strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetFileNum(tmpIntr(y)) & "</td>" & vbCrLf & _
		'		"<td class='tblgrn2'><nobr>" & GetIntr(tmpIntr(y)) & "</td>" & vbCrLf & _
		'		"<td class='tblgrn2'><nobr>" & Z_FormatNumber(myMile,2) & "</td>" & vbCrLf & _
		'		"<td class='tblgrn2'><nobr>$" & Z_FormatNumber(AmtRate(myMile),2) & "</td>" & vbCrLf & _
		'		"<td class='tblgrn2'><nobr>$" & Z_FormatNumber(myToll,2) & "</td>" & vbCrLf & _
		'		"<td class='tblgrn2'><nobr>$" & Z_FormatNumber(mileTOT,2) & "</td></tr>" & vbCrLf
								
			CSVBody = CSVBody & GetFileNum(tmpIntr(y)) & "," & GetIntr(tmpIntr(y)) & "," & Z_FormatNumber(myMile,2) & "," & _
				Z_FormatNumber( fltMyRate , 2) & "," & Z_FormatNumber(myToll,2) & "," & Z_FormatNumber(mileTOT,2) & vbCrLf
			y = y + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	
	rsRep.Close
	Set rsRep = Nothing
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
	strLog = Now & vbTab & "Mileage ran by " & Session("UsrName") & "."
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	

	Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
	Prt.WriteLine "LANGUAGE BANK - REPORT"
	Prt.WriteLine strMSG
	Prt.WriteLine CSVHead
	Prt.WriteLine CSVBody
	Prt.Close	
	Set Prt = Nothing
	
	tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV)
	Set fso = Nothing
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
	<title>Mileage Report Result</title>
	<link href='style.css' type='text/css' rel='stylesheet'>
	<script language='JavaScript'>
function exportMe()
{
	document.frmResult.action = "printreport.asp?csv=1"
	document.frmResult.submit();
}
function PassMe(xxx)
{
	window.opener.document.frmReport.hideID.value = xxx;
	window.opener.SubmitAko();
	self.close();
}
	</script>
</head>
<body>
<div style="width: 300px; text-align: center; margin: 0px auto 20px;" >
	<img src='images/LBISLOGO.jpg' align='center' style="width: 287px; height: 67px;" />
	340 Granite Street 3<sup>rd</sup> Floor, Manchester, NH 03102<br />
	Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
</div>
		<form method='post' name='frmResult'>
			<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>
				<tr><td valign='top' >
							<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
								<tr style="background-color: #f58426;">
									<td colspan='<%=ctr + 7%>' align='center'>
<b><%=strMSG%></b>
									</td>
								</tr>
							</table>
<%=strBody%>
								<tr><td>&nbsp;</td></tr>
							</table>
<div style="page-break-after: always;">&nbsp;</div>
<div>
<%
arrCtr = 0
If IsArray(arrBody) And y > 0 Then
	'Response.Write "Limit: " & LBound(arrBody) & "<br />" & vbCrLf
	Do Until arrCtr = UBound(arrBody) + 1
		Response.Write arrBody(arrCtr)
		arrCtr = arrCtr + 1
	Loop
'Else
'	Response.Write strBody
'	should output 'nothing found' here sometime.
End If
%>
</div>
						<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
								<tr><td colspan='<%=ctr + 4%>' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
										<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'"
												onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring%>';">
									</td>
								</tr>
								<tr><td colspan='<%=ctr + 4%>' align='center' height='100px' valign='bottom'>
										* If needed, please adjust the page orientation of your printer to landscape to view all columns in a single page   
									</td>
								</tr>
							</table>	
						</td>
					</tr>
				</table>
				<p align="center">
					<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
				</p>

		</form>
	</body>
</html>
