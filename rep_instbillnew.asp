<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
DIM ts0, ts1
ts0 = Now

Function Z_YMDDate(dtDate)
DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	lngTmp = Z_CLng(DatePart("m", dtDate))
	If lngTmp < 10 Then Z_YMDDate = "0"
	Z_YMDDate = Z_YMDDate & lngTmp & "-"
	lngTmp = Z_CLng(DatePart("d", dtDate))
	If lngTmp < 10 Then Z_YMDDate = Z_YMDDate & "0"
	Z_YMDDate = Z_YMDDate & lngTmp
	strTmp = DatePart("yyyy", dtDate)
	Z_YMDDate = strTmp & "-" & Z_YMDDate
End Function

DIM tmpReport, tmpDate, tmpTime, RepCSV, strMSG, CSVHead, strHead, rsInst, strIDT, sqlDT, sqlInst, x, y, z, kulay
Server.ScriptTimeout = 360000


tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
tmpdate = replace(date, "/", "") 
tmpTime = replace(FormatDateTime(time, 3), ":", "")


'INSTITUTION BILLING
RepCSV =  "InstBillReq" & tmpdate & "-" & tmpTime & ".csv" 
RepCSVBill = "InstBillReqNew" & tmpdate & "-" & tmpTime & ".csv" 
RepCSVBillL = "InstBillReqNewL" & tmpdate & "-" & tmpTime & ".csv"
RepCSVBillSigma = "InstBillReqNewSigma" & tmpdate & "-" & tmpTime & ".csv"
RepCSVBillCourts = "InstBillReqCourts"	& tmpdate & "-" & tmpTime & ".csv"
	
Set rsRep = Server.CreateObject("ADODB.RecordSet")
strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _ 
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf & _
		"<td class='tblgrn'>Comment</td>" & vbCrlf & _
		"<td class='tblgrn'>DOB</td>" & vbCrlf & _
		"<td class='tblgrn'>Customer ID</td>" & vbCrlf & _
		"<td class='tblgrn'>DHHS</td>" & vbCrlf 

' EMERGENCY RATE
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT * FROM EmergencyFee_T"
rsRate.Open sqlRate, g_strCONN, 3,1
If Not rsRate.EOF Then
	tmpFeeL = rsRate("FeeLegal")
	tmpFeeO = rsRate("FeeOther")
End If
rsRate.Close
Set rsRate = Nothing
ctr = 11	
CSVHead = "Request ID,Institution, Department, Appointment Date, Client Last Name, Client First Name" & _
		", Language, Interpreter Last Name, Interpreter First Name, Appointment Start Time" & _
		", Appointment End Time, Hours, Rate, Travel Time, Mileage, Emergency Surcharge, Total" & _
		", Comments, DOB, Customer ID, DHHS, Requesting Person, User"
	
sqlRep = "SELECT claimant, judge, meridian, nhhealth, wellsense" & _
		", billingTrail, ReqID, HPID, syscom, processed, hasmed" & _
		", outpatient, medicaid, vermed, autoacc, wcomp, drg, pid" & _
		", RTRIM(COALESCE([cfname], '')) AS [cfname]" & _
		", RTRIM(COALESCE([clname], '')) AS [clname]" & _
		", COALESCE([docnum], '') AS [docnum]" & _
		", r.[index] AS myindex, r.InstID AS myinstID, [status]" & _
		", AStarttime, AEndtime, Billable, DOB, emerFEE" & _
		", [class], TT_Inst, M_Inst, DeptID, LangID, appDate, InstRate" & _
		", bilComment, custID, ccode, billgroup, IntrID, appTimeFrom, appTimeTo" & _
		", l.[Language], d.[Dept]" & _
		", d.distcode, i.[Last Name], i.[First Name] " & _
		"FROM [request_T] AS r " & _
		"INNER JOIN [interpreter_T] AS i ON r.[IntrID]=i.[index] " & _
		"INNER JOIN [dept_T] AS d ON r.[DeptID]=d.[index] " & _
		"INNER JOIN [language_T] AS l ON r.[LangID]=l.[index] " & _
		"WHERE r.[instID] <> 479  " & _
		"AND (Status = 1 OR Status = 4)  " & _
		"AND Processed IS NULL " & _
		"AND ProcessedMedicaid IS NULL "
' add date filter
'ORDER BY CustID ASC, AppDate DESC

strMSG = "Institution Billing request report"
If tmpReport(1) <> "" Then
	sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
	strMSG = strMSG & " from " & tmpReport(1)
End If
If tmpReport(2) <> "" Then
	sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
	strMSG = strMSG & " to " & tmpReport(2)
End If
strMSG = strMSG & ". * - Cancelled Billable."
If tmpReport(9) = "" Then tmpReport(9) = 0
If tmpReport(9) <> 0 Then
	If tmpReport(6) = "" Then tmpReport(6) = 0
	If tmpReport(6) <> 0 Then 
		sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
	End If
	If tmpReport(7) = "" Then tmpReport(7) = 0
	If tmpReport(7) <> "0" Then
		tmpCli = Split(tmpReport(7), ",")
		sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
	End If
	If tmpReport(8) = "" Then tmpReport(8) = 0
	If tmpReport(8) <> 0 Then 
		sqlRep = sqlRep & " AND [Class] = " & tmpReport(8)
	End If
End If
sqlRep = sqlRep & " ORDER BY CustID ASC, AppDate DESC" '" ORDER BY AppDate DESC"
' Response.Write "<pre>" & sqlRep & "</pre>"
' Response.End
rsRep.Open sqlRep, g_strCONN, 1, 3

If Not rsRep.EOF Then 
	x = 0
	tmpCID = ""
	Do Until rsRep.EOF
		BillHours = 0
		IncludeReq = True
		docnum = ""
		kulay = "#FFFFFF"
		If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
		CB = ""
		If rsRep("status") = 4 Then CB = "*"
		strIntrName = rsRep("Last Name") & ",  " & rsRep("First Name")
		strCliName =  rsRep("Clname") & ", " & rsRep("Cfname")
		strATime =  cTime(rsRep("AStarttime")) & " -  " & cTime(rsRep("AEndtime"))
		'totHrs =  DateDiff("n", CDate(rsRep("AStarttime")) , CDate(rsRep("AEndtime")))
		BillHours =  rsRep("Billable")
		'check if previously drg
		If Not rsRep("drg") Then
			deptdrg = False
			If Z_PrevDRG(rsRep("DeptID")) Then deptdrg = True
		Else
			deptdrg = True
		End If
		PayMed = PaytoMedicaid(rsRep("outpatient"), rsRep("hasmed"), rsRep("vermed") _
				, rsRep("autoacc"), rsRep("wcomp"), deptdrg, rsRep("IntrID") _
				, rsRep("medicaid"), rsRep("meridian"), rsRep("nhhealth"), rsRep("wellsense"))
		If PayMed AND BillHours <= 4 Then IncludeReq = False
		If PayMed Then IncludeReq = False
		If PayMed And rsRep("vermed") = 0 Then IncludeReq = False
		If IncludeReq Then
			If rsRep("emerFEE") = True Then
				If rsRep("class") = 3 Or rsRep("class") = 5 Then
					tmpPay = (BillHours * tmpFeeL) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
				ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
					tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst")) + tmpFeeO
				End If
			Else
				tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			End If
			totalPay = Z_FormatNumber(tmpPay, 2)
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
					"<td class='tblgrn2'><nobr>" & CB & rsRep("myindex") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetInst2(rsRep("myinstID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Replace(rsRep("Dept"), " - ", "") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & strCliName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("Language") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & strIntrName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & strATime & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
			If rsRep("emerFEE") = True Then 
				If rsRep("class") = 3 Or rsRep("class") = 5 Then
					strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeL & "</td>" & vbCrLf
				Else
					strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
				End If
			Else
				strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
			End If
			strBody = strBody & "<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf 
			If rsRep("emerFEE") = True Then 
				If rsRep("class") = 3 Or rsRep("class") = 5 Then
					strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
				ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
					strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeO & "</td>" & vbCrLf
				End If
			Else
				strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
			End If
			bilcomment = Z_fixNull(rsRep("bilComment") & rsRep("syscom") & rsRep("billingTrail"))
			strBody = strBody & "<td class='tblgrn2'><nobr><b>$" & totalPay & "</b></td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & bilcomment & "</td><td class='tblgrn2'><nobr>" & _
					rsRep("DOB") & "</td>" & vbCrLf & "<td class='tblgrn2'><nobr>" & rsRep("CustID") & "</td>"
			If rsRep("myinstID") = 108 Then
				strBody = strBody & "<td class='tblgrn2'><nobr>" & GetUserID(rsRep("DeptID")) & "</td><tr>" & vbCrLf 
			Else
				strBody = strBody & "<td class='tblgrn2'><nobr>&nbsp;</td><tr>" & vbCrLf 
			End If
			
			CSVBodyLine = CB & rsRep("myindex") & "," & GetInst2(rsRep("myinstID")) & "," & _
					Replace(rsRep("Dept"), " - ", "") & "," & rsRep("appDate") & _
					"," & rsRep("Clname") & "," & rsRep("Cfname") &  "," & _
					rsRep("Language") & "," & rsRep("Last Name") & _
					"," & rsRep("First Name") & ","  & cTime(rsRep("AStarttime")) & _
					"," & cTime(rsRep("AEndtime")) & "," & BillHours
					
			'''''MIP IMPORT'''''
			If rsRep("myinstID") <> 108 And rsRep("myinstID") <> 30 Then 'exclude DHHS and LSS
				'If rsRep("class") <> 3 Then
				tmpCID = Trim(UCase(rsRep("custID")))
				If (rsRep("langID")=52 Or rsRep("langID")=109 Or rsRep("langID")=81) Then
					tmpccode = "LB 70 Rate ASL"
				Else
					tmpccode = "LB " & rsRep("InstRate") & " Rate"
				End If
				If tmpCID <> tmpCID2 Or tmpCID2 = "" Then
					If rsRep("class") <> 5 Then
						If rsRep("myInstID") = 240 Then
							CSVBodyBillSigma = CSVBodyBillSigma & """" & "HOTC" & """,""" & _
									rsRep("billgroup") & """,""" & rsRep("custID") & """" & vbCrLf
						Else
							CSVBodyBill = CSVBodyBill & """" & "HOTC" & """,""" & _
									rsRep("billgroup") & """,""" & rsRep("custID") & """" & vbCrLf
						End If
					Else
						CSVBodyBillL = CSVBodyBillL & """" & "HOTC" & """,""" & rsRep("billgroup") & """,""" & rsRep("custID") & """" & vbCrLf
					End If
				End If	
				If rsRep("class") <> 5 Then
					If rsRep("myinstID") = 19 Or rsRep("myinstID") = 22 Or rsRep("myinstID") = 27 _
							Or rsRep("myinstID") = 229 Or rsRep("myinstID") = 265 Or rsRep("myinstID") = 268 _
							Or rsRep("myinstID") = 269 Or rsRep("myinstID") = 273 Or rsRep("myinstID") = 289 _
							Or rsRep("myinstID") = 300 Or rsRep("myinstID") = 308 Or rsRep("myinstID") = 398 _
							Or rsRep("myinstID") = 427 Or rsRep("myinstID") = 431 Then
						'DHMC and Elliot
						CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""0"",""" & tmpccode & """,""" & _
								rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - " & _
								Replace(rsRep("Dept"), " - ", "") & """,""" & date & """,""" & _
								rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
					ElseIf rsRep("myinstID") = 15 Or rsRep("myinstID") = 33 Or rsRep("myinstID") = 41 _
							Or rsRep("myinstID") =  70 Or rsRep("DeptID") = 645 Then
						'concord hosp, fam and mchc and com council nashua, riverbend
						CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""0"",""" & tmpccode & """,""" & _
								rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - " & _
								rsRep("Language") & """,""" & date & """,""" & _
								rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
					ElseIf rsRep("deptID") = 1058 Then 'man welfare
						CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""0"",""" & tmpccode & """,""" & _
								rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - MW" & _
								rsRep("myindex") & """,""" & date & """,""" & _
								rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
					ElseIf rsRep("myinstID") = 39 Or rsRep("myinstID") = 52 Or rsRep("myinstID") = 168 _
							Or rsRep("myinstID") = 199 Or rsRep("myinstID") = 724 _
							Or rsRep("myinstID") = 683 Or rsRep("myinstID") = 717 Then
					' [39] SAU # 37 Manchester School District,
					' [52] Exeter Hospital,
					' [168] Concord Head Start,
					' [199] Southern New Hampshire Services,
					' [724] City of Worcester 
					' added 180716:
					' 		[683] Core Physicians with Exeter Hospital
					' 		[717] Frisbie Memorial Hospital
						reqp = GetReq(rsRep("reqID"))
						If Z_Czero(rsRep("HPID")) > 0 Then reqp = reqp & " / " & GetReqHPID(rsRep("HPID"))
						CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""0"",""" & tmpccode & """,""" & _
								rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & _
								" (Req: " & reqp & ")"",""" & date & """,""" & _
								rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
					ElseIf rsRep("myinstID") = 860 Then 'umass med
						CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""0"",""" & tmpccode & """,""" & _
								rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - " & _
								rsRep("Language") & " - " & Replace(rsRep("DeptID"), " - ", "") & """,""" & date & """,""" & _
								rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
					Else
						If rsRep("myInstID") = 240 Then
							CSVBodyBillSigma = CSVBodyBillSigma & """" & "DOTC" & """,""0"",""" & tmpccode & """,""" & _
									rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " – Interpretation" & _
									""",""" & date & """,""" & rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
						Else 
							If rsRep("custID") = "New England Heart In" Then
								CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""0"",""" & tmpccode & """,""" & _
										rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " " & _
										rsRep("DOB") & """,""" & date & """,""" & _
										rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
							Else
								If rsRep("class") = 3 Then
									If rsRep("emerFEE") = True Then
										tmpccode = "LB 60 Rate"

										If (rsRep("langID")=52 Or rsRep("langID")=109 Or rsRep("langID")=81) Then tmpccode = "LB 70 Rate ASL"

									End if
								End If
								If rsRep("custID") = "UMass Community Serv" Or rsRep("custID") = "Seven Hills Found" Or _
										rsRep("custID") = "Saint Vincent Hosp" Or rsRep("custID") = "Spear Management Grp" _
										Or rsRep("custID") = "BayPath Elder Svc" Then
									CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""0"",""" & tmpccode & """,""" & _
											rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & _
											" – Interpretation" &  """,""" & date & """,""" & _
											"langbankma" & """,""" & BillHours & """" & vbCrLf
								Else
									CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""0"",""" & tmpccode & """,""" & _
											rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & _
											" – Interpretation" &  """,""" & date & """,""" & _
											rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
								End If
							End If
						End If
					End If
				Else
					If rsRep("emerFEE") = True Then tmpccode = "LB 60 Rate"
						docnum = ""
						If Z_Fixnull(rsRep("docnum")) <> "" Then docnum = " - " & rsrep("docnum")
						If rsRep("myinstID") = 757 Or rsRep("myinstID") = 777 Then 'SSA
							apptadr = Z_GetApptAddr(rsRep("myindex"))
							CSVBodyBillL = CSVBodyBillL & """" & "DOTC" & """,""" & _
									"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & _
									" " & rsRep("Clname") & docnum & " – Interpretation" & " - " & rsRep("judge") & _
									" - " & rsRep("claimant") & " - " & rsRep("appdate") & " - " & CTime(rsRep("appTimeFrom")) & _
									" - " & CTime(rsRep("appTimeto")) & " - " & rsRep("Language") & _
									" - " & GetDept(rsRep("deptID")) & " - " & apptadr & " - " & CTime(rsRep("AStarttime")) & _
									" - " & CTime(rsRep("AEndtime")) & """,""" & date & """,""" & _
									rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
						ElseIf rsRep("myinstID") = 126 Then 'NH legal
							CSVBodyBillL = CSVBodyBillL & """" & "DOTC" & """,""" & _
								"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & docnum & " – Interpretation (" & strIntrName & ")" & """,""" & date & """,""" & _
								rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
						Else
							CSVBodyBillL = CSVBodyBillL & """" & "DOTC" & """,""" & _
								"0" & """,""" & tmpccode & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & docnum & " – Interpretation" & """,""" & date & """,""" & _
								rsRep("distcode") & """,""" & BillHours & """" & vbCrLf
						End If
					End If
						
					If Z_CZero(rsRep("TT_Inst")) <> 0 Then 'new billing rules for Travel
						strTT = rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname")
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & MyTravelTime(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("TT_Inst"), strTT, rsRep("distcode")) & """" & vbCrLf
							Else
								If rsRep("myInstID") = 273 Or _
									ClassInt(rsRep("DeptID")) = 3 Then 'dart and court
										CSVBodyBill = CSVBodyBill & """" & MyTravelTimeDarthCourt(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("TT_Inst"), strTT, rsRep("distcode")) & """" & vbCrLf
								Else
									CSVBodyBill = CSVBodyBill & """" & MyTravelTime(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("TT_Inst"), strTT, rsRep("distcode")) & """" & vbCrLf
								End If
							End If
						Else
							CSVBodyBillL = CSVBodyBillL & """" & MyTravelTime(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("TT_Inst"), strTT, rsRep("distcode")) & """" & vbCrLf
						End If
					End If
					
					If Z_CZero(rsRep("M_Inst")) <> 0 Then 'new billing rules for Mileage
						strM = rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname")
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & MyMileages(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("M_Inst"), strM, rsRep("distcode")) & """" & vbCrLf
							Else
								CSVBodyBill = CSVBodyBill & """" & MyMileages(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("M_Inst"), strM, rsRep("distcode")) & """" & vbCrLf
							End If
						Else
							CSVBodyBillL = CSVBodyBillL & """" & MyMileages(rsRep("IntrID"), rsRep("LangID"), rsRep("myinstID"), rsRep("DeptID"), rsRep("M_Inst"), strM, rsRep("distcode")) & """" & vbCrLf
						End If
					End If
					
					If rsRep("emerFEE") = True Then 
						If rsRep("class") <> 5 Then
							If rsRep("myInstID") = 240 Then
								CSVBodyBillSigma = CSVBodyBillSigma & """" & "DOTC" & """,""" & _
									"0" & """,""" & "LB Emer Fee" & """,""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & " - Emer Fee" & """,""" & date & """,""" & _
									rsRep("distcode") & """,""" & "1" & """" & vbCrLf
							Else 
								If rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
									If rsRep("custID") = "UMass Community Serv" Or rsRep("custID") = "Seven Hills Found" _
											Or rsRep("custID") = "Saint Vincent Hosp" Or rsRep("custID") = "Spear Management Grp" _
											Or rsRep("custID") = "BayPath Elder Svc" Then
										CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""0"",""" & "LB Emer Fee" & _
												""",""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & _
												" - Emer Fee" & """,""" & date & """,""" & _
												rsRep("distcode") & """,""" & "1" & """" & vbCrLf
									Else
										CSVBodyBill = CSVBodyBill & """" & "DOTC" & """,""0"",""" & "LB Emer Fee" & _
												""",""" & rsRep("appDate") & " " & rsRep("Cfname") & " " & rsRep("Clname") & _
												" - Emer Fee" & """,""" & date & """,""" & _
												rsRep("distcode") & """,""" & "1" & """" & vbCrLf
									End If
								End If
							End If
						End If
					End If
					tmpCID2 = tmpCID
				End If
				'''''''''''''''''''
				If rsRep("emerFEE") = True Then 
					If rsRep("class") = 3 Or rsRep("class") = 5 Then
						CSVBodyLine = CSVBodyLine & "," & tmpFeeL
					Else
						CSVBodyLine = CSVBodyLine & "," & rsRep("InstRate")
					End If
				Else
					CSVBodyLine = CSVBodyLine & "," & rsRep("InstRate")
				end if
				
				CSVBodyLine = CSVBodyLine & ",""" & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & ""","
				
				If rsRep("emerFEE") = True Then 
					If rsRep("class") = 3 Or rsRep("class") = 5 Then
						CSVBodyLine = CSVBodyLine & "0.00"
					ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
						CSVBodyLine = CSVBodyLine & tmpFeeO
					End If
				Else
					CSVBodyLine = CSVBodyLine & "0.00"
				end if
				bilcommentcsv = Replace(Z_fixNull(rsRep("bilComment") & rsRep("syscom") & rsRep("billingTrail")), "<br>", " / ")
				CSVBodyLine = CSVBodyLine & ",""" & totalPay & """,""" & bilcommentcsv & """,""" & rsRep("DOB") & """,""" & rsRep("CustID")
				If rsRep("myinstID") = 108 Then
					CSVBodyLine = CSVBodyLine & """,""" & GetUserID(rsRep("DeptID")) 
				Else
					CSVBodyLine = CSVBodyLine & """,""" & ""  
				End If
				If rsRep("class") = 3 Then
					If Z_CZero(rsRep("HPID")) > 0 Then 
						reqname = GetReqHPID(rsRep("HPID"))
					Else
						reqname = GetReq(rsRep("reqID"))
					End If
					CSVBodyCourt = CSVBodyCourt & CSVBodyLine & """,""" & reqname & """,""" & Z_GetLoginHP(Z_GetUIDHP(rsRep("HPID"))) & """" & vbCrLf
				End If
				CSVBody = CSVBody & CSVBodyLine & """" & vbCrLf
			
				'TODO: reinstate the next 2 timestamps for live
				'rsRep("billingTrail") = rsRep("billingTrail") & "<br>Billed to Institution " & Date
				'rsRep("Processed") = Date

				x = x + 1
				rsRep.Update
			End If
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='13' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
	strLog = Now & vbTab & "Billing ran by " & Session("UsrName") & "."
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing

tmpBills = Request("Bill")
If Request("csv") <> 1 Then
	'CONVERT TO CSV
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
	Prt.WriteLine "LANGUAGE BANK - REPORT"
	Prt.WriteLine strMSG
	Prt.WriteLine CSVHead
	Prt.WriteLine CSVBody
	Prt.Close	
	Set Prt = Nothing

If Z_CZero(tmpReport(0)) = 3 Or Z_CZero(tmpReport(0)) = 16 Then 'additional csv for billing
		Set Prt2 = fso.CreateTextFile(RepPath &  RepCSVBill, True)
		'Prt.WriteLine "LANGUAGE BANK - REPORT"
		'Prt.WriteLine strMSG
		Prt2.WriteLine CSVBodyBill
		Prt2.Close	
		Set Prt2 = Nothing
		fso.CopyFile RepPath & RepCSVBill, BackupStr
		
		
		Set Prt2 = fso.CreateTextFile(RepPath &  RepCSVBillL, True)
		'Prt.WriteLine "LANGUAGE BANK - REPORT"
		'Prt.WriteLine strMSG
		Prt2.WriteLine CSVBodyBillL
		Prt2.Close	
		Set Prt2 = Nothing
		fso.CopyFile RepPath & RepCSVBillL, BackupStr
		
		Set Prt3 = fso.CreateTextFile(RepPath &  RepCSVBillSigma, True)
		'Prt.WriteLine "LANGUAGE BANK - REPORT"
		'Prt.WriteLine strMSG
		Prt3.WriteLine CSVBodyBillSigma
		Prt3.Close	
		Set Prt3 = Nothing
		fso.CopyFile RepPath & RepCSVBillSigma, BackupStr
		
		Set Prt = fso.CreateTextFile(RepPath &  RepCSVBillCourts, True)
		Prt.WriteLine "LANGUAGE BANK - REPORT"
		Prt.WriteLine strMSG
		Prt.WriteLine CSVHead
		Prt.WriteLine CSVBodyCourt
		Prt.Close	
		Set Prt = Nothing
		fso.CopyFile RepPath & RepCSVBillCourts, BackupStr
	End If
	If Z_CZero(tmpReport(0)) = 39  Or Z_CZero(tmpReport(0)) = 40 Then 'medicaid CSV
		Set Prt2 = fso.CreateTextFile(RepPath &  RepCSVBill, True)
		'Prt.WriteLine "LANGUAGE BANK - REPORT"
		Prt2.WriteLine CSVHeadBill
		Prt2.Write CSVBodyBill
		Prt2.Close	
		Set Prt2 = Nothing
		fso.CopyFile RepPath & RepCSVBill, BackupStr
		
		Set Prt3 = fso.CreateTextFile(RepPath &  RepCSVBillMeri, True)
		'Prt.WriteLine "LANGUAGE BANK - REPORT"
		Prt3.WriteLine CSVHeadBill
		Prt3.WriteLine CSVBodyBillMer
		Prt3.Close	
		Set Prt3 = Nothing
		fso.CopyFile RepPath & RepCSVBillMeri, BackupStr
		
		'NEW CSV
		Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillLB, True) 'LB
		If Z_CZero(tmpReport(0)) = 39 Then
			Prt2.WriteLine "LANGUAGE BANK - REPORT"
			Prt2.WriteLine strMSG
			Prt2.WriteLine CSVHead
		Else
			Prt2.WriteLine CSVHeadBill
		End If
		Prt2.WriteLine csvBodyLB
		Prt2.Close	
		Set Prt2 = Nothing
		fso.CopyFile RepPath & RepCSVBillLB, BackupStr
		
		Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillMHP, True) 'MHP
		If Z_CZero(tmpReport(0)) = 39 Then
			Prt2.WriteLine "LANGUAGE BANK - REPORT"
			Prt2.WriteLine strMSG
			Prt2.WriteLine CSVHead
		Else
			Prt2.WriteLine CSVHeadBill
		End If
		Prt2.WriteLine csvBodyMHP
		Prt2.Close	
		Set Prt2 = Nothing
		fso.CopyFile RepPath & RepCSVBillMHP, BackupStr
		
		Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillNHHF, True) 'NHHF
		If Z_CZero(tmpReport(0)) = 39 Then
			Prt2.WriteLine "LANGUAGE BANK - REPORT"
			Prt2.WriteLine CSVHead
		Else
			Prt2.WriteLine CSVHeadBill
		End If
		Prt2.WriteLine csvBodyNHHF
		Prt2.Close	
		Set Prt2 = Nothing
		fso.CopyFile RepPath & RepCSVBillNHHF, BackupStr
		
		Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillWSHP, True) 'WSHP
		If Z_CZero(tmpReport(0)) = 39 Then
			Prt2.WriteLine "LANGUAGE BANK - REPORT"
			Prt2.WriteLine strMSG
			Prt2.WriteLine CSVHead
		Else
			Prt2.WriteLine CSVHeadBill
		End If
		Prt2.WriteLine csvBodyWSHP
		Prt2.Close	
		Set Prt2 = Nothing
		fso.CopyFile RepPath & RepCSVBillWSHP, BackupStr
	End If

	tmpstring = "CSV/" & repCSV 'add for RepCSVBill
	tmpstring2 = "CSV/" & RepCSVBill
	tmpstring3 = "CSV/" & RepCSVBillL
	tmpstring4 = "CSV/" & RepCSVBillSigma
	tmpstring5 = "CSV/" & RepCSVBillMeri
	tmpstringMED = "CSV/" & RepCSVBillLB
	tmpstringMHP = "CSV/" & RepCSVBillMHP
	tmpstringNHHF = "CSV/" & RepCSVBillNHHF
	tmpstringWSHP = "CSV/" & RepCSVBillWSHP
	tmpstringcourts = "CSV/" & RepCSVBillCourts

	' corrections!
	tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV)
	tmpstring2 = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBill)
	tmpstring3 = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillL)
	tmpstring4 = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillSigma)
	tmpstring5 = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillMeri)
	tmpstringMED = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillLB)
	tmpstringMHP = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillMHP)
	tmpstringNHHF = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillNHHF)
	tmpstringWSHP = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillWSHP)
	tmpstringcourts = "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillCourts)

	'COPY FILE TO BACKUP
	fso.CopyFile RepPath & RepCSV, BackupStr
	
	Set fso = Nothing
	'EXPORT CSV
	'If Request("bill") <> 1 Then
	tmpstring = "CSV/" & repCSV 'add for RepCSVBill
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
	<title>Language Bank - Report Result</title>
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
<div style="width: 300px; text-align: center; margin: 0px auto 20px;">
	<img src='images/LBISLOGO.jpg' align='center' style="width: 287px; height: 67px;" />
	340 Granite Street 3<sup>rd</sup> Floor, Manchester, NH 03102<br />
	Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
</div>
		<form method='post' name='frmResult'>
			<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>
				<tr><td valign='top' >
							<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
								<tr bgcolor='#f58426'>
									<td colspan='<%=ctr + 7%>' align='center'>
<b><%=strMSG%></b>
									</td>
								</tr>
<tr><%=strHead%></tr>
<%=strBody%>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='<%=ctr + 7%>' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
										<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'"
											onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring%>';">
										<input class='btn' type='button' value='LB Billing CSV' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring2%>';">
										<input class='btn' type='button' value='LB Legal CSV' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring3%>';">
										<input class='btn' type='button' value='LB Comp Sigma CSV' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring4%>';">
										<input class='btn' type='button' value='CSV Export Courts' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstringcourts%>';">
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
		</form>
<%
ts1 = Now
Response.Write "<div class=""debug"" >Time elapsed: " & DateDiff("s", ts0, ts1) & "s</div>"
%>
	</body>
</html>
