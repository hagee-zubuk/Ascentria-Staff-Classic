<!DOCTYPE html>
<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
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
Function Z_Time24(dtDate)
	If Not IsDate(dtDate) Then Exit Function
	lngTmp = DatePart("h", dtDate)
	If lngTmp < 10 Then Z_Time24 = Z_Time24 & "0"
	Z_Time24 = Z_Time24 & lngTmp & ":"
	lngTmp = DatePart("n", dtDate)
	If lngTmp < 10 Then Z_Time24 = Z_Time24 & "0"
	Z_Time24 = Z_Time24 & lngTmp
End Function
%>
<%
DIM tmpIntr(), tmpTown(), tmpIntrName(), tmpLang(), tmpClass(), tmpBill(), tmpAhrs(), tmpApp(), tmpInst(), tmpDept(), tmpAmt(), tmpFac(), tmpMonthYr(), tmpCtr(), tmpMonthYr2(), tmpMonthYr3()
DIM tmpMonthYr4(), tmpHrs(), tmpHHrs(), tmpMile(), tmpToll(), arrTS(), arrAuthor(), arrPage(), tmpTrain(), tmpIHTrain(), tmpbhrs(), arrBody(), tmpHrs2(), tmpHrs3(), tmpHrs4() , tmpHrs5(), tmpZip()
DIM tmpHrsHP(), tmpHrsHP2()
DIM blnWeaponized

blnWeaponized = CBool(Z_FixNull(Request("tag")) = "111")	' because just one '1' isn't committed enough?? LOL!
Server.scripttimeout = 3600000	' 1 hour!

tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
tmpdate = Replace(Date, "/", "") 
tmpTime = Replace(FormatDateTime(Time, 3), ":", "")
tmpOver =Z_CLng(Request("override"))
ctr = 10

tmpDate = Z_YMDDate(Date)
Set rsRep = Server.CreateObject("ADODB.RecordSet")

RepCSV 			= "ALLXBillReq" & tmpdate & "-" & tmpTime & ".csv" 
RepCSVBill 		= "MedicaidXBillReqNew" & tmpdate & "-" & tmpTime & ".csv" 
RepCSVBillMeri 	= "MeridianXBillReqNew" & tmpdate & "-" & tmpTime & ".csv" 
RepCSVBillLB 	= "LBXBill" & tmpdate & "-" & tmpTime & ".csv"
RepCSVBillMHP 	= "MHPXBill" & tmpdate & "-" & tmpTime & ".csv"
RepCSVBillNHHF 	= "NHHFXBill" & tmpdate & "-" & tmpTime & ".csv"
RepCSVBillWSHP 	= "WSHPXBill" & tmpdate & "-" & tmpTime & ".csv"
RepCSVBillAHC 	= "AHCXBill" & tmpdate & "-" & tmpTime & ".csv"
CSVHeadBill 	= """" & FixDateFormat(date) & "LB Billing" & """"
csvBodyLB 		= ""
csvBodyAHC 		= ""
csvBodyMHP 		= ""
csvBodyNHHF 	= ""
csvBodyWSHP 	= ""

strHead = "<th class='tblgrn'>Request ID</th>" & vbCrlf & _ 
		"<th class='tblgrn'>Institution</th>" & vbCrlf & _
		"<th class='tblgrn'>Department</th>" & vbCrlf & _
		"<th class='tblgrn'>Appointment Date</th>" & vbCrlf & _
		"<th class='tblgrn'>Client Name</th>" & vbCrlf & _
		"<th class='tblgrn'>Medicaid</th>" & vbCrlf & _
		"<th class='tblgrn'>MCO</th>" & vbCrlf & _
		"<th class='tblgrn'>Language</th>" & vbCrlf & _
		"<th class='tblgrn'>Interpreter Name</th>" & vbCrlf & _
		"<th class='tblgrn'>Appointment Time</th>" & vbCrlf & _
		"<th class='tblgrn'>Hours</th>" & vbCrlf & _
		"<th class='tblgrn'>Rate</th>" & vbCrlf & _
		"<th class='tblgrn'>Travel Time</th>" & vbCrlf & _
		"<th class='tblgrn'>Mileage</th>" & vbCrlf & _
		"<th class='tblgrn'>Emergency Surcharge</th>" & vbCrlf & _
		"<th class='tblgrn'>Total</th>" & vbCrlf & _
		"<th class='tblgrn'>Comment</th>" & vbCrlf & _
		"<th class='tblgrn'>DOB</th>" & vbCrlf & _
		"<th class='tblgrn'>DHHS</th>" & vbCrlf 
	
CSVHead = "Request ID, Institution, Department, Appointment Date, Client Last Name" & _
		", Client First Name, Medicaid, MCO, Language, Interpreter Last Name" & _
		", Interpreter First Name, Appointment Start Time, Appointment End Time, Hours, Rate" & _
		", Travel Time, Mileage, Emergency Surcharge, Total, Comments, DOB, DHHS"
	' add vermed = 1 AND if medicaid billing is go 
sqlRep = "SELECT req.[syscom], itr.[wid], req.[medicaid], req.[meridian]" & _
		", req.[nhhealth], req.[wellsense], req.[vermed]" & _
		", req.[autoacc], req.[wcomp], dep.[drg], itr.[pid]" & _
		", req.[index] as myindex, req.[amerihealth]" & _
		", req.[status], itr.[Last Name], itr.[First Name]" & _
		", req.[Clname], req.[Cfname], req.[Billable], req.[DOB]" & _
		", req.[AStarttime], req.[AEndtime], req.[M_Inst]" & _
		", req.[emerFEE], dep.[class], req.[TT_Inst]" & _
		", req.[InstID] as myinstID, req.[DeptID], req.[processedmedicaid]" & _
		", req.[LangID], req.[appDate], req.[InstRate]" & _
		", req.[bilComment], dep.[custID], dep.[ccode]" & _
		", dep.[billgroup], req.[IntrID], req.[billingTrail] " & _
		"FROM [request_T] AS req " & _
		"INNER JOIN [interpreter_T] AS itr ON req.[intrID]=itr.[index] " & _
		"INNER JOIN [dept_T] AS dep ON req.[deptid]=dep.[index] " & _
		"WHERE req.[instID] <> 479 " & _
		"AND req.[outpatient] = 1 " & _
		"AND req.[hasmed] = 1 " & _
		"AND req.[vermed] = 1 " & _
		"AND req.[autoacc] <> 1 " & _
		"AND req.[wcomp] <> 1 " & _
		"AND dep.[drg] = 1 " & _
		"AND (medicaid <> '' OR NOT medicaid IS NULL OR meridian <> '' OR NOT meridian IS NULL) " & _
		"AND (req.[Status]=1 OR req.[Status]=4) " & _
		"AND req.[ProcessedMedicaid] IS NULL " & _
		"AND req.[Processed] IS NULL "
strMSG = "Medicaid/MCO Billing request report (simulated)"
	
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
		sqlRep = sqlRep & " AND Class = " & tmpReport(8)
	End If
End If
sqlRep = sqlRep & " ORDER BY CustID ASC, AppDate DESC"
'
'response.write sqlRep
'
If blnWeaponized Then
	rsRep.Open sqlRep, g_strCONN, 1, 3
Else
	rsRep.Open sqlRep, g_strCONN, 3, 1
End If
'EMERGENCY RATE
tmpFeeL = 0
tmpFeeL = 0
Set rsRate = Server.CreateObject("ADODB.RecordSet")
sqlRate = "SELECT * FROM EmergencyFee_T"
rsRate.Open sqlRate, g_strCONN, 3,1
If Not rsRate.EOF Then
	tmpFeeL = rsRate("FeeLegal")
	tmpFeeO = rsRate("FeeOther")
End If
rsRate.Close
Set rsRate = Nothing

If Not rsRep.EOF Then 
	x = 0
	tmpCID = ""
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
		CB = ""
		If rsRep("status") = 4 Then CB = "*"
		strIntrName = rsRep("Last Name") & ",  " & rsRep("First Name")
		strCliName =  rsRep("Clname") & ", " & rsRep("Cfname")
		strATime =  cTime(rsRep("AStarttime")) & " -  " & cTime(rsRep("AEndtime"))
		'totHrs =  DateDiff("n", CDate(rsRep("AStarttime")) , CDate(rsRep("AEndtime")))
		BillHours =  rsRep("Billable")
		tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))

		totalPay = Z_FormatNumber(tmpPay, 2)
		hmoused = 0
		hmo = Z_FixNull(rsRep("medicaid")) 
		If hmo = "" Then
			hmo = Trim(Ucase(Z_FixNull(rsRep("amerihealth"))))
			hmoused = 4
		End If
		If hmo = "" Then
			hmo = Trim(Ucase(Z_FixNull(rsRep("meridian"))))
			hmoused = 1
		End If
		If hmo = "" Then
			hmo = Trim(Ucase(Z_FixNull(rsRep("nhhealth"))))
			hmoused = 2
		End If
		If hmo = "" Then 
			hmo = Trim(Ucase(Z_FixNull(rsRep("wellsense")))) 
			hmoused = 3
		End If
		strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'>" & CB & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & GetInst2(rsRep("myinstID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strCliName & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'>" & rsRep("medicaid") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & hmo & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strIntrName & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strATime & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'>" & BillHours & "</td>" & vbCrLf
		strBody = strBody & "<td class='tblgrn2'>$" & rsRep("InstRate") & "</td>" & vbCrLf
		strBody = strBody & "<td class='tblgrn2'>$" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>$" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf 
		strBody = strBody & "<td class='tblgrn2'>$0.00</td>" & vbCrLf
		'bilcomment = Replace(Z_fixNull(rsRep("bilComment") & rsRep("syscom")), "<br>Ap", "Ap")
		bilcomment = Z_fixNull(rsRep("bilComment") & rsRep("syscom"))
		If (Left(bilcomment, 4) = "<br>") Then bilComment = Right(bilcomment, Len(bilcomment) - 4)
		strBody = strBody & "<td class='tblgrn2'><b>$" & totalPay & "</b></td>" & vbCrLf & _
				"<td class='tblgrn2'>" & bilcomment & "</td><td class='tblgrn2'><nobr>" & rsRep("DOB") & "</td>"
		If rsRep("myinstID") = 108 Then
			strBody = strBody & "<td class='tblgrn2'>" & GetUserID(rsRep("DeptID")) & "</td><tr>" & vbCrLf 
		Else
			strBody = strBody & "<td class='tblgrn2'>&nbsp;</td><tr>" & vbCrLf 
		End If

		bilcommentcsv = Replace(bilcomment, "<br>", " / ")

		CSVBody = CSVBody & """" & CB & rsRep("myindex") & """,""" & GetInst2(rsRep("myinstID")) & """,""" & _
				Replace(GetMyDept(rsRep("DeptID")), " - ", "") & """,""" & rsRep("appDate") & """,""" & rsRep("Clname") & """,""" & _
				rsRep("Cfname") & """,""" & rsRep("medicaid") & """,""" & hmo &  """,""" & GetLang(rsRep("LangID")) & """,""" & rsRep("Last Name") & _
				""",""" & rsRep("First Name") & ""","""  & cTime(rsRep("AStarttime")) & """,""" & cTime(rsRep("AEndtime")) & """,""" & BillHours & _
				""",""" & rsRep("InstRate") & """,""" & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & """,""" & "0.00" & _
				""",""" & totalPay & """,""" & bilcommentcsv & """,""" & rsRep("DOB")
		If rsRep("myinstID") = 108 Then
			CSVBody = CSVBody & """,""" & GetUserID(rsRep("DeptID")) & """" & vbCrLf 
		Else
			CSVBody = CSVBody & """" & vbCrLf 
		End If
		csvBodyMCO = """" & CB & rsRep("myindex") & """,""" & GetInst2(rsRep("myinstID")) & """,""" & _
				Replace(GetMyDept(rsRep("DeptID")), " - ", "") & """,""" & rsRep("appDate") & """,""" & rsRep("Clname") & """,""" & _
				rsRep("Cfname") & """,""" & rsRep("medicaid") & """,""" & hmo &  """,""" & GetLang(rsRep("LangID")) & """,""" & rsRep("Last Name") & _
				""",""" & rsRep("First Name") & ""","""  & cTime(rsRep("AStarttime")) & """,""" & cTime(rsRep("AEndtime")) & """,""" & BillHours & _
				""",""" & rsRep("InstRate") & """,""" & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & """,""" & "0.00" & _
				""",""" & totalPay & """,""" & bilcommentcsv & """,""" & rsRep("DOB")
		If rsRep("myinstID") = 108 Then
			csvBodyMCO = csvBodyMCO & """,""" & GetUserID(rsRep("DeptID")) & """" & vbCrLf 
		Else
			csvBodyMCO = csvBodyMCO & """" & vbCrLf 
		End If
		mycode = Progcode(hmoused)
		'NEW CSV
		If mycode = "LB" Then 'medicaid
			csvBodyLB = csvBodyLB & csvBodyMCO
		ElseIf mycode = "AHC" Then 'meridian
			csvBodyAHC = csvBodyMHP & csvBodyMCO
		ElseIf mycode = "MHP" Then 'meridian
			csvBodyMHP = csvBodyMHP & csvBodyMCO
		ElseIf mycode = "NHHF" Then 'healthy fam
			csvBodyNHHF = csvBodyNHHF & csvBodyMCO
		ElseIf mycode = "WSHP" Then 'wellsense
			csvBodyWSHP = csvBodyWSHP & csvBodyMCO
		End If
		' ' ' ' ' ' ' '
		If BillHours >= 2 Then
			If hmoused <> 1 Then 'not meridian
				CSVBodyBill = CSVBodyBill & GetLBCode(BillHours, hmo, rsRep("wid"), rsRep("appDate"), mycode)
			Else
				CSVBodyBillMer = CSVBodyBillMer & GetLBCode(BillHours, hmo, rsRep("wid"), rsRep("appDate"), mycode)
			End If
		Else
			If hmoused <> 1 Then 'not meridian
				tmpLbhrs = BillHours
				CSVBodyBill = CSVBodyBill & GetLBCode(tmpLbhrs, hmo, rsRep("wid"), rsRep("appDate"), mycode)
			Else
				tmpLbhrs = BillHours
				CSVBodyBillMer = CSVBodyBillMer & GetLBCode(tmpLbhrs, hmo, rsRep("wid"), rsRep("appDate"), mycode)
			End If
		End If
		x = x + 1

		If blnWeaponized Then
			rsRep("billingTrail") = rsRep("billingTrail") & "<br>Billed to Medicaid/MCO " & Date
			rsRep("processedmedicaid") = Date
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

'CONVERT TO CSV
Set fso = CreateObject("Scripting.FileSystemObject")
Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
Prt.WriteLine "LANGUAGE BANK - REPORT"
Prt.WriteLine strMSG
Prt.WriteLine CSVHead
Prt.WriteLine CSVBody
Prt.Close	
Set Prt = Nothing
	
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
Prt2.WriteLine "LANGUAGE BANK - REPORT"
Prt2.WriteLine "LB/Medicaid," & strMSG
Prt2.WriteLine CSVHead
Prt2.WriteLine csvBodyLB
Prt2.Close	
Set Prt2 = Nothing
fso.CopyFile RepPath & RepCSVBillLB, BackupStr

Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillAHC, True) 'AHC
Prt2.WriteLine "LANGUAGE BANK - REPORT"
Prt2.WriteLine "AHC," & strMSG
Prt2.WriteLine CSVHead
Prt2.WriteLine csvBodyAHC
Prt2.Close	
Set Prt2 = Nothing
fso.CopyFile RepPath & RepCSVBillAHC, BackupStr
		
Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillMHP, True) 'MHP
Prt2.WriteLine "LANGUAGE BANK - REPORT"
Prt2.WriteLine "MHP," & strMSG
Prt2.WriteLine CSVHead
Prt2.WriteLine csvBodyMHP
Prt2.Close	
Set Prt2 = Nothing
fso.CopyFile RepPath & RepCSVBillMHP, BackupStr
		
Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillNHHF, True) 'NHHF
Prt2.WriteLine "LANGUAGE BANK - REPORT"
Prt2.WriteLine "NHHF," &CSVHead
Prt2.WriteLine csvBodyNHHF
Prt2.Close	
Set Prt2 = Nothing
fso.CopyFile RepPath & RepCSVBillNHHF, BackupStr

Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillWSHP, True) 'WSHP
Prt2.WriteLine "LANGUAGE BANK - REPORT"
Prt2.WriteLine "WSHP," &strMSG
Prt2.WriteLine CSVHead
Prt2.WriteLine csvBodyWSHP
Prt2.Close	
Set Prt2 = Nothing
fso.CopyFile RepPath & RepCSVBillWSHP, BackupStr

'COPY FILE TO BACKUP
fso.CopyFile RepPath & RepCSV, BackupStr
Set fso = Nothing
'EXPORT CSV
tmpstring 		= "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV)
tmpstring2 		= "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBill)
tmpstring5 		= "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillMeri)
tmpstringMED 	= "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillLB)
tmpstringMHP 	= "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillMHP)
tmpstringNHHF 	= "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillNHHF)
tmpstringWSHP 	= "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillWSHP)
tmpstringAHC 	= "dl_csv.asp?FN=" & Z_DoEncrypt(RepCSVBillAHC)

If blnWeaponized Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile(AdminLog, 8, True)
	strLog = Now & vbTab & "Medicaid Billing ran by " & Session("UsrName") & "."
	LogMe.WriteLine strLog
	Set LogMe = Nothing
End If
%>
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
function CloseMe() {
	window.close();
}
	</script>
<style>
@media print {
	tbody td { border-bottom: 1px dotted #bbb; }
	thead th { border-bottom: 1px dotted #888; background-color: #bbb; }
	thead {display: table-header-group;}
	tfoot {display: none;}
}
</style>
</head>
<body>
<div style="width: 300px; text-align: center; margin: 0px auto 20px;">
	<img src='images/LBISLOGO.jpg' align='center' style="width: 287px; height: 67px;" />
	340 Granite Street 3<sup>rd</sup> Floor, Manchester, NH 03102<br />
	Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
</div>
	<form method='post' name='frmResult'>
<b><%=strMSG%></b>		
<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center' style="width: 100%;">
<thead>
<tr><%=strHead%></tr>
</thead>
<tbody>
<%=strBody%>
</tbody>
<tfoot>
					<tr><td>&nbsp;</td></tr>
					<tr><td colspan='8' align='right' valign='bottom'><nobr>
						<input class='btn' type='button' value='Meridian Billing CSV'
								onmouseover="this.className='hovbtn'"
								onmouseout="this.className='btn'"
								onclick="document.location='<%=tmpstring5%>';" />
						<input class='btn' type='button' value='LB CSV'
								onmouseover="this.className='hovbtn'"
								onmouseout="this.className='btn'"
								onclick="document.location='<%=tmpstringMED%>';" />
						<input class='btn' type='button' value='AHC CSV'
								onmouseover="this.className='hovbtn'"
								onmouseout="this.className='btn'"
								onclick="document.location='<%=tmpstringAHC%>';" />
						<input class='btn' type='button' value='MHP CSV'
								onmouseover="this.className='hovbtn'"
								onmouseout="this.className='btn'"
								onclick="document.location='<%=tmpstringMHP%>';" />
						<input class='btn' type='button' value='NHHF CSV'
								onmouseover="this.className='hovbtn'"
								onmouseout="this.className='btn'"
								onclick="document.location='<%=tmpstringNHHF%>';" />
						<input class='btn' type='button' value='WSHP CSV'
								onmouseover="this.className='hovbtn'"
								onmouseout="this.className='btn'"
								onclick="document.location='<%=tmpstringWSHP%>';" />
						</nobr>
						</td><td colspan='8' align='left' valign='bottom'>
						<input class='btn' type='button' value='Print'
								onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'"
								onclick='print({bShrinkToFit: true});' />
						</td></tr>
					<tr><td colspan='4' align='right' valign='bottom'>
						&nbsp;
						</td><td colspan='12' align='left' valign='bottom'>
						* If needed, please adjust the page orientation of your printer to landscape to view all columns in a single page   
						</td></tr>
</tfoot>						
</table>
</form>
</body>
</html>
