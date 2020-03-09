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


Set rsRep = Server.CreateObject("ADODB.RecordSet")
strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _ 
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Medicaid</td>" & vbCrlf & _
		"<td class='tblgrn'>MCO</td>" & vbCrlf & _
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
		"<td class='tblgrn'>DHHS</td>" & vbCrlf 

	strMSG = "Medicaid/MCO Billing request report "



RepCSV 			= "ALLXBillReq" & tmpdate & "-" & tmpTime & ".csv" 
RepCSVBill 		= "MedicaidXBillReqNew" & tmpdate & "-" & tmpTime & ".csv" 
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

	
CSVHead = "Request ID, Institution, Department, Appointment Date, Client Last Name" & _
		", Client First Name, Medicaid, MCO, Language, Interpreter Last Name" & _
		", Interpreter First Name, Appointment Start Time, Appointment End Time, Hours, Rate" & _
		", Travel Time, Mileage, Emergency Surcharge, Total, Comments, DOB, DHHS"
CSVSimH = CSVHead
	' add vermed = 1 AND if medicaid billing is go 
sqlRep = "SELECT req.[index] as myindex, req.[billingTrail], req.[syscom]" & _
		", req.[medicaid], req.[amerihealth], req.[meridian], req.[nhhealth], req.[wellsense]" & _
		", req.[status], req.[vermed], req.[autoacc], req.[wcomp]" & _
		", dep.[dept], dep.[drg], dep.[class], dep.[billgroup], dep.[ccode], dep.[custID]" & _
		", itr.[pid], itr.[wid], itr.[Last Name], itr.[First Name]" & _
		", COALESCE(ins.[facility], 'N/A') AS [institution]" & _
		", lan.[language]" & _
		", COALESCE(usr.[user], 'N/A') AS [user]" & _
		", req.[Clname], req.[Cfname], req.[AStarttime], req.[AEndtime]" & _
		", req.[Billable], req.[DOB], req.[emerFEE], req.[TT_Inst], req.[M_Inst]" & _
		", req.[InstID] as myinstID" & _
		", req.[DeptID], req.[LangID], req.[appDate], req.[InstRate], req.[bilComment]" & _
		", req.[IntrID], req.[ProcessedMedicaid] " & _
	"FROM [request_T] AS req " & _
		"INNER JOIN [interpreter_T] AS itr ON req.[intrID]=itr.[index] " & _
		"INNER JOIN [dept_T] AS dep ON req.[deptID]=dep.[index] " & _
		"LEFT  JOIN [institution_T] AS ins ON req.[instID]=ins.[index] " & _
		"LEFT  JOIN [language_T] AS lan ON req.[langID]=lan.[index] " & _
		"LEFT  JOIN [interpreterSQL].dbo.[Appointment_T] AS app ON req.[HPID]=app.[index] " & _
		"LEFT  JOIN [interpreterSQL].dbo.[User_T] AS usr ON req.[instID]=usr.[InstID] AND app.[UID]=usr.[index] " & _
	"WHERE req.[instID] <> 479 " & _
		"AND req.[outpatient] = 1 " & _
		"AND req.[hasmed] = 1 " & _
		"AND req.[vermed] = 1 " & _
		"AND req.[autoacc] <> 1 " & _
		"AND req.[wcomp] <> 1 " & _
		"AND dep.[drg] = 1  " & _
		"AND (req.[medicaid] <> ''  " & _
				"OR NOT req.[medicaid] IS NULL " & _
				"OR req.[meridian] <> ''  " & _
				"OR NOT req.[meridian] IS NULL)  " & _
		"AND req.[Status] IN (1, 4) " & _
		"AND req.[ProcessedMedicaid] IS NULL " & _
		"AND req.[Processed] IS NULL "
strMSG = "Medicaid/MCO Billing request report (simulated)"
	
If tmpReport(1) <> "" Then
	sqlRep = sqlRep & " AND req.[appDate] >= '" & tmpReport(1) & "'"
	strMSG = strMSG & " from " & tmpReport(1)
End If
If tmpReport(2) <> "" Then
	sqlRep = sqlRep & " AND req.[appDate] <= '" & tmpReport(2) & "'"
	strMSG = strMSG & " to " & tmpReport(2)
End If
strMSG = strMSG & ". * - Cancelled Billable."
If tmpReport(9) = "" Then tmpReport(9) = 0
If tmpReport(9) <> 0 Then
	If tmpReport(6) = "" Then tmpReport(6) = 0
	If tmpReport(6) <> 0 Then 
		sqlRep = sqlRep & " AND req.[LangID] = " & tmpReport(6)
	End If
	If tmpReport(7) = "" Then tmpReport(7) = 0
	If tmpReport(7) <> "0" Then
		tmpCli = Split(tmpReport(7), ",")
		sqlRep = sqlRep & " AND req.[Clname] = '" & Trim(tmpCli(0)) & "' AND req.[Cfname] = '" & Trim(tmpCli(1)) & "'"
	End If
	If tmpReport(8) = "" Then tmpReport(8) = 0
	If tmpReport(8) <> 0 Then 
		sqlRep = sqlRep & " AND dep.[Class] = " & tmpReport(8)
	End If
End If
sqlRep = sqlRep & " ORDER BY dep.[CustID] ASC, req.[AppDate] DESC"

'Response.Write "<pre>" & sqlRep & "</pre>" & vbCrLf
'Response.End

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
		BillHours =  rsRep("Billable")
		tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))

		totalPay = Z_FormatNumber(tmpPay, 2)
		hmoused = 0
		hmo = Z_FixNull(rsRep("medicaid")) 
		If Z_FixNull(rsRep("meridian")) <> "" Then 
			hmo = Z_FixNull(rsRep("meridian"))
			hmoused = 1
		End If
		If Z_FixNull(rsRep("nhhealth")) <> "" Then 
			hmo = Z_FixNull(rsRep("nhhealth"))
			hmoused = 2
		End If
		If Z_FixNull(rsRep("wellsense")) <> "" Then 
			hmo = Z_FixNull(rsRep("wellsense")) 
			hmoused = 3
		End If
		If Z_FixNull(rsRep("amerihealth")) <> "" Then 
			hmo = Z_FixNull(rsRep("amerihealth")) 
			hmoused = 4
		End If
		strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & CB & rsRep("myindex") & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("institution") & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Dept") & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strCliName & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("medicaid") & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & hmo & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Language") & "</nobr></td>" & vbCrLf & _
				"<td class='tblgrn2'>" & strIntrName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & strATime & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & BillHours & "</td>" & vbCrLf
		strBody = strBody & "<td class='tblgrn2'>$" & rsRep("InstRate") & "</td>" & vbCrLf
		strBody = strBody & "<td class='tblgrn2'>$" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>$" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf 
		strBody = strBody & "<td class='tblgrn2'>$0.00</td>" & vbCrLf
			
		bilcomment = Z_fixNull(rsRep("bilComment") & rsRep("syscom"))
		If (Left(bilcomment, 4) = "<br>") Then bilComment = Right(bilcomment, Len(bilcomment) - 4)

		strBody = strBody & "<td class='tblgrn2'><b>$" & totalPay & "</b></td>" & vbCrLf & _
				"<td class='tblgrn2'>" & bilcomment & "</td><td class='tblgrn2'><nobr>" & rsRep("DOB") & "</td>"
		If rsRep("myinstID") = 108 Then
			strBody = strBody & "<td class='tblgrn2'>" & rsRep("user") & "</td><tr>" & vbCrLf 
		Else
			strBody = strBody & "<td class='tblgrn2'>&nbsp;</td><tr>" & vbCrLf 
		End If
		
		bilcommentcsv = Replace(bilcomment, "<br>", " / ")

		csvBodyLin = """" & CB & rsRep("myindex") & """,""" & rsRep("institution") & """,""" & _
				rsRep("Dept") & """,""" & rsRep("appDate") & """,""" & rsRep("Clname") & """,""" & _
				rsRep("Cfname") & """,""" & rsRep("medicaid") & """,""" & hmo &  """,""" & rsRep("language") & """,""" & rsRep("Last Name") & _
				""",""" & rsRep("First Name") & ""","""  & cTime(rsRep("AStarttime")) & """,""" & cTime(rsRep("AEndtime")) & """,""" & BillHours & _
				""",""" & rsRep("InstRate") & """,""" & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & """,""" & "0.00" & _
				""",""" & totalPay & """,""" & bilcommentcsv & """,""" & rsRep("DOB")
		If rsRep("myinstID") = 108 Then
			csvBodyLin = csvBodyLin & """,""" & rsRep("user") & """" & vbCrLf 
		Else
			csvBodyLin = csvBodyLin & """" & vbCrLf 
		End If

		CSVBody = CSVBody & csvBodyLin
		
		mycode = Progcode(hmoused)
		csvBodyMCO = GetLBCode(BillHours, hmo, rsRep("wid"), rsRep("appDate"), mycode)


		If BillHours >= 2 Then
			'If BillHours = 2 Then
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
			'NEW CSV
			If mycode = "LB" Then 'medicaid
				csvBodyLB = csvBodyLB & csvBodyMCO
			ElseIf mycode = "MHP" Then 'meridian
				csvBodyMHP = csvBodyMHP & csvBodyMCO
			ElseIf mycode = "NHHF" Then 'healthy fam
				csvBodyNHHF = csvBodyNHHF & csvBodyMCO
			ElseIf mycode = "WSHP" Then 'wellsense
				csvBodyWSHP = csvBodyWSHP & csvBodyMCO
			ElseIf mycode = "AHC" Then 'amerihealth
				csvBodyAHC = csvBodyAHC & csvBodyMCO			
			End If
			''''''''
		Else
			If mycode = "LB" Then 'medicaid
				csvBodyLB = csvBodyLB & csvBodyLin
			ElseIf mycode = "MHP" Then 'meridian
				csvBodyMHP = csvBodyMHP & csvBodyLin
			ElseIf mycode = "NHHF" Then 'healthy fam
				csvBodyNHHF = csvBodyNHHF & csvBodyLin
			ElseIf mycode = "WSHP" Then 'wellsense
				csvBodyWSHP = csvBodyWSHP & csvBodyLin
			ElseIf mycode = "AHC" Then 'amerihealth
				csvBodyAHC = csvBodyAHC & csvBodyLin			
			End If
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
		
'NEW CSV
Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillLB, True) 'LB
If blnWeaponized Then
	Prt2.WriteLine CSVHeadBill
Else
	Prt2.WriteLine "LANGUAGE BANK - REPORT"
	Prt2.WriteLine strMSG
	Prt2.WriteLine CSVSimH
End If
Prt2.WriteLine csvBodyLB
Prt2.Close	
Set Prt2 = Nothing
fso.CopyFile RepPath & RepCSVBillLB, BackupStr

Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillAHC, True) 'AHC
If blnWeaponized Then
	Prt2.WriteLine CSVHeadBill
Else
	Prt2.WriteLine "LANGUAGE BANK - REPORT"
	Prt2.WriteLine strMSG
	Prt2.WriteLine CSVSimH
End If
Prt2.WriteLine csvBodyAHC
Prt2.Close	
Set Prt2 = Nothing
fso.CopyFile RepPath & RepCSVBillAHC, BackupStr
		
Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillMHP, True) 'MHP
If blnWeaponized Then
	Prt2.WriteLine CSVHeadBill
Else
	Prt2.WriteLine "LANGUAGE BANK - REPORT"
	Prt2.WriteLine strMSG
	Prt2.WriteLine CSVSimH
End If
Prt2.WriteLine csvBodyMHP
Prt2.Close	
Set Prt2 = Nothing
fso.CopyFile RepPath & RepCSVBillMHP, BackupStr
		
Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillNHHF, True) 'NHHF
If blnWeaponized Then
	Prt2.WriteLine CSVHeadBill
Else
	Prt2.WriteLine "LANGUAGE BANK - REPORT"
	Prt2.WriteLine strMSG
	Prt2.WriteLine CSVSimH
End If
Prt2.WriteLine csvBodyNHHF
Prt2.Close	
Set Prt2 = Nothing
fso.CopyFile RepPath & RepCSVBillNHHF, BackupStr

Set Prt2 = fso.CreateTextFile(RepPath & RepCSVBillWSHP, True) 'WSHP
If blnWeaponized Then
	Prt2.WriteLine CSVHeadBill
Else
	Prt2.WriteLine "LANGUAGE BANK - REPORT"
	Prt2.WriteLine strMSG
	Prt2.WriteLine CSVSimH
End If
Prt2.WriteLine csvBodyWSHP
Prt2.Close	
Set Prt2 = Nothing
fso.CopyFile RepPath & RepCSVBillWSHP, BackupStr

'COPY FILE TO BACKUP
fso.CopyFile RepPath & RepCSV, BackupStr
Set fso = Nothing
'EXPORT CSV
tmpstring 		= "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV)
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
tbody td { border-bottom: 1px dotted #bbb; }
@media print {
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
					<tr><td colspan='18' align='center' valign='bottom'><nobr>
						<input class='btn' type='button' value='CSV Export'
								onmouseover="this.className='hovbtn'"
								onmouseout="this.className='btn'"
								onclick="document.location='<%=tmpstring%>';" />
<!--						<input class='btn' type='button' value='Meridian Billing CSV'
								onmouseover="this.className='hovbtn'"
								onmouseout="this.className='btn'"
								onclick="document.location='<%=tmpstring5%>';" /> -->
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
