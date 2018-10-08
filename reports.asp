<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
'USER CHECK
If Request.Cookies("LBUSERTYPE") <> 1 And Request.Cookies("LBUSERTYPE") <> 5 Then 
	Session("MSG") = "Invalid account."
	Response.Redirect "default.asp"
End If
tmpPage = "document.frmReport."
tmpRep = Request("rep")
If Request("rep") = "" Then tmpRep = 0
tmpSele = Request("sel")
If Request("sel") = "" Then tmpSele = 0
'ON ERROR
If Session("MSG") <> "" Then
	tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
	'x = 1
	'Do Until x = 44
	'	TypeSel & x = ""
	'	If tmpReport(0) = x Then TypeSel & x  = "selected"
	'	x = x + 1
	'Loop
	TypeSel1 = ""
	TypeSel2 = ""
	TypeSel3 = ""
	TypeSel4 = ""
	TypeSel5 = ""
	TypeSel6 = ""
	TypeSel7 = ""
	TypeSel8 = ""
	TypeSel9 = ""
	TypeSel10 = ""
	TypeSel11 = ""
	TypeSel12 = ""
	TypeSel13 = ""
	TypeSel14 = ""
	TypeSel15 = ""
	TypeSel16 = ""
	TypeSel17 = ""
	TypeSel18 = ""
	TypeSel19 = ""
	TypeSel20 = ""
	TypeSel21 = ""
	TypeSel22 = ""
	TypeSel23 = ""
	TypeSel24 = ""
	TypeSel25 = ""
	TypeSel26 = ""
	TypeSel27 = ""
	TypeSel28 = ""
	TypeSel29 = ""
	TypeSel30 = ""
	TypeSel31 = ""
	TypeSel32 = ""
	TypeSel33 = ""
	TypeSel34 = ""
	TypeSel35 = ""
	TypeSel36 = ""
	TypeSel37 = ""
	TypeSel38 = ""
	TypeSel39 = ""
	TypeSel40 = ""
	TypeSel41 = ""
	TypeSel42 = ""
	TypeSel43 = ""
	TypeSel44 = ""
	TypeSel45 = ""
	TypeSel46 = ""
	TypeSel47 = ""
	TypeSel48 = ""
	TypeSel49 = ""
	TypeSel50 = ""
	TypeSel51 = ""
	TypeSel52 = ""
	TypeSel53 = ""
	TypeSel54 = ""
	TypeSel55 = ""
	TypeSel56 = ""
	TypeSel57 = ""
	TypeSel58 = ""
	TypeSel59 = ""
	TypeSel60 = ""
	TypeSel61 = ""
	TypeSel62 = ""
	TypeSel63 = ""
	TypeSel64 = ""
	TypeSel65 = ""
	TypeSel66 = ""
	TypeSel67 = ""
	TypeSel68 = ""
	TypeSel69 = ""
	TypeSel70 = ""
	TypeSel71 = ""
	TypeSel72 = ""
	TypeSel73 = ""
	TypeSel74 = ""

	If tmpReport(0) = 1 Then TypeSel1 = "selected"
	If tmpReport(0) = 2 Then TypeSel2 = "selected"
	If tmpReport(0) = 3 Then TypeSel3 = "selected"
	If tmpReport(0) = 4 Then TypeSel4 = "selected"
	If tmpReport(0) = 5 Then TypeSel5 = "selected"
	If tmpReport(0) = 6 Then TypeSel6 = "selected"
	If tmpReport(0) = 7 Then TypeSel7 = "selected"
	If tmpReport(0) = 8 Then TypeSel8 = "selected"
	If tmpReport(0) = 9 Then TypeSel9 = "selected"
	If tmpReport(0) = 10 Then TypeSel10 = "selected"
	If tmpReport(0) = 11 Then TypeSel11 = "selected"
	If tmpReport(0) = 12 Then TypeSel12 = "selected"
	If tmpReport(0) = 13 Then TypeSel13 = "selected"
	If tmpReport(0) = 14 Then TypeSel14 = "selected"
	If tmpReport(0) = 15 Then TypeSel15 = "selected"
	If tmpReport(0) = 16 Then TypeSel16 = "selected"
	If tmpReport(0) = 17 Then TypeSel17 = "selected"
	If tmpReport(0) = 18 Then TypeSel18 = "selected"
	If tmpReport(0) = 19 Then TypeSel19 = "selected"
	If tmpReport(0) = 20 Then TypeSel20 = "selected"
	If tmpReport(0) = 21 Then TypeSel21 = "selected"
	If tmpReport(0) = 22 Then TypeSel22 = "selected"
	If tmpReport(0) = 23 Then TypeSel23 = "selected"
	If tmpReport(0) = 24 Then TypeSel24 = "selected"
	If tmpReport(0) = 25 Then TypeSel25 = "selected"
	If tmpReport(0) = 26 Then TypeSel26 = "selected"
	If tmpReport(0) = 27 Then TypeSel27 = "selected"
	If tmpReport(0) = 28 Then TypeSel28 = "selected"
	If tmpReport(0) = 29 Then TypeSel29 = "selected"
	If tmpReport(0) = 30 Then TypeSel30 = "selected"
	If tmpReport(0) = 31 Then TypeSel31 = "selected"
	If tmpReport(0) = 32 Then TypeSel32 = "selected"
	If tmpReport(0) = 33 Then TypeSel33 = "selected"
	If tmpReport(0) = 34 Then TypeSel34 = "selected"
	If tmpReport(0) = 35 Then TypeSel35 = "selected"
	If tmpReport(0) = 36 Then TypeSel36 = "selected"
	If tmpReport(0) = 37 Then TypeSel37 = "selected"
	If tmpReport(0) = 38 Then TypeSel38 = "selected"
	If tmpReport(0) = 39 Then TypeSel39 = "selected"
	If tmpReport(0) = 40 Then TypeSel40 = "selected"
	If tmpReport(0) = 41 Then TypeSel41 = "selected"
	If tmpReport(0) = 42 Then TypeSel42 = "selected"
	If tmpReport(0) = 43 Then TypeSel43 = "selected"
	If tmpReport(0) = 44 Then TypeSel44 = "selected"
	If tmpReport(0) = 45 Then TypeSel45 = "selected"
	If tmpReport(0) = 46 Then TypeSel46 = "selected"
	If tmpReport(0) = 47 Then TypeSel47 = "selected"
	If tmpReport(0) = 48 Then TypeSel48 = "selected"
	If tmpReport(0) = 49 Then TypeSel49 = "selected"
	If tmpReport(0) = 50 Then TypeSel50 = "selected"
	If tmpReport(0) = 51 Then TypeSel51 = "selected"
	If tmpReport(0) = 52 Then TypeSel52 = "selected"
	If tmpReport(0) = 53 Then TypeSel53 = "selected"
	If tmpReport(0) = 54 Then TypeSel54 = "selected"
	If tmpReport(0) = 55 Then TypeSel55 = "selected"
	If tmpReport(0) = 56 Then TypeSel56 = "selected"
	If tmpReport(0) = 57 Then TypeSel57 = "selected"
	If tmpReport(0) = 58 Then TypeSel58 = "selected"
	If tmpReport(0) = 59 Then TypeSel59 = "selected"
	If tmpReport(0) = 60 Then TypeSel60 = "selected"
	If tmpReport(0) = 61 Then TypeSel61 = "selected"
	If tmpReport(0) = 62 Then TypeSel62 = "selected"
	If tmpReport(0) = 63 Then TypeSel63 = "selected"
	If tmpReport(0) = 64 Then TypeSel64 = "selected"
	If tmpReport(0) = 65 Then TypeSel65 = "selected"
	If tmpReport(0) = 66 Then TypeSel66 = "selected"
	If tmpReport(0) = 67 Then TypeSel67 = "selected"
	If tmpReport(0) = 68 Then TypeSel68 = "selected"
	If tmpReport(0) = 69 Then TypeSel69 = "selected"
	If tmpReport(0) = 70 Then TypeSel70 = "selected"
	If tmpReport(0) = 71 Then TypeSel71 = "selected"	
	If tmpReport(0) = 72 Then TypeSel72 = "selected"
	If tmpReport(0) = 73 Then TypeSel73 = "selected"
	If tmpReport(0) = 74 Then TypeSel74 = "selected"
		
	tmpRepFrom = tmpReport(1)
	tmpRepTo = tmpReport(2)
	tmpInst = Z_Cdbl(tmpReport(3))
	tmpIntr = tmpReport(4)
	tmpTown = tmpReport(5)
End If
strIntr = ""
strTown = ""
strLang = ""
strCli = ""
'GET INTERPRETER
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE Active = 1 ORDER BY [Last Name], [First Name]"
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	tmpSel = ""
	If tmpIntr = rsIntr("index") Then tmpSel = "selected"
	strIntr = strIntr & "<option " & tmpSel & " value='" & rsIntr("index") & "'>" & rsIntr("Last Name") & ", " & rsIntr("First Name") & "</option>" & vbCrLf 
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsInt = Nothing
strIntr = strIntr & "<option value='0'>---inactive interpreter----</option>" & vbCrLf 
'GET INACTIVE INTERPRETERS
Set rsIntr = Server.CreateObject("ADODB.RecordSet")
sqlIntr = "SELECT * FROM interpreter_T WHERE Active = 0 ORDER BY [Last Name], [First Name]"
rsIntr.Open sqlIntr, g_strCONN, 3, 1
Do Until rsIntr.EOF
	tmpSel = ""
	If tmpIntr = rsIntr("index") Then tmpSel = "selected"
	strIntr = strIntr & "<option " & tmpSel & " value='" & rsIntr("index") & "'>" & rsIntr("Last Name") & ", " & rsIntr("First Name") & "</option>" & vbCrLf 
	rsIntr.MoveNext
Loop
rsIntr.Close
Set rsInt = Nothing
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T ORDER BY [Facility]"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF
	tmpSel = ""
	If tmpInst = rsInst("index") Then tmpSel = "selected"
		InstName = rsInst("Facility")
		strInst = strInst	& "<option " & tmpSel & " value='" & rsInst("Index") & "'>" &  InstName & "</option>" & vbCrlf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
todaydate = Cdate(date)
%>
<html>
	<head>
		<title>Language Bank - Reports</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function CriSel(xxx)
		{
			document.frmReport.txtRepFrom.disabled = true;
			document.frmReport.txtRepTo.disabled = true;
			document.frmReport.cal1.disabled = true;
			document.frmReport.cal2.disabled = true;
			document.frmReport.selInst.disabled = true;
			document.frmReport.selIntr.disabled = true;
			document.frmReport.selTown.disabled = true;
			document.frmReport.txtZip.disabled = true;
			document.frmReport.txtState.disabled = true;
			document.frmReport.txtRepFrom.value = "";
			document.frmReport.txtRepTo.value = "";
			document.frmReport.selInst.value = -1;
			document.frmReport.selIntr.value = -1;
			document.frmReport.selTown.value = -1;
			document.frmReport.selIntrStat.value = 0;
			document.frmReport.selIntrStat.disabled = true;
			document.frmReport.txtZip.value = "";
			document.frmReport.txtState.value = "";
			if (xxx == 1 || xxx == 59 || xxx == 60)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = false;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtZip.value = "";
				document.frmReport.txtState.value = "";
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
			}
		else if (xxx == 33 || xxx == 55)
			{
				document.frmReport.txtRepFrom.disabled = true;
				document.frmReport.txtRepTo.disabled = true;
				document.frmReport.cal1.disabled = true;
				document.frmReport.cal2.disabled = true;
				document.frmReport.selInst.disabled = false;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtZip.value = "";
				document.frmReport.txtState.value = "";
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
			}
			else if (xxx == 2 ||xxx == 9 || xxx == 11 ||xxx == 12 || xxx == 14 || xxx == 6 || xxx == 47)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtZip.value = "";
				document.frmReport.txtState.value = "";
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
			}
			else if (xxx == 17 || xxx == 3 ||xxx == 10 ||xxx == 15 || xxx == 16 ||xxx == 21 ||xxx == 22 || xxx == 19 || xxx == 23 || xxx == 28 || xxx == 29 ||xxx == 30 ||xxx == 32 ||xxx == 38 ||xxx == 39 ||xxx == 40 ||xxx == 41 || xxx == 42 || xxx == 44 || xxx == 45 || xxx == 46 || xxx == 48 || xxx == 49 || xxx == 50 || xxx == 51 || xxx == 52 || xxx == 57 || xxx == 58 || xxx == 64 || xxx == 65 || xxx == 66 || xxx == 67 || xxx == 68 || xxx == 69 || xxx == 71 || xxx == 72 )
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
				if (xxx == 21 || xxx == 22)
				{
					document.frmReport.selIntrStat.disabled = false;
				}
				if (xxx == 38) {
					document.frmReport.selIntr.disabled = false;
					document.frmReport.selInst.disabled = false;
				}
				if (xxx == 64 || xxx == 65) {
					document.frmReport.txtZip.disabled = false;
				}
				if (xxx == 66 || xxx == 69) {
					document.frmReport.txtState.disabled = false;
				}
				if (xxx == 69) {
					document.frmReport.selInst.disabled = false;
				}
			}
			else if (xxx == 4 || xxx == 20 || xxx == 53 || xxx == 31  || xxx == 73 || xxx == 74)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = false;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
				document.frmReport.txtZip.value = "";
				document.frmReport.txtState.value = "";
			}
			else if (xxx == 5)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = false;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.txtZip.value = "";
				document.frmReport.txtState.value = "";
			}
			else if(xxx == 7 || xxx == 18)
			{
				document.frmReport.txtRepFrom.disabled = true;
				document.frmReport.txtRepTo.disabled = true;
				document.frmReport.cal1.disabled = true;
				document.frmReport.cal2.disabled = true;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
				document.frmReport.txtZip.value = "";
				document.frmReport.txtState.value = "";
			}
			else if(xxx == 8 || xxx == 24)
			{
				document.frmReport.txtRepFrom.disabled = true;
				document.frmReport.txtRepTo.disabled = true;
				document.frmReport.cal1.disabled = true;
				document.frmReport.cal2.disabled = true;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
				document.frmReport.txtZip.value = "";
				document.frmReport.txtState.value = "";
			}
			else if (xxx == 25 || xxx == 34 || xxx == 35 || xxx == 56 || xxx == 61)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = true;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = true;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = true;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
				document.frmReport.txtZip.value = "";
				document.frmReport.txtState.value = "";
			}
			else if (xxx == 26 || xxx == 36)
			{
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = true;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = true;
				document.frmReport.selInst.disabled = true;
				document.frmReport.selIntr.disabled = false;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
				document.frmReport.selIntrStat.value = 0;
				document.frmReport.selIntrStat.disabled = true;
				document.frmReport.txtZip.value = "";
				document.frmReport.txtState.value = "";
			}
			else if (xxx == 13 || xxx == 27 || xxx == 43) {
				document.frmReport.txtRepFrom.disabled = false;
				document.frmReport.txtRepTo.disabled = false;
				document.frmReport.cal1.disabled = false;
				document.frmReport.cal2.disabled = false;
				document.frmReport.selInst.disabled = true;
				if (xxx == 13) {
					document.frmReport.selInst.disabled = false;
				}
				document.frmReport.selIntr.disabled = false;
				document.frmReport.selTown.disabled = true;
				document.frmReport.txtRepFrom.value = "";
				document.frmReport.txtRepTo.value = "";
				document.frmReport.selInst.value = -1;
				document.frmReport.selIntr.value = -1;
				document.frmReport.selTown.value = -1;
				document.frmReport.selIntrStat.value = 0;
				document.frmReport.selIntrStat.disabled = true;
				document.frmReport.txtZip.value = "";
				document.frmReport.txtState.value = "";
			}
			document.frmReport.chkAddnl.disabled = ( (xxx == 47) || (xxx == 73) || xxx == 74 );
		}
		
		
		function RepGen() {
			if (document.frmReport.selRep.value == -1) {
				alert("Error: Please select a report type.");
				return;
			}
			if (document.frmReport.chkAddnl.checked == true) {
				if (document.frmReport.selLang.value == -1 && document.frmReport.selCli.value == -1 && document.frmReport.selClass.value == -1)	{
					alert("Error: Please select a filter.");
					return;
				}
			}
			if (document.frmReport.chkAddnl.checked == true) {
				if (document.frmReport.selRep.value == 7 || document.frmReport.selRep.value == 8 || 
						document.frmReport.selRep.value == 11 || document.frmReport.selRep.value == 12 || 
						document.frmReport.selRep.value == 13 || document.frmReport.selRep.value == 14 || 
						document.frmReport.selRep.value == 15 || document.frmReport.selRep.value == 17 || 
						document.frmReport.selRep.value == 18 || document.frmReport.selRep.value == 19 || 
						document.frmReport.selRep.value == 6 || document.frmReport.selRep.value == 23 || 
						document.frmReport.selRep.value == 24 || document.frmReport.selRep.value == 26 || 
						document.frmReport.selRep.value == 27 || document.frmReport.selRep.value == 29 || 
						document.frmReport.selRep.value == 17 || document.frmReport.selRep.value == 36 || 
						document.frmReport.selRep.value == 42 || document.frmReport.selRep.value == 64 ||
						document.frmReport.selRep.value == 65 || document.frmReport.selRep.value == 66 ||
						document.frmReport.selRep.value == 67 || document.frmReport.selRep.value == 68 ||
						document.frmReport.selRep.value == 69 || document.frmReport.selRep.value == 71 ||
						document.frmReport.selRep.value == 72 || document.frmReport.selRep.value == 73
						) {
					alert("Error: Filter is not applicable with this report type.") 
					return;
				}
			}
			if (document.frmReport.selRep.value == 10 || document.frmReport.selRep.value == 19 || 
					document.frmReport.selRep.value == 23 || document.frmReport.selRep.value == 27 || 
					document.frmReport.selRep.value == 43 || document.frmReport.selRep.value == 47 ||
					document.frmReport.selRep.value == 49 || document.frmReport.selRep.value == 50 || 
					document.frmReport.selRep.value == 52 || document.frmReport.selRep.value == 57 || 
					document.frmReport.selRep.value == 58 || document.frmReport.selRep.value == 28 ||
					document.frmReport.selRep.value == 64 || document.frmReport.selRep.value == 65 ||
					document.frmReport.selRep.value == 66 || document.frmReport.selRep.value == 67 ||
					document.frmReport.selRep.value == 68 || document.frmReport.selRep.value == 69 ||
					document.frmReport.selRep.value == 71 || document.frmReport.selRep.value == 72 ||
					document.frmReport.selRep.value == 73
					) {
				if (document.frmReport.txtRepFrom.value == "" || document.frmReport.txtRepTo.value == "") {
					alert("Error: Timeframe is required.");
					return;
				}
			}
			if (document.frmReport.selRep.value == 28 || document.frmReport.selRep.value == 67 || document.frmReport.selRep.value == 71) {
				var datefrom = new Date(document.frmReport.txtRepFrom.value + " 00:00:00");
				var dateto = new Date(document.frmReport.txtRepTo.value + " 00:00:00");
				if (datefrom.getDay() != 0 || dateto.getDay() != 6) {
					alert("Error: 'From' date needs to be a Sunday and/or 'To' date needs to be a Saturday.");
					return;
				}
			}
			if (document.frmReport.selRep.value == 43) {
				if (document.frmReport.selIntr.value == 0) {
					alert("Error: Specific Interpreter is required.");
					return;
				}
				var datefrom = new Date(document.frmReport.txtRepFrom.value + " 00:00:00");
				var dateto = new Date(document.frmReport.txtRepTo.value + " 00:00:00");
				if (datefrom.getDay() != 0 || dateto.getDay() != 6) {
					alert("Error: 'From' date needs to be a Sunday and/or 'To' date needs to be a Saturday.");
					return;
				}
			}
			if (document.frmReport.selRep.value == 17 || document.frmReport.selRep.value == 25 ||
					document.frmReport.selRep.value == 34 || document.frmReport.selRep.value == 35 ||
					document.frmReport.selRep.value == 48 || document.frmReport.selRep.value == 56 ||
					document.frmReport.selRep.value == 61 || document.frmReport.selRep.value == 72 ) {
				if (document.frmReport.txtRepFrom.value == "") {
					alert("Error: Timeframe is required.");
					return;
				}
			}
			if (document.frmReport.selRep.value == 21 || document.frmReport.selRep.value == 22 ||
					document.frmReport.selRep.value == 48 || document.frmReport.selRep.value == 72
				) {
					if (document.frmReport.txtRepTo.value == "") {
						alert("Error: 'To:' date is required.")
						return; 
					} else {
						var currentTime = new Date();
						var month = currentTime.getMonth() + 1;
						var day = currentTime.getDate();
						var year = currentTime.getFullYear();
						var datetoday = new Date(month + "/" + day + "/" + year);
						var todate = new Date(document.frmReport.txtRepTo.value);
						var todateyear = todate.getFullYear();
						var todatemonth = todate.getMonth() + 1;
						var todateday = todate.getDate();
						var newtodate = new Date(todatemonth + "/" + todateday + "/" + todateyear);
 						if (datetoday < newtodate) {
 							alert("Error: 'To:' date should be today or in the past");
 							return;
 						}
 					}
			}
			document.frmReport.action = "action.asp?ctrl=5";
			document.frmReport.submit();
		}
		function PopMe(zzz, xxx)
		{
			if (zzz !== 0)
			{
				newwindow = window.open('printreport.asp','','height=800,width=900,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
				if (window.focus) {newwindow.focus()}
			}
		}
		function FilterMe()
		{
			if (document.frmReport.chkAddnl.checked == true)
			{
				document.frmReport.selLang.disabled = false;
				document.frmReport.selCli.disabled = false;
				document.frmReport.selClass.disabled = false;
			}
			else
			{
				document.frmReport.selLang.value = -1;
				document.frmReport.selCli.value = -1;
				document.frmReport.selClass.value = -1;
				document.frmReport.selLang.disabled = true;
				document.frmReport.selCli.disabled = true;
				document.frmReport.selClass.disabled = true;
			}
		}
		function TypeDef(xxx)
		{
			document.frmReport.tadef.value = "";
			if (xxx == 1)
			{
				document.frmReport.tadef.value = "Invoice Report definition here.";
			}
			if (xxx == 8)
			{
				document.frmReport.tadef.value = "List of Interpreters.";
			}
			if (xxx == 7)
			{
				document.frmReport.tadef.value = "List of Requesting Persons";
			}
			if (xxx == 2)
			{
				document.frmReport.tadef.value = "List of Canceled appointments";
			}
			if (xxx == 3)
			{
				document.frmReport.tadef.value = "Bills completed requests for institutions.";
			}
			if (xxx == 4)
			{
				document.frmReport.tadef.value = "per institution Report definition here.";
			}
			if (xxx == 5)
			{
				document.frmReport.tadef.value = "per town Report definition here.";
			}
			if (xxx == 6)
			{
				document.frmReport.tadef.value = "usage Report definition here.";
			}
			if (xxx == 9)
			{
				document.frmReport.tadef.value = "Missed Report definition here.";
			}
			if (xxx == 10)
			{
				document.frmReport.tadef.value = "Language Bank Statistics. NOT YET DONE";
			}
			if (xxx == 11)
			{
				document.frmReport.tadef.value = "Pending requests";
			}
			if (xxx == 12)
			{
				document.frmReport.tadef.value = "Completed requests";
			}
			if (xxx == 13)
			{
				document.frmReport.tadef.value = "Missed requests";
			}
			if (xxx == 14)
			{
				document.frmReport.tadef.value = "Canceled requests";
			}
			if (xxx == 15)
			{
				document.frmReport.tadef.value = "Canceled (Billable) requests";
			}
			if (xxx == 16)
			{
				document.frmReport.tadef.value = "Simulates billing report. This report will not tag requests as billed.";
			}
			if (xxx == 17)
			{
				document.frmReport.tadef.value = "KPI report. Select any date of the month you wish to have a report.";
			}
			if (xxx == 18)
			{
				document.frmReport.tadef.value = "Pending Court requests for the past 30 days.";
			}
			if (xxx == 19)
			{
				document.frmReport.tadef.value = "Completed and Canceled - Billable Court appointments";
			}
			if (xxx == 20)
			{
				document.frmReport.tadef.value = "Audit report for Institutions.";
			}
			if (xxx == 21)
			{
				document.frmReport.tadef.value = "Pays completed requests for interpreters.";
			}
			if (xxx == 22)
			{
				document.frmReport.tadef.value = "Simulates payroll report. This report will not tag requests as paid.";
			}
			if (xxx == 23)
			{
				document.frmReport.tadef.value = "Cancelled Court Appointments. ";
			}
			if (xxx == 24)
			{
				document.frmReport.tadef.value = "List of ACTIVE interpreters. ";
			}
			if (xxx == 25)
			{
				document.frmReport.tadef.value = "Weekly report. Select any date of the week you wish to have a report.";
			}
			if (xxx == 26)
			{
				document.frmReport.tadef.value = "Mileage report. Select any date of the Month you wish to have a report. You can also select a specific interpreter.";
			}
			if (xxx == 27)
			{
				document.frmReport.tadef.value = "Timesheet report. You can also select a specific interpreter.";
			}
			if (xxx == 28)
			{
				document.frmReport.tadef.value = "Total Hours report (pay period).";
			}
			if (xxx == 29)
			{
				document.frmReport.tadef.value = "Billable Hours report.";
			}
			if (xxx == 30)
			{
				document.frmReport.tadef.value = "Pending, completed, billed, canceled-billable appointments.";
			}
			if (xxx == 31)
			{
				document.frmReport.tadef.value = "Emergency appointments.";
			}
			if (xxx == 32)
			{
				document.frmReport.tadef.value = "Duration Complete Report. Gives number of days an appointment takes to be completed.";
			}
			if (xxx == 33)
			{
				document.frmReport.tadef.value = "Instiution Report.";
			}
			if (xxx == 34)
			{
				document.frmReport.tadef.value = "On-call schedule Report. Select any date of the month you wish to have a report.";
			}
			if (xxx == 35)
			{
				document.frmReport.tadef.value = "Hours Billable Report. Select any date of the month you wish to have a report.";
			}
			if (xxx == 36)
			{
				document.frmReport.tadef.value = "Simulates Mileage report. Select any date of the Month you wish to have a report. You can also select a specific interpreter.";
			}
			if (xxx == 37)
			{
				document.frmReport.tadef.value = "Inactive Interpreters report";
			}
			if (xxx == 38)
			{
				document.frmReport.tadef.value = "Tardiness report";
			}
			if (xxx == 39)
			{
				document.frmReport.tadef.value = "Simulates medicaid report. This report will not tag requests as billed.";
			}
			if (xxx == 40)
			{
				document.frmReport.tadef.value = "Bills completed requests for medicaid.";
			}
			if (xxx == 41)
			{
				document.frmReport.tadef.value = "Travel Time and Mileage report";
			}
			if (xxx == 42)
			{
				document.frmReport.tadef.value = "Billed Hours report";
			}
			if (xxx == 43)
			{
				document.frmReport.tadef.value = "Weekly Timesheet report. Please select a specific interpreter.";
			}
			if (xxx == 44)
			{
				document.frmReport.tadef.value = "Costs of foreign language and ASL interpreter services, by court location and jurisdiction and for the entire NHJB.";
			}
			if (xxx == 45)
			{
				document.frmReport.tadef.value = "Costs of foreign language and ASL interpreter services, by language.";
			}
			if (xxx == 46)
			{
				document.frmReport.tadef.value = "Frequency of NHJB use of interpreters in different foreign languages and in ASL.";
			}
			if (xxx == 47)
			{
				document.frmReport.tadef.value = "Frequency of language usage.";
			}
			if (xxx == 48)
			{
				document.frmReport.tadef.value = "Interpreters without appointments within the set date range.";
			}
			if (xxx == 49)
			{
				document.frmReport.tadef.value = "Alen's report.";
			}
			if (xxx == 50)
			{
				document.frmReport.tadef.value = "Alen's report 2.";
			}
			if (xxx == 51)
			{
				document.frmReport.tadef.value = "Frequency of language by class report.";
			}
			if (xxx == 52)
			{
				document.frmReport.tadef.value = "Lynda's report.";
			}
			if (xxx == 53)
			{
				document.frmReport.tadef.value = "Audit report for medicaid appointments.";
			}
			if (xxx == 54)
			{
				document.frmReport.tadef.value = "Emergency Reports.";
			}
			if (xxx == 55)
			{
				document.frmReport.tadef.value = "Active Instiution Report.";
			}
			if (xxx == 56)
			{
				document.frmReport.tadef.value = "WellSense Health Plan Language Frequency Report. Select any date of the month you wish to have a report.";
			}
			if (xxx == 57)
			{
				document.frmReport.tadef.value = "Number of appointments and total billables per Institution for MA.";
			}
			if (xxx == 58)
			{
				document.frmReport.tadef.value = "Number of appointments and total billables per Institution for NH";
			}
			if (xxx == 59)
			{
				document.frmReport.tadef.value = "'Happened' appointments. Medicaid/MCO only.";
			}
			if (xxx == 60)
			{
				document.frmReport.tadef.value = "'Did not Happened' appointments. Medicaid/MCO only.";
			}
			if (xxx == 61)
			{
				document.frmReport.tadef.value = "Monthly Total Hours. Select any date of the month you wish to have a report.";
			}
			if (xxx == 62)
			{
				document.frmReport.tadef.value = "Payroll export.";
			}
			if (xxx == 63)
			{
				document.frmReport.tadef.value = "FWA Traning report.";
			}
			if (xxx == 64)
			{
				document.frmReport.tadef.value = "Interpreter Activity by zip.";
			}
			if (xxx == 65)
			{
				document.frmReport.tadef.value = "Activity by zip ";
			}
			if (xxx == 66)
			{
				document.frmReport.tadef.value = "Activity by state ";
			}
			if (xxx == 67)
			{
				document.frmReport.tadef.value = "No Hours report (pay period).";
			}
			if (xxx == 68)
			{
				document.frmReport.tadef.value = "Elliot Report (Medicaid appointments, pending, billed, Completed unbilled).";
			}
			if (xxx == 69)
			{
				document.frmReport.tadef.value = "Language statistics.";
			}
			if (xxx == 70)
			{
				document.frmReport.tadef.value = "Interpreter I list";
			}
			if (xxx == 71)
			{
				document.frmReport.tadef.value = "Total Hours Weekly Report (NEW)";
			}
			if (xxx == 72)
			{
				document.frmReport.tadef.value = "Creation of appointment";
			}
			if (xxx == 73) document.frmReport.tadef.value = "Pending Appointments with Medicaid information";
			if (xxx == 74) document.frmReport.tadef.value = "Interpreter responses to appointment availability";
		}
		function CalendarView(strDate)
		{
			document.frmReport.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmReport.submit();
		}
		function SubmitAko()
			{
				document.frmReport.action = 'reqconfirm.asp?ID=' + document.frmReport.hideID.value;
				document.frmReport.submit();
			}
		-->
		</script>
		<body onload='CriSel(document.frmReport.selRep.value); PopMe(<%=tmpRep%>,<%= tmpSele%>);FilterMe();'>
			<form method='post' name='frmReport' action='reports.asp'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" class='bgstyle2' border='0'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<!-- #include file="_greetme.asp" -->
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='4' cellPadding='0' align='center' border='0' class="defborder">
								<tr>
									<td colspan='2' align='center'>
										<b>Report Query</b>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Type:
									</td>
									<td>
										<select class='seltxt' name='selRep'  style='width:200px;' onchange='CriSel(document.frmReport.selRep.value); TypeDef(document.frmReport.selRep.value);'>
											<option value='-1'>&nbsp;</option>
											<option value='30' <%=TypeSel30%>>Activity Report</option>
											<option value='66' <%=TypeSel66%>>Activity by State Report</option>
											<option value='65' <%=TypeSel65%>>Activity by Zip Code Report</option>
											<option value='72' <%=TypeSel72%>>Appointment Creation Report</option>
											<option value='20' <%=TypeSel20%>>Audit Report - Institutions</option>
											<option value='53' <%=TypeSel53%>>Audit Report - Medicaid</option>
											<option value='57' <%=TypeSel57%>>Billable Appointments Report (MA)</option>
											<option value='58' <%=TypeSel58%>>Billable Appointments Report (NH)</option>
											<option value='29' <%=TypeSel29%>>Billable Hours Report</option>
											<option value='42' <%=TypeSel42%>>Billed Hours Report</option>
											<option value='15' <%=TypeSel15%>>Canceled (Billable) Appointment Report</option>
											<option value='14' <%=TypeSel14%>>Canceled Appointment Report</option>
											<option value='12' <%=TypeSel12%>>Completed Appointment Report</option>
											<option value='44' <%=TypeSel44%>>Court Appointment Cost</option>
											<option value='19' <%=TypeSel19%>>Court Appointment Report</option>
											<option value='23' <%=TypeSel23%>>Court Cancelled Appointment Report</option>
											<option value='45' <%=TypeSel45%>>Court Language Cost</option>
											<option value='46' <%=TypeSel46%>>Court Language Frequency</option>
											<option value='18' <%=TypeSel18%>>Court Pending Appointment Report</option>
											<option value='32' <%=TypeSel32%>>Duration Complete Report</option>
											<option value='68' <%=TypeSel68%>>Elliot Report</option>
											<option value='31' <%=TypeSel31%>>Emergency Report</option>
											<option value='63' <%=TypeSel63%>>FWA Training Report</option>
											<option value='59' <%=TypeSel59%>>Happened Appointments Report</option>
											<option value='35' <%=TypeSel35%>>Hours Billable Report</option>
											<option value='33' <%=TypeSel33%>>Institution Report</option>
											<option value='55' <%=TypeSel55%>>Institution Report (ACTIVE ONLY)</option>
											<option value='64' <%=TypeSel64%>>Interpreter Activity by Zip Code Report</option>
											<option value='70' <%=TypeSel70%>>Interpreter I</option>
											<option value='48' <%=TypeSel48%>>Interpreter Inactivity Report</option>
											<option value='1' <%=TypeSel1%>>Invoice Report</option>
											<option value='17' <%=TypeSel17%>>KPI Report</option>
											<option value='51' <%=TypeSel51%>>Language Frequency by Class Report</option>
											<option value='69' <%=TypeSel69%>>Language Statistics Report</option>
											<option value='47' <%=TypeSel47%>>Language Use Report</option>
											<option value='24' <%=TypeSel8%>>List - Interpreter (ACTIVE ONLY)</option>
											<option value='37' <%=TypeSel37%>>List - Interpreter (INACTIVE ONLY)</option>
											<option value='8' <%=TypeSel8%>>List - Interpreter</option>
											<option value='7' <%=TypeSel7%>>List - Requesting Person</option>
											<option value='13' <%=TypeSel13%>>Missed Appointment Report</option>
											<option value='67' <%=TypeSel67%>>No Hours Report (pay period)</option>
											<option value='60' <%=TypeSel260%>>Not Happened Appointments Report</option>
											<option value='34' <%=TypeSel34%>>On-Call Schedule Report</option>
											<option value='11' <%=TypeSel11%>>Pending Appointment Report</option>
											<option value='4' <%=TypeSel4%>>Per-Institution Report</option>
											<option value='5' <%=TypeSel5%>>Per-Town Report</option>
											<option value='10' <%=TypeSel10%>>Statistics</option>
											<option value='38' <%=TypeSel38%>>Tardiness Report</option>
											<option value='27' <%=TypeSel27%>>Timesheet Report</option>
											<option value='61' <%=TypeSel61%>>Total Hours Report (Monthly)</option>
											<option value='28' <%=TypeSel28%>>Total Hours Report (pay period)</option>
											<option value='71' <%=TypeSel71%>>Total Hours Report (NEW)</option>
											<option value='6' <%=TypeSel6%>>Usage Report</option>
											<option value='25' <%=TypeSel25%>>Weekly Report</option>
											<option value='43' <%=TypeSel43%>>Weekly Timesheet Report</option>
											<option value='56' <%=TypeSel56%>>WellSense Health Plan Language Report</option>
											<option value='41' <%=TypeSel41%>>Travel Time and Mileage Report</option>
											<option value='16' <%=TypeSel16%>>*Pre-Billing Report</option>
											<option value='39' <%=TypeSel39%>>*Pre-Medicaid Report</option>
											<option value='36' <%=TypeSel36%>>*Pre-Mileage Report</option>
											<% If Request.Cookies("UID") = 8 Or Request.Cookies("UID") = 2 Then %> <!-- finance only or me//-->
												<option value='3' <%=TypeSel3%>>*Billing Report</option>
												<option value='40' <%=TypeSel40%>>*Medicaid Report</option>
											<% End If %>
											<option value='26' <%=TypeSel26%>>*Mileage Report</option>
											<option value='49' <%=TypeSel49%>>**Alen Report</option>
											<option value='50' <%=TypeSel50%>>**Alen Report 2</option>
											<option value='52' <%=TypeSel52%>>**Lynda Report</option>
											<option value='73' <%=TypeSel73%>>* Pending Appts w/ Medicaid</option>
											<option value='74' <%=TypeSel74%>>Interpreter Appt Response</option>
											<!--<option value='21' <%=TypeSel21%>>Payroll Report</option>-->
											<!--<option value='22' <%=TypeSel22%>>Pre-Payroll Report</option>-->
										</select>
									</td>
								</tr>
								<tr>
									<td align='right' valign='top'>
										Description: 
									</td>
									<td>
										<textarea class='def' name='tadef' readonly ></textarea>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td align='right'>
										Criteria:
									</td>
									<td>
										( leave blank to select all )
									</td>
								</tr>
								<tr>
									<td align='right'>Timeframe:</td>
									<td>
										&nbsp;From:<input class='main' size='10' maxlength='10' name='txtRepFrom' readonly value='<%=tmpRepFrom%>'>
										<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;"
											onclick="showCalendarControl(document.frmReport.txtRepFrom);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
										&nbsp;To:<input class='main' size='10' maxlength='10' name='txtRepTo' readonly value='<%=tmpRepTo%>'>
										<input type="button" value="..." title='Calendar' name="cal2" style="width: 19px;"
											onclick="showCalendarControl(document.frmReport.txtRepTo);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
									</td>
								</tr>
								<tr>
									<td align='right'>
										Institution:
									</td>
									<td>
										<select class='seltxt' name='selInst'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strInst%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Interpreter:
									</td>
									<td>
										<select class='seltxt' name='selIntr'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strIntr%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Town:
									</td>
									<td>
										<select class='seltxt' name='selTown'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strTown%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>
										State:
									</td>
									<td>
										<input class='main'  name='txtState' size='3' maxlength='2'>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Zip Code:
									</td>
									<td>
										<input class='main'  name='txtZip' style='width:200px;'>
										<br><span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*use semi-colon(;) as separator if more than 1 zip code.</span>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td align='right'>
										Filter:
									</td>
									<td>
										<input type='checkbox' name='chkAddnl' id='chkAddnl' value='1' onclick='FilterMe();'>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Language:
									</td>
									<td>
										<select class='seltxt' name='selLang' id='selLang' style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strLang%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Client:
									</td>
									<td>
										<select class='seltxt' name='selCli' id='selCli'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<%=strCli%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Classification:
									</td>
									<td>
										<select class='seltxt' style='width: 200px;' name='selClass' id='selClass'>
											<option value='0'>&nbsp;</option>
											<option value='1' <%=SocSer%>>Social Services</option>
											<option value='2' <%=Priv%>>Private</option>
											<option value='3' <%=Court%>>Court</option>
											<option value='4' <%=Med%>>Medical</option>
											<option value='5' <%=legal%>>Legal</option>
										</select>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td align='right'><nobr>Interpreter:</td>
									<td>
										<select class='seltxt' style='width: 200px;' name='selIntrStat'>
											<option value='0'>&nbsp;---All---&nbsp;</option>
											<option value='1'>Employee</option>
											<option value='2'>Outside Consultant</option>
										</select>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<input class='btn' type='button' style='width: 200px;' value='Generate' 
											onmouseover="this.className='hovbtn'" 
											onmouseout="this.className='btn'" 
											onclick='RepGen();'>
										<input type='hidden' name='hideID'>
									</td>
								</tr>
								<tr>
									<td colspan='2' align='center'>
										<span class='error'><%=Session("MSG")%></span>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
		</body>
	</head>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>