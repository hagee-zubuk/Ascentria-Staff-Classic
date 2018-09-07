<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Request.Cookies("LBUSERTYPE") <> 1 Then 
	Session("MSG") = "Invalid account."
	Response.Redirect "default.asp"
End If
strDT1 = Request("dtStart")
strDT2 = Request("dtEnd")
dt1 = Z_CDate(strDT1)
dt2 = Z_CDate(strDT2)

strSQL = "EXEC [dbo].[spReqsMissingVForms] '" & Z_MDYDate(dt1) & "','" & Z_MDYDate(dt2) & "'"
Set rsF = Server.CreateObject("ADODB.RecordSet")
rsF.Open strSQL, g_strCONN, 3, 1

strZ = "<h2 style=""width: 450px; float: left;"">Scan period from " & Z_MDYDate(dt1) & " to " & Z_MDYDate(dt2) & "</h2>" & vbCrLf
csvF = "vf" & Z_GenerateGUID() & ".csv"
strZ = strZ & "<a href=""dl_csv.asp?NF=" & Z_DoEncrypt("VFormMissing.csv") & "&FN=" & Z_DoEncrypt(csvF) & """>" & _
		"<button style=""margin-top: 12px; width: 120px; float: left;"" type=""button"" id=""btnDL"">download CSV</button></a>" & vbCrLf
csvF = RepPath & csvF
' strZ = strZ & "<p>Saving to: <code>" & csvF & "</code></p>" & vbCrLf

If (Not rsF.EOF) Then
	lngI = 0
	strZ = strZ & "<table class=""resultsss"">" & vbCrLf
	strZ = strZ & "<tr><th>Req#</th><th>Status</th><th>Date</th><th>Institution</th><th>Department</th><th>Language</th><th>Uploads</th>" & _
			"<th>Interpreter</th><th colspan=""3"">Interpreter Contact</th></tr>" & vbCrLf
	Set fso  = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(csvF) Then
		fso.DeleteFile csvF
	End If
	Set oCSV = fso.OpenTextFile(csvF, 8, True)	

	oCSV.WriteLine """ReqID"",""Status"",""App Date"",""Inst"",""Dept"",""Lang"",""Uploads"",""Interpreter"",""e-MaIL"",""Phone"",""Phone"""
	Do While Not rsF.EOF
		oCSV.Write rsF("index") & ",""" & rsF("Status") & """," & Z_MDYDate(CDate(rsF("appDate"))) & ",""" & rsF("Facility") & """,""" & rsF("dept") & ""","""
		oCSV.Write rsF("Language") & """," & rsF("uploads") & ",""" & rsF("interpreter") & """,""" & rsF("email") & ""","""
		oCSV.WriteLine rsF("phone1") & """,""" & rsF("phone2") & """"
		strZ = strZ & "<tr><td>" & rsF("index") & "</td><td>" & Replace(rsF("Status"),"CANCELED/BILLABLE", "CANC/BILL") & "</td><td>" & Z_MDYDate(CDate(rsF("appDate"))) & "</td><td>"
		strZ = strZ & rsF("Facility") & "</td><td>" & rsF("dept") & "</td><td>" & rsF("Language") & "</td><td align=""center"">" & rsF("uploads") & "</td><td>"
		strZ = strZ & rsF("interpreter") & "</td><td>" & rsF("email") & "</td><td>" & rsF("phone1") & "</td><td>" & rsF("phone2") & "</td></tr>" & vbCrLF
		rsF.MoveNext
		lngI = lngI + 1
	Loop
	strZ = strZ & "</table><p>" & lngI & " records found"
	Set oCSV = Nothing
	Set fso  = Nothing
Else
strZ = strZ & "<br /><p>No records found.</p>"	
End If

rsF.Close
Set rsF = Nothing

%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
	<title>Language Bank - Missing Verification Form Report</title>
	<link href="CalendarControl.css" type="text/css" rel="stylesheet">
	<script src="CalendarControl.js" language="javascript"></script>
	<link href="css/jquery-ui.min.css" type="text/css" rel="stylesheet" />
	<link href="style.css" type="text/css" rel="stylesheet" />
	<script src="js/jquery-3.3.1.min.js" ></script>
	<script src="js/jquery-ui.min.js" ></script>
	<script language="JavaScript"><!--
function CalendarView(strDate)
	{
		document.frmConfirm.action = 'calendarview2.asp?appDate=' + strDate;
		document.frmConfirm.submit();
	}
	function MyEdits()
{
	document.frmConfirm.action = 'admin.asp?edits=1';
	document.frmConfirm.submit();
}
		//--></script>
<style>
.pickdate {
	text-align: center;
}
.err {
	background-color: yellow;
	color: red;
	font-weight: bold;
}
table.resultsss {
	border-collapse: collapse;
	clear: both;
}
table.resultsss,
.resultsss th,
.resultsss td {
   border: 1px solid #ccc;
}
.resultsss th {
	background-color: yellowgreen;
}
.resultsss td {
	padding: 1px 3px;
}
</style>		
<head>
<body>
		<table cellSpacing="0" cellPadding="0" height="100%" width="100%" border="0" class="bgstyle2">
			<tr><td valign="top" style="height: 100px;">
					<!-- #include file="_header.asp" -->
				</td>
			</tr>
			<tr><td valign="top" style="height: 60px;">
					<table cellSpacing='2' cellPadding='0' width="100%" border='0' align='center' >
						<!-- #include file="_greetme.asp" -->
						<tr><td align="center"><div class="err" id="err"></div></td></tr>
					</table>
<a href="rep.novform.asp">&lt;&lt;&nbsp;back</a>					
				</td></tr>
			<tr><td style="vertical-align: top; height: 250px;">
				
<%=strZ%>
				</td></tr>
		</table>
		<div style="width: 100%; position: relative; bottom: 0px;">
					<!-- #include file="_footer.asp" -->
		</div>
	</body>
</html>

<script>
function isValidDate(strz) {
	timestamp = Date.parse(strz) ;
	if (isNaN(timestamp) == false) { 
    	return(true);
	}
	return(false);
}

$( document ).ready(function() {
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
	Response.Write "alert(""" & tmpMSG & """);"
	Session("MSG") = ""
End If
%>
		$( "#dtStart").datepicker();
    	$( "#dtEnd" ).datepicker();
    	$( "#btnGO" ).click( function () {
    			var strErrs = "";
    			if (! isValidDate( $("#dtEnd").val() )) {
    				strErrs += "End date is invalid.\n";
    			}
    			if (! isValidDate( $("#dtStart").val() )) {
    				strErrs += "Start date is invalid.\n";
    			}
    			if (strErrs.length > 0) {
    				$("#err").html(strErrs);
    				return false;
    			}
    			$( "#frmMVF" ).submit();
    		} );
    	console.log( "ready!" );
	});	
</script>