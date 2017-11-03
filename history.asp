<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
'dispalys history of request
Set rsApp = Server.CreateObject("ADODB.RecordSet")
sqlApp = "SELECT * FROM Request_T WHERE [index] = " & Request("ReqID")
rsApp.Open sqlApp, g_strCONN, 3, 1
If Not rsApp.EOF Then
	tmpTS = "Appointment created on <b>" & rsApp("Timestamp") & "</b"
	tmpSR = "Request email has not been sent to Requesting Person."
	If rsApp("SentReq") <> "" Then tmpSR = "Request email was last sent to Requesting Person on <b>" & rsApp("SentReq")& "</b>"
	tmpSI = "Request email has not been sent to Interpreter."
	If rsApp("SentIntr") <> "" Then tmpSI = "Request email was last sent to Interpreter on <b>" & rsApp("SentIntr")& "</b>"
	tmpP = "Request has not been printed."
	If rsApp("Print") <> "" Then tmpP = "Request email was last printed on <b>" & rsApp("Print")& "</b>"
End If
rsApp.Close
Set rsApp = Nothing
Set rsHist = Server.CreateObject("ADODB.RecordSet")
sqlHist = "SELECT * FROM History_T WHERE reqID = " & Request("ReqID")
rsHist.Open sqlHist, g_strCONNHist, 3, 1
If Not rsHist.EOF Then
	tmpCreate = rsHist("creator")
	tmpDate = rsHist("date")
	tmpDateTS = rsHist("dateTS")
	tmpDateU = rsHist("dateU")
	tmpStime = rsHist("Stime")
	tmpStimeTS = rsHist("StimeTS")
	tmpStimeU = rsHist("StimeU")
	tmploc = rsHist("location")
	tmplocTS = rsHist("locationTS")
	tmplocU = rsHist("locationU")
	If rsHist("interID") <> "-1" Then
	 	tmpIntrTS = rsHist("interTS")
		tmpIntrU = rsHist("interU")
		tmpIntr = "Interpreter entered on " &  tmpIntrTS & " by " & tmpIntrU
	Else
		tmpIntr = "Interpreter not yet assigned."
	End If
	tmpCancel = rsHist("cancelTS") 
	tmpCancelTS = rsHist("cancelTS") 
	tmpCancelU = rsHist("cancelU") 
End If
rsHist.Close
Set rsHist = Nothing 
%>
<html>
	<head>
		<title>Language Bank - History</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function MyHist(xxx)
		{
			newwindow4 = window.open('dhistory.asp?ReqID=' + xxx,'name','height=400,width=800,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow4.focus()}
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmHist' action=''>
			<table cellpadding='0' cellspacing='0' border='0' align='left' width='100%'> 
				<tr>
					<td height='25px'>&nbsp;</td>
					<td class='header' colspan='6'>
						<nobr>HISTORY --&gt&gt
					</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left'>
						<%=tmpTS %>
					</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left'>
						<%=tmpSR%>
					</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left'>
						<%=tmpSI%>
					</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left'>
						<%=tmpP%>
					</td>
				</tr>
				<tr><td colspan='6'><hr align='center' width='75%'></td></tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left'>
						Created by: <%=tmpCreate %>
					</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left'>
						Appointment Date entered on  <%=tmpDateTS%> by <%=tmpDateU%>
					</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left'>
						Appointment Time entered on  <%=tmpStimeTS%> by <%=tmpStimeU%>
					</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left'>
						Appointment location entered on  <%=tmplocTS%> by <%=tmplocU%>
					</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='left'>
						<%=tmpIntr%>
					</td>
				</tr>
				<% If tmpCancel <> "" Then %>
					<tr>
						<td height='35px'>&nbsp;</td>
						<td align='left'>
							Appointment CANCELED by <%=tmpcancelU%> on <%=tmpcancelTS%>
						</td>
					</tr>
				<% End If %>
				<tr>
					<td align='center' colspan='10' class='RemME'>
						<a style='text-decoration: none;' href="JavaScript: MyHist(<%=Request("ReqID")%>);">[View Detailed History]</a>
						
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>