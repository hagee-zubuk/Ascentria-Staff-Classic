<!DOCTYPE html>
<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_googleDMA.asp" -->
<%
' all that went before should mother-fathering go away.'

' GET mileage cap for interpreters
Set rsmile = Server.CreateObject("ADODB.Recordset")
sqlMile = "SELECT [milediff] FROM [travel_t]"
rsmile.Open sqlMile, g_strCONN, 3, 1
If Not rsmile.EOF Then tmpmilecap = Z_czero(rsmile("milediff"))
rsmile.Close
Set rsmile = Nothing

' GET ADDRESS AND ZIP of Intrpreter
Set rsIntr = Server.CreateObject("ADODB.REcordSet")
strItrID = Request("selIntr")
sqlIntr = "SELECT * FROM Interpreter_T WHERE [index] = " & strItrID
rsIntr.Open sqlIntr, g_strCONN, 1, 3
If Not rsIntr.EOF Then
	tmpIntrAdd = rsIntr("address1") & ", " & rsIntr("City") & ", " &  rsIntr("state") & ", " & rsIntr("Zip Code")
	tmpIntrZip = rsIntr("Zip Code")
	tmpAvail = rsIntr("Availability")
End If
rsIntr.Close
Set rsIntr = Nothing

'GET ADDRESS AND ZIP of DEPARTMENT/CLIENT
intReqID = Request("RID")

Set oGDM = New acaDistanceMatrix
oGDM.DBCONN = g_strCONN
Call oGDM.FetchMileageV2(intReqID, strItrID, tmpIntrAdd, tmpIntrZip, TRUE)
'// Call oGDM.FetchMileageFromReqID(rsReq("index"), TRUE)
fltRealTT	= oGDM.fltRealTT
fltRealM	= oGDM.fltRealM
fltActTT	= oGDM.fltActTT
tmpDeptAddr = oGDM.ApptAddr 	'= strDstAdr
tmpZipInst	= oGDM.ApptZIP		'= strDstZIP

%>
<html lang="en">
<head>
	<meta charset="utf-8">
	<title>Email Interpreter</title>
	<link rel='stylesheet' href='style.css' type='text/css' >
	<script src="js/jquery-3.3.1.min.js"></script>
</head>
<body onload='document.getElementById("btnOK").disabled = true;'>
		<form id='frmMile' name='frmMile' method='post'>
			<center>
			<table>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right'>Intepreter:</td>
					<td>
						<b><%=GetIntr(Request("selIntr"))%></b>
					</td>
				</tr>
				<tr>
					<td align='right' valign='top'>Availablity:</td>
					<td>
						<textarea readonly><%=tmpAvail%></textarea>
					</td>
				</tr>
				<tr>
					<td align='right'>Mileage:</td>
					<td>
						<input class='main' size='5' readonly name='txtMile' value="<%=fltRealM%>">&nbsp;miles
					</td>
				</tr>
				<tr>
					<td align='right'>Travel Time:</td>
					<td>
						<input class='main' size='5' readonly name='txtTravel'  value="<%=fltRealTT%>">&nbsp;hrs
					</td>
				</tr>	
				<tr>
					<td colspan='2' align='center'>
						<input class="btn" type="button" name="btnOK" id="btnOK" value="OK" style="width: 100px;" disabled
								onmouseover="this.className='hovbtn'"
								onmouseout="this.className='btn'"
								/>
						<input class='btn' type='button' value='Back' style='width: 100px;' onclick='document.location="emailIntr.asp?ID=<%=Request("RID")%>";'
								onmouseover="this.className='hovbtn'"
								onmouseout="this.className='btn'"
								/>
					</td>
				</tr>
			</table>
			
			<input type='hidden' name='selIntr' value='<%=Request("selIntr")%>' 	/>
			<input type='hidden' name='adr1'  value='<%=tmpIntrAdd%>' 				/>
			<input type='hidden' name='adr2'  value='<%=tmpDeptaddr%>' 				/>
			<input type='hidden' name='zip1'  value='<%=tmpIntrZip%>' 				/>
			<input type='hidden' name='zip2'  value='<%=tmpZipInst%>' 				/>
			<input type='hidden' name='ID'  value='<%=intReqID%>' 					/>
			<tr>
									<td valign="top"><div id="output" style="display: none;"></div></td>
								</tr>
		</form>
	</body>
</html>
<!-- #include file="_closeSQL.asp" -->
<script>
	$('#btnOK').click(function(){
		$('#frmMile').attr('action', 'emailIntr.asp');
		$('#frmMile').submit();
	});
	$(document).ready(function(){
		console.log("ready");
		$('#btnOK').prop('disabled', false);
	})
	
</script>