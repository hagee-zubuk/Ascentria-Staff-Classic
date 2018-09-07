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
dt1 = DateDiff("d", -14, Date)
dt2 = DateDiff("d", -3, Date)
strDT1 = Z_MDYDate(dt1)
strDT2 = Z_MDYDate(dt2)
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
				</td></tr>
			<tr><td style="vertical-align: top; height: 250px; text-align: center;">
<form action="rep.novform_zz.asp" name="frmMVF" id="frmMVF" method="POST">
	<div id="tblInput" >
				<table cellSpacing="2" cellPadding="0" border="0" style="text-align: center; margin: 2px auto; width: 400px;">
						<tr><td colspan="2"><h1>Missing Verification Form Report</h1></td></tr>
						<tr><td align="right">Start Date:</td><td>
								<input type="text" id="dtStart" name="dtStart" class="pickdate" value="<%=strDT1%>" />
							</td></tr>
						<tr><td align="right">End Date:</td><td>
								<input type="text" id="dtEnd" name="dtEnd" class="pickdate" value="<%=strDT2%>" />
							</td></tr>
						<tr><td>&nbsp;</td><td><button type="button" name="btnGO" id="btnGO">Generate</button></td></tr>
				</table>
	</div>
</form>
<img src="images/ajax-loader.gif" title="Loading" alt="Wait..." style="display: none; text-align: center; width: 66px; height: 66px;" id="imgLoading" />
				</td></tr>
		</table>
		<div style="width: 100%; position: absolute; bottom: 0px;">
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
    			$("#imgLoading").show();
    			$("#tblInput").hide();
    			if (! isValidDate( $("#dtEnd").val() )) {
    				strErrs += "End date is invalid.\n";
    			}
    			if (! isValidDate( $("#dtStart").val() )) {
    				strErrs += "Start date is invalid.\n";
    			}
    			if (strErrs.length > 0) {
    				$("#err").html(strErrs);
    				$("#tblInput").show();
    				$("#imgLoading").hide();
    				return false;
    			}
    			$( "#frmMVF" ).submit();
    		} );
    	console.log( "ready!" );
	});	
</script>