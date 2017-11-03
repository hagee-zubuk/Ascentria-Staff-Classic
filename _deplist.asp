<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	rsDept.Open "SELECT [index] AS deptID, dept FROM dept_T WHERE instID = " & Request("instID") & " ORDER BY dept", g_strCONN, 3, 1
	Do Until rsDept.EOF
	  kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		stroption = stroption & "<tr bgcolor='" & kulay & "' onclick=""PassMe('" & rsDept("deptID") & "','" & rsDept("Dept") & "');""><td align='left'>" & rsDept("Dept") & "</td></tr>" & vbCrLf
	  rsDept.MoveNext
	Loop
	rsDept.Close
	Set rsDept = Nothing 
%>
<html>
	<head>
		<title>Department List</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function PassMe(xxx, yyy) {
			opener.document.frmTbl.selDept.value =  xxx;
			if (yyy == 0) {yyy = '';}
			opener.document.frmTbl.txtDept.value =  yyy;
			self.close();
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmSec' >
			<table align="center" border="0" width="100%">
				<tr>
					<td class='header' colspan='2'>
						<nobr>Department List --&gt&gt
					</td>
				</tr>
				<tr>
					<td>(to select, click on the department)</td>
				</tr>
				<tr>
					<td align="center">
						<table border='0' cellspacing='0' cellpadding='0'>
						<tr bgcolor="#F5F5F5" onclick="PassMe(0,0)"><td align="left"><i>--- remove selection ---</i></td></tr>
							<%=stroption%>
						</table>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align="center">
						
						<input type="button" name="btnClose" value="Close" onclick="self.close();" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
						
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>