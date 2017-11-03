<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
Function CleanMe(xxx)
	If xxx = "" Then Exit Function
	If IsNull(xxx) Then Exit Function
	tmpString = Replace(xxx, "'", " ")
	CleanMe = Replace(tmpString, ",", " ")
End Function
'GET INSTITUTION LIST
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM institution_T ORDER BY Facility"
rsReq.Open sqlReq, g_strCONN, 3, 1
Do Until rsReq.EOF
	tmpInst = CleanMe(rsReq("facility"))
	tmpName = rsReq("Index") & " -- " & tmpInst
	jscriptArr = jscriptArr & "'" & tmpName & "' , " 
	rsReq.MoveNext
Loop
rsReq.Close
Set rsReq = Nothing
tmpTitle = "Institution"
'CLEAN ARRAY
If jscriptArr <> "" Then jscriptArr =  Trim(Left(jscriptArr, Len(jscriptArr) - 2))
%>
<html>
	<head>
		<title>Language Bank - Find <%=tmpTitle%></title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language="javascript"  type="text/javascript" src="AutoComplete.js"></script>
		<script language="javascript"  type="text/javascript" src="common.js"></script>
		<script language='JavaScript'>
		<!--
		function PassMe(xxx)
		{
			var ReqID = xxx.split(" -- ");
			window.opener.document.frmMain.selInst.value = ReqID[0];
			window.opener.document.frmMain.selInst.focus();
			self.close();
		}
		var customarray = new Array(<%=jscriptArr%>);
		
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmFind' action='javascript: PassMe(document.frmFind.txtFind.value);'>
			<table cellpadding='0' cellspacing='0' border='0' align='left' height='95%' width='100%'>
				<tr>
					<td height='25px'>&nbsp;</td>
					<td class='header' colspan='3'><nobr>FIND <%=tmpTitle%> --&gt&gt</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td colspan='2'>
						<input class='main' size='50' maxlength='50' name='txtFind'>
						<input class='btn' type='submit' style='width: 50px;' value='Insert' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='javascript: PassMe(document.frmFind.txtFind.value);'>
					</td>
				</tr>
				<tr>
					<td colspan='3' align='right' valign='bottom'>
						<font size='1'><i><u>* Type first few letters of desired institution.</u></i></font>
					</td>
				</tr>
			</table>
		</form>
		<script>
			var obj = actb(document.frmFind.txtFind, customarray);
		</script>
	</body>
</html>