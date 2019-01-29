<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function SelClass(idx, cls)
	SelClass = ""
	If idx = cls Then SelClass = " selected "
End Function
Function Z_YMDDate(dtDate)
	DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	Z_YMDDate = DatePart("yyyy", dtDate) & "-"
	lngTmp = Z_CLng(DatePart("m", dtDate))
	If lngTmp < 10 Then Z_YMDDate = Z_YMDDate & "0"
	Z_YMDDate = Z_YMDDate & lngTmp & "-"
	lngTmp = Z_CLng(DatePart("d", dtDate))
	If lngTmp < 10 Then Z_YMDDate = Z_YMDDate & "0"
	Z_YMDDate = Z_YMDDate & lngTmp
End Function

	'strSQL = "SELECT TOP 30 [timestamp], [filename], [staff], [rid], [type], [uid] FROM [uploads] WHERE [type]=0 ORDER BY [uid] DESC"
	strSCol = Z_FixNull( Request("sCol") )
	strSDir = Z_FixNull( Request("sDir") )
	strZTag = Z_FixNull( Request("fZTg") )
	If (strSDir = "") Then strSDir = "DESC"
	If (strSCol = "") Then strSCol = "upl.[timestamp]"
	If (strZTag = "") Then
		strLabel = "Latest Uploads"
		strCnt = "TOP 10"
		strFilt = " AND upl.[type]=0 "
	Else
		strLabel = "Search Results"
		strCnt = ""
		strFilt = ""
		tmpTod8 = ""
		tmpFromd8 = ""
		tmpUTod8 = ""
		tmpUFromd8 = ""
		If Z_FixNull(Request("txtFromd8")) <> "" Then
			tmpFromd8 = Z_CDate(Z_FixNull(Request("txtFromd8")))
			If tmpFromd8 > Z_CDate("1/1/2010") Then
				If Z_FixNull(Request("tmpTod8")) <> "" Then
					tmpTod8 = Z_CDate(Z_FixNull(Request("txtTod8")))
					strFilt = strFilt & " AND req.[appDate] <= '" & Z_YMDDate(tmpTod8) & "' "
				End If
				strFilt = strFilt & " AND req.[appDate] >= '" & Z_YMDDate(tmpFromd8) & "' "
			End If
		Else
			If Z_FixNull(Request("tmpTod8")) <> "" Then 'tmpFromd8 is empty
				tmpTod8 = Z_CDate(Z_FixNull(Request("txtTod8")))
				If tmpTod8 > Z_CDate("1/1/2010") Then
					strFilt = strFilt & " AND req.[appDate] <= '" & Z_YMDDate(tmpTod8) & "' "
				End If
			End If
		End If
		If Z_FixNull(Request("txtUFromd8")) <> "" Then
			tmpUFromd8 = Z_CDate(Z_FixNull(Request("txtUFromd8")))
			If tmpUFromd8 > Z_CDate("1/1/2010") Then
				strFilt = strFilt & " AND upl.[timestamp] >= '" & Z_YMDDate(tmpUFromd8) & "' "
				If Z_FixNull(Request("tmpUTod8")) <> "" Then
					tmpUTod8 = Z_CDate(Z_FixNull(Request("txtUTod8")))
					strFilt = strFilt & " AND upl.[timestamp] <= '" & Z_YMDDate(tmpUTod8) & "' "
				End If	
			End If
		Else
			If Z_FixNull(Request("tmpUTod8")) <> "" Then 'tmpFromd8 is empty
				tmpUTod8 = Z_CDate(Z_FixNull(Request("txtUTod8")))
				If tmpUTod8 > Z_CDate("1/1/2010") Then
					strFilt = strFilt & " AND upl.[timestamp] <= '" & Z_YMDDate(tmpUTod8) & "' "
				End If
			End If
		End If
	End If

	lngInst = 0
	If Z_FixNull(Request("selInst")) <> "" Then
		lngInst = Z_CLng(Request("selInst"))
		If lngInst > 0 Then
			strFilt = strFilt & " AND req.[InstID]  = " & lngInst
		End If
	End If
	lngDept = 0
	If Z_FixNull(Request("selDept")) <> "" Then
		lngDept = Z_CLng(Request("selDept"))
		If lngDept > 0 Then
			strFilt = strFilt & " AND req.[DeptID]  = " & lngDept
		End If
	End If
	lngLang = 0
	If Z_FixNull(Request("selLang")) <> "" Then
		lngLang = Z_CLng(Request("selLang"))
		If lngLang > 0 Then
			strFilt = strFilt & " AND req.[LangID]  = " & lngLang
		End If
	End If
	lngIntr = 0
	If Z_FixNull(Request("selIntr")) <> "" Then
		lngIntr = Z_CLng(Request("selIntr"))
		If lngIntr > 0 Then
			strFilt = strFilt & " AND req.[IntrID]  = " & lngIntr
		End If
	End If
	lngClass = 0
	If Z_FixNull(Request("selClass")) <> "" Then
		lngClass = Z_CLng(Request("selClass"))
		If lngClass > 0 Then
			strFilt = strFilt & " AND dep.[class]  = " & lngClass
		End If
	End If
	
	strSQL = "SELECT " & strCnt & " upl.*" & _
			", itr.[First Name], itr.[Last Name]" & _
			", dep.[dept], ins.[facility] AS [institution]" & _
			", req.[appdate], req.[apptimefrom], req.[apptimeto]" & _
			", req.[cfname], req.[clname], lan.[language] " & _
			"FROM [langbankuploads].dbo.[uploads] AS upl " & _
			"INNER JOIN [langbank].dbo.[request_T] AS req ON upl.[RID]=req.[index] " & _
			"INNER JOIN [langbank].dbo.[dept_T] AS dep ON req.[deptid] = dep.[index] " & _
			"INNER JOIN [langbank].dbo.[institution_T] AS ins ON req.[InstID] = ins.[index] " & _
			"INNER JOIN [langbank].dbo.[interpreter_T] AS itr ON req.[intrid] = itr.[index] " & _
			"INNER JOIN [langbank].dbo.[language_T] AS lan ON req.[langid] = lan.[index] " & _
			"WHERE upl.[uid]>0 " & strFilt  & _
			" ORDER BY " & strSCol & " " & strSDir
	'Response.Write strSQL
	Set rsUploads = Server.CreateObject("ADODB.Recordset")			
	rsUploads.Open strSQL, g_strCONNupload, 3, 1
	lngC = 0
	Do While (Not rsUploads.EOF) And (lngC < 50)
		strVform = strVform & "<tr><td>&nbsp;&nbsp;&nbsp;<img src=""images/zoom.gif"" alt=""Q"" title=""view files"" " & _
				"onclick=""listfiles(" & rsUploads("rid") & ");"" />" & rsUploads("rid") & "</td>" 
		strVform = strVform & "<td style=""text-align: center;""><div id=""vw" & rsUploads("uid") & _
				""" class=""ico_view"" onclick=""viewfile(" & rsUploads("uid") & ");"">" &  _
				"<img src=""images/zzz-dl.png"" title=""view " & rsUploads("uid") & """ "
		strVform = strVform & "alt=""[]"" /></div>" & FormatDateTime(rsUploads("timestamp"), 2) & "<br />" & _
				FormatDateTime(rsUploads("timestamp"), 4) & "</td>" & _
				"<td>" & rsUploads("first name") & " " & rsUploads("last name") & "/ "
		strVform = strVform & rsUploads("language") & "</td>" & _
				"<td>" & rsUploads("institution") & "<br />" & rsUploads("dept") & "</td>" 
		strVform = strVform & "<td>" & UCase(rsUploads("cfname") & " " & rsUploads("clname")) & "</td>" & _
				"<td>" & FormatDateTime(rsUploads("appdate"), 2) & " " & FormatDateTime(rsUploads("apptimefrom"), 4) & "</td>" & _
				"</tr>" & vbCrLf
		rsUploads.MoveNext
		lngC = lngC + 1
	Loop
	rsUploads.Close
	Set rsUploads = Nothing

strDebug = ""'strSQL

' ******  initialize lookups '
	Set rsTmp = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT [index], [Language] FROM [language_T] ORDER BY [Language]"
	strLang = ""
	rsTmp.Open strSQL, g_strCONN, 3, 1
	Do Until rsTmp.EOF
		strLang = strLang & "<option value=""" & rsTmp("index") & """"
		If rsTmp("index") = lngLang Then strLang = strLang & " selected"
		strLang = strLang & ">" & rsTmp("Language") & "</option>" & vbCrLf
		rsTmp.MoveNext
	Loop
	rsTmp.Close

	strInst = ""
	strSQL = "SELECT [Index], [Facility] FROM [langbank].dbo.[institution_T] WHERE [active]=1 ORDER BY [Facility]"
	rsTmp.Open strSQL, g_strCONN, 3, 1
	Do Until rsTmp.EOF
		strInst = strInst & "<option value=""" & rsTmp("index") & """"
		If rsTmp("index") = lngInst Then strInst = strInst & " selected"
		strInst = strInst & ">" & rsTmp("Facility") & "</option>" & vbCrLf
		rsTmp.MoveNext
	Loop
	rsTmp.Close

	strIntr = ""
	strSQL = "SELECT [Index], [First Name] + ' ' + [Last Name] AS [name] FROM [langbank].dbo.[interpreter_T] ORDER BY [First Name]"
		rsTmp.Open strSQL, g_strCONN, 3, 1
	Do Until rsTmp.EOF
		strIntr = strIntr & "<option value=""" & rsTmp("index") & """"
		If rsTmp("index") = lngIntr Then strIntr = strIntr & " selected"
		strIntr = strIntr & ">" & rsTmp("name") & "</option>" & vbCrLf
		rsTmp.MoveNext
	Loop
	rsTmp.Close
	If lngInst > 0 And lngDept > 0 Then
		strDept = "<option value=""0""></option>"
		strSQL = "SELECT [index], [dept] FROM [Dept_T] WHERE [InstID]=" & lngInst & " ORDER BY [dept]"
		rsTmp.Open strSQL, g_strCONN, 3, 1
		Do Until rsTmp.EOF
			strDept = strDept & "<option value=""" & rsTmp("index") & """"
			If rsTmp("index") = lngDept Then strDept = strDept & " selected"
			strDept = strDept & ">" & rsTmp("dept") & "</option>" & vbCrLf
			rsTmp.MoveNext
		Loop
		rsTmp.Close
	Else
		strDept = "<option value=""-1"">** select institution first **</option>"
	End If
	Set rsTmp = Nothing

%>
<!doctype html>
<html lang="en">
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width,initial-scale=1">
	<title>V-Forms QA/QC</title>
	<meta name="description" content="Verification Form QA">
	<meta name="author" content="Hagee@zubuk">
 	<link rel="stylesheet" type="text/css" href="css/normalize.css" />
 	<link rel="stylesheet" type="text/css" href="css/skeleton.css" />
 	<link rel="stylesheet" type="text/css" href="css/jquery-ui.min.css" />
	<link rel="stylesheet" type="text/css" href="style.css" />
	<script langauge="javascript" type="text/javascript" src="js/jquery-3.3.1.min.js"></script>
	<script langauge="javascript" type="text/javascript" src="js/jquery-ui.min.js"></script>
  <!--[if lt IE 9]>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html5shiv/3.7.3/html5shiv.js"></script>
  <![endif]-->
	<style>
div.formatsmall { display: inline-block; }
.ui-autocomplete-loading { background: white url("images/ui-anim_basic_16x16.gif") right center no-repeat; }
.smallertable th { text-align: center; }
.smallertable td { padding-top: 3px; padding-bottom: 7px; vertical-align: text-top; }
.smallertable td:first-child	{ text-align: center; font-size: 130%; }
.smallertable td:last-child		{ text-align: center; }
table.filtertable { background-color: #fbeec7; padding: 5px; border-radius: 4px; width: 100%; }
.filtertable td { padding: 1px 5px 2px; vertical-align: text-top; }
.ico_view { display: inline-block; float: left; }
	</style>

</head>
<body>
<!-- #include file="_header.asp" -->
<div class="container">
	<div class="row">
	</div>

	<div class="row" style="margin-top: 40px;">
		<div class="twelve columns">
			<h4 id="lblResults"><%= strLabel %></h4>
<%
If Session("MSG") <> "" Then
	Response.Write "<div class=""err"">" & Session("MSG") & "</div>"
	Session("MSG") = ""
End If
%>
			<table class="u-full-width smallertable" id="results">
  				<thead>
    				<tr><th class="hdrSort" id="colReq" value="index">Request ID</th>
    					<th class="hdrSort" id="colUpl" value="timestamp" style="font-size: 80%;">Upload<br />Time/Date</th>
						<th class="hdrSort" id="colItr" value="first name">Interpreter/Language</th>
						<th class="hdrSort" id="colDep">Institution/Department</th>
						<th class="hdrSort" id="colCli">Client</th>
						<th class="hdrSort" id="colApp" style="font-size: 80%;">Appointment<br />Date</th>
						</tr>
  				</thead>
  				<tbody>
<%=strVform%>
				</tbody>
			</table>
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="eleven columns">
<form name="frmUploads" id="frmUploads" action="vform.qc.asp" method="post">
<input type="hidden" name="sCol" id="sCol" />
<input type="hidden" name="sDir" id="sDir" />
<input type="hidden" name="fFld" id="fFld" />
<input type="hidden" name="fSrc" id="fSrc" />
<input type="hidden" name="fZTg" id="fZTg" value="1" />

			<table class="filtertable"><tbody>
				<tr><td colspan="4" style="text-align: center;"><h5>Filter</h5></td>
					<td><!-- button class="button button-secondary" name="btnUpload" id="btnUpload" >VForm Upload</button --></td></tr>
				<tr><td>
					&nbsp;<b>Appt. Date Range:</b>
					</td><td>
					<input size="10" maxlength="10" name="txtFromd8" value="<%=tmpFromd8%>" />
					&nbsp;-&nbsp;
					<input size="10" maxlength="10" name="txtTod8" value="<%=tmpTod8%>" />
					<div class='formatsmall'>mm/dd/yyyy</div>
					</td><td>
					&nbsp;<b>Institution:</b>
					</td><td>
						<select style="width: 150px;" name="selInst" id="selInst">
							<option value='-1'>&nbsp;</option>
<%=strInst%>
						</select>
					</td><td rowspan="4">
						<button class="button button-primary" name="btnApply" id="btnApply" style="margin-top: 20px;">Apply</button>
					</td></tr>
				<tr><td>
					&nbsp;<b>Request ID Range:</b>
					</td><td>
					<input size="7" maxlength="7" name="txtFromID" value="<%=tmpFromID%>" />
					&nbsp;-&nbsp;
					<input size="7" maxlength="7" name="txtToID" value="<%=tmpToID%>" />
					</td><td>
					&nbsp;<b>Department:</b>
					</td><td>
						<select style="width: 150px;" name="selDept" id="selDept">
<%=strDept%>
						</select>
					<input type="hidden" name="txtInst" id="txtInst" readonly value="<%=tmpdepttxt%>" />
					</td></tr>
				<tr><td>
					&nbsp;<b>Upload Date Range:</b>
					</td><td>
					<input size="10" maxlength="10" name="txtUFromd8" value="<%=tmpUFromd8%>" />
					&nbsp;-&nbsp;
					<input size="10" maxlength="10" name="txtUTod8" value="<%=tmpUTod8%>" />
					<div class='formatsmall'>mm/dd/yyyy</div>
					</td><td>
					&nbsp;<b>Language:</b>
					</td><td>
					<select style="width: 150px;" name="selLang">
						<option value='-1'>&nbsp;</option>
<%=strLang%>
					</select>
				</td></tr>
				<tr><td>
					&nbsp;<b>Interpreter:</b>
					</td><td>
					<select name="selIntr" >
						<option value='-1'>&nbsp;</option>
<%=strIntr%>
					</select>
					</td><td>
					&nbsp;<b>Classification:</b>
					</td><td>
					<select style="width: 100px;" name="selClass" id="selClass">
						<option value='-1'>&nbsp;</option>
						<option value='1' <%=SelClass(1, lngClass)%> >Social Services</option>
						<option value='2' <%=SelClass(2, lngClass)%> >Private</option>
						<option value='3' <%=SelClass(3, lngClass)%> >Court</option>
						<option value='4' <%=SelClass(4, lngClass)%> >Medical</option>
						<option value='5' <%=SelClass(5, lngClass)%> >Legal</option>
						<option value='6' <%=SelClass(6, lngClass)%> >Mental Health</option>
					</select>
				</td></tr>
			</tbody>
			</table>
</form>	

<%= strDebug%>

		</div>
	</div>
</div><!-- container -->
</body>
</html>
<script language="javascript" type="text/javascript"><!--
function viewfile(zzz) {
	console.log('view the file:' + zzz);
	var zzw = window.open("vform.view.asp?uid="+zzz,
		"viewfile",
		"height=400,width=600,menubar=0,toolbar=0,location=0,personalbar=0,directories=0,status=0,dependent=1,");
}
function listfiles(zzz) {
	var zzw = window.open("viewuploads.asp?reqid="+zzz,
		"viewuploads",
		"height=300,width=830,menubar=0,toolbar=0,location=0,personalbar=0,directories=0,status=0,dependent=1,");
	console.log('list files for RID=' + zzz);
}
$( document ).ready(function() {
	$("#selInst").on("change", function() {
		$("#txtInst").val($("#selInst :selected").text());
		// alert( $("#txtInst").val() + "\n" + this.value );
		var inst = this.value;
		$.get("vform.updatedepts.asp?q="+inst, function(data) {
			$("#selDept option").remove();
    		$("#selDept").html(data);
		});
	});
	
	$("hdrSort").click(function () {

	});
	$("#btnUpload").click(function () { 
		var zzw = window.open("vform.upload.asp",
			"upload_vf",
			"height=300,width=830,menubar=0,toolbar=0,location=0,personalbar=0,directories=0,status=0,dependent=1,");
	});
	$("#btnApply").click(function () {
		$("#frmUploads").submit();
	});
});
// --></script>