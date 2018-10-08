<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
strUsr = Request.Cookies("LBUsrName")
lngUID = Z_CLng(Request.Cookies("UID"))

Set rsSurv = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT y.[index]" & _
		", y.[iid]" & _
		", i.[First Name] + ' ' + i.[Last Name] AS [inter_name]" & _
		", y.[txtDate]" & _
		", y.[uid]" & _
		", y.[ts]" & _
		", COALESCE(r.[guid], '') AS [guid]" & _
		", r.[generated], r.[viewed]" & _
		", COALESCE(r.[release], 0) AS [release]" & _
		", COALESCE(r.[signature], '') AS [signature]" & _
		", COALESCE(m.[index], 0) AS [med_ix]" & _
		", u.[Fname] + ' ' + u.[Lname] AS [reviewer] " & _
	"FROM [survey2018] AS y " & _
		"INNER JOIN [interpreter_T] AS i ON y.[iid]=i.[index] " & _
		"INNER JOIN [user_T]        AS u ON y.[uid]=u.[index] " & _
		"LEFT JOIN  [surveyreports] AS r ON y.[iid]=r.[iid] " & _
		"LEFT JOIN  [survey2018med] AS m ON y.[iid]=m.[iid] " & _
	"ORDER BY i.[First Name], u.[Fname]"
rsSurv.Open strSQL, g_strCONN, 3, 1
strFulTbl = ""
strLast = ""
lngIdx = 0
Do While Not rsSurv.EOF
	strTbl = "<tr><td>"
	blnRel = rsSurv("release")
	If (Z_CDate(rsSurv("signature"))>CDate("2018-01-01")) Then 
		blnSigned = True
	Else
		blnSigned = False
	End If
	
	If strLast <> rsSurv("inter_name") Then
		If ( blnSigned ) Then
			strTbl = strTbl & "<div class=""icon ui-icon-check"" title=""signed""></div></td><td>"
		Else
			strTbl = strTbl & "</td><td>"
		End If
	
		strTbl = strTbl & rsSurv("inter_name")
		strLast = rsSurv("inter_name")
		lngIdx = lngIdx + 1
		If ( blnRel ) Then strTbl = strTbl & " (released)"
		If ( blnSigned ) Then strTbl = strTbl & "<br />(signed: " & Z_MDYDate( rsSurv("signature")) & ")"
		strTbl = strTbl & "</td><td style=""text-align: center;"">"

		strTbl = strTbl & "<a href=""survey.print.asp?ix=" & rsSurv("iid") & _
					""" title=""PDF summary""><img src=""images/pdf-dl.png"" " & _
					"alt=""PDF"" /></a>"
		
		strTbl = strTbl & "</td><td style=""text-align: center;"">"
		strTbl = strTbl & "<a href=""survey.report.asp?ix=" & rsSurv("iid") & _
				""" title=""view summary""><div class=""icon ui-icon-note""></div></a>"
		' check if released or not
		strTbl = strTbl & "<div id=""divRel" & rsSurv("iid") & """>"
		If ( blnRel ) Then
			strTbl = strTbl & "<div class=""icon ui-icon-check""></div>"
		Else
			strTbl = strTbl & "<a href=""#"" class=""font-small"" id=""lnkRel"" onclick=""release('" & rsSurv("iid") & "')"">release</a></div>"
		End If
		strTbl = strTbl & "</td><td style=""text-align: center;""><a title=""Medical Checklist"
		If ( CLng(rsSurv("med_ix")) > 0 ) Then strTbl = strTbl & " completed"
		strTbl = strTbl & """ href=""survey2018-medical.asp?iid=" & rsSurv("iid") & """><img src=""images/medical.png"" " & _
				"alt=""Med"" /><div class=""icon ui-icon-"
		If ( CLng(rsSurv("med_ix")) > 0 ) Then
			strTbl = strTbl & "check"
		Else
			strTbl = strTbl & "help"
		End If
		strTbl = strTbl &"""></div></a>"
	Else
		strTbl = strTbl & "</td><td>&mdash;</td><td colspan=""2"">&nbsp;</td><td>"
	End If

	If (Z_IsOdd(lngIdx)) Then
		strTbl = Replace(strTbl, "tr", "tr class=""yellow""")
	End If
	strTbl = strTbl & "</td><td>" & rsSurv("ts") & "</td>"
	
	blnZZ =  CLng(rsSurv("uid"))'' - lngUID )
	strTbl = strTbl & "<td>" & rsSurv("reviewer") '& " ([" & rsSurv("uid") & "] =? [" & lngUID & "]) ~~ " & blnZZ

'	If (CLng(rsSurv("uid")) = lngUID) Then
'		strTbl = strTbl & "<a href=""survey.edit.asp?ix=" & rsSurv("index") & """ title=""Edit""><div class=""icon ui-icon-pencil""></div></a>"
'	End If
	strTbl = strTbl & "</td><td><a href=""survey.view.asp?ix=" & rsSurv("index") & """ title=""View this entry""><div class=""icon ui-icon-document""></div></a></td><td>"
	If (Not blnRel ) Then
		strTbl = strTbl & "<a href=""survey.dele.asp?ix=" & rsSurv("index") & _
				""" title=""Delete this entry""><div class=""icon ui-icon-trash""></div></a>"
	End If
	strTbl = strTbl & "</td></tr>" & vbCrLF
	rsSurv.MoveNext
	strFulTbl = strFulTbl & strTbl
Loop
%>
<!doctype html>
<html lang="en">
<head>
	<meta charset="utf-8">
	<meta name="viewport" content="width=device-width,initial-scale=1">
	<title>Interpreter Survey</title>
	<meta name="description" content="LanguageBank Internal Interpreter Survey 2018">
	<meta name="author" content="Hagee@zubuk">
 	<link rel="stylesheet" href="css/normalize.css" />
 	<link rel="stylesheet" href="css/skeleton.css" />
 	<link rel="stylesheet" href="css/jquery-ui.min.css" />
	<link rel="stylesheet" href="css/survey.css" />
	<script langauge="javascript" type="text/javascript" src="js/jquery-3.3.1.min.js"></script>
	<script langauge="javascript" type="text/javascript" src="js/jquery-ui.min.js"></script>
  <!--[if lt IE 9]>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html5shiv/3.7.3/html5shiv.js"></script>
  <![endif]-->
	<style>
.ui-autocomplete-loading { background: white url("images/ui-anim_basic_16x16.gif") right center no-repeat; }
	</style>
</head>
<body>
<div class="container">
	<div class="row">
		<div class="twelve columns" id="logobar">
			<img id="logo" src="images/lb-logo.jpg" alt="The Language Bank" title="" />
			<h1>Interpreter Performance Evaluation</h1>
		</div>
	</div>
	<div class="row" style="margin-top: 40px;">
		<div class="twelve columns">
<%
If Session("MSG") <> "" Then
	Response.Write "<div class=""err"">" & Session("MSG") & "</div>"
	Session("MSG") = ""
End If
%>
			<table class="u-full-width smallertable">
  				<thead>
    				<tr><th colspan="2">Interpreter</th>
    					<th colspan="3" style="text-align: center; font-size: 75%;">Summary</th>
    					<th>Date</th>
    					<th>Reviewer</th>
    					<th colspan="2">&nbsp;</th></tr>
  				</thead>
  				<tbody>
<%=strFulTbl%>
				</tbody>
			</table>
		</div>
	</div>
	<div class="row">
		<div class="one column">&nbsp;</div>
		<div class="eleven columns">
<a href="survey2018.asp" target="_BLANK">Open Survey Form</a>
		</div>
	</div>
</div><!-- container -->
</body>
</html>
<script language="javascript" type="text/javascript"><!--
function release(vIID) {
	if(!(vIID > 0)) {
		console.log("Yay! I'm exiting instead.");
	}
	console.log("Releasing: " + vIID);
	if(confirm("Click OK to release results to interpreter")) {
		document.location="survey.release.asp?iid="+vIID;
	}
}
$( document ).ready(function() {
});
// --></script>