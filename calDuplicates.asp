<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
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
%>
<%
dtDate = Z_FixNull( Request("dt") )
dtDate = Z_CDate(dtDate)
If dtDate < CDate("1/1/2017") Then
	dtDate = Date
End If
txtDate = Z_YMDDate(dtDate)

sqlDup = "SELECT z.[Clname], z.[Cfname], r.[appTimeFrom], r.[appTimeTo], r.[index], l.[language] " & _
		"FROM (	SELECT COUNT([index]) AS qq, [Clname], [Cfname] " & _
			"FROM [request_T] " & _
			"WHERE [appDate]='" & txtDate & "' " & _
			"AND ([Status]<2 OR [Status]>3) " & _
			"GROUP BY [Clname], [Cfname] " & _
		") AS z " & _
		"INNER JOIN [request_T] AS r ON z.[Clname]=r.[Clname] AND z.[Cfname]=r.[Cfname] AND [appDate]='" & txtDate & "' " & _
		"INNER JOIN [language_T] AS l ON r.[langID]=l.[index] " & _
		"WHERE qq > 1 " & _
		"ORDER BY [clname], [cfname]"
Set rsDup = Server.CreateObject("ADODB.Recordset")
rsDup.Open sqlDup, g_strConn, 3, 1
If Not rsDup.EOF Then
	strResult = ""
	strLastClname = ""
	'strLastTmFrom = CDate("1/1/2017")
	Do While Not rsDup.EOF
		If (strLastClname <> rsDup("Clname")) Then
			strResult = strResult & "<tr><td><b>" & rsDup("Clname") & ", " & rsDup("Cfname") & "</b>"
			strLastClname = rsDup("Clname")
		Else
			strResult = strResult & "<tr><td>&nbsp;"
		End If
		strResult = strResult & "</td><td>" & rsDup("language") & "</td><td>" & _
				Ctime(rsDup("appTimeFrom")) & " - " & _
				Ctime(rsDup("appTimeTo")) & "</td></tr>" & vbCrLf
		rsDup.MoveNext
	Loop
	strResult = "<table style=""width: 100%;""><thead><tr><th>Name</th><th>Language</th>" & _
			"<th>Time</th></tr></thead><tbody>" & strResult & "</tbody></table>"
Else
	strResult = "<p style=""text-align: center;""><i>no results</i></p>"
End If

%>
<!doctype html>
<html lang="en">
<head>
	<meta charset="utf-8">
	<title>Duplicate Appointments</title>
	<meta name="description" content="Possible duplicate appointment list" />
	<meta name="author" content="Hagee" />
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/normalize/8.0.0/normalize.min.css"
			integrity="sha256-oSrCnRYXvHG31SBifqP2PM1uje7SJUyX0nTwO2RJV54=" crossorigin="anonymous" />
	<script src="https://code.jquery.com/jquery-3.3.1.min.js" crossorigin="anonymous"
			integrity="sha256-FgpCb/KJQlLNfOu91ta32o/NMZxltwRo8QtmkMRdAu8="></script>
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/2.0.4/skeleton.min.css"
			integrity="sha256-2YQRJMXD7pIAPHiXr0s+vlRWA7GYJEK0ARns7k2sbHY=" crossorigin="anonymous" />
	<!-- link href='style.css' type='text/css' rel='stylesheet' / -->
</head>
<body>
	<div class="container">
		<div class="row">
			<div class="twelve columns">
				<h5>Possible Duplicates</h5>
<%=strResult%>
			</div>
		</div>
		<div class="row">
			<div class="three columns"><button name="btnClose" id="btnClose" class="button-primary">Close Window</button></div>
		</div>
	</div>
</body>
<script language="javascript"><!--
$(document).ready(function() {
	$('#btnClose').click(function(){ window.close(); });
});
--></script>
</html>