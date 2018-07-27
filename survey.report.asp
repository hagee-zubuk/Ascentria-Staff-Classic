<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Set rsSurv = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT y.[index]" & _
	", y.[rdoPunct], y.[rdoProfb], y.[rdoProcG], y.[rdoTeamW], y.[rdoProDv], y.[rdoReliasTrng]" & _
	", y.[txtGoals], y.[txtStrengths], y.[txtImprovement], y.[txtComments]" & _
	", y.[iid]" & _
	", COALESCE(m.[index], 0) AS [med_ix]" & _
	", i.[First Name] + ' ' + i.[Last Name] AS [inter_name]" & _
	"FROM [survey2018]				AS y " & _
	"INNER JOIN [interpreter_T]		AS i ON y.[iid]=i.[index] " & _
	"INNER JOIN [user_T]			AS u ON y.[uid]=u.[index] " & _
	"LEFT JOIN  [survey2018med]		AS m ON y.[iid]=m.[iid] " & _
	"WHERE y.[iid]=" & lngID
rsSurv.Open strSQL, g_strCONN, 3, 1
If rsSurv.EOF Then
	rsSurv.Close
	Set rsSurv = Nothing
	Session("MSG") = "survey response index was not found"
	Response.Redirect "survey.list.asp"
End If
lngIdx = 0
avgPunct = 0
avgProfb = 0
avgProcG = 0
avgTeamW = 0
avgProDv = 0
txtGoals = ""
txtStrng = ""
txtImprv = ""
txtComnt = ""
avgReliasTrng = "N"
Do While Not rsSurv.EOF
	txtInterpreter = rsSurv("inter_name")
	avgPunct = avgPunct + Z_CLng(rsSurv("rdoPunct"))
	avgProfb = avgProfb + Z_CLng(rsSurv("rdoPunct"))
	avgProcG = avgProcG + Z_CLng(rsSurv("rdoProcG"))
	avgTeamW = avgTeamW + Z_CLng(rsSurv("rdoTeamW"))
	avgProDv = avgProDv + Z_CLng(rsSurv("rdoProDv"))
	If rsSurv("rdoReliasTrng") = "Y" Then avgReliasTrng = "Y"
	If Len(Z_FixNull(rsSurv("txtGoals"))) > 0 Then txtGoals = txtGoals & rsSurv("txtGoals") & vbCrLf
	If Len(Z_FixNull(rsSurv("txtStrengths"))) > 0 Then txtStrng = txtStrng & rsSurv("txtStrengths") & vbCrLf
	If Len(Z_FixNull(rsSurv("txtImprovement"))) > 0 Then txtImprv = txtImprv & rsSurv("txtImprovement") & vbCrLf
	If Len(Z_FixNull(rsSurv("txtComments"))) > 0 Then txtComnt = txtComnt & rsSurv("txtComments") & vbCrLf

	lngMedIx = CLng(rsSurv("med_ix"))
	' iterate!
	rsSurv.MoveNext
	lngIdx = lngIdx + 1
Loop
rsSurv.Close
Set rsSurv = Nothing

If lngIdx <= 0 Then
	Session("MSG") = "not enough survey resources to create a report -- must have at least one!"
	Response.Redirect "survey.list.asp"
End If

avgPunct = avgPunct / lngIdx
avgProfb = avgProfb / lngIdx
avgProcG = avgProcG / lngIdx
avgTeamW = avgTeamW / lngIdx
avgProDv = avgProDv / lngIdx

styPunct = Int(avgPunct)
styProfb = Int(avgProfb)
styProcG = Int(avgProcG)
styTeamW = Int(avgTeamW)
styProDv = Int(avgProDv)

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
	<!-- script langauge="javascript" type="text/javascript" src="js/jquery-ui.min.js"></script -->
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
			<div class="no-print u-full-width">
				<a href="survey.list.asp" title="go back to the list of responses">&lt;&lt;&nbsp;back</a>
			</div>
	<div class="row" id="intrbar">
		<div class="five columns">
			<b>Interpreter Name</b>&nbsp;&nbsp;<div style="display: inline-block;font-weight: bold; font-size: 150%;"><%=txtInterpreter%></div>
		</div>
		<div class="seven columns align-right no-print">
<%
If lngMedIx > 0 Then
%>
			<button type="button" class="button button-secondary"id="btnMedFm" name="btnMedFm">Med Competency Checklist</button>
<%
End If
%>
			<button type="button" class="button button-primary"id="btnRelease" name="btnRelease">Release</button>
		</div>
	</div>
	<div class="row">
		<div class="twelve columns">
			<table class="u-full-width">
  				<thead>
    				<tr><th colspan="2" class="yellow">Performance Criteria</th></tr>
  				</thead>
  				<tbody>
  					<tr><td><h5>Punctuality</h5>
							</td>
						<td class="resp rr<%=styPunct%>"><%=avgPunct%></td>
					</tr>
					<tr><td><h5>Professional Behavior</h5>
							</td>
						<td class="resp rr<%=styProfb%>"><%=avgProfb%></td>
					</tr>
					<tr><td><h5>Adherence to LB Procedural Guidelines</h5>
							</td>
						<td class="resp rr<%=styProcG%>"><%=avgProcG%></td>
					</tr>
					<tr><td><h5>Team Work Ethics</h5>
							</td>
						<td class="resp rr<%=styTeamW%>"><%=avgTeamW%></td>
					</tr>
					<tr><td><h5>Professional Development</h5></td>
						<td class="resp rr<%=styProDv%>"><%=avgProDv%></td>
					</tr>
				</tbody>
			</table>
			Completed the required trainings in Relias (Yes or No):  <div class="resp"><%=avgReliasTrng%></div>
			<br /><br />
			<p>
				<label>Goals:</label>
				<pre class="resp"><%=txtGoals%></pre>
			</p>
			<p>
				<label>Existing Strengths:</label>
				<pre class="resp"><%=txtStrng%></pre>
			</p>
			<p>
				<label>Areas Needing Improvement:</label>
				<pre class="resp"><%=txtImprv %></pre>
			</p>
			<p>
				<label>Comments:</label>
				<pre class="resp"><%=txtComnt %></pre>
			</p>
  		</div>
  	</div>
  	<div class="row">
		<div class="twelve columns align-right no-print">
  			<button type="button" class="button button-primary"id="btnClos" name="btnClos">Back</button>
  		</div>
	</div>
</div>
</body>
</html>
<script language="javascript" type="text/javascript"><!--
$( document ).ready(function() {
	$('#btnClos').click(function(){
		document.location="survey.list.asp";
	});
	$('#btnMedFm').click(function() {
		document.location="survey2018-medical.asp?iid=<%=lngID%>";
	});
	$('#intrbar').sticky({topSpacing:0});
	console.log( "ready!" );
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
	alert("<%=tmpMSG%>");
<%
End If
%>
});
// --></script>