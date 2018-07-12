<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<%
lngID = Request("ix")
If lngID < 1 Then
	Session("MSG") = "survey response index is missing"
	Response.Redirect "survey.list.asp"
End If

Set rsSurv = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT TOP 1 y.[index]" & _
	", y.[rdoPunct], y.[rdoProfb], y.[rdoProcG], y.[rdoTeamW], y.[rdoProDv], y.[rdoReliasTrng]" & _
	", y.[txtGoals], y.[txtStrengths], y.[txtImprovement], y.[txtComments]" & _
	", y.[iid]" & _
	", i.[First Name] + ' ' + i.[Last Name] AS [inter_name]" & _
	", y.[txtDate]" & _
	", y.[uid]" & _
	", u.[Fname] + ' ' + u.[Lname] AS [reviewer] " & _
	"FROM [survey2018] AS y " & _
	"INNER JOIN [interpreter_T] AS i ON y.[iid]=i.[index] " & _
	"INNER JOIN [user_T] AS u ON y.[uid]=u.[index] " & _
	"WHERE y.[index]=" & lngID
rsSurv.Open strSQL, g_strCONN, 3, 1
If rsSurv.EOF Then
	Session("MSG") = "survey response index was not found"
	Response.Redirect "survey.list.asp"
End If
%>
<!-- #include file="_Security.asp" -->
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
	<script langauge="javascript" type="text/javascript" src="js/jquery.sticky.js"></script>
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
	<div class="row" id="intrbar">
		<div class="seven columns">
			<label for="txtName">Interpreter Name</label><input name="txtName" id="txtName" readonly="true" value="<%=rsSurv("inter_name")%>"
				tabstop="-1" class="u-full-width" />
			<label for="txtName">Reviewer</label><input name="txtRevw" id="txtRevw" readonly="true" value="<%=rsSurv("reviewer")%>"
				tabstop="-1" class="u-full-width" />
		</div>
		<div class="five columns">
			<label for="txtDate">Date</label><input name="txtDate" id="txtDate" tabstop="-1" readonly="true" value="<%=Z_MDYDate(rsSurv("txtDate"))%>" />
			<p><b>&#x2605;</b>&nbsp;Higher values are better.</p>
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
						<td class="resp rr<%=rsSurv("rdoPunct")%>"><%=rsSurv("rdoPunct")%></td>
					</tr>
					<tr><td><h5>Professional Behavior</h5>
							</td>
						<td class="resp rr<%=rsSurv("rdoProfb")%>"><%=rsSurv("rdoProfb")%></td>
					</tr>
					<tr><td><h5>Adherence to LB Procedural Guidelines</h5>
							</td>
						<td class="resp rr<%=rsSurv("rdoProcG")%>"><%=rsSurv("rdoProcG")%></td>
					</tr>
					<tr><td><h5>Team Work Ethics</h5>
							</td>
						<td class="resp rr<%=rsSurv("rdoTeamW")%>"><%=rsSurv("rdoTeamW")%></td>
					</tr>
					<tr><td><h5>Professional Development</h5></td>
						<td class="resp rr<%=rsSurv("rdoProDv")%>"><%=rsSurv("rdoProDv")%></td>
					</tr>
				</tbody>
			</table>
			Completed the required trainings in Relias (Yes or No):  <div class="resp"><%=rsSurv("rdoReliasTrng")%></div>
			<br /><br />
			<p>
				<label>Goals:</label>
				<pre class="resp"><%=rsSurv("txtGoals")%></pre>
			</p>
			<p>
				<label>Existing Strengths:</label>
				<pre class="resp"><%=rsSurv("txtStrengths")%></pre>
			</p>
			<p>
				<label>Areas Needing Improvement:</label>
				<pre class="resp"><%=rsSurv("txtImprovement")%></pre>
			</p>
			<p>
				<label>Comments:</label>
				<pre class="resp"><%=rsSurv("txtComments")%></pre>
			</p>
  		</div>
  	</div>
  	<div class="row">
		<div class="twelve columns align-right">
  			<button type="button" class="button button-primary"id="btnClos" name="btnClos">Close</button>
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
<%
rsSurv.Close
Set rsSurv = Nothing
%>