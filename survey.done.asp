<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
strMesg = Session("SURVEY")
Session("SURVEY") = ""

lngID = Z_CLng(Request("iid"))

blnGotMed = TRUE
If lngID > 0 Then
	Set rsSurv = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT COUNT([index]) AS [cnt] FROM [survey2018med] WHERE [iid]=" & lngID ' UID doesn't matter in this case -- " AND [uid]=" & lngUID
	rsSurv.Open strSQL, g_strCONN, 1, 3
	blnGotMed = FALSE
	If Not rsSurv.EOF Then
		cnt = rsSurv("cnt")
		If cnt > 0 Then
			blnGotMed = TRUE
		End If
	End If
	rsSurv.Close
	Set rsSurv = Nothing
End If
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
	<div class="row" style="margin-top: 50px;">
		<div class="twelve columns">
			<p><%=strMesg%></p>
			<h4>Thank you for your response. You may close this window or <a href="survey2018.asp">click here to fill in another one</a>.</h4>
<%
If Not blnGotMed Then
%>
<h4>No Medical Checklist form was found for this user. <a href="survey2018-medical.asp?iid=<%=lngID%>">Click here to fill one in</a>.</h4>
<%
End If
%>
			<p>back to <a href="survey.list.asp">survey list</a></p>
		</div>
	</div>
</div>
</body>
</html>