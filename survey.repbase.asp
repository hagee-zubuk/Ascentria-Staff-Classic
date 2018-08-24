<%
lngID = Request("ix")
If lngID < 1 Then
	Session("MSG") = "survey response index is missing"
	Response.Redirect "survey.list.asp"
End If

Set rsSurv = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT y.[index]" & _
	", y.[rdoPunct], y.[rdoProfb], y.[rdoProcG], y.[rdoTeamW], y.[rdoProDv], y.[rdoReliasTrng]" & _
	", y.[txtGoals], y.[txtStrengths], y.[txtImprovement], y.[txtComments]" & _
	", y.[iid]" & _
	", COALESCE(m.[index], 0) AS [med_ix]" & _
	", COALESCE(r.[release], 0) AS [release]" & _
	", COALESCE(r.[signature], '') AS [signature]" & _
	", i.[First Name] + ' ' + i.[Last Name] AS [inter_name]" & _
	"FROM [survey2018]				AS y " & _
	"INNER JOIN [interpreter_T]		AS i ON y.[iid]=i.[index] " & _
	"INNER JOIN [user_T]			AS u ON y.[uid]=u.[index] " & _
	"LEFT JOIN  [survey2018med]		AS m ON y.[iid]=m.[iid] " & _
	"LEFT JOIN  [surveyreports] 	AS r ON y.[iid]=r.[iid] " & _
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
dtSig = ""
avgReliasTrng = "N"
blnRel = False
Do While Not rsSurv.EOF
	blnRel = CBool(rsSurv("release"))
	dtSig = Z_MDYDate( rsSurv("signature") )
	txtInterpreter = rsSurv("inter_name")
	avgPunct = avgPunct + Z_CLng(rsSurv("rdoPunct"))
	avgProfb = avgProfb + Z_CLng(rsSurv("rdoPunct"))
	avgProcG = avgProcG + Z_CLng(rsSurv("rdoProcG"))
	avgTeamW = avgTeamW + Z_CLng(rsSurv("rdoTeamW"))
	avgProDv = avgProDv + Z_CLng(rsSurv("rdoProDv"))
	If rsSurv("rdoReliasTrng") = "Y" Then avgReliasTrng = "Y"
	If Len(Z_FixNull(rsSurv("txtGoals"))) > 0 Then txtGoals = txtGoals & Z_FixNull(rsSurv("txtGoals") ) & vbCrLf
	If Len(Z_FixNull(rsSurv("txtStrengths"))) > 0 Then txtStrng = txtStrng & Z_FixNull(rsSurv("txtStrengths") ) & vbCrLf
	If Len(Z_FixNull(rsSurv("txtImprovement"))) > 0 Then txtImprv = txtImprv & Z_FixNull(rsSurv("txtImprovement") ) & vbCrLf
	If Len(Z_FixNull(rsSurv("txtComments"))) > 0 Then txtComnt = txtComnt & Z_FixNull(rsSurv("txtComments") ) & vbCrLf

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
avgOvral = (avgPunct + avgProfb + avgProcG + avgTeamW + avgProDv) / 5

styPunct = Int(avgPunct)
styProfb = Int(avgProfb)
styProcG = Int(avgProcG)
styTeamW = Int(avgTeamW)
styProDv = Int(avgProDv)

If (Z_CDate(dtSig) < CDate("2018-01-01")) Then dtSig = "_______________"

txtGoals = Replace(txtGoals, " ", vbCrLf)
txtStrng = Replace(txtStrng, " ", vbCrLf)
txtImprv = Replace(txtImprv, " ", vbCrLf)
txtComnt = Replace(txtComnt, " ", vbCrLf)

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
<%=g_strTopBackLink%>
	<div class="row" 
<% If Not g_blnHideCtls Then Response.Write	"id=""intrbar""" %>
	>
		<div class="five columns">
			<b>Interpreter Name</b>&nbsp;&nbsp;<div style="display: inline-block;font-weight: bold; font-size: 150%;"><%=txtInterpreter%></div>
		</div>
<% If Not g_blnHideCtls Then
%>
		<div class="seven columns align-right no-print">
<%
	If lngMedIx > 0 Then
%>
			<button type="button" class="button button-secondary"id="btnMedFm" name="btnMedFm">Med Competency Checklist</button>
<%
	End If
	If Not blnRel Then
%>

			<button type="button" class="button button-primary"id="btnRelease" name="btnRelease">Release</button>
<%
	End If
%>
		</div>
<%
End If
%>		
	</div>
	<div class="row">
		<div class="twelve columns">
			<table class="u-full-width smallertable" style="width: 90%; margin-left: auto; margin-right: auto;">
				<thead></thead><tbody>
				</tbody>
<tr><td colspan="2">Criteria</td><td colspan="2" style="border-left: 1px dotted #999;">Rating Scale</td></tr>
<tr><td valign="top" colspan="2"><p>Punctuality</p>
	<p>Professional&nbsp;Behavior</p>
	<p>Adherence to LB Procedural Guidelines</p>
<!-- /td><td valign="top" -->
	<p>Team Work Ethics</p>
	<p>Professional&nbsp;Development</p>
</td><td colspan="2" style="border-left: 1px dotted #999;">
	<p>Outstanding - 4<br />
	Employee <i>consistently exceeds</i>the expectations of position. Their colleages recognize their exceelence and their unique contribution to the organization. They serve as a role model for others. They require little or no supervision and generate output that is exceptionally high in quality and quantity. They accept high level or responsibility for own performance.</p>
	<p>Above Average - 3<br />
	Employee <i>frequently exceeds</i> expectations, provides significant and measureable contribution well beyond their position responsibilities. The employee demonstrates a desire and ability to excel in performance.</p>
	<p>Satisfactory - 2<br />
	Employee <i>meets</i> expectations. The employee is a productive and valued member of the team.</p>
	<p>Needs Improvement - 1<br />
	Employee is <i>struggling</i> to meet the basic responsibilty of their position and is not meeting job expectations or the employee is new in their position and is still developing.</p>
</td></tr>
			</table>
		</div>
	</div>
	<div class="row">
		<div class="twelve columns">
			<table class="u-full-width smallertable">
  				<thead>
    				<tr><th colspan="2" class="yellow"><h4>Performance Criteria</h4></th></tr>
  				</thead>
  				<tbody>
  					<tr><td><p style="font-size: 11pt; font-weight: bold; margin-left: 20px;">Punctuality</p>
							</td>
						<td class="resp rr<%=styPunct%>"><%=avgPunct%></td>
					</tr>
					<tr><td><p style="font-size: 11pt; font-weight: bold; margin-left: 20px;">Professional Behavior</p>
							</td>
						<td class="resp rr<%=styProfb%>"><%=avgProfb%></td>
					</tr>
					<tr><td><p style="font-size: 11pt; font-weight: bold; margin-left: 20px;">Adherence to LB Procedural Guidelines</p>
							</td>
						<td class="resp rr<%=styProcG%>"><%=avgProcG%></td>
					</tr>
					<tr><td><p style="font-size: 11pt; font-weight: bold; margin-left: 20px;">Team Work Ethics</p>
							</td>
						<td class="resp rr<%=styTeamW%>"><%=avgTeamW%></td>
					</tr>
					<tr><td><p style="font-size: 11pt; font-weight: bold; margin-left: 20px;">Professional Development</p></td>
						<td class="resp rr<%=styProDv%>"><%=avgProDv%></td>
					</tr>
					<tr><td>Completed the required trainings in Relias:  <div class="resp"><%=avgReliasTrng%></div>
						</td><td></td></tr>
					<tr><td><p style="font-size: 11pt; font-weight: bold; margin-left: 20px;">Overall Rating</p></td>
						<td class="resp" style="border: 1px solid #888 !important;"><%=avgOvral%></td>
					</tr>
				</tbody>
			</table>
			
			<div style="page-break-before:always"></div>

			<p>
				<label>Goals:</label>
				<div class="befixd" style="word-wrap: break-word !important; width: 100% !important;"><p><%=txtGoals%></p></div>
			</p>
			<p>
				<label>Existing Strengths:</label>
				<div class="befixd" style="word-wrap: break-word !important; width: 100% !important;"><p><%=txtStrng%></p></div>
			</p>
			<p>
				<label>Areas Needing Improvement:</label>
				<div class="befixd" style="word-wrap: break-word !important; width: 100% !important;"><p><%=txtImprv %></p></div>
			</p>
			<p>
				<label>Comments:</label>
				<div class="befixd" style="word-wrap: break-word !important; width: 100% !important;"><p><%=txtComnt %></p></div>
			</p>

  		</div>
  	</div>
  	<div class="row">
		<div class="twelve columns" style="margin-top 30px; border-top: 1px dotted #444;">
			
  			<p>My signature below does not necessarily indicate agreement, but that I've read and undertood this performance appraisal. I also understand that I may respond to it in writing and understand that my comments will be included with this appraisal in my personnel file.</p>
  			<table class="u-full-width smaller">
  				<thead></thead>
  				<tbody>
  					<tr><td>Signature of Employee/Interpreter</td><td>______________________________________________</td><td>Date: <%=dtSig%></td></tr>
  					<tr><td>Signature of Supervisor</td><td>______________________________________________</td><td>Date: _______________</td></tr>
  					<tr><td>Signature of Director/VP</td><td>______________________________________________</td><td>Date: _______________</td></tr>
  				</tbody>
  			</table>
  		</div>
	</div>