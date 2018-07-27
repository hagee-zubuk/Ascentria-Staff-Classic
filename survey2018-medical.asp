<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function ChkBoxState(aaa)
	ChkBoxState = ""
	valA = Z_CLng(aaa)
	If valA = 1 Then
		ChkBoxState = " checked "
	End If
End Function


txtName = ""
lngID = Z_CLng(Request("iid"))
blnLoad = FALSE
blnNm = ""
If lngID > 1 Then
	' an interpreter is specified -- you have to look it up!
	strSQL = "SELECT [index] AS [id], [First Name] + ' ' + [Last Name] AS [intr] " & _
			"FROM [interpreter_T] " & _
			"WHERE [index]= " & lngID
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open strSQL, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		txtName = Z_FixNull(rsIntr("intr"))
		lngID = Z_FixNull( rsIntr("id") )
	End If
	rsIntr.Close
	Set rsIntr = Nothing
	blnLoad = TRUE
	blnNm = " readonly=""readonly"" "
	strSQL = "SELECT * FROM [survey2018med] WHERE [iid]=" & lngID 
	Set rsSurv = Server.CreateObject("ADODB.RecordSet")
	rsSurv.Open strSQL, g_strCONN, 3, 1
	If rsSurv.EOF Then
		blnA1 = ""
		blnA2 = ""
		blnA3 = ""
		blnA4 = ""
		blnA5 = ""
		blnA6 = ""
		blnB1 = ""
		blnC1 = ""
		blnC2 = ""
		blnC3 = ""
		blnC4 = ""
		blnD1 = ""
		blnD2 = ""
		blnD3 = ""
		blnD4 = ""
		blnD5 = ""
		blnD6 = ""
		blnD7 = ""
		blnD8 = ""
		blnE1 = ""
		blnE2 = ""
		blnE3 = ""
		blnE4 = ""
		blnE5 = ""
		blnF1 = ""
		blnF2 = ""
		blnF3 = ""
		blnF4 = ""
		blnF5 = ""
		blnF6 = ""
		blnF7 = ""
		blnG1 = ""
		blnG2 = ""
		blnH1 = ""
		blnH2 = ""
		blnH3 = ""
		blnH4 = ""
		blnH5 = ""
		blnH6 = ""
		blnH7 = ""
		blnH8 = ""
		blnH9 = ""
		blnH10 = ""
		blnH11 = ""
		blnH12 = ""
		blnH13 = ""
	Else
		blnA1 = ChkBoxState(rsSurv("chkA1"))
		blnA2 = ChkBoxState(rsSurv("chkA2"))
		blnA3 = ChkBoxState(rsSurv("chkA3"))
		blnA4 = ChkBoxState(rsSurv("chkA4"))
		blnA5 = ChkBoxState(rsSurv("chkA5"))
		blnA6 = ChkBoxState(rsSurv("chkA6"))
		blnB1 = ChkBoxState(rsSurv("chkB1"))
		blnC1 = ChkBoxState(rsSurv("chkC1"))
		blnC2 = ChkBoxState(rsSurv("chkC2"))
		blnC3 = ChkBoxState(rsSurv("chkC3"))
		blnC4 = ChkBoxState(rsSurv("chkC4"))
		blnD1 = ChkBoxState(rsSurv("chkD1"))
		blnD2 = ChkBoxState(rsSurv("chkD2"))
		blnD3 = ChkBoxState(rsSurv("chkD3"))
		blnD4 = ChkBoxState(rsSurv("chkD4"))
		blnD5 = ChkBoxState(rsSurv("chkD5"))
		blnD6 = ChkBoxState(rsSurv("chkD6"))
		blnD7 = ChkBoxState(rsSurv("chkD7"))
		blnD8 = ChkBoxState(rsSurv("chkD8"))
		blnE1 = ChkBoxState(rsSurv("chkE1"))
		blnE2 = ChkBoxState(rsSurv("chkE2"))
		blnE3 = ChkBoxState(rsSurv("chkE3"))
		blnE4 = ChkBoxState(rsSurv("chkE4"))
		blnE5 = ChkBoxState(rsSurv("chkE5"))
		blnF1 = ChkBoxState(rsSurv("chkF1"))
		blnF2 = ChkBoxState(rsSurv("chkF2"))
		blnF3 = ChkBoxState(rsSurv("chkF3"))
		blnF4 = ChkBoxState(rsSurv("chkF4"))
		blnF5 = ChkBoxState(rsSurv("chkF5"))
		blnF6 = ChkBoxState(rsSurv("chkF6"))
		blnF7 = ChkBoxState(rsSurv("chkF7"))
		blnG1 = ChkBoxState(rsSurv("chkG1"))
		blnG2 = ChkBoxState(rsSurv("chkG2"))
		blnH1 = ChkBoxState(rsSurv("chkH1"))
		blnH2 = ChkBoxState(rsSurv("chkH2"))
		blnH3 = ChkBoxState(rsSurv("chkH3"))
		blnH4 = ChkBoxState(rsSurv("chkH4"))
		blnH5 = ChkBoxState(rsSurv("chkH5"))
		blnH6 = ChkBoxState(rsSurv("chkH6"))
		blnH7 = ChkBoxState(rsSurv("chkH7"))
		blnH8 = ChkBoxState(rsSurv("chkH8"))
		blnH9 = ChkBoxState(rsSurv("chkH9"))
		blnH10 = ChkBoxState(rsSurv("chkH10"))
		blnH11 = ChkBoxState(rsSurv("chkH11"))
		blnH12 = ChkBoxState(rsSurv("chkH12"))
		blnH13 = ChkBoxState(rsSurv("chkH13"))
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
	<link href="https://fonts.googleapis.com/css?family=Roboto+Condensed" rel="stylesheet">
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
h1 { font-size: 12pt; font-family: 'Roboto Condensed', sans-serif; }
td.indent-1 { padding-left: 25px; }
td.indent-2 { padding-left: 55px; }
td input[type="checkbox"] { margin: 8px 10px 5px; }
	</style>
</head>
<body>
<div class="container">
	<div class="row">
		<div class="twelve columns" id="logobar">
			<img id="logo" src="images/lb-logo.jpg" alt="The Language Bank" title="" />
			<h1>Medical&nbsp;Interpreter<br />Competency&nbsp;Checklist</h1>
		</div>
	</div>
	<div class="no-print u-full-width">
		<a href="survey.list.asp" title="go back to the list of responses">&lt;&lt;&nbsp;back</a>
	</div>
	<form id="frmA" name="frmA" method="post" action="survey-med.save.asp">	
	<div class="row" id="intrbar">
		<div class="five columns">
			<label for="txtName">Interpreter Name</label><input name="txtName" id="txtName" value="<%=txtName%>" <%=blnNm %>
				placeholder="type in an interpreter name" autocomplete="off" autofocus="true" class="u-full-width" />
		</div>
		<div class="four columns">
			<label for="txtDate">Date</label><input name="txtDate" id="txtDate" tabstop="-1" readonly="true" value="<%=Z_MDYDate(Date)%>" />
		</div>
		<div class="three columns align-right">
			<input type="hidden" name="IID" id="IID" readonly="true" value="<%=lngID%>" />
			<br />
			<button type="submit" class="button button-primary" id="btnSave2" style="display: none;">Save</button>
		</div>
	</div>
	<div class="row">
		<div class="ten columns" style="font-size: 9pt;">
			COMPETENCY<br />
			<p>(For further details on competency requirements, please refer
			to Manual of Orientation for Medical Interpreters and Guidelines
			for Establishing Competency)</p>
		</div>
		<div class="two columns" style="font-size: 9pt; text-align: center;vertical-align: bottom;">
			Check if feedback is required			
		</div>
	</div>
	<div class="row">
		<div class="twelve columns">
			<table class="u-full-width smallertable">
  				<thead></thead>
  				<tbody>
    				<tr><th>A. INTRODUCTION/ROLE OF INTERPRETER:  The interpreter...</th><th>&nbsp;</th>
				    </tr>
  					<tr><td class="indent-1">
						Introduces self, explains role of interpreter to patient, and establishes rapport with patient.</td>
						<td><input type="checkbox" id="chkA1" name="chkA1" value="1" <%=blnA1%>/></td>
					</tr>
					<tr><td class="indent-1">
						Ascertains whether the patient has prior experience working with interpreters.</td>
						<td><input type="checkbox" id="chkA2" name="chkA2" value="1" <%=blnA2%>/></td>
					</tr>
					<tr><td class="indent-1">
						Encourages patient to ask for clarification of any issue as it arises during the visit.</td>
						<td><input type="checkbox" id="chkA3" name="chkA3" value="1" <%=blnA3%>/></td>
					</tr>
					<tr><td class="indent-1">
						Relays to the patient legal requirements and essential information regarding informed consent, confidentiality, and security of medical communication.</td>
						<td><input type="checkbox" id="chkA4" name="chkA4" value="1" <%=blnA4%>/></td>
					</tr>
					<tr><td class="indent-1">
						Asks the provider to introduce him/herself to the patient using his/her full title and to state the provider’s goal for the visit.</td>
						<td><input type="checkbox" id="chkA5" name="chkA5" value="1" <%=blnA5%>/></td>
					</tr>
					<tr><td class="indent-1">
						Relays to both the health professional and the patient that if either desires a confidential conversation that they do not want the interpreter to hear, that the interpreter must leave the room given the requirement that interpreters translate everything that is said by either the patient or healthcare professional.</td>
						<td><input type="checkbox" id="chkA6" name="chkA6" value="1" <%=blnA6%>/></td>
					</tr>
					<tr><th>B. MANAGEMENT OF PHYSICAL SPACE: The interpreter...</th><th>&nbsp;</th>
					</tr>
					<tr><td class="indent-1">
						Effectively arranges the spatial configuration of the interview to encourage direct face-to-face contact by the patient and provider of care.</td>
						<td><input type="checkbox" id="chkB1" name="chkB1" value="1" <%=blnB1%>/></td>
					</tr>
					<tr><th>C. CULTURAL UNDERSTANDING: The interpreter...</th><th>&nbsp;</th>
					</tr>
  					<tr><td class="indent-1">
  						Understands the rules of cultural etiquette with respect to status, age, gender, hierarchy, and level of acculturation.</td>
						<td><input type="checkbox" id="chkC1" name="chkC1" value="1" <%=blnC1%>/></td>
					</tr>
					<tr><td class="indent-1">
						Demonstrates an understanding of potential barriers to communication including cultural differences, ethnic issues, gender issues, lack of education or differences between patient or provider life experience.</td>
						<td><input type="checkbox" id="chkC2" name="chkC2" value="1" <%=blnC2%>/></td>
					</tr>
					<tr><td class="indent-1">
						Anticipates the need for and reassesses patient and provider comfort levels and addresses any perceived barriers that may impact on the success of the interaction between provider and patient.</td>
						<td><input type="checkbox" id="chkC3" name="chkC3" value="1" <%=blnC3%>/></td>
					</tr>
					<tr><td class="indent-1">
						Shares any relevant cultural information with both patient and provider to facilitate understanding between all parties.</td>
						<td><input type="checkbox" id="chkC4" name="chkC4" value="1" <%=blnC4%>/></td>
					</tr>
					<tr><th>D. INTERPRETATION SKILLS: The interpreter...</th><th>&nbsp;</th>
					</tr>
					<tr><td class="indent-1">Understands the vital role of accurate interpretation and understands the risks of inaccurate interpretation in a medical situation.
					</tr>
					<tr><td class="indent-1">
						Considers and selects the most effective mode of interpretation prior to the start of the interpretation service (e.g., consecutive, simultaneous, or first/third person) and adjusts mode as needed during clinical interview.</td>
						<td><input type="checkbox" id="chkD1" name="chkD1" value="1" <%=blnD1%>/></td>
					</tr>
					<tr><td class="indent-1">
						Ensures that he/she understands the message prior to transmission.</td>
						<td><input type="checkbox" id="chkD2" name="chkD2" value="1" <%=blnD2%>/></td>
					</tr>
					<tr><td class="indent-1">
						Understands his/her limitations of medical knowledge, refrains from making assumptions, and demonstrates willingness to obtain clarification of medical terms and concepts as necessary.</td>
						<td><input type="checkbox" id="chkD3" name="chkD3" value="1" <%=blnD3%>/></td>
					</tr>
					<tr><td class="indent-1">
						Accurately transmits information between patient and provider, transmitting the message completely, utilizing communication aids (e.g., pictures, drawings, or gestures) to supplement communication</td>
						<td><input type="checkbox" id="chkD4" name="chkD4" value="1" <%=blnD4%>/></td>
					</tr>
					<tr><td class="indent-1">
						Ensures that the listener (patient/family) understands what is being conveyed after transmission of the information.</td>
						<td><input type="checkbox" id="chkD5" name="chkD5" value="1" <%=blnD5%>/></td>
					</tr>
					<tr><td class="indent-1" colspan="2">
						Manages the flow of communication in order to insure accuracy of transmission and enhance rapport between patient and provider.  Specifically:</td>
					</tr>
					<tr><td class="indent-2">
						Manages the conversation so that only one person talks at a time.</td>
						<td><input type="checkbox" id="chkD6" name="chkD6" value="1" <%=blnD6%>/></td>
					</tr>
					<tr><td class="indent-2">
						Interrupts the other speaker to allow the other party to speak when necessary.</td>
						<td><input type="checkbox" id="chkD7" name="chkD7" value="1" <%=blnD7%>/></td>
					</tr>
					<tr><td class="indent-2">
						Indicates clearly when he/she is speaking on his/her own behalf.</td>
						<td><input type="checkbox" id="chkD8" name="chkD8" value="1" <%=blnD8%>/></td>
					</tr>
					<tr><th>E. COMMUNICATION SKILLS: The interpreter...</th><th>&nbsp;</th></tr>
					<tr><td class="indent-1">
						Is cognizant of the changing tone and emotional content of medical conversations, and remains alert to internal conflicts that may emerge between provider and patient.</td>
						<td><input type="checkbox" id="chkE1" name="chkE1" value="1" <%=blnE1%>/></td>
					</tr>
					<tr><td class="indent-1">
						When strong feelings or conflict arise between the provider and the patient, the interpreter does not take sides in the conflict and remains calm while acknowledging the tension between patient and provider. He/she manages the situation effectively through use of clarification.</td>
						<td><input type="checkbox" id="chkE2" name="chkE2" value="1" <%=blnE2%>/></td>
					</tr>
					<tr><td class="indent-1">
						Manages his/her own internal personal conflicts by clearly separating his/her own values and beliefs from those of the patient and provider of care.</td>
						<td><input type="checkbox" id="chkE3" name="chkE3" value="1" <%=blnE3%>/></td>
					</tr>
					<tr><td class="indent-1">
						Is able to acknowledge openly to the patient/provider that the topic is difficult for interpreter.</td>
						<td><input type="checkbox" id="chkE4" name="chkE4" value="1" <%=blnE4%>/></td>
					</tr>
					<tr><td class="indent-1">
						Actively identifies his/her own mistakes, corrects him/herself as quickly as possible, communicates that to both patient/provider, and accepts the feedback and restates new understanding for the record.</td>
						<td><input type="checkbox" id="chkE5" name="chkE5" value="1" <%=blnE5%>/></td>
					</tr>
					<tr><th>F. ROLE AS FACILITATOR: The interpreter...</th><th>&nbsp;</th></tr>
					<tr><td class="indent-1">
						Encourages the provider to give the patient appropriate instructions and makes certain that the patient understands both the instructions and what he/she must do next.</td>
						<td><input type="checkbox" id="chkF1" name="chkF1" value="1" <%=blnF1%>/></td>
					</tr>
					<tr><td class="indent-1">
						Ascertains from the patient whether he/she has any final questions for the provider.</td>
						<td><input type="checkbox" id="chkF2" name="chkF2" value="1" <%=blnF2%>/></td>
					</tr>
					<tr><td class="indent-1">
						Assesses whether the patient will need interpretation services after the medical visit is concluded.</td>
						<td><input type="checkbox" id="chkF3" name="chkF3" value="1" <%=blnF3%>/></td>
					</tr>
					<tr><td class="indent-1">
						Ensures that the patient understands to contact the Provider of Record or OnCall provider, or organization telephone service after hours with any concerns or questions.</td>
						<td><input type="checkbox" id="chkF4" name="chkF4" value="1" <%=blnF4%>/></td>
					</tr>
					<tr><td class="indent-1">
						Explains after hours process to patients with limited English proficiency.</td>
						<td><input type="checkbox" id="chkF5" name="chkF5" value="1" <%=blnF5%>/></td>
					</tr>
					<tr><td class="indent-1">
						Ensures appropriate referrals are made, including place, date and time, and ensures interpretive services are scheduled.</td>
						<td><input type="checkbox" id="chkF6" name="chkF6" value="1" <%=blnF6%>/></td>
					</tr>
					<tr><td class="indent-1">
						Ensures that any concerns raised (before or after the interview) are addressed and referred to clinical personnel who can assist with resolution of such concerns.</td>
						<td><input type="checkbox" id="chkF7" name="chkF7" value="1" <%=blnF7%>/></td>
					</tr>
					<tr><th>G. ADMINISTRATIVE TASKS: The interpreter...</th><th>&nbsp;</th></tr>
					<tr><td class="indent-1">
						Completes appropriate documentation as indicated or requested by clinical personnel.</td>
						<td><input type="checkbox" id="chkG1" name="chkG1" value="1" <%=blnG1%>/></td>
					</tr>
					<tr><td class="indent-1">
						Appropriately signs, dates, and indicates time of day on all notes.</td>
						<td><input type="checkbox" id="chkG2" name="chkG2" value="1" <%=blnG2%>/></td>
					</tr>
					<tr><th>H. ETHICAL STANDARDS: In each of the following areas, the interpreter...</th><th>&nbsp;</th></tr>
					<tr><td class="indent-1" colspan="2">CONFIDENTIALITY:</td></tr>
					<tr><td class="indent-2">
						Is aware of and observes all relevant organizational policies and state/federal laws regarding release of confidential medical information.</td>
						<td><input type="checkbox" id="chkH1" name="chkH1" value="1" <%=blnH1%>/></td>
					</tr>
					<tr><td class="indent-2">
						Understands that protection of patient confidentiality is NOT limited to the potential for sharing personal medical information outside of the organization, but also includes a prohibition against sharing any of the patient's personal information with anyone on the health care team or in the healthcare organization who does not have a specific need to know that information.</td>
						<td><input type="checkbox" id="chkH2" name="chkH2" value="1" <%=blnH2%>/></td>
					</tr>
					<tr><td class="indent-1" colspan="2">IMPARTIALITY:</td></tr>
					<tr><td class="indent-2">
						Is aware and able to identify any personal bias, belief, or conflict of interest that may interfere with his/her ability to impartially interpret in any given situation, and discloses this to the provider so that another interpreter can step in to provide the service.</td>
						<td><input type="checkbox" id="chkH3" name="chkH3" value="1" <%=blnH3%>/></td>
					</tr>
					<tr><td class="indent-1" colspan="2">PROFESSIONAL INTEGRITY:</td></tr>
					<tr><td class="indent-2">
						Acts as a conduit of information, not as an information source, unless specifically trained or licensed to supply that particular information. Therefore, the interpreter refrains from counseling or advising the patient at any time.</td>
						<td><input type="checkbox" id="chkH4" name="chkH4" value="1" <%=blnH4%>/></td>
					</tr>
					<tr><td class="indent-2">
						Refrains from any contact with the patient outside of employment, avoiding personal benefit</td>
						<td><input type="checkbox" id="chkH5" name="chkH5" value="1" <%=blnH5%>/></td>
					</tr>
					<tr><td class="indent-2">
						Engages in ongoing professional development.</td>
						<td><input type="checkbox" id="chkH6" name="chkH6" value="1" <%=blnH6%>/></td>
					</tr>
					<tr><td class="indent-2">
						Maintains professional dress and demeanor at all times.</td>
						<td><input type="checkbox" id="chkH7" name="chkH7" value="1" <%=blnH7%>/></td>
					</tr>
					<tr><td class="indent-2">
						Is consistently observed to be free of prejudice or critical comments or judgement of the patient.</td>
						<td><input type="checkbox" id="chkH8" name="chkH8" value="1" <%=blnH8%>/></td>
					</tr>
					<tr><td class="indent-1" colspan="2">PROFESSIONAL DISTANCE:</td></tr>
					<tr><td class="indent-2">
						Can explain the meaning of “distance” in this context, and its implications and consequences.</td>
						<td><input type="checkbox" id="chkH9" name="chkH9" value="1" <%=blnH9%>/></td>
					</tr>
					<tr><td class="indent-2">
						Refrains from becoming personally involved in the patient's life.</td>
						<td><input type="checkbox" id="chkH10" name="chkH10" value="1" <%=blnH10%>/></td>
					</tr>
					<tr><td class="indent-2">
						Does not create any expectations that the interpreter role cannot fulfill.</td>
						<td><input type="checkbox" id="chkH11" name="chkH11" value="1" <%=blnH11%>/></td>
					</tr>
					<tr><td class="indent-2">
						Actively promotes patient self-sufficiency.</td>
						<td><input type="checkbox" id="chkH12" name="chkH12" value="1" <%=blnH12%>/></td>
					</tr>
					<tr><td class="indent-2">
						Monitors own personal agenda of service, and is aware of transference and countertransference issues, discussing them with the team leader or with his/her supervisor when any boundary issue or potential overreaching of mission could occur.</td>
						<td><input type="checkbox" id="chkH13" name="chkH13" value="1" <%=blnH13%>/></td>
					</tr>
				</tbody>
			</table>
		</div>
	</td>
  	<div class="row">
		<div class="twelve columns align-right">
  			<button type="button" class="button button-primary" style="display: none;" id="btnSave" name="btnSave">Save Checklist</button>
  		</div>
	</div>

	</form>

</div>
</body>
</html>
<script language="javascript" type="text/javascript"><!--
$( document ).ready(function() {
<%
If blnLoad Then
%>
	$('#btnSave2').show();
	$('#btnSave').show();
<%
Else
%>	
	$('#txtName').autocomplete({
		source: "ajx_intrsearch.asp",
		minlength: 3,
		select: function(event, ui) {
			inm = ui.item.value;
			iid = ui.item.id;
			if (iid > 0) {
				$('#IID').val(iid);
				$('#txtName').prop('disabled', true);
				setScaleVisibility(false);
				$('#btnSave2').show();
				$('#btnSave').show();
			}
		}
	});
<%
End If
%>
	$('#btnSave').click(function(){ submitme(); });
	$('#btnSave2').click(function(){ submitme(); });
	console.log( "ready!" );
});
// --></script>