<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
strUsr = Request.Cookies("LBUsrName")
lngUID = Request.Cookies("UID")
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
<form id="frmA" name="frmA" method="post" action="survey.save.asp">	
	<div class="row" id="intrbar">
		<div class="five columns">
			<label for="txtName">Interpreter Name</label><input name="txtName" id="txtName"
				placeholder="type in an interpreter name" autocomplete="off" autofocus="true" class="u-full-width" />
		</div>
		<div class="four columns">
			<label for="txtDate">Date</label><input name="txtDate" id="txtDate" tabstop="-1" readonly="true" value="<%=Z_MDYDate(Date)%>" />
		</div>
		<div class="three columns align-right">
			<br />
			<button type="submit" class="button button-primary" id="btnSave2" style="display: none;">Save Survey</button>
		</div>
	</div>
	<div class="row">
		<div class="twelve columns">
			<input type="hidden" name="IID" id="IID" readonly="true" value="" />
			<p>It is required that the interpreter's job performance be evaluated annually.</p>
			<p>The interpreter's performance is based on the following criteria and rating scale: <a href="#ratingscale" class="smaller" id="showhide" name="showhide">(hide this)</a> </p>
		</div>
	</div>
	<div class="row" id="ratingscale">
		<div class="six columns">
			<h4>CRITERIA</h4>
			<ul>
				<li>Punctuality</li>
				<li>Professional Behavior</li>
				<li>Adherence to LB Procedural Guidelines</li>
				<li>Team Work Ethics</li>
				<li>Professional Development</li>
			</ul>
		</div>
		<div class="six columns">
			<h4>RATING SCALE</h4>
			<ul>
				<li><h5>Outstanding  - 4</h5>
					<p>Employee consistently exceeds the expectations of 
						their position. Their colleagues recognize their 
						excellence and their unique contribution to the
						organization. They serve as a role model for others.
						They require little or no supervision and generate
						output that is exceptionally high in quality and
						quantity. They accept high level of responsibility
						for own performance.</p>
					</li>
				<li><h5>Above Average - 3</h5>
					<p>Employee frequently exceeds expectations, provides
						significant and measureable contribution well beyond
						their position responsibilities.  The employee
						demonstrates a desire and ability to excel in
						performance.</p>
					</li>
				<li><h5>Satisfactory - 2</h5>
					<p>Employee meets expectations. The employee is a
						productive and valued member of the team.</p>
					</li>
				<li><h5>Needs Improvement - 1</h5>
					<p>Employee is struggling to meet the basic responsibility
						of their position and is not meeting job expectations
						or the employee is new in their position and is still
						developing </p>
					</li>
			</ul>
		</div>
	</div>
	<div id="diavolo" style="display: none;">
	<div class="row">
		<div class="one column">&nbsp;</div><div class="eleven columns">
			<p>
				<b>&#x2605;</b>&nbsp;Please note that the statements listed under a given criteria are examples of relevant behaviors and not an exhaustive list.<br />
				<b>&#x2605;</b>&nbsp;Performance under these criteria and subcategories can affect the rating in an overarching catgeory, and, therefore, your total rating.<br /></p>
		</div>
	</div>
	<div class="row">
		<div class="twelve columns">
			<table class="u-full-width">
  				<thead>
    				<tr><th>Performance Criteria</th><th>1</th><th>2</th><th>3</th><th>4</th>
				    </tr>
  				</thead>
  				<tbody>
  					<tr><td><h5>Punctuality</h5>
							<p>Arrives at assigned appointments on time, or, preferably, five (5) minutes early.</p>
							</td>
						<td><input type="radio" name="rdoPunct" id="rdoPunct1" value="1" /></td>
						<td><input type="radio" name="rdoPunct" id="rdoPunct2" value="2" /></td>
						<td><input type="radio" name="rdoPunct" id="rdoPunct3" value="3" /></td>
						<td><input type="radio" name="rdoPunct" id="rdoPunct4" value="4" /></td>
					</tr>
					<tr><td><h5>Professional Behavior</h5>
						<p>Maintains impartiality and keeps personal opinions/feelings/beliefs out of the triadic setting.<br />
						Does not stay with the patient alone (behind closed doors) at any time.<br />
						Practices transparency when asking for clarification/repetition, and avoids side conversation.<br />
						Dresses Professionally - Business Casual Dress Code.<br />
						Communicates with providers and other staff in a professional manner.<br />
						Withdraws from conflicts of interest or any other situations that may interfere with impartiality.<br />
						Maintains professional boundaries by facilitating appropriate resources without becoming personally involved.<br />
						</p></td>
						<td><input type="radio" name="rdoProfb" id="rdoProfb1" value="1" /></td>
						<td><input type="radio" name="rdoProfb" id="rdoProfb2" value="2" /></td>
						<td><input type="radio" name="rdoProfb" id="rdoProfb3" value="3" /></td>
						<td><input type="radio" name="rdoProfb" id="rdoProfb4" value="4" /></td>
					</tr>
					<tr><td><h5>Adherence to LB Procedural Guidelines</h5>
						<p>Holds pre-sessions to introduce themselves and briefly explain their role to clients and consumers.<br />
						Follows Language Bank guidelines and Ascentria company policies/procedures.<br />
						Sends signed V-Forms to the Language Bank within 5 business days<br />
						Enters time into the database within 48 hours of job completion.<br />
						Calls the Language Bank when running late to an appointment.<br />
						</p></td>
						<td><input type="radio" name="rdoProcG" id="rdoProcG1" value="1" /></td>
						<td><input type="radio" name="rdoProcG" id="rdoProcG2" value="2" /></td>
						<td><input type="radio" name="rdoProcG" id="rdoProcG3" value="3" /></td>
						<td><input type="radio" name="rdoProcG" id="rdoProcG4" value="4" /></td>
					</tr>
					<tr><td><h5>Team Work Ethics</h5>
						<p>Treats Language Bank staff with respect and courtesy.<br />
						Develops productive/cooperative relations with other team members<br />
						Commits to the success of the entire team<br />
						Works with Language Bank staff in a flexible manner to cover appointments.<br />
						Informs Language Bank via the database of any known periods of unavailability<br />
						</p></td>
						<td><input type="radio" name="rdoTeamW" id="rdoTeamW1" value="1" /></td>
						<td><input type="radio" name="rdoTeamW" id="rdoTeamW2" value="2" /></td>
						<td><input type="radio" name="rdoTeamW" id="rdoTeamW3" value="3" /></td>
						<td><input type="radio" name="rdoTeamW" id="rdoTeamW4" value="4" /></td>
					</tr>
					<tr><td><h5>Professional Development</h5>
						<p>Engages in ongoing professional development, attending educational events offered by any reputable source.
						</p></td>
						<td><input type="radio" name="rdoProDv" id="rdoProDv1" value="1" /></td>
						<td><input type="radio" name="rdoProDv" id="rdoProDv2" value="2" /></td>
						<td><input type="radio" name="rdoProDv" id="rdoProDv3" value="3" /></td>
						<td><input type="radio" name="rdoProDv" id="rdoProDv4" value="4" /></td>
					</tr>
				</tbody>
			</table>
			<p>
			<label for="rdoReliasTrng">Completed the required trainings in Relias (Yes or No)</label>
			<div class="rdoChoice"><input type="radio" name="rdoReliasTrng" id="rdoReliasTrngY" value="Y" />&nbsp;Yes<br /></div>
			<div class="rdoChoice"><input type="radio" name="rdoReliasTrng" id="rdoReliasTrngN" value="N" />&nbsp;No<br /></div>
			</p>
			<p>
				<label for="txtGoals">Goals:</label>
				<textarea class="u-full-width" name="txtGoals" id="txtGoals"></textarea>
			</p>
			<p>
				<label for="txtStrengths">Existing Strengths:</label>
				<textarea class="u-full-width" name="txtStrengths" id="txtStrengths"></textarea>
			</p>
			<p>
				<label for="txtImprovement">Areas Needing Improvement:</label>
				<textarea class="u-full-width" name="txtImprovement" id="txtImprovement"></textarea>
			</p>
			<p>
				<label for="txtComments">Comments:</label>
				<textarea class="u-full-width" name="txtComments" id="txtComments"></textarea>
			</p>
  		</div>
  	</div>
  	<div class="row">
		<div class="twelve columns align-right">
  			<!-- <%=Request.Cookies("LBUsrName")%> -->
  			<button type="button" class="button button-primary" style="display: none;" id="btnSave" name="btnSave">Save Survey</button>
  		</div>
	</div>

	</div>
</form>
</div>
</body>
</html>
<script language="javascript" type="text/javascript"><!--
var blnScaleVis = false;
var sticky;

function setScaleVisibility(zzVis) {
	if (zzVis) {
		blnScaleVis = true;
		$('#ratingscale').show('blind', {}, 400);
		$('#showhide').html("(hide this)")
	} else {
		blnScaleVis = false;
		$('#ratingscale').hide('blind', {}, 500);
		$('#showhide').html("(show)")
	}
	return(blnScaleVis);
}

function submitme() {
	$('#frmA').submit();
}
$( document ).ready(function() {
	sticky = $('#intrbar').offsetTop;
	$('#diavolo').hide();
	setScaleVisibility(true);
	$('#showhide').click(function() {
		setScaleVisibility(!blnScaleVis);
	});
	$('#intrbar').sticky({topSpacing:0});
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
				$('#diavolo').show('fade', {}, 400);
				$('#btnSave2').show();
				$('#btnSave').show();
			}
		}
	});
	$('#btnSave').click(function(){ submitme(); });
	$('#btnSave2').click(function(){ submitme(); });
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