<%@Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
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
<!-- #include file="survey2018-medbase.asp" -->
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