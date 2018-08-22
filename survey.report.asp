<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
g_strTopBackLink = "<div class=""no-print u-full-width"">" & _
				"<a href=""survey.list.asp"" title=""go back to the list of responses"">&lt;&lt;&nbsp;back</a>" & _
				"</div>"
g_blnHideCtls = False
%>
<!-- #include file="survey.repbase.asp" -->
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