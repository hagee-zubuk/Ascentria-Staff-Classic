<%@Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- inc file="_Security.asp" -->
<%
g_strTopBackLink = "<br />"
g_blnHideCtls = True
%>
<!-- #include file="survey.repbase.asp" -->
<%
If lngMedIx > 0 Then
%>
<div style="page-break-before:always"></div>
<!-- #include file="survey2018-medbase.asp" -->
<%
End If
%>
</div>
</body>
</html>
