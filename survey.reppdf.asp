<%@Language=VBScript%>
<%
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
%>
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
