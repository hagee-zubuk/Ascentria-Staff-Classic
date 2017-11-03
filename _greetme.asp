<%
'yellow strip that dispalys name of user and annoucements
%>
<tr>
	<td colspan='14' align='center' class='greet'><nobr> --- Welcome&nbsp;&nbsp;<%=Request.Cookies("LBUsrName")%> ---</td>
</tr>
<% If strANN <> "" Then %>
<tr>
	<td colspan='14' align='center' class='greet2'><marquee scrollamount="3"><nobr> >>> ANNOUNCEMENT: <%=strANN%> <<<</marquee></td>
</tr>
<% End If %>