<%
'redirect user to default page when not logged in
If Request("PDF") <> 1 Then
	tmpUser = Request.Cookies("LBUSER")
	If tmpUser = "" Then
		Session("MSG") = "ERROR: Cookies has expired or was not found.<br> Please sign in again."
		Response.redirect "default.asp"
	End If
	tmpName	= Request.Cookies("LBUsrName")
	If tmpName = "" Then
		Session("MSG") = "ERROR: Session has expired.<br> Please sign in again."
		Response.redirect "default.asp"
	End If
End If
%>