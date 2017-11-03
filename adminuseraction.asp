<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Request("ctrl") = 1 Then
	'CHECK ENTRIES
	If Request("txtUserPword") <> Request("txtUserPword2") Then
		Session("MSG") = Session("MSG") & "Error: Password confirmation is different from your assigned password."
	End If
	If Request("txtUserPword")= "" Then
		Session("MSG") = Session("MSG") & "Error: Password cannot be blank."
	End If
	If Request("selUser") = 0 Then
		Set rsUser = CreateObject("ADODB.RecordSet")
		sqlUser = "SELECT * FROM user_T WHERE ucase(username) = '" & ucase(Request("txtUserLname")) & "'"
		rsUser.Open sqlUser, g_strCONN, 1, 3
		If Not rsUser.EOF Then
			Session("MSG") = Session("MSG") & "Error: Username already exists."
		End If
		rsUser.Close
		Set rsUser = Nothing 
	End If
	If Request("selType") = 2 Then
		Set rsUser = CreateObject("ADODB.RecordSet")
		sqlUser = "SELECT * FROM user_T WHERE IntrID = " & Request("selIntr2")
		rsUser.Open sqlUser, g_strCONN, 1, 3
		If Not rsUser.EOF Then
			Session("MSG") = Session("MSG") & "Error: Interpreter already assigned."
		End If
		rsUser.Close
		Set rsUser = Nothing 
	End If
	If Session("MSG") = "" Then
		If Request("selUser") <> 0 Then
			tmpID = Request("userID")
			Set rsUser = CreateObject("ADODB.RecordSet")
			sqlUser = "SELECT * FROM user_T WHERE index = " & Request("selUser")
			rsUser.Open sqlUser, g_strCONN, 1, 3
			If Not rsUser.EOF Then 
				rsUser("lname") = Request("txtUserLname")
				rsUser("fname") = Request("txtUserfname")
				rsUser("username") = Request("txtUserUname")
				rsUser("password") = Z_DoEncrypt(Request("txtUserPword"))
				rsUser("Type") = Request("selType")
				rsUser("IntrID") = Request("selIntr2")
			End If
			rsUser.Close
			Set rsUser = Nothing
		Else
			Set rsUser = CreateObject("ADODB.RecordSet")
			sqlUser = "SELECT * FROM user_T "
			rsUser.Open sqlUser, g_strCONN, 1, 3
			rsUser.AddNew
			tmpID = rsUser("index")
			rsUser("lname") = Request("txtUserLname")
			rsUser("fname") = Request("txtUserfname")
			rsUser("username") = Request("txtUserUname")
			rsUser("password") = Z_DoEncrypt(Request("txtUserPword"))
			rsUser("Type") = Request("selType")
			rsUser("IntrID") = Request("selIntr2")
			rsUser.Update
			rsUser.Close
			Set rsUser = Nothing
		End If
		Session("MSG") = "User updated/saved."
	End If
	response.redirect "adminusers.asp?userID=" & tmpID
ElseIf Request("ctrl") = 2 Then
	If Request("selUser") <> 0 Then
		Set rsUser = CreateObject("ADODB.RecordSet")
		sqlUser = "DELETE FROM User_T WHERE index = " & Request("selUser")
		rsUser.Open sqlUser, g_strCONN, 1, 3
		Set rsUser = Nothing
		Session("MSG") = "User Deleted."
	End If
	response.redirect "adminusers.asp"
End If
%>
