<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
server.scripttimeout = 360000

If Request("ctrl") = 1 Then
	myholdate = Request("selMonth1") & "/" & Request("selDay1") & "/" & Request("txtYear")
	If IsDate(myholdate) Then
			Set rsHol = Server.CreateObject("ADODB.RecordSet")
			sqlHol = "SELECT * FROM holiday_T WHERE holdate = '" & myholdate & "'"
			rsHol.Open sqlHol, g_strCONN, 1, 3
			If rsHol.EOF Then
				rsHol.Addnew
				rsHol("Holdate") = myholdate
				rsHol.Update
				Session("MSG") = "Holiday date saved."
			Else
				Session("MSG") = "ERROR: New Date already exists."
			End If
			rsHol.Close
			Set rsHol = Nothing
	Else
		Session("MSG") = "ERROR: New Date is not a Date."
	End If
ElseIf Request("ctrl") = 2 Then
	Set tblTowns = Server.CreateObject("ADODB.RecordSet")
		sqlTowns = "SELECT * FROM holiday_T"
		tblTowns.Open sqlTowns, g_strCONN, 1, 3
		If Not tblTowns.EOF Then
			If Request("ctr") <> "" Then 
				ctr = Request("ctr")
				For i = 0 to ctr 
					tmpctr = Request("chk" & i)
					If tmpctr <> "" Then
						strTmp = "index= " & tmpctr 
						tblTowns.Movefirst
						tblTowns.Find(strTmp)
						If Not tblTowns.EOF Then
							tblTowns.Delete
							tblTowns.Update
						End If
					End If
				Next
			End If 
		End If
		tblTowns.Close
		Set tblTowns = Nothing	
		Session("MSG") = "Checked holiday dates deleted."
End If
response.redirect "holiday.asp"
%>