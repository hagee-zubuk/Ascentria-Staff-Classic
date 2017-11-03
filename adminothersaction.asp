<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
If Request("ctrl") = 1 Then
	'add/edit
	'language
	If Request("selLang") = 0 Then
		If Request("txtLang") <> "" Then
			Set rsLang = Server.CreateObject("ADODB.RecordSet")
			sqlLang = "SELECT * FROM Language_T WHERE UCase([language]) = '" & Ucase(Request("txtLang")) & "'"
			rsLang.Open sqlLang, g_strCONN, 1, 3
			If rsLang.EOF Then
				rsLang.AddNew
				rsLang("Language") = trim(Request("txtLang"))
				rsLang.Update
				Session("MSG") = Session("MSG") & Request("txtLang") & " language saved.<br>" 
			Else
				Session("MSG") = Session("MSG") & Request("txtLang") & " already exists in the database.<br>" 
			End If
			rsLang.CLose
			Set rsLang = Nothing 
		End If
	Else
		If Request("txtLang") <> Request("LangName") Then
			If Request("txtLang") <> "" Then
				Set rsLang = Server.CreateObject("ADODB.RecordSet")
				sqlLang = "SELECT * FROM Language_T WHERE UCase([language]) = '" & Ucase(Request("LangName")) & "'"
				rsLang.Open sqlLang, g_strCONN, 1, 3
				If NOT rsLang.EOF Then
					rsLang("Language") = trim(Request("txtLang"))
					rsLang.Update
					Session("MSG") = Session("MSG") & Request("txtLang") & " language saved.<br>" 
				Else
					Session("MSG") = Session("MSG") & Request("txtLang") & " already exists in the database.<br>" 
				End If
				rsLang.CLose
				Set rsLang = Nothing 
			End If
		End If
	End If
	'inst rates
	If Request("selRate") = 0 Then
		If Z_Czero(Request("txtRate")) <> 0 Then
			Set rsRate = Server.CreateObject("ADODB.RecordSet")
			sqlRate = "SELECT * FROM rate_T WHERE rate = " & Request("txtRate")
			rsRate.Open sqlRate, g_strCONN, 1, 3
			If rsRate.EOF Then
				rsRate.AddNew
				rsRate("Rate") = Request("txtRate")
				rsRate.Update
				Session("MSG") = Session("MSG") & "Institution Rate of " & Request("txtRate") & " saved.<br>" 
			Else
				Session("MSG") = Session("MSG") & "Institution Rate of " & Request("txtRate") & " already exists in the database.<br>" 
			End If
			rsRate.CLose
			Set rsRate = Nothing
		End If
	Else
		If Request("txtRate") <> Request("RateReas") Then
			If Z_Czero(Request("txtRate")) <> 0 Then
				Set rsRate = Server.CreateObject("ADODB.RecordSet")
				sqlRate = "SELECT * FROM rate_T WHERE rate = " & Request("RateReas")
				rsRate.Open sqlRate, g_strCONN, 1, 3
				If NOT rsRate.EOF Then
					rsRate("Rate") = Request("txtRate")
					rsRate.Update
					Session("MSG") = Session("MSG") & "Institution Rate of " & Request("txtRate") & " saved.<br>" 
				Else
					Session("MSG") = Session("MSG") & "Institution Rate of " & Request("txtRate") & " already exists in the database.<br>" 
				End If
				rsRate.CLose
				Set rsRate = Nothing
			End If
		End If
	End If
	'intr rates
	If Request("selRate2") = 0 Then
		If Z_Czero(Request("txtRate2")) <> 0 Then
			Set rsRate = Server.CreateObject("ADODB.RecordSet")
			sqlRate = "SELECT * FROM rate2_T WHERE rate2 = " & Request("txtRate2")
			response.write sqlrate
			rsRate.Open sqlRate, g_strCONN, 1, 3
			If rsRate.EOF Then
				rsRate.AddNew
				rsRate("Rate2") = Request("txtRate2")
				rsRate.Update
				Session("MSG") = Session("MSG") & "Interpreter Rate of " & Request("txtRate2") & " saved.<br>" 
			Else
				Session("MSG") = Session("MSG") & "Interpreter Rate of " & Request("txtRate2") & " already exists in the database.<br>" 
			End If
			rsRate.CLose
			Set rsRate = Nothing
		End If
	Else
		If Request("txtRate2") <> Request("RateReas2") Then
			If Z_Czero(Request("txtRate2")) <> 0 Then
				Set rsRate = Server.CreateObject("ADODB.RecordSet")
				sqlRate = "SELECT * FROM rate2_T WHERE rate2 = " & Request("RateReas2")
				rsRate.Open sqlRate, g_strCONN, 1, 3
				If NOT rsRate.EOF Then
					rsRate("Rate2") = Request("txtRate2")
					rsRate.Update
					Session("MSG") = Session("MSG") & "Interpreter Rate of " & Request("txtRate2") & " saved.<br>" 
				Else
					Session("MSG") = Session("MSG") & "Interpreter Rate of " & Request("txtRate2") & " already exists in the database.<br>" 
				End If
				rsRate.CLose
				Set rsRate = Nothing
			End If
		End If
	End If
	'interpreter mileage rate
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM MileageRate_T " 
	rsRate.Open sqlRate, g_strCONN, 1, 3
	rsRate("mileageRate") = Z_Czero(Request("txtMR"))
	rsRate.Update
	rsRate.CLose
	Set rsRate = Nothing
	'Cancel Reason
	If Request("selCancel") = 0 Then
		If Request("txtCancel") <> "" Then
			Set rsLang = Server.CreateObject("ADODB.RecordSet")
			sqlLang = "SELECT * FROM cancel_T WHERE UCase(reason) = '" & Ucase(Request("txtCancel")) & "'"
			rsLang.Open sqlLang, g_strCONN, 1, 3
			If rsLang.EOF Then
				rsLang.AddNew
				rsLang("reason") = trim(Request("txtCancel"))
				rsLang.Update
				Session("MSG") = Session("MSG") & "'" & Request("txtCancel") & "' as Cancel reason saved.<br>" 
			Else
				Session("MSG") = Session("MSG") & "'" & Request("txtCancel") & "' as Cancel reason already exists in the database.<br>" 
			End If
			rsLang.CLose
			Set rsLang = Nothing 
		End If
	Else
		If Request("txtCancel") <> Request("CancelReas") Then
			If Request("txtLang") <> "" Then
				Set rsLang = Server.CreateObject("ADODB.RecordSet")
				sqlLang = "SELECT * FROM cancel_T WHERE UCase(reason) = '" & Ucase(Request("CancelReas")) & "'"
				rsLang.Open sqlLang, g_strCONN, 1, 3
				If NOT rsLang.EOF Then
					rsLang("reason") = trim(Request("txtCancel"))
					rsLang.Update
					Session("MSG") = Session("MSG") & "'" & Request("txtCancel") & "' as Cancel reason saved.<br>" 
				Else
					Session("MSG") = Session("MSG") & "'" & Request("txtCancel") & "' as Cancel reason already exists in the database.<br>" 
				End If
				rsLang.CLose
				Set rsLang = Nothing 
			End If
		End If
	End If
	'Missed Reason
	If Request("selMissed") = 0 Then
		If Request("txtMissed") <> "" Then
			Set rsLang = Server.CreateObject("ADODB.RecordSet")
			sqlLang = "SELECT * FROM missed_T WHERE UCase(reason) = '" & Ucase(Request("txtMissed")) & "'"
			rsLang.Open sqlLang, g_strCONN, 1, 3
			If rsLang.EOF Then
				rsLang.AddNew
				rsLang("reason") = trim(Request("txtMissed"))
				rsLang.Update
				Session("MSG") = Session("MSG") & "'" & Request("txtMissed") & "' as Missed reason saved.<br>" 
			Else
				Session("MSG") = Session("MSG") & "'" & Request("txtMissed") & "' as Missed reason already exists in the database.<br>" 
			End If
			rsLang.CLose
			Set rsLang = Nothing 
		End If
	Else
		If Request("txtMissed") <> Request("MissedReas") Then
			If Request("txtLang") <> "" Then
				Set rsLang = Server.CreateObject("ADODB.RecordSet")
				sqlLang = "SELECT * FROM missed_T WHERE UCase(reason) = '" & Ucase(Request("MissedReas")) & "'"
				rsLang.Open sqlLang, g_strCONN, 1, 3
				If NOT rsLang.EOF Then
					rsLang("reason") = trim(Request("txtMissed"))
					rsLang.Update
					Session("MSG") = Session("MSG") & "'" & Request("txtMissed") & "' as Missed reason saved.<br>" 
				Else
					Session("MSG") = Session("MSG") & "'" & Request("txtMissed") & "' as Missed reason already exists in the database.<br>" 
				End If
				rsLang.CLose
				Set rsLang = Nothing 
			End If
		End If
	End If
	'interpreter mileage Cap
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM travel_T "
	rsRate.Open sqlRate, g_strCONN, 1, 3
	rsRate("milediff") = Z_Czero(Request("txtMile"))
	rsRate.Update
	rsRate.CLose
	Set rsRate = Nothing
	'institution mileage Cap (court and legal)
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM travelInstCourt_T " 
	rsRate.Open sqlRate, g_strCONN, 1, 3
	rsRate("milediffcourt") = Z_Czero(Request("txtMileCourt"))
	rsRate.Update
	rsRate.CLose
	Set rsRate = Nothing
	'institution mileage Cap (others)
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM travelInst_T " 
	rsRate.Open sqlRate, g_strCONN, 1, 3
	rsRate("milediffInst") = Z_Czero(Request("txtMileInst"))
	rsRate.Update
	rsRate.CLose
	Set rsRate = Nothing
	'emergency fee
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T " 
	rsRate.Open sqlRate, g_strCONN, 1, 3
	rsRate("FeeLegal") = Z_Czero(Request("txtFeel"))
	rsRate("FeeOther") = Z_Czero(Request("txtFeeO"))
	rsRate.Update
	rsRate.CLose
	Set rsRate = Nothing
	response.redirect "adminothers.asp"
ElseIf Request("ctrl") = 2 Then
	'delete
	Session("MSG") = "DELETE FUNCTION UNDER CONSTRUCTION"
	If Request("selLang") <> 0 Then
		Session("MSG") = Session("MSG") & GetLang(Request("selLang")) & " Language deleted.<br>" 
		Set rsLang = Server.CreateObject("ADODB.RecordSet")
		sqlLang = "DELETE FROM Language_T WHERE index = " & Request("selLang")
		rsLang.Open sqlLang, g_strCONN, 1, 3
		Set rsLang = Nothing
	End If
	If Request("selRate") <> 0 Then
		Session("MSG") = Session("MSG") & "$" & Request("selRate") & " Institution Rate deleted.<br>" 
		Set rsLang = Server.CreateObject("ADODB.RecordSet")
		sqlLang = "DELETE FROM Rate_T WHERE rate = " & Request("selRate")
		rsLang.Open sqlLang, g_strCONN, 1, 3
		Set rsLang = Nothing
	End If
	If Request("selRate2") <> 0 Then
		Session("MSG") = Session("MSG") & "$" & Request("selRate2") & " Interpreter Rate deleted.<br>" 
		Set rsLang = Server.CreateObject("ADODB.RecordSet")
		sqlLang = "DELETE FROM Rate2_T WHERE rate2 = " & Request("selRate2")
		rsLang.Open sqlLang, g_strCONN, 1, 3
		Set rsLang = Nothing
	End If
	If Request("selCancel") <> 0 Then
		Session("MSG") = Session("MSG") & "'" & GetCanReason(Request("selCancel")) & "' Cancellation reason deleted.<br>" 
		Set rsLang = Server.CreateObject("ADODB.RecordSet")
		sqlLang = "DELETE FROM cancel_T WHERE index = " & Request("selCancel")
		rsLang.Open sqlLang, g_strCONN, 1, 3
		Set rsLang = Nothing
	End If
	If Request("selMissed") <> 0 Then
		Session("MSG") = Session("MSG") & "'" & GetMisReason(Request("selCancel")) & "' Missed reason deleted.<br>" 
		Set rsLang = Server.CreateObject("ADODB.RecordSet")
		sqlLang = "DELETE FROM missed_T WHERE index = " & Request("selMissed")
		rsLang.Open sqlLang, g_strCONN, 1, 3
		Set rsLang = Nothing
	End If
	response.redirect "adminothers.asp"
End If
%>