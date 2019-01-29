<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
'ovrd_ttm_do.asp?ReqID=
'tmpBilTInst = Z_FormatNumber( rsConfirm("TT_Inst"), 2)
'tmpBilTIntr = Z_FormatNumber( Z_CZero(rsConfirm("actTT")) * Z_CZero(rsConfirm("intrrate")) , 2)
'tmpBilMInst = Z_FormatNumber( rsConfirm("M_Inst"), 2)
'tmpBilMIntr = Z_FormatNumber( rsConfirm("actMil") * tmpmilerate, 2)
strReqID = Z_FixNull(Request("ReqID"))
strPostBack = Z_FixNull(Request("postback"))
If strPostBack = "" Then strPostBack="reqconfirm.asp?ID=" & strReqID

Set rsMain = Server.CreateObject("ADODB.RecordSet")
sqlMain = "SELECT RealTT, RealM, M_Inst, TT_Inst, actTT, actMil, M_Intr" & _
		", BillInst, PayIntr, Toll, LBConfirmToll" & _
		", TT_Intr, InstActTT, InstActMil, IntrRate FROM request_T WHERE [index] = " & Request("ReqID")
rsMain.Open sqlMain, g_strCONN, 1, 3
If Not rsMain.EOF Then
	rsMain("RealTT")	= Z_CZero(Request("RealTT"))
	rsMain("RealM")		= Z_CZero(Request("RealM"))

	rsMain("M_Inst")	= Z_CZero(Request("M_Inst"))
	rsMain("TT_Inst")	= Z_CZero(Request("TT_Inst"))
	' interpreter'
	rsMain("actTT")		= Z_CZero(Request("actTT"))
	rsMain("actMil")	= Z_CZero(Request("actMil"))
	
	rsMain("M_Intr") 	= Z_CZero(Request("M_Intr"))
	rsMain("TT_Intr")	= Z_CZero(Request("TT_Intr"))
	
	rsMain("InstActTT") = Z_CZero(Request("InstActTT"))
	rsMain("InstActMil") = Z_CZero(Request("InstActMil"))
	
	rsMain("Billinst") = Request("BillInst")
	rsMain("PayIntr") = Request("PayIntr")
	rsMain("LBconfirmToll") = Request("LBconfirmToll")
	tmpToll = Z_CZero(Request("Toll"))
	If tmpToll = 0 Then tmpToll = Null
	rsMain("Toll") = tmpToll

	Session("MSG") = "Travel Time and Mileage Saved."
	rsMain.Update
	rsMain.Close
	Set rsMain = Nothing
	SaveHist strReqID, "ovrd_ttm.asp" 
Else
	Session("MSG") = "Appointment record not found or inaccessible."
	rsMain.Close
	Set rsMain = Nothing
End If

Response.Redirect strPostBack
%>
