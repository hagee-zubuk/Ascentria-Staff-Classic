<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
server.scripttimeout = 360000
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) = 2 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", "''")
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
Function GetPrime(xxx)
	GetPrime = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM requester_T WHERE [index] = " & xxx
	rsRP.Open sqlRP, g_strCONN, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime = rsRP("Email")
		ElseIf rsRP("prime") = 1 Then
			'GetPrime = rsRP("Phone")
			GetPrime = ""
		ElseIf rsRP("prime") = 2 Then
			GetPrime = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
tmpPage = "document.frmMain."
tmpInst = "-1"
tmpIntr = "-1"
'default
selRPEmail = ""
selRPPhone = ""
selRPFax = "checked"
'default
selIntrFax = "checked"
selIntrP2 = ""
selIntrP1 = ""
selIntrEmail = ""
AHMemId = ""
tmpTS = Now
tmpDept = 0
tmpReqP = 0
tmpHPID = 0
If Session("MSG") <> "" Then
on error resume next
	tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUESTW4")), "|")
	tmplName = tmpEntry(1)
	tmpfName = tmpEntry(2)
	chkClient = ""
	If tmpEntry(3) <> "" Then chkClient = "checked"
	tmpCAdrI = tmpEntry(4)
	tmpAddr = tmpEntry(5)
	chkUClientadd = ""
	If tmpEntry(6) <> "" Then chkUClientadd = "checked"
	tmpCFon = tmpEntry(7)
	tmpCity = tmpEntry(8)
	tmpState = tmpEntry(9)
	tmpZip = tmpEntry(10)
	tmpCAFon = tmpEntry(11)
	tmpDir = tmpEntry(12)
	tmpSC = tmpEntry(13)
	tmpDOB = Z_FormatTime(tmpEntry(14))
	tmpLang = tmpEntry(15)
	tmpAppDate = Z_FormatTime(tmpEntry(16))
	tmpAppTFrom = Z_FormatTime(tmpEntry(17))
	tmpAppTTo = Z_FormatTime(tmpEntry(18))
	tmpAppLoc = tmpEntry(19)
	tmpDoc = tmpEntry(20)
	tmpCRN = tmpEntry(21)
	tmpCom = tmpEntry(22)
	tmpJudge = tmpEntry(28)
	tmpClaim = tmpEntry(29)
	'tmpGender = tmpEntry(23)
	'tmpMinor = tmpEntry(24)
	chkcall = ""
	If tmpEntry(30) <> "" Then chkcall = "CHECKED"
	chkleave = ""
	If tmpEntry(31) <> "" Then chkleave = "CHECKED"
	tmpGender	= tmpEntry(23)
	tmpMale = ""
	tmpFemale = ""
	If tmpGender = 0 Then 
		tmpMale = "SELECTED"
	ElseIf tmpGender = 1 Then 
		tmpFemale = "SELECTED"
	End If
	chkMinor = ""
	If tmpEntry(24) <> "" Then chkMinor = "CHECKED"
	chkout = ""
	If tmpEntry(25) <> "" Then chkout = "CHECKED"
	chkmed = ""
	If tmpEntry(26) <> "" Then chkmed = "CHECKED"
	MCNum = tmpEntry(27)
End If
If Request("tmpID") <> "" Then
	Set rsW1 = Server.CreateObject("ADODB.RecordSet")
	sqlW1 = "SELECT * FROM Wrequest_T WHERE [index] = " & Request("tmpID")
	rsW1.Open sqlW1, g_strCONNW, 1, 3
	If Not rsw1.EOF Then
		myInst = rsW1("InstID")
		tmpEmer = ""
		If rsW1("Emergency") = True Then tmpEmer = "checked"
		tmpEmerFee = ""
		If rsW1("EmerFee") = True Then tmpEmerFee = "checked"
		tmpInstRate = rsW1("InstRate")
		tmpDept = rsW1("DeptID")
		tmpReqP = rsW1("ReqID")
		tmpAppNum = rsW1("AppNum")
		
		tmpPDAmount = rsW1("PDamount")
	End If
	rsW1.Close
	Set rsW1 = Nothing
End If
'GET TEMP DATA
Set rsWdata = Server.CreateObject("ADODB.RecordSet")
sqlWdata = "SELECT * FROM Wrequest_T WHERE [index] = " & Request("tmpID")
rsWdata.Open sqlWdata, g_strCONNW, 1, 3
If Not rsWdata.EOF Then
	tmpInst = rsWdata("instID")
	tmpEmer = ""
	If rsWdata("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	tmpInstRate = Z_FormatNumber(rsWdata("InstRate"), 2)	
	myAppNum = rsWdata("AppNum")
End If
rsWdata.Close
Set rsWdata = Nothing
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE [index] = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpIname = rsInst("Facility") 
	PubDef = 0
	If rsInst("PD") Then PubDef = 1
End If
rsInst.Close
Set rsInst = Nothing 
'GET allowed mco
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM mco_T"
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	Do Until rsInst.EOF 
		If rsInst("mco") = "AmeriHealth" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "document.frmMain.rdoMed_Ame.disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "document.frmMain.rdoMed_Ame.disabled = true; " & vbCrLf 
			End If
		End If
		If rsInst("mco") = "Medicaid" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "document.frmMain.rdoMed_Med.disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "document.frmMain.rdoMed_Med.disabled = true; " & vbCrLf 
			End If
		End If
		If rsInst("mco") = "Meridian" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "document.frmMain.rdoMed_Mer.disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "document.frmMain.rdoMed_Mer.disabled = true; " & vbCrLf 
			End If
		End If
		If rsInst("mco") = "NHhealth" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "document.frmMain.rdoMed_NHH.disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "document.frmMain.rdoMed_NHH.disabled = true; " & vbCrLf 
			End If
		End If
		If rsInst("mco") = "WellSense" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "document.frmMain.rdoMed_Wel.disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "document.frmMain.rdoMed_Wel.disabled = true; " & vbCrLf 
			End If
		End If
		rsInst.MoveNext
	Loop
End If
rsInst.Close
Set rsInst = Nothing
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
	deptclass = rsDept("class")
	mydrg = rsDept("drg")
	tmpDname = rsDept("dept") 
	tmpDeptaddr = rsDept("address") & ", " & rsDept("InstAdrI") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	tmpBaddr = rsDept("Baddress") & ", " & rsDept("BCity") & ", " &  rsDept("Bstate") & ", " & rsDept("Bzip")
	tmpBContact = rsDept("Blname")
	tmpZipInst = ""
	If rsDept("zip") <> "" Then tmpZipInst = rsDept("zip")
	If tmpDeptaddrG = "" Then 
		'tmpDeptaddr = rsDept("InstAdrI") & " " & rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
		tmpDeptaddrG = rsDept("address") & ", " & rsDept("City") & ", " &  rsDept("state") & ", " & rsDept("zip")
	End If
	
End If
rsDept.Close
Set rsDept = Nothing 
'GET REQUESTING PERSON
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T WHERE [index] = " & tmpReqP
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
	tmpRP = rsReq("Lname") & ", " & rsReq("Fname") 
	Fon = rsReq("phone") 
	If rsReq("pExt") <> "" Then Fon = Fon & " ext. " & rsReq("pExt")
	Fax = rsReq("fax")
	email = rsReq("email")
	Pcon = GetPrime(tmpReqP)
	aFon = rsReq("aphone") 
End If
rsReq.Close
Set rsReq = Nothing
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT * FROM language_T WHERE [index] <> 95 ORDER BY [Language]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	tmpL = ""
	If tmpLang = "" Then tmpLang = -1
	If CInt(tmpLang) = rsLang("index") Then tmpL = "selected"
	strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("language") & "</option>" & vbCrlf
	strLangChk = strLangChk & "if (xxx == """ & Trim(rsLang("Language")) & """){ " & vbCrLf & _
		"return " & rsLang("index") & ";}"
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
'CREATE date and time info
If tmpAppNum > 1 Then
	AppCtr = 2
	Do Until AppCtr = tmpAppNum + 1
		strAppDate = strAppDate & "<tr><td>&nbsp;</td><td>" & vbCrLf & _
			"<input class='main' size='10' maxlength='10' name='txtAppDate" & AppCtr & "'  readonly value=''>" & vbCrLf & _
			"<input type='button' value='...' title='Calendar' name='cal" & AppCtr & "' style='width: 19px;'" & vbCrLf & _ 
			"onclick='showCalendarControl(document.frmMain.txtAppDate" & AppCtr & ");' class='btnLnk' onmouseover=""this.className='hovbtnLnk'"" onmouseout=""this.className='btnLnk'"">" & vbCrLf & _
			"<input type='hidden' name='mydate" & AppCtr & "' value='" & tmpAppDate & AppCtr & "'>" & vbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf & _
			"*Appointment Time:" & vbCrLf & _
			"&nbsp;From:<input class='main' size='5' maxlength='5' name='txtAppTFrom" & AppCtr & "' value='' onKeyUp=""javascript:return maskMe(this.value,this,'2,6',':');"" onBlur=""Javascript:return maskMe(this.value,this,'2,6',':');"">" & vbCrLf & _
			"&nbsp;To:<input class='main' size='5' maxlength='5' name='txtAppTTo" & AppCtr & "' value='' onKeyUp=""javascript:return maskMe(this.value,this,'2,6',':');"" onBlur=""javascript:return maskMe(this.value,this,'2,6',':');"">" & vbCrLf & _
			"<span class='formatsmall' onmouseover=""this.className='formatbig'"" onmouseout=""this.className='formatsmall'"">24-hour format</span>" & vbCrLf & _
			"<input type='hidden' name='mystime" & AppCtr & "' value='" & tmpAppTFrom & AppCtr & "'>" & vbCrLf & _
			"</td></tr>" & vbCrLf
			
			
		strDateChk = strDateChk & "	if (document.frmMain.txtAppDate" & AppCtr & ".value == """")" & vbCrLf & _
			"{alert('ERROR: Appointment Date(" & AppCtr & ") is Required.'); " & vbCrLf & _
			"	return;" & vbCrLf & _
			"} " & vbCrLf
			
		strTimeChk = strTimeChk & "	if (Trim(document.frmMain.txtAppTFrom" & AppCtr & ".value) == """")" & vbCrLf & _
			"{alert('ERROR: Appointment Time (From:)(" & AppCtr & ") is Required.'); " & vbCrLf & _
			"	return;" & vbCrLf & _
			"} "	
			
		strTimeChk = strTimeChk & "	if (Trim(document.frmMain.txtAppTTo" & AppCtr & ".value) == """")" & vbCrLf & _
			"{alert('ERROR: Appointment Time (To:)(" & AppCtr & ") is Required.'); " & vbCrLf & _
			"	return;" & vbCrLf & _
			"} "	 
			
		AppCtr = AppCtr + 1
	Loop
Else
	strAppDate = ""
End If
tmpFilename = Z_GenerateGUID()
Do Until GUIDExists(tmpFilename) = False
	tmpFilename = Z_GenerateGUID()
Loop
If mydrg Then 'secondary insur
	Const adOpenForwardOnly = 0
	Const adOpenKeyset      = 1
	Const adOpenDynamic     = 2
	Const adOpenStatic      = 3
	mySheet = "Alphabetical Order"
	my1stCell = "B3"
	myLastCell = "B900"
	my1stCell2 = "A3"
	myLastCell2 = "A900"
	strHeader = "HDR=NO;"
	myXlsFile = secinsPath & "CARRIER CODE LIST.xls"
	Set objExcel = CreateObject( "ADODB.Connection" )
	 objExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
	    myXlsFile & ";Extended Properties=""Excel 8.0;IMEX=1;" & _
	    strHeader & """"
	 Set objRS = CreateObject( "ADODB.Recordset" )
	    strRange = mySheet & "$" & my1stCell & ":" & myLastCell
	    objRS.Open "Select * from [" & strRange & "]", objExcel, adOpenStatic
	 Set objRS2 = CreateObject( "ADODB.Recordset" )
	    strRange2 = mySheet & "$" & my1stCell2 & ":" & myLastCell2
	    objRS2.Open "Select * from [" & strRange2 & "]", objExcel, adOpenStatic
	 i = 0
	    Do Until objRS.EOF
	
	      '  If IsNull( objRS.Fields(0).Value ) Or Trim( objRS.Fields(0).Value ) = "" Then Exit Do
	
	        For j = 0 To objRS.Fields.Count - 1
	            If Not IsNull( objRS.Fields(j).Value ) Or Trim(objRS.Fields(j).Value) <> "" Then
	           
	            	stroption =stroption & "<option value='" & objRS2.Fields(j).Value & "'>" & objRS.Fields(j).Value & "</option>" & vbCrlf
	               'arrData( j, i ) = Trim( objRS.Fields(j).Value )
	            End If
	        Next
	        ' Move to the next row
	        objRS.MoveNext
	        objRS2.MoveNext
	        ' Increment the array "row" number
	        i = i + 1
	    Loop
	 ' Close the file and release the objects
	 	objRS2.Close
    objRS.Close
    objExcel.Close
    Set objRS    = Nothing
    Set objRS2   = Nothing
    Set objExcel = Nothing
End If
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Interpreter Request Form - Appointment Information</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<script language='JavaScript'>
		<!--
		function Left(str, n){
			if (n <= 0)
			    return "";
			else if (n > String(str).length)
			    return str;
			else
			    return String(str).substring(0,n);
		}
		function RTrim(str)
    {
            var whitespace = new String(" \t\n\r");

            var s = new String(str);

            if (whitespace.indexOf(s.charAt(s.length-1)) != -1) {
               

                var i = s.length - 1;       
                while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1)
                    i--;


              
                s = s.substring(0, i+1);
            }

            return s;
    }
    function LTrim(str)
    {
            var whitespace = new String(" \t\n\r");

            var s = new String(str);

            if (whitespace.indexOf(s.charAt(0)) != -1) {
                
                var j=0, i = s.length;

                while (j < i && whitespace.indexOf(s.charAt(j)) != -1)
                    j++;

                s = s.substring(j, i);
            }

            return s;
    }
    function Trim(str)
    {
            return RTrim(LTrim(str));
    }
		function bawal(tmpform)
		{
			var iChars = ",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  		}
			  	else
		  		{
		  			tmp = tmp + tmpform.value.charAt(i);
		  		}
		  	}
		}
		function bawal2(tmpform)
		{
			var iChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz0123456789-,.\'"; //",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		tmp = tmp + tmpform.value.charAt(i);
		  		}
			  	else
		  		{
		  			alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  			
		  		}
		  	}
		}
		function CalendarView(strDate)
		{
			document.frmMain.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmMain.submit();
		}
		function maskMe(str,textbox,loc,delim)
		{
			var locs = loc.split(',');
			for (var i = 0; i <= locs.length; i++)
			{
				for (var k = 0; k <= str.length; k++)
				{
					 if (k == locs[i])
					 {
						if (str.substring(k, k+1) != delim)
					 	{
					 		str = str.substring(0,k) + delim + str.substring(k,str.length);
		     			}
					}
				}
		 	}
			textbox.value = str
		}
		function bawalletters(tmpform) {
			var iChars = "0123456789";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		tmp = tmp + tmpform.value.charAt(i);
		  		}
			  	else
		  		{
		  			alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  			
		  		}
		  	}
		}
		function WSubmit(xxx)
		{
			<% If mydrg Then %>
				if (document.frmMain.chkmed.checked == true) {
						if (document.frmMain.txtDOB.value == "") {
							alert("Please input client's date of birth.")
							return;
						}
						if (document.frmMain.rdoMed_Mer.checked == false &&
								document.frmMain.rdoMed_NHH.checked == false &&
								document.frmMain.rdoMed_Wel.checked == false &&
								document.frmMain.rdoMed_Med.checked == false &&
								document.frmMain.rdoMed_Ame.checked == false
								) {
							alert("Please select a Medicaid/MCO.")
							return;
						}
						if (document.frmMain.rdoMed_Ame.checked == true) {
							var strAHMid = Trim(document.frmMain.AHMemId.value);
							if (strAHMid == "") {
								alert("Please input client's AmeriHealth member ID.")
							} else if ((strAHMid.length < 8) || (strAHMid.length > 9) ) {
								alert("Client's AmeriHealth member ID is invalid.");
								return;
							}
							document.frmMain.AHMemId.value = strAHMid;
						}
						if (Trim(document.frmMain.MHPnum.value) == "" && document.frmMain.rdoMed_Mer.checked == true) {
							alert("Please input client's Meridian Health Plan number.")
							return;
						}
						if (Trim(document.frmMain.NHHFnum.value) == "" && document.frmMain.rdoMed_NHH.checked == true) {
							alert("Please input client's NH Healthy Families number.")
							return;
						}
						else {
							if (Trim(document.frmMain.NHHFnum.value) != "") {
								var chrmed = Trim(document.frmMain.NHHFnum.value);
								if (chrmed.length != 11) {
									alert("Invalid NH Healthy Families number length(11).");
									return;
								}
							}
						}
						if (Trim(document.frmMain.WSHPnum.value) == "" && document.frmMain.rdoMed_Wel.checked == true) {
							alert("Please input client's Well Sense Health Plan number.");
							return;
						}
						else {
							if (Trim(document.frmMain.WSHPnum.value) != "") {
								var chrmed = Trim(document.frmMain.WSHPnum.value);
								if (chrmed.length != 9) {
									alert("Invalid Well Sense Health Plan number length(9).");
									return;
								}
								var str = Left(document.frmMain.WSHPnum.value, 2);
								var res = str.toUpperCase(); 
								if (res != 'NH') {
									alert("Well Sense number MUST contain NH (eg: NHXXXXXXX).");
									return;
								}
							}
						}
						if (Trim(document.frmMain.MCnum.value) == "" && document.frmMain.rdoMed_Med.checked == true) {
							alert("Please input client's Medicaid number.");
							return;
						}
						else {
							if (Trim(document.frmMain.MCnum.value) != "") {
								var chrmed = Trim(document.frmMain.MCnum.value);
								if (chrmed.length != 11) {
									alert("Invalid Medicaid number length(11).");
									return;
								}
							}
						}
						if (document.frmMain.chkawk.checked == false) {
							alert("Acknowledge statement is required.");
							return;
						}
					}
			<% End If %>
			if (Trim(document.frmMain.txtCliAdd.value) != "" || Trim(document.frmMain.txtCliCity.value) != "" || Trim(document.frmMain.txtCliState.value) != "" || Trim(document.frmMain.txtCliZip.value) != "") {
				if (document.frmMain.chkClientAdd.checked == false) {
					alert("Alternate Appointment Address detected. If you wish to make this address as the appointment address, please check the checkbox beside it.");
					return;
				}
			}
			if (document.frmMain.chkClientAdd.checked == true)
			{
				if (Trim(document.frmMain.txtCliAdd.value) == "" || Trim(document.frmMain.txtCliCity.value) == "" || Trim(document.frmMain.txtCliState.value) == "" || Trim(document.frmMain.txtCliZip.value) == "")
				{
					alert("Please input Alternate Appointment's full address.");
					return;
				}
			}
			if (document.frmMain.txtClilname.value == "" && document.frmMain.txtClifname.value == "")
			{
				alert("ERROR: Client is Required."); 
				return;
			}
			if ((document.frmMain.chkcall.checked == true || document.frmMain.chkleave.checked == true) && document.frmMain.txtCliFon.value == "") {
				alert("Please input client's phone number.");
				return;
			}
			if (document.frmMain.txtCliFon.value != "") {
				document.frmMain.chkcall.checked = true;
			}
			if (document.frmMain.selLang.value == 0)
			{
				alert("ERROR: Language is Required."); 
				return;
			}
			if (document.frmMain.txtAppDate.value == "")
			{
				alert("ERROR: Appointment Date is Required."); 
				return;
			}
			<%=strDateChk%>
			if (Trim(document.frmMain.txtAppTFrom.value) == "")
			{
				alert("ERROR: Appointment Time (From:) is Required."); 
				return;
			}
			if (document.frmMain.txtAppTFrom.value == "24:00")
			{
				alert("ERROR: Appointment Time (From:) is invalid (24:00 not accepted)."); 
				return;
			}
			if (Trim(document.frmMain.txtAppTTo.value) == "")
			{
				alert("ERROR: Appointment Time (To:) is Required."); 
				return;
			}
			if (document.frmMain.txtAppTTo.value == "24:00")
			{
				alert("ERROR: Appointment Time (To:) is invalid (24:00 not accepted)."); 
				return;
			}
			<%=strTimeChk%>
			<% If PubDef = 1 Then %>
				if (document.frmMain.txtDocNum.value == "")
				{
					alert("ERROR: Docket Number is Required."); 
					return;
				}
				if (document.frmMain.txtPDamount.value == "")
				{
					alert("ERROR: Amount requested from court is Required."); 
					return;
				}
			<% End If %>
			<% If deptclass = 4 Or deptclass = 6 Then %>
				//if (document.frmMain.mrrec.value == "")
				//{
				//	alert("ERROR: Please provide Patient MR#.")
				//	return;
				//}
			<% End If %>
			var ans = window.confirm("Submit Appointment to Database?");
			if (ans){
				document.frmMain.action = "waction.asp?ctrl=4";
				document.frmMain.submit();
			}
		}
		function WBack(xxx)
		{
			var ans = window.confirm("Any changes made in this page will not be saved.");
			if (ans){
				document.frmMain.action = "wMain3.asp?tmpID=" + xxx;
				document.frmMain.submit();
			}
		}
		<% If mydrg Then %>
			function OutPatient() {
				if (document.frmMain.chkout.checked == true) {
					document.frmMain.chkmed.disabled = false;
				} else {
					document.frmMain.chkmed.checked = false;
					document.frmMain.chkmed.disabled = true;

					document.frmMain.rdoMed_Ame.disabled = true;
					document.frmMain.rdoMed_Med.disabled = true;
					document.frmMain.rdoMed_Mer.disabled = true;
					document.frmMain.rdoMed_NHH.disabled = true;
					document.frmMain.rdoMed_Wel.disabled = true;

					document.frmMain.rdoMed_Ame.checked = false;
					document.frmMain.rdoMed_Med.checked = false;
					document.frmMain.rdoMed_Mer.checked = false;
					document.frmMain.rdoMed_NHH.checked = false;
					document.frmMain.rdoMed_Wel.checked = false;

					document.frmMain.AHMemId.value = "";
					document.frmMain.MHPnum.value = "";
					document.frmMain.NHHFnum.value = "";
					document.frmMain.WSHPnum.value = "";
					document.frmMain.MCnum.value = "";
					document.frmMain.chkawk.disabled = true;
					document.frmMain.AHMemId.disabled = true;
					document.frmMain.MHPnum.disabled = true;
					document.frmMain.NHHFnum.disabled = true;
					document.frmMain.WSHPnum.disabled = true;
					document.frmMain.MCnum.disabled = true;
				}
			}
			function HasMedicaid(dept) {
			if (document.frmMain.chkmed.checked == true) {
				//document.frmMain.MCnum.disabled = false;
				<%=allowMCO%>
				document.frmMain.chkawk.disabled = false;
			} else {
				document.frmMain.rdoMed_Ame.disabled = true;
				document.frmMain.rdoMed_Med.disabled = true;
				document.frmMain.rdoMed_Mer.disabled = true;
				document.frmMain.rdoMed_NHH.disabled = true;
				document.frmMain.rdoMed_Wel.disabled = true;

				document.frmMain.rdoMed_Ame.checked = false;
				document.frmMain.rdoMed_Med.checked = false;
				document.frmMain.rdoMed_Mer.checked = false;
				document.frmMain.rdoMed_NHH.checked = false;
				document.frmMain.rdoMed_Wel.checked = false;

				document.frmMain.AHMemId.value = "";
				document.frmMain.MHPnum.value = "";
				document.frmMain.NHHFnum.value = "";
				document.frmMain.WSHPnum.value = "";
				document.frmMain.MCnum.value = "";

				document.frmMain.chkawk.disabled = true;
				document.frmMain.AHMemId.disabled = true;
				document.frmMain.MCnum.disabled = true;
				document.frmMain.MHPnum.disabled = true;
				document.frmMain.NHHFnum.disabled = true;
				document.frmMain.WSHPnum.disabled = true;
			}
		}
		<% End If %>
		function uploadFile()
		{
			var tmpfname = "<%=tmpFilename%>";
			newwindow = window.open('upload.asp?hfname=' + tmpfname ,'name','height=150,width=400,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
		}
		function PDchk() {
			if (document.frmMain.h_PD.value == 1) {
				document.frmMain.btnUp.disabled = false;
			}
			else {
				document.frmMain.btnUp.disabled = true;
			}				
		}	
		function SelPlan() {
				document.frmMain.AHMemId.disabled = true;
				document.frmMain.MHPnum.disabled = true;
				document.frmMain.NHHFnum.disabled = true;
				document.frmMain.WSHPnum.disabled = true;
				document.frmMain.MCnum.disabled = true;
				if (document.frmMain.rdoMed_Mer.checked == true) {
					document.frmMain.MHPnum.disabled = false;
					document.frmMain.NHHFnum.value = "";
					document.frmMain.WSHPnum.value = "";
					document.frmMain.MCnum.disabled = false;
				}
				if (document.frmMain.rdoMed_NHH.checked == true) {
					document.frmMain.NHHFnum.disabled = false;
					document.frmMain.MHPnum.value = "";
					document.frmMain.WSHPnum.value = "";
					//document.frmMain.MCnum.value = "";
					document.frmMain.MCnum.disabled = false;
				}
				if (document.frmMain.rdoMed_Wel.checked == true) {
					document.frmMain.WSHPnum.disabled = false;
					document.frmMain.NHHFnum.value = "";
					document.frmMain.MHPnum.value = "";
					//document.frmMain.MCnum.value = "";
					document.frmMain.MCnum.disabled = false;
				}
				if (document.frmMain.rdoMed_Med.checked == true) {
					//document.frmMain.MCnum.disabled = false;
					document.frmMain.NHHFnum.value = "";
					document.frmMain.WSHPnum.value = "";
					document.frmMain.MHPnum.value = "";
					document.frmMain.MCnum.disabled = false;
				}
				if (document.frmMain.rdoMed_Ame.checked == true) {
					document.frmMain.MHPnum.value = "";
					document.frmMain.NHHFnum.value = "";
					document.frmMain.WSHPnum.value = "";
					document.frmMain.AHMemId.disabled = false;
				}
			}
			function Chkdrg(tmpdept) {
				<% If Not myDRG Then %>
					//document.frmMain.chkmed.checked = false;
					document.frmMain.MCnum.value = "";
					//document.frmMain.chkmed.disabled = true;
					document.frmMain.MCnum.disabled = true;
					document.frmMain.chkacc.disabled = true;
					document.frmMain.chkcomp.disabled = true;
					//document.frmMain.selIns.disabled = true;
					document.frmMain.chkacc.checked = false;
					document.frmMain.chkcomp.checked = false;
					//document.frmMain.btnSec.disabled = true;
					document.frmMain.selIns.value = "";
					document.frmMain.chkout.checked = false;
					document.frmMain.chkout.disabled = true;
						document.frmMain.chkawk.disabled = true;
				<% Else %>
					document.frmMain.chkout.disabled = false;
					document.frmMain.chkmed.disabled = false;
					//document.frmMain.MCnum.disabled = false;
					document.frmMain.chkacc.disabled = false;
					document.frmMain.chkcomp.disabled = false;
					document.frmMain.chkawk.disabled = false;
					//document.frmMain.selIns.disabled = false;
					//document.frmMain.btnSec.disabled = false;
					OutPatient();
					HasMedicaid(tmpdept);
				<% End If %>
			}
			function DpwedeMed() {
				if (document.frmMain.chkacc.checked == true || document.frmMain.chkcomp.checked == true) {
					alert("This appointment is not eligible for Medicaid/MCO.");
					document.frmMain.chkout.checked = false;
					return;
				}
			}
			function DpwedeIba() {
				if (document.frmMain.chkout.checked == true) {
					if (document.frmMain.chkacc.checked == true || document.frmMain.chkcomp.checked == true) {
						alert("This appointment is not eligible for Auto Accident and/or Worker's compensation.");
						document.frmMain.chkacc.checked = false;
						document.frmMain.chkcomp.checked = false;
						return;
					}
				}
			}
			function chkleavemsg() {
				if (document.frmMain.chkcall.checked == false) {
					document.frmMain.chkleave.checked = false;
				}
			}
		//-->
		</script>
		</head>
		<body onload='PDchk();
			<% If mydrg Then %> 
				 Chkdrg(<%=Z_CZero(tmpdept)%>); SelPlan();
			<% End If %>
			'>
			<form method='post' name='frmMain'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<form name='frmService' method='post' action=''>
								<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
									<!-- #include file="_greetme.asp" -->
									<tr>
										<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - 4 / 4</td>
									</tr>
									<tr>
										<td align='center' colspan='10'><nobr>(*) required</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td align='left'>
											<div name="dErr" style="width:100%; height:55px;OVERFLOW: auto;">
												<table border='0' cellspacing='1'>		
													<tr>
														<td><span class='error'><%=Session("MSG")%></span></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr>
										<td class='header' colspan='10'><nobr>Contact Information</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
									<td align='right'>Institution:</td>
									<td class='confirm'><%=tmpIname%></td>
									</tr>
									<tr>
										<td align='right'>Department:</td>
										<td class='confirm'><%=tmpDname%></td>
									</tr>
									<tr>
										<td align='right'>Address:</td>
										<td class='confirm'><%=tmpDeptaddr%></td>
									</tr>
									<tr>
										<td align='right'>Billed To:</td>
										<td class='confirm'><%=tmpBContact%></td>
									</tr>
									<tr>
										<td align='right'>Billing Address:</td>
										<td class='confirm'><%=tmpBaddr%></td>
									</tr>
									<% If Request.Cookies("LBUSERTYPE") = 1 Then %>
										<tr>
											<td align='right' width='15%'>Rate:</td>
											<td class='confirm'><%=tmpInstRate%></td>
										</tr>
									<% End If %>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'>Requesting Person:</td>
										<td class='confirm'><%=tmpRP%></td>
									</tr>
									<tr>
										<td align='right'>Phone:</td>
										<td class='confirm'><%=fon%></td>
									</tr>
									<tr>
										<td align='right'>Fax:</td>
										<td class='confirm'><%=fax%></td>
									</tr>
									<tr>
										<td align='right'>E-Mail:</td>
										<td class='confirm'><%=email%></td>
									</tr>
										<tr>
										<td align='right'>Alternate Phone:</td>
										<td class='confirm'><%=afon%></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' class='header'><nobr>Appointment Information</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									
									<tr>
										<td align='right' valign='top'>&nbsp;</td>
										<td colspan='2'>
											<input type='checkbox' name='chkblock' value='1'>
												&nbsp;Block Schedule
												&nbsp;&nbsp;
												<input type='checkbox' name='chkClient' value='1' <%=chkClient%>>&nbsp;LSS Client
										</td>
									</tr>
									
									<tr>
										<td align='right'>*Client Last Name:</td>
										<td>
											<input class='main' size='20' maxlength='25' name='txtClilname' value="<%=tmplname%>" onkeyup='bawal2(this);'>&nbsp;First Name:
											<input class='main' size='20' maxlength='25' name='txtClifname' value="<%=tmpfname%>" onkeyup='bawal2(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtCliAddrI' value='<%=tmpCAdrI%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'><nobr>Alternate Appointment Address:</td>
										<td colspan='3'><nobr>
											<input class='main' size='50' maxlength='50' name='txtCliAdd' value='<%=tmpAddr%>' onkeyup='bawal(this);'>
											<input type='checkbox' name='chkClientAdd' value='1' <%=chkUClientadd%>>CHECK this box and FILL these fields if appointment address is different from department address
											<br>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td>
											<input class='main' size='25' maxlength='25' name='txtCliCity' value='<%=tmpCity%>' onkeyup='bawal(this);'>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtCliState' value='<%=tmpState%>' onkeyup='bawal(this);'>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtCliZip' value='<%=tmpZip%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>Client Phone:</td>
										<td><input class='main' size='12' maxlength='12' name='txtCliFon' value='<%=tmpCFon%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td colspan="3" align="left">
											<input type='checkbox' name='chkcall' value='1' <%=chkcall%>  onclick='chkleavemsg();'>
											Language Bank Interpreter to provide courtesy reminder call (Please note that this is ONLY courtesy reminder call and patient/client may still not show up to his/her appointment).
											<br><br>
										</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td colspan="3" align="left">
											<input type='checkbox' name='chkleave' value='1' <%=chkleave%> onclick='chkleavemsg();'>
											If a patient/client does not answer the phone and his answering machine/voice mail picks up a call or family member answers the phone, can interpreter provide/give full appointment<br>
											info (date, time, location, name of hospital/clinic/department, providers name) on patient/client voice message or give this info to patient/client�s family member?
											<br><br>
										</td>
									</tr>
									<tr>
										<td align='right' valign='top'>Alter. Phone:</td>
										<td align='left'>
											<textarea name='txtAlter' class='main' onkeyup='bawal(this);' ><%=tmpCAFon%></textarea>
										</td>
									</tr>
									<tr>
										<td align='right'>Gender:</td>
										<td>
											<select class='seltxt' name='selGender' style='width: 75px;'>
												<option value ='-1'> &nbsp; </option>
												<option value='0' <%=tmpMale%>>Male</option>
												<option value='1' <%=tmpfeMale%>>Female</option>
											</select>
											&nbsp;&nbsp;
											Minor:
											<input type='checkbox' name='chkMinor' value='1' <%=chkMinor%>>
										</td>
									</tr>
									<tr>
										<td align='right'>Directions / Landmarks:</td>
										<td><input class='main' size='50' maxlength='50' name='txtCliDir' value='<%=tmpDir%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right' valign='top'>Special Circumstances/Precautions:</td>
										<td>
											<textarea name='txtCliCir' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpSC%></textarea>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Precautions (infections, safety, etc.) for this appointment</span>
										</td>
									
									</tr>
									<tr>
										<td align='right'>DOB:</td>
										<td>
											<input class='main' size='11' maxlength='10' name='txtDOB' value='<%=tmpDOB%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										</td>
										
									</tr>
									<tr>
											<td align='right'>Patient MR #:</td>
											<td>
												<input class='main' size='50' maxlength='50' name='mrrec' value='<%=mrrec%>' onkeyup='bawal(this);'>
											</td>
										</tr>
									<tr>
										<td align='right'>*Language:</td>
										<td>
											<select class='seltxt' name='selLang'  style='width:100px;' onchange=''>
												<option value='0'>&nbsp;</option>
												<%=strLang%>
											</select>
											<input type='hidden' name='myLang' value='<%=tmpLang%>'>
										</td>
									</tr>
										<tr>
											<td align='right'>*Appointment Date:</td>
											<td>
												<input class='main' size='10' maxlength='10' name='txtAppDate'  readonly value='<%=tmpAppDate%>'>
												<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;"
												onclick="showCalendarControl(document.frmMain.txtAppDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
												<input type='hidden' name='mydate' value='<%=tmpAppDate%>'>
												&nbsp;&nbsp;&nbsp;&nbsp;
												*Appointment Time:
												
													&nbsp;From:<input class='main' size='5' maxlength='5' name='txtAppTFrom' value='<%=tmpAppTFrom%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
													&nbsp;To:<input class='main' size='5' maxlength='5' name='txtAppTTo' value='<%=tmpAppTTo%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
													<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">24-hour format</span>
													<input type='hidden' name='mystime' value='<%=tmpAppTFrom%>'>
												
											</td>
										</tr>
										<%=strAppDate%>
					
									<tr><td>&nbsp;</td></tr>
									<tr><td>&nbsp;</td></tr>
									<% If mydrg Then %>
										<tr>
											<td align='right'><b>For Medicaid/MCO billing:</b></td>
											<td><b>(also fill in)</b></td>
										</tr>
										<tr>
											<td align='right'>Auto Accident:</td>
											<td><input type='checkbox' name='chkacc' value='1' <%=chkacc%> onclick="DpwedeIba();"></td>
										</tr>
										<tr>
											<td align='right'>Worker's Compensation:</td>
											<td><input type='checkbox' name='chkcomp' value='1' <%=chkcomp%> onclick="DpwedeIba();"></td>
										</tr>
										<tr>
											<td align='right'>Outpatient:</td>
											<td><input type='checkbox' name='chkout' value='1' <%=chkout%> onclick="DpwedeMed(); OutPatient();"></td>
										</tr>
										<tr>
											<td align='right'>Has Medicaid/MCO:</td>
											<td>
												<input type='checkbox' name='chkmed' value='1' <%=chkmed%> onclick="HasMedicaid(<%=tmpDept%>);">
												Medicaid:<input type='text' class='main' maxlength='14' name='MCnum' value="<%=MCNum%>">
											</td>
										</tr>
										<tr><td align='right'></td><td colspan='3'>
							<!-- START :: new for 2019-11-22: AMERIHEALTH AND AMERIHEALTH MEMBER ID  -->
							<input type='radio' name='radiomed' <%=radiomed5%> id="rdoMed_Ame" value='5' onclick='SelPlan();'>
							AmeriHealth
							<input type='text' class='main' maxlength='9' minlength="8" placeholder="member ID"
								name='AHMemId' value="<%=AHMemId%>" /><br/>
							<!-- END :: for 2019-11-22: AMERIHEALTH AND AMERIHEALTH MEMBER ID  -->

							<input type='radio' name='radiomed' <%=radiomed1%> id="rdoMed_Mer" value='1' onclick='SelPlan();'>
							Meridian Health Plan
							<input type='text' class='main' maxlength='14' name='MHPnum' value="<%=MHPnum%>"><br>

							<input type='radio' name='radiomed' <%=radiomed2%> id="rdoMed_NHH" value='2' onclick='SelPlan();'>
							NH Healthy Families
							<input type='text' class='main' maxlength='14' name='NHHFnum' value="<%=NHHFnum%>" onkeyup='bawalletters(this);'><br>

							<input type='radio' name='radiomed' <%=radiomed3%> id="rdoMed_Wel" value='3' onclick='SelPlan();'>
							Well Sense Health Plan
							<input type='text' class='main' maxlength='14' name='WSHPnum' value="<%=WSHPnum%>"><span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">(Well Sense number MUST contain NH (eg: NHXXXXXXX).)</span><br>

							<input type='radio' name='radiomed' <%=radiomed4%> id="rdoMed_Med" value='4' onclick='SelPlan();'>
							Medicaid
							<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">(Directly Billed to Medicaid/Straight Medicaid/Non-MCO)</span> 
							<br>

												<br /></td></tr>
											<tr>
												<td>&nbsp;</td>
												<td colspan="3" align="left">
													<input type='checkbox' name='chkawk' value='1' <%=chkawk%> >
													Acknowledgement Statement:<br> On behalf of my organization/institution, I/we agree to accept financial responsibility for this appointment and agree to pay Language Bank for interpretation services provided to us, if MCO or Medicaid declines to pay/cover this appointment.<br><br>
													 I acknowledge that appointment entered is NOT Auto Accident or Workers Compensation case. On behalf of my organization/institution, I/we agree to reimburse/pay Language Bank if the state or MCO request repayment (if case is to be Auto Accident or Workers Compensation case). 
													<br><br>
												</td>
												<!--<td><input type='text' class='main' maxlength='14' name='MCnum' value="<%=MCNum%>"></td>//-->
											</tr>
										
										<tr>
											<td align='right'>Secondary Insurance:</td>
											<td>	
												<select class='seltxt' name='selIns'  style='width:200px;' onchange=''>
													<option value='0'>&nbsp;</option>
													<%=stroption%>
												</select>
											</td>
										</tr>
										<tr><td>&nbsp;</td></tr>
									<% End If %>
									<tr>
										<td align='right'><b>For Legal appointments:</b></td>
										<td><b>(also fill in)</b></td>
									</tr>
									<tr>
										<td align='right'>Judge:</td>
										<td><input class='main' size='50' maxlength='50' name='txtJudge' value='<%=tmpJudge%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Claimant:</td>
										<td><input class='main' size='50' maxlength='50' name='txtClaim' value='<%=tmpClaim%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<% If tmpInst = 757 Or tmpInst = 777 Then %>
											<td align='right'>Delivery Ticket:</td>
										<% Else %>	
											<td align='right'>Docket Number:</td>
										<% End If %>
										<td><input class='main' size='50' maxlength='50' name='txtDocNum' value='<%=tmpDoc%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr>
										<td align='right'>Court Room No:</td>
										<td><input class='main' size='12' maxlength='12' name='txtCrtNum' value='<%=tmpCRN%>' onkeyup='bawal(this);'></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'><b>For Public Defender:</b></td>
										<td><b>(also fill in)</b></td>
									</tr>
									<tr>
										<td align='right'>Amount requested from court:</td>
										<td>
											$<input class='main' size='8' maxlength='7' name='txtPDamount' value='<%=tmpPDAmount%>'>
										</td>
									</tr>
									<tr>
										<td align='right'>Form 604A:</td>
										<td class='RemME'>
											<input type="button" name="btnUp" value="UPLOAD" onclick="uploadFile();" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" <%=disUpload%>>
											<!--<input  class='main' type="file" name="F1" size="20" class='btn'>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*PDF format only</span>//-->
											<input type="hidden" name="h_tmpfilename" value='<%=tmpFilename%>'>
											<%=tmpfileuploaded%>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>	
										<td align='right' valign='top'>Appointment Comment:</td>
										<td colspan='3' >
											<textarea name='txtcom' class='main' onkeyup='bawal(this);' style='width: 375px;'><%=tmpCom%></textarea>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
											<input class='btn' type='button' value='<<' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WBack(<%=Request("tmpID")%>);'>
											<input class='btn' type='Reset' value='Clear' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
											<input class='btn' type='button' value='Cancel' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="window.location='calendarview2.asp'">
											<input class='btn' type='button' value='Submit' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WSubmit(<%=Request("tmpID")%>);'>
											<input type='hidden' name='h_PD' value='<%=PubDef%>'>
											<input type='hidden' name='tmpID' value='<%=Request("tmpID")%>'>
											<input type='hidden' name='tmpInst' value='<%=tmpInst%>'>
											<input type='hidden' name='tmpDep' value='<%=tmpDept%>'>
											<input type='hidden' name='tmpReqP' value='<%=tmpReqP%>'>
										</td>
									</tr>
									
								</table>
							</form>
						</td>
					</tr>
					<tr>
						<td valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
		</body>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>
