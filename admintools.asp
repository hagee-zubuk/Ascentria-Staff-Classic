<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
server.scripttimeout = 360000
'USER CHECK
If Cint(Request.Cookies("LBUSERTYPE")) <> 1 Then
	Session("MSG") = "Error: Please sign-in as user with admin rights."
	Response.Redirect "default.asp"
End If
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function
tmpPage = "document.frmAdmin."
selRPEmail = ""
selRPPhone = ""
selRPFax = "checked"
'GET REQUESTING PERSON LIST
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T ORDER BY Lname, Fname"
rsReq.Open sqlReq, g_strCONN, 3, 1
Do Until rsReq.EOF
	tmpReqP = Request("ReqID")
	ReqSel = ""
	If tmpReqP = "" Then tmpReqP = -1
	tmpReqName = CleanMe(rsReq("lname")) & ", " & CleanMe(rsReq("fname"))
	If CInt(tmpReqP) = rsReq("index") Then ReqSel = "selected"
	strReq2 = strReq2 & "<option " & ReqSel & " value='" & rsReq("Index") & "'>" & rsReq("Lname") & ", " & rsReq("Fname") & "</option>" & vbCrLf
	strReq = strReq & "{var ChoiceReq = document.createElement('option');" & vbCrLf & _
			"ChoiceReq.value = " & rsReq("index") & ";" & vbCrLf & _
			"ChoiceReq.appendChild(document.createTextNode(""" & tmpReqName & """));" & vbCrLf & _
			"document.frmAdmin.selReq.appendChild(ChoiceReq);}" & vbCrLf
			
		strReqI = strReqI & "if (Req == " & rsReq("Index") & ") " & vbCrLf & _
		"{document.frmAdmin.txtphone.value = """ & rsReq("Phone") &"""; " & vbCrLf & _
		"document.frmAdmin.txtReqExt.value = """ & rsReq("pExt") &"""; " & vbCrLf & _
		"document.frmAdmin.txtfax.value = """ & rsReq("Fax") &"""; " & vbCrLf & _
		"document.frmAdmin.txtemail.value = """ & rsReq("Email") &"""; " & vbCrLf & _
		"document.frmAdmin.txtReqFname.value = """ & rsReq("Fname") &"""; " & vbCrLf & _
		"document.frmAdmin.txtReqLname.value = """ & rsReq("Lname") &"""; " & vbCrLf
		If rsReq("prime") = 0 Then
			strReqI = strReqI & "document.frmAdmin.radioPrim1[2].checked = true;" & vbCrLf 
		ElseIf rsReq("prime") = 1 Then
			strReqI = strReqI & "document.frmAdmin.radioPrim1[0].checked = true;" & vbCrLf 
		ElseIf rsReq("prime") = 2 Then
			strReqI = strReqI & "document.frmAdmin.radioPrim1[1].checked = true;" & vbCrLf 
		End If
		strReqI = strReqI & "document.frmAdmin.ReqName.value = """ & rsReq("Lname") & ", " & rsReq("Fname") &"""; }" & vbCrLf
	rsReq.MoveNext
Loop
rsReq.Close
Set rsReq = Nothing
'GET REQUESTING PERSON INFO
'Set rsReqI = Server.CreateObject("ADODB.RecordSet")
'sqlReqI = "SELECT * FROM requester_T ORDER BY Lname, Fname"
'rsReqI.Open sqlReqI, g_strCONN, 3, 1
'Do Until rsReqI.EOF
'
'		rsReqI.MoveNext
'Loop
'rsReqI.Close
'Set rsReqI = Nothing
'REQUESTING PERSON CHECKER
'Set rsReqCHK = Server.CreateObject("ADODB.RecordSet")
'sqlReqCHK = "SELECT * FROM requester_T"
'rsReqCHK.Open sqlReqCHK, g_strCONN, 3, 1
'Do Until rsReqCHK.EOF
'	strReqCHK = strReqCHK & "if (document.frmAdmin.txtReqLname.value == """ & Trim(rsReqCHK("lname")) & """ && document.frmAdmin.txtReqFname.value == """ & Trim(rsReqCHK("Fname")) & """) " & vbCrLf & _
'		"{var ans = window.confirm(""Requester's name already exists. Click on Cancel to rename. Click on OK to continue.""); " & vbCrLf & _
'		"{if (ans){ " & vbCrLf & _
'		"pnt = 1; " & vbCrLf & _
'		"} " & vbCrLf & _
'		"else " & vbCrLf & _
'		"{ " & vbCrLf & _
'		"return; " & vbCrLf & _
'		"} " & vbCrLf & _
'		"} " & vbCrLf & _
'		"} " & vbCrLf & _
'		"else " & vbCrLf & _
'		"{pnt = 1; " & vbCrLf & _
'		"} " & vbCrLf
'	rsReqCHK.MoveNext 
'Loop
'rsReqCHK.Close
'Set rsReqCHK = Nothing
'GET INSTITUTION LIST
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT Facility, [index] FROM institution_T ORDER BY Facility"
rsInst.Open sqlInst, g_strCONN, 3, 1
Do Until rsInst.EOF
	tmpInst = Request("InstID")
	tmpDO = ""
	If tmpInst = "" Then tmpInst = -1
	If Cint(tmpInst) = rsInst("index") Then tmpDO = "selected"
	InstName = rsInst("Facility")
	strInst = strInst	& "<option " & tmpDO & " value='" & rsInst("Index") & "'>" &  InstName & "</option>" & vbCrlf
	
	strInstI = strInstI & "if (Inst == " & rsInst("Index") & ") " & vbCrLf & _
		"{document.frmAdmin.txtNewInst.value = """ & rsInst("Facility") &"""; " & vbCrLf & _
		"document.frmAdmin.InstName.value = """ & rsInst("Facility") & """; }" & vbCrLf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing
'GET INSTITUTION INFO
'Set rsInst = Server.CreateObject("ADODB.RecordSet")
'sqlInst = "SELECT Facility FROM institution_T ORDER BY Facility"
'rsInst.Open sqlInst, g_strCONN, 3, 1
'Do Until rsInst.EOF
'	strInstI = strInstI & "if (Inst == " & rsInst("Index") & ") " & vbCrLf & _
'		"{document.frmAdmin.txtNewInst.value = """ & rsInst("Facility") &"""; " & vbCrLf & _
'		"document.frmAdmin.InstName.value = """ & rsInst("Facility") & """; }" & vbCrLf
'	rsInst.MoveNext
'Loop
'rsInst.Close
'Set rsInst = Nothing

'GET DEPARTMENTS
Set rsDept2 = Server.CreateObject("ADODB.RecordSet")
sqlDept2 = "SELECT * FROM dept_T ORDER BY Dept"
rsDept2.Open sqlDept2, g_strCONN, 3, 1
Do Until rsDept2.EOF
	tmpDepart = Request("DeptID")
	tmpDpt = ""
	If Z_CZero(tmpDepart) = rsDept2("index") Then tmpDpt = "selected"
	DeptName = rsDept2("Dept")
	'If rsInst("Department") <> "" Then InstName = rsInst("Facility") & " - " & rsInst("Department")
	strDept2 = strDept2	& "<option " & tmpDpt & " value='" & rsDept2("Index") & "'>" &  DeptName & "</option>" & vbCrlf
	
	strDept = strDept & "if (dept == " & rsDept2("Index") & ") " & vbCrLf & _
		"{document.frmAdmin.txtInstAddr.value = """ & rsDept2("address") &"""; " & vbCrLf & _
		"document.frmAdmin.selDept.value = " & rsDept2("Index") & "; " & vbCrLf & _
		"document.frmAdmin.txtNewDept.value = """ & rsDept2("dept") & """; " & vbCrLf & _
		"document.frmAdmin.txtInstCity.value = """ & rsDept2("city") &"""; " & vbCrLf & _
		"document.frmAdmin.txtInstState.value = """ & rsDept2("state") &"""; " & vbCrLf & _
		"document.frmAdmin.txtInstZip.value = """ & rsDept2("zip") &"""; " & vbCrLf & _
		"document.frmAdmin.txtInstAddrI.value = """ & rsDept2("InstAdrI") &"""; " & vbCrLf & _
		"document.frmAdmin.txtBlname.value = """ & rsDept2("BLname") &"""; " & vbCrLf & _
		"document.frmAdmin.txtBillAddr.value = """ & rsDept2("Baddress") &"""; " & vbCrLf & _
		"document.frmAdmin.txtBillCity.value = """ & rsDept2("Bcity") &"""; " & vbCrLf & _
		"document.frmAdmin.txtBillState.value = """ & rsDept2("Bstate") &"""; " & vbCrLf & _
		"document.frmAdmin.txtBillZip.value = """ & rsDept2("Bzip") &"""; " & vbCrLf & _
		"document.frmAdmin.DepartName.value = """ & rsDept2("dept") &"""; " & vbCrLf & _
		"document.frmAdmin.selClass.value = """ & rsDept2("Class") &"""; }" & vbCrLf 
	rsDept2.MoveNext
Loop
rsDept2.Close
Set rsDept2 = Nothing
'GET AVAILABLE DEPARTMENTS
Set rsInstDept = Server.CreateObject("ADODB.RecordSet")
sqlInstDept = "SELECT * FROM institution_T ORDER BY Facility"
rsInstDept.Open sqlInstDept, g_strCONN, 3, 1
Do Until rsInstDept.EOF
	InstDept = rsInstDept("Index")
	strInstDept = strInstDept & "if (inst == " & InstDept & "){" & vbCrLf
	Set rsDeptInst = Server.CreateObject("ADODB.RecordSet")
	sqlDeptInst = "SELECT * FROM dept_T WHERE InstID = " &  InstDept & " ORDER BY Dept"
	rsDeptInst.Open sqlDeptInst, g_strCONN, 3, 1
	If Not rsDeptInst.EOF Then
		Do Until rsDeptInst.EOF
			strInstDept = strInstDept & "if (dept != " & rsDeptInst("index") & ")" & vbCrLf & _
				"{var ChoiceInst = document.createElement('option');" & vbCrLf & _
				"ChoiceInst.value = " & rsDeptInst("index") & ";" & vbCrLf & _
				"ChoiceInst.appendChild(document.createTextNode(""" & rsDeptInst("Dept") & """));" & vbCrLf & _
				"document.frmAdmin.selDept.appendChild(ChoiceInst);} " & vbCrlf
			rsDeptInst.MoveNext
		Loop
	End If
	rsDeptInst.Close
	Set rsDeptInst = Nothing
	rsInstDept.MoveNext
	strInstDept = strInstDept & "}"
Loop
rsInstDept.Close
Set rsInstDept = Nothing
'GET DEPT INFO
'Set rsDept = Server.CreateObject("ADODB.RecordSet")
'sqlDept = "SELECT * FROM dept_T ORDER BY dept"
'rsDept.Open sqlDept, g_strCONN, 3, 1
'Do Until rsDept.EOF
'	
'	rsDept.MoveNext
'Loop
'rsDept.Close
'Set rsDept = Nothing
'GET AVAILABLE REQUESTING PERSON PER DEPARTMENT
Set rsInstReq = Server.CreateObject("ADODB.RecordSet")
sqlInstReq = "SELECT * FROM dept_T ORDER BY dept"
rsInstReq.Open sqlInstReq, g_strCONN, 3, 1
Do Until rsInstReq.EOF
	InstReq = rsInstReq("index")
	strInstReqDept = strInstReqDept & "if (dept == " & InstReq & "){" & vbCrLf
	Set rsReqInst = Server.CreateObject("ADODB.RecordSet")
	sqlReqInst = "SELECT * FROM requester_T, reqdept_T WHERE  ReqID = requester_T.index AND DeptID = " & InstReq & " ORDER BY lname, fname"
	rsReqInst.Open sqlReqInst, g_strCONN, 3, 1
	Do Until rsReqInst.EOF
		tmpReqName = CleanMe(rsReqInst("lname")) & ", " & CleanMe(rsReqInst("fname"))
		strInstReqDept = strInstReqDept	& "if(req != "& rsReqInst("requester_T.index") & ")" & vbCrLf & _
			"{var ChoiceReq = document.createElement('option');" & vbCrLf & _
			"ChoiceReq.value = " & rsReqInst("requester_T.index") & ";" & vbCrLf & _
			"ChoiceReq.appendChild(document.createTextNode(""" & tmpReqName & """));" & vbCrLf & _
			"document.frmAdmin.selReq.appendChild(ChoiceReq);}" & vbCrLf
		rsReqInst.MoveNext
	Loop
	rsReqInst.Close
	Set rsReqInst = Nothing
	rsInstReq.MoveNext
	strInstReqDept = strInstReqDept & "}"
Loop
rsInstReq.Close
Set rsLangIntr = Nothing
%>
<html>
	<head>
		<title>Language Bank - Administrator</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script type='text/javascript' language='JavaScript'>
		<!--
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
		function ReqInfo(Req)
		{
			if (Req == -1)
			{document.frmAdmin.txtphone.value = "";
			document.frmAdmin.txtReqExt.value = "";
			document.frmAdmin.txtfax.value = "";
			document.frmAdmin.txtemail.value = "";
			document.frmAdmin.txtReqFname.value = "";
			document.frmAdmin.txtReqLname.value = "";}
			<%=strReqI%>
			chkPrim();
		}
		function InstInfo(Inst)
		{
			if (Inst == -1)
			{document.frmAdmin.txtNewInst.value = "";}
			<%=strInstI%>
		}
		function ReqChoice(dept, req)
		{
			 var i;
			for(i=document.frmAdmin.selReq.options.length-1;i>=1;i--)
			{
				if (req != "undefined")
				{
					if (document.frmAdmin.selReq.options[i].value != req)
					{
						document.frmAdmin.selReq.remove(i);
					}
				}
				else
				{
					document.frmAdmin.selReq.remove(i);
				}
			}
			<%=strInstReqDept%>
		}
		function Ucase(xxx)
		{
			var xxx;
			return xxx.toUpperCase();
			
		}
		function DeptChoice(inst, dept)
		{
			var i;
			for(i=document.frmAdmin.selDept.options.length-1;i>=1;i--)
			{
				if (dept != "undefined")
				{
					if (document.frmAdmin.selDept.options[i].value != dept)
					{
						document.frmAdmin.selDept.remove(i);
					}
				}
				else
				{
					document.frmAdmin.selReq.remove(i);
				}
			}
			<%=strInstDept%>
		}
		function DeptInfo(dept)
		{
			if (dept == 0)
			{
				document.frmAdmin.selDept.value =0;
				document.frmAdmin.txtNewDept.value = "";
				document.frmAdmin.txtInstAddr.value = "";
				document.frmAdmin.txtInstCity.value = "";
				document.frmAdmin.txtInstState.value = "";
				document.frmAdmin.txtInstZip.value = "";
				document.frmAdmin.txtInstAddrI.value = "";
				document.frmAdmin.txtBlname.value = "";
				document.frmAdmin.txtBillAddr.value = "";
				document.frmAdmin.txtBillCity.value = "";
				document.frmAdmin.txtBillState.value = "";
				document.frmAdmin.txtBillZip.value = "";
			}
			<%=strDept%>
		}
		function CalendarView(strDate)
		{
			document.frmAdmin.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmAdmin.submit();
		}
		function ReqShowMe()
		{
			if (document.frmAdmin.chkAll.checked == true) 
			{
				for(i=document.frmAdmin.selReq.options.length-1;i>=1;i--)
				{
					document.frmAdmin.selReq.remove(i);
				}
				<%=strReq%>
			}
			else
			{
				ReqChoice(document.frmAdmin.selDept.value);
			}
		}
		function KillMe()
		{
			var ans = window.confirm("DELETE?");
			if (ans)
			{
				document.frmAdmin.action = "adminInstaction.asp?ctrl=2";
				document.frmAdmin.submit();
			}
		}
		function SaveMe()
		{
			document.frmAdmin.action = "adminInstaction.asp?ctrl=1";
			document.frmAdmin.submit();
		}
		function chkPrim()
		{
			if (document.frmAdmin.radioPrim1[0].checked == true)
			{
				document.frmAdmin.txtPRim1.value = "Phone";
			}
			if (document.frmAdmin.radioPrim1[1].checked == true)
			{
				document.frmAdmin.txtPRim1.value = "Fax";
			}
			if (document.frmAdmin.radioPrim1[2].checked == true)
			{
				document.frmAdmin.txtPRim1.value = "E-Mail";
			}
		}
		-->
		</script>
	</head>
	<body onload='ReqInfo(document.frmAdmin.selReq.value);InstInfo(document.frmAdmin.selInst.value);
		 DeptInfo(document.frmAdmin.selDept.value);'>
		<form method='post' name='frmAdmin'>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td height='100px' valign='top' colspan='2'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<!-- #include file="_greetme.asp" -->
				<tr>
					<td align='left' valign='top' rowspan='2'>
						<div id='admin' name="dAdmin" style="position: relative; width:100%;height:450px;OVERFLOW: auto;">
							<table cellSpacing='2' cellPadding='0' border='0' width="80%" style='width:100%;'>

								<tr>
									<td class='header' colspan='10'><nobr>Institution</td>
								</tr>
								<tr>
									<td align='right'>Institution:</td>
									<td width='350px'>
										<select class='seltxt' name='selInst'  style='width:250px;' onfocus='DeptChoice(document.frmAdmin.selInst.value); DeptInfo(document.frmAdmin.selDept.value);' onchange='DeptChoice(document.frmAdmin.selInst.value); DeptInfo(document.frmAdmin.selDept.value); InstInfo(document.frmAdmin.selInst.value);'>
											<option value='0'>&nbsp;</option>
											<%=strInst%>
										</select>
										<input type='checkbox' name='chkDelInst' value='1' onclick=''>
										Delete
										&nbsp;<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='* Leave blank to add new'>
									</td>
								</tr>
								<tr>
									<td align='right'>Institution:</td>
									<td>
										<input size='50' class='main' maxlength='50' name='txtNewInst' value='<%=tmpNewInst%>' onkeyup='bawal(this);'>
										<input type='hidden' name='InstName'>
									</td>
								</tr>
								<tr>
									<td align='right' width='18%'>Department:</td>
									<td>
										<select class='seltxt' name='selDept'  style='width:250px;' onfocus='DeptInfo(document.frmAdmin.selDept.value);ReqChoice(document.frmAdmin.selDept.value); '  onchange='DeptInfo(document.frmAdmin.selDept.value);ReqChoice(document.frmAdmin.selDept.value); '>
											<option value='0'>&nbsp;</option>
											<%=strDept2%>
										</select>
										<input type='checkbox' name='chkDelDept' value='1' onclick=''>
										Delete
										&nbsp;<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='* Leave blank to add new'>
									</td>
								</tr>
								<tr>
									<td align='right'>Department:</td>
									<td>
										<input size='50' class='main' maxlength='50' name='txtNewDept' value='<%=tmpNewDept%>' onkeyup='bawal(this);'>
										<input type='hidden' name='DepartName'>
									</td>
								</tr>
								<tr>
									<td align='right'>Classification:</td>
									<td>
										<select class='seltxt' name='selClass'>
											<option value='1' <%=SocSer%>>Social Services</option>
											<option value='2' <%=Priv%>>Private</option>
											<option value='3' <%=Court%>>Court</option>
											<option value='4' <%=Med%>>Medical</option>
											<option value='5' <%=Legal%>>Legal</option>
										</select>
									</td>
								</tr>
								<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtInstAddrI' value='<%=tmpNewInstAddrI%>' onkeyup='bawal(this);'>
										</td>
									</tr>
								<tr>
								<tr>
									<td align='right'>Appointment Address:</td>
									<td>
										<input class='main' size='50' maxlength='50' name='txtInstAddr' value='<%=tmpNewInstAddr%>' onkeyup='bawal(this);'>
										<br>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
										</td>
								</tr>
								<tr>
									<td align='right'>City:</td>
									<td colspan='5'>
										<input class='main' size='25' maxlength='25' name='txtInstCity' value='<%=tmpNewInstCity%>' onkeyup='bawal(this);'>&nbsp;State:
										<input class='main' size='2' maxlength='2' name='txtInstState' value='<%=tmpNewInstState%>' onkeyup='bawal(this);'>&nbsp;Zip:
										<input class='main' size='10' maxlength='10' name='txtInstZip' value='<%=tmpNewInstZip%>' onkeyup='bawal(this);'>
									</td>
								</tr>
								<td align='right'>Billed To:</td>
									<td>
										<input class='main' size='50' maxlength='50' name='txtBlname' value='<%=tmpBLname%>' onkeyup='bawal(this);'>
									</td>
								</tr>
								<tr>
									<td align='right'>Billing Address:</td>
									<td><input class='main' size='50' maxlength='50' name='txtBillAddr' value='<%=tmpBillInstAddr%>' onkeyup='bawal(this);'></td>
								</tr>
								<tr>
									<td align='right'>City:</td>
									<td colspan='5'>
										<input class='main' size='25' maxlength='25' name='txtBillCity' value='<%=tmpBillInstCity%>' onkeyup='bawal(this);'>&nbsp;State:
										<input class='main' size='2' maxlength='2' name='txtBillState' value='<%=tmpBillInstState%>' onkeyup='bawal(this);'>&nbsp;Zip:
										<input class='main' size='10' maxlength='10' name='txtBillZip' value='<%=tmpBillInstZip%>' onkeyup='bawal(this);'>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
																<tr>
									<td class='header' colspan='10'><nobr>Requesting Person</td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<select class='seltxt' name='selReq'  style='width:250px;' onchange='ReqInfo(document.frmAdmin.selReq.value);'>
											<option value='0'>&nbsp;</option>
											<%=strReq2%>
										</select>
										<input type='checkbox' name='chkAll' onclick='ReqShowMe(); ReqInfo(document.frmAdmin.selReq.value);'>
											Show All
										<input type='checkbox' name='chkDelReq' value='1' onclick=''>
										Delete
										&nbsp;<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='* Leave blank to add new'>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Name:
									</td>
									<td>
										<input class='main' size='20' maxlength='20' name='txtReqLname' value='<%=tmpNewReqLN%>' onkeyup='bawal(this);'>
										,&nbsp;
										<input class='main' size='20' maxlength='20' name='txtReqFname' value='<%=tmpNewReqFN%>' onkeyup='bawal(this);'>
										<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" name='txtformat2' readonly value='last name, first name'>
										<input type='hidden' name='ReqName'>
									</td>
								</tr>
								<tr>
										<td align='right'>Primary:</td>
										<td>
											<input class='main'  name='txtPRim1'  readonly size='6'>
										</td>
									</tr>
								<tr>
									<td align='right'>
										<input type='radio' name='radioPrim1' value='1' <%=selRPPhone%> onclick='chkPrim();'>
										Phone:</td>
									<td>
										
										<input class='main' size='12' maxlength='12' name='txtphone' value='<%=tmpNewReqPhone%>' onkeyup='bawal(this);'>
										&nbsp;Ext:<input class='main' size='12' maxlength='12' name='txtReqExt' value='<%=tmpReqExt%>' onkeyup='bawal(this);'>
									</td>
								</tr>
								<tr>
									<td align='right'>
										<input type='radio' name='radioPrim1' value='2'  <%=selRPFax%> onclick='chkPrim();'>
										Fax:</td>
									<td><input class='main' size='12' maxlength='12' name='txtfax' value='<%=tmpNewReqFax%>' onkeyup='bawal(this);'></td>
								</tr>
								<tr>
									<td align='right'>
									<input type='radio' name='radioPrim1' value='0' <%=selRPEmail%> onclick='chkPrim();'>	
										E-Mail:</td>
									<td><input class='main' size='50' maxlength='50' name='txtemail' value='<%=tmpNewReqeMail%>' onkeyup='bawal(this);'></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								
							</table>
						</div>
					</td>
					<td valign='top' align='left'>
						<br>&nbsp;
						<input class='btn' type='button' style="width:250px;" value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveMe();'>
						<br>&nbsp;
						<input class='btn' type='button' style="width:250px;" value='Delete' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='KillMe();'>
						<br>&nbsp;
					</td>
				</tr>
				<tr>
					<td valign='top' align='left'>
						<% If Session("MSG") <> "" Then %>
							<div name="dErr" style="position: relative; left: 10px; width:250px; height:200px;OVERFLOW: auto;">
								<table border='0' cellspacing='2'>		
									<tr>
										<td><span class='error'><%=Session("MSG")%></span></td>
									</tr>
								</table>
							</div>
						<% End If %>
					</td>
				</tr>
				<tr>
					<td height='50px' valign='bottom' colspan='2'>
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