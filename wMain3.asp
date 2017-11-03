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
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
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
tmpTS = Now
tmpDept = 0
tmpReqP = "-1"
tmpHPID = 0
If Request("tmpID") <> "" Then
	Set rsW1 = Server.CreateObject("ADODB.RecordSet")
	sqlW1 = "SELECT InstID, Emergency, EmerFee, InstRate, DeptID, ReqID FROM Wrequest_T WHERE [index] = " & Request("tmpID")
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
	End If
	rsW1.Close
	Set rsW1 = Nothing
End If
'GET TEMP DATA
Set rsWdata = Server.CreateObject("ADODB.RecordSet")
sqlWdata = "SELECT instID, Emergency, DeptID, InstRate FROM Wrequest_T WHERE [index] = " & Request("tmpID")
rsWdata.Open sqlWdata, g_strCONNW, 1, 3
If Not rsWdata.EOF Then
	tmpInst = rsWdata("instID")
	tmpEmer = ""
	If rsWdata("Emergency") = True Then tmpEmer = "(EMERGENCY)" 
	tmpInstRate = Z_FormatNumber(rsWdata("InstRate"), 2)	
	tmpDept = rsWdata("DeptID")
End If
rsWdata.Close
Set rsWdata = Nothing
'GET INSTITUTION
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT Facility FROM institution_T WHERE [index] = " & tmpInst
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpIname = rsInst("Facility") 
End If
rsInst.Close
Set rsInst = Nothing 
'GET DEPARTMENT
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqlDept = "SELECT * FROM dept_T WHERE [index] = " & tmpDept
rsDept.Open sqlDept, g_strCONN, 3, 1
If Not rsDept.EOF Then
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
checkAll = ""
'If Request("show") <> 1 Then
	'GET REQUESTING PERSON INFO
	Set rsReqI = Server.CreateObject("ADODB.RecordSet")
	sqlReqI = "SELECT requester_T.[Index] as myindex, Phone, pExt, Fax, Email, prime, lname, fname, aphone FROM requester_T, reqdept_T WHERE DeptID = " & tmpDept & " AND ReqID = requester_T.[index] ORDER BY Lname, Fname"
	rsReqI.Open sqlReqI, g_strCONN, 3, 1
	Do Until rsReqI.EOF
		strJScript3 = strJScript3 & "if (Req == " & rsReqI("myindex") & ") " & vbCrLf & _
			"{document.frmMain.txtphone.value = """ & rsReqI("Phone") &"""; " & vbCrLf & _
			"document.frmMain.selReq.value = " & rsReqI("myindex") & "; " & vbCrLf & _
			"document.frmMain.txtReqExt.value = """ & rsReqI("pExt") &"""; " & vbCrLf & _
			"document.frmMain.txtfax.value = """ & rsReqI("Fax") &"""; " & vbCrLf & _
			"document.frmMain.txtaphone.value = """ & rsReqI("aPhone") &"""; " & vbCrLf & _
			"document.frmMain.txtemail.value = """ & rsReqI("Email") &"""; " & vbCrLf
			If rsReqI("prime") = 0 Then
				strJScript3 = strJScript3 & "document.frmMain.radioPrim1[2].checked = true;" & vbCrLf 
			ElseIf rsReqI("prime") = 1 Then
				strJScript3 = strJScript3 & "document.frmMain.radioPrim1[0].checked = true;" & vbCrLf 
			ElseIf rsReqI("prime") = 2 Then
				strJScript3 = strJScript3 & "document.frmMain.radioPrim1[1].checked = true;" & vbCrLf 
			End If
			strJScript3 = strJScript3 & "}"
			
		ReqSel = ""
		If tmpReqP = "" Then tmpReqP = -1
		If CInt(tmpReqP) = rsReqI("myindex") Then ReqSel = "selected"
		tmpReqName = CleanMe(rsReqI("lname")) & ", " & CleanMe(rsReqI("fname"))
		strReq2 = strReq2 & "<option " & ReqSel & " value='" & rsReqI("myindex") & "'>" & rsReqI("Lname") & ", " & rsReqI("Fname") & "</option>" & vbCrLf
		rsReqI.MoveNext
	Loop
	rsReqI.Close
	Set rsReqI = Nothing
'Else
	
'End If
'REQUESTING PERSON CHECKER
'Set rsReqCHK = Server.CreateObject("ADODB.RecordSet")
'sqlReqCHK = "SELECT * FROM requester_T"
'rsReqCHK.Open sqlReqCHK, g_strCONN, 3, 1
'Do Until rsReqCHK.EOF
'	strReqCHK = strReqCHK & "if (document.frmMain.txtReqLname.value == """ & Trim(rsReqCHK("lname")) & """ && document.frmMain.txtReqFname.value == """ & Trim(rsReqCHK("Fname")) & """) " & vbCrLf & _
'		"{var ans = window.confirm(""Requester's name already exists. Click on Cancel to rename. Click on OK to continue.""); " & vbCrLf & _
'		"{if (!ans){ " & vbCrLf & _
'		"return; " & vbCrLf & _
'		"} " & vbCrLf & _
'		"} " & vbCrLf & _
'		"} " & vbCrLf
'	rsReqCHK.MoveNext 
'Loop
'rsReqCHK.Close
'Set rsReqCHK = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Language Bank - Interpreter Request Form - Contact Information</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<script language='JavaScript'>
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
		function CalendarView(strDate)
		{
			document.frmMain.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmMain.submit();
		}
		function chkPrim()
		{
			if (document.frmMain.radioPrim1[0].checked == true)
			{
				document.frmMain.txtPRim1.value = "Phone";
			}
			if (document.frmMain.radioPrim1[1].checked == true)
			{
				document.frmMain.txtPRim1.value = "Fax";
			}
			if (document.frmMain.radioPrim1[2].checked == true)
			{
				document.frmMain.txtPRim1.value = "E-Mail";
			}
		}
		function textboxchangeReq() 
		{
			//if (document.frmMain.btnNewReq.value == 'NEW')
			//{
				
			//}
			//else
			//{
				
				document.frmMain.selReq.disabled = false;
				document.frmMain.txtReqLname.value = "";
				document.frmMain.txtReqFname.value = "";
				document.frmMain.txtReqLname.style.visibility = 'hidden';
				document.frmMain.txtReqFname.style.visibility = 'hidden';
				document.frmMain.txtcoma2.style.visibility = 'hidden';
				document.frmMain.txtformat2.style.visibility = 'hidden';
				ReqInfo(document.frmMain.selReq.value);
				document.frmMain.HnewReq.value = 'NEW';
			//}
		}
		function hideNewReq() 
		{
			if (document.frmMain.txtReqLname.value == "" && document.frmMain.txtReqFname.value == "")
			{	
				document.frmMain.txtReqLname.style.visibility = 'hidden';
				document.frmMain.txtReqFname.style.visibility = 'hidden';
				document.frmMain.txtcoma2.style.visibility = 'hidden';
				document.frmMain.txtformat2.style.visibility = 'hidden';
				
				document.frmMain.txtReqLname.value = "";
				document.frmMain.txtReqFname.value = "";
				document.frmMain.HnewReq.value = 'NEW';
			}
			else
			{
				document.frmMain.txtReqLname.style.visibility = 'visible';
				document.frmMain.txtReqFname.style.visibility = 'visible';
				document.frmMain.txtcoma2.style.visibility = 'visible';
				document.frmMain.txtformat2.style.visibility = 'visible';
				Document.frmMain.selReq.disabled = true;
				document.frmMain.txtReqLname.value = '<%=tmpNewReqLN%>';
				document.frmMain.txtReqFname.value = '<%=tmpNewReqFN%>';
				document.frmMain.txtemail.value = '<%=tmpNewReqeMail%>';
				document.frmMain.txtReqExt.value = '<%=tmpReqExt%>';
				document.frmMain.txtphone.value = '<%=tmpNewReqPhone%>';
				document.frmMain.txtfax.value = '<%=tmpNewReqFax%>';
				document.frmMain.HnewReq.value = 'BACK';
			}
		}
		function ReqInfo(Req)
		{
			if (Req == 0)
			{
				if  (document.frmMain.txtReqLname.value == "" || document.frmMain.txtReqFname.value == "")
				{
					hideNewReq();
					document.frmMain.txtphone.value = ""; 
					document.frmMain.txtReqExt.value = ""; 
					document.frmMain.radioPrim1[1].checked = true;
					document.frmMain.txtfax.value = ""; 
					document.frmMain.txtemail.value = ""; 
				}
				else
				{
					document.frmMain.txtphone.value = ""; 
					document.frmMain.txtReqExt.value = ""; 
					document.frmMain.radioPrim1[1].checked = true;
					document.frmMain.txtfax.value = ""; 
					document.frmMain.txtemail.value = ""; 
				}
			}
			<%=strJScript3%>
			chkPrim();
		}
		function ReqChoice(dept, req)
		{
			 var i;
			for(i=document.frmMain.selReq.options.length-1;i>=1;i--)
			{
				if (req != "undefined")
				{
					if (document.frmMain.selReq.options[i].value != req)
					{
						document.frmMain.selReq.remove(i);
					}
				}
				else
				{
					document.frmMain.selReq.remove(i);
				}
			}
			<%=strInstReqDept%>
		}
		
		function WNext()
		{
			if (document.frmMain.selReq.value == 0 && document.frmMain.HnewReq.value == 'NEW')
			{
				alert("ERROR: Requesting Person is Required."); 
				return;
			}
			if (document.frmMain.radioPrim1[2].checked == true && document.frmMain.txtemail.value == "")
			{
				alert("ERROR: Please supply an E-mail address to requesting person."); 
				document.frmMain.txtemail.focus();
				return;
			}
			if (document.frmMain.radioPrim1[0].checked == true && document.frmMain.txtphone.value == "")
			{
				alert("ERROR: Please supply a Phone Number to requesting person."); 
				document.frmMain.txtphone.focus();
				return;
			}
			if (document.frmMain.radioPrim1[1].checked == true && document.frmMain.txtfax.value == "")
			{
				alert("ERROR: Please supply a Fax Number to requesting person."); 
				document.frmMain.txtfax.focus();
				return;
			}
			//CHECK VALID FAX
			if (document.frmMain.radioPrim1[1].checked == true && document.frmMain.txtfax.value != "")
			{
				var tmpFax =  document.frmMain.txtfax.value
				tmpFax = tmpFax.replace("-", "")
				if (tmpFax.length < 10) 
				{
					alert("ERROR: Please include area code in Fax Number to requesting person."); 
					document.frmMain.txtfax.focus();
					return;
				}
			}
			<%=strReqCHK%>
			document.frmMain.action = "waction.asp?ctrl=3";
			document.frmMain.submit();
		}
		function WBack(xxx)
		{
			var ans = window.confirm("Any changes made in this page will not be saved.");
			if (ans){
				document.frmMain.action = "wMain2.asp?tmpID=" + xxx;
				document.frmMain.submit();
			}
		}
		//-->
		</script>
		</head>
		<body onload='ReqInfo(<%=tmpReqP%>); hideNewReq();'>
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
										<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form - 3 / 4</td>
									</tr>
									<tr>
										<td align='center' colspan='10'><nobr>(*) required</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td  align='left'>
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
									<% If Request.Cookies("LBUSERTYPE") <> 4 Then %>
									<tr>
										<td align='right' width='15%'>Rate:</td>
										<td class='confirm'><%=tmpInstRate%></td>
									</tr>
									<% End If %>
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
									<tr>
										<td align='right'>*Requesting Person:</td>
										<td width='200px'>
											<nobr>
											<select id='selReq' class='seltxt' name='selReq'  style='width:250px;' onfocus='JavaScript:ReqInfo(document.frmMain.selReq.value);' onchange='JavaScript:ReqInfo(document.frmMain.selReq.value);'>
												<option value='0'>&nbsp;</option>
												<%=strReq2%>
											</select>
											<!--<input class='btnLnk' type='button' name='btnNewReq' <%=HpLock%> value='NEW' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'"
											onclick='textboxchangeReq();'>
											<input type='checkbox' <%=HpLock%> name='chkAll' <%=checkAll%> onclick='ShowAll(<%=Request("tmpID")%>)'>
											Show All//-->
											<input type='hidden' name='HnewReq'>
											<input type='hidden' name='hideReq' value='<%=tmpReqP%>'>
										</td>
									</tr>
									<tr>
										<td align='right'>&nbsp;</td>
										<td align='left'>
											<input class='main' size='20' maxlength='20' name='txtReqLname' value='<%=tmpNewReqLN%>' onkeyup='bawal(this);'>
											<input class='trans' style='width: 5px;' name='txtcoma2' readonly value=', '>
											<input class='main' size='20' maxlength='20' name='txtReqFname' value='<%=tmpNewReqFN%>' onkeyup='bawal(this);'>
											<input class='transmall' onmouseover="this.className='trans'" onmouseout="this.className='transmall'" size='22' name='txtformat2' readonly value='last name, first name'>
										</td>
									</tr>
									<tr>
										<td align='right'><b>*Contact Numbers:</b></td>
										<td align='left'><b>(any of the following)</b></td>
									</tr>
									<tr>
										<td align='right'>Primary:</td>
										<td>
											<input class='main2'  name='txtPRim1'  readonly size='6'>
										</td>
									</tr>
									<tr>
										<td align='right'>
											<input type='radio' name='radioPrim1' value='1' <%=selRPPhone%> onclick='chkPrim();'>
											Phone:
										</td>
										<td>
											<input class='main' size='12' maxlength='12' name='txtphone' value='<%=tmpNewReqPhone%>' onkeyup='bawal(this);' Readonly>
											&nbsp;Ext:<input class='main' size='12' maxlength='12' name='txtReqExt' value='<%=tmpReqExt%>' onkeyup='bawal(this);' Readonly>
										</td>
										<td align='right'>
											<input type='radio' name='radioPrim1' value='2'  <%=selRPFax%> onclick='chkPrim();'>
											Fax:
										</td>
										<td width='300px'><input class='main' size='12' maxlength='12' name='txtfax' value='<%=tmpNewReqFax%>' onkeyup='bawal(this);' Readonly></td>
									</tr>
									<tr>
										<td align='right'>
											<input type='radio' name='radioPrim1' value='0' <%=selRPEmail%> onclick='chkPrim();'>
											E-Mail:
										</td>
										<td><input class='main' size='50' maxlength='50' name='txtemail' value='<%=tmpNewReqeMail%>' onkeyup='bawal(this);' Readonly></td>
										<td>&nbsp;</td>
										
									</tr>
									<tr>
										<td align='right'>
											
											Alternate Phone:
										</td>
										<td><input class='main' size='12' maxlength='12' name='txtaphone' value='<%=tmpNewReqaphone%>' onkeyup='bawal(this);' Readonly></td>
										<td>&nbsp;</td>
										
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
											<input class='btn' type='button' value='<<' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WBack(<%=Request("tmpID")%>);'>
											<input class='btn' type='Reset' value='Clear' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
											<input class='btn' type='button' value='Cancel' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="window.location='calendarview2.asp'">
											<input class='btn' type='button' value='>>' style='width: 50px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='WNext();'>
											<input type='hidden' name='tmpID' value='<%=Request("tmpID")%>'>
											<input type='hidden' name='tmpInst' value='<%=tmpInst%>'>
											<input type='hidden' name='tmpDep' value='<%=tmpDept%>'>
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
