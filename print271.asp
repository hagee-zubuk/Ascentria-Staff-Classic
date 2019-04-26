<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function CleanREF(strRef)
    If strRef <> "" Then
        If Right(strRef, 1) = "," Then CleanREF = Left(strRef, Len(strRef) - 1)
    End If
End Function
Function FixDate(dte)
    If dte = "" Then Exit Function
    Yr = Left(dte, 4)
    d = Mid(dte, 7, 2)
    mth = Mid(dte, 5, 2)
    FixDate = mth & "/" & d & "/" & Yr
End Function
Function ConvertEB(strEB)
    If strEB = "" Or IsNull(strEB) Then
        ConvertEB = "ERROR"
        Exit Function
    End If
    ConvertEB = "YES"
    If strEB = "6" Or strEB = "7" Or strEB = "8" Then ConvertEB = "NO"
End Function
Function ConvertERROR(aaaNum)
		If Z_Czero(aaaNum) = 0 Then Exit Function
    If aaaNum = 71 Then
        ConvertERROR = "Patient Birth does not match that for the patient on the Database."
    ElseIf aaaNum = 72 Then
        ConvertERROR = "Invalid/Missing Subscriber/Insured ID"
    ElseIf aaaNum = 73 Then
        ConvertERROR = "Invalid/Missing Subscriber/Insured Name"
    ElseIf aaaNum = 74 Then
        ConvertERROR = "Invalid/Missing Subscriber/Insured Gender Code"
    ElseIf aaaNum = 75 Then
        ConvertERROR = "Subscriber/Insured not found"
    Else
        ConvertERROR = "Unknown Error: " & aaaNum 
    End If
End Function
server.scripttimeout = 360000
strHead = "<td class='tblgrn'>Last Name</td>" & vbCrlf & _
			"<td class='tblgrn'>First Name</td>" & vbCrlf & _
			"<td class='tblgrn'>DOB</td>" & vbCrlf & _
			"<td class='tblgrn'>Service Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Eligible</td>" & vbCrlf & _
			"<td class='tblgrn'>Manage Care</td>" & vbCrlf & _
			"<td class='tblgrn'>Secondary Insurance</td>" & vbCrlf & _
			"<td class='tblgrn'>Notes</td>" & vbCrlf
csvHead = "Last Name,First Name,DOB,Service Date,Eligible,Secondary Insurance,Notes"
If Request("fname") <> "none" Then
	Dim arrST(), arrSE(), arrNMLN(), arrNMFN(), arrDMG(), arrEB(), arrDTP(), arrREF(), arrAAA(), arrMC(), arrMCP()
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set oFile = fso.OpenTextFile(f271Str & Request("fname"), 1, False) 'read file
	Do Until oFile.AtEndOfStream
    oLine = oFile.ReadLine
    arrLine = Split(oLine, "~")
    maxLine = UBound(arrLine)
	 Loop
	 Set oFile = Nothing
	 Set fso = Nothing
	 If InStr(arrLine(0),"ISA*00*          *00*          *ZZ*026000618      *ZZ*NH100496       *") > 0 Then
	  ctrLine = 0
	  ctr = 0
	  Do Until ctrLine = maxLine
       arrSeg = Split(arrLine(ctrLine), "*")
       If arrSeg(0) = "ST" Then 'TRANSACTION SET
           ReDim Preserve arrST(ctr)
           ReDim Preserve arrAAA(ctr)
           ReDim Preserve arrNMLN(ctr)
           ReDim Preserve arrNMFN(ctr)
           ReDim Preserve arrDMG(ctr)
           ReDim Preserve arrEB(ctr)
           ReDim Preserve arrREF(ctr)
           ReDim Preserve arrREF(ctr)
           ReDim Preserve arrSE(ctr)
           ReDim Preserve arrDTP(ctr)
           ReDim Preserve arrMC(ctr)
           ReDim Preserve arrMCP(ctr)
           arrST(ctr) = ctr
       ElseIf arrSeg(0) = "NM1" Then 'NAME
           If arrSeg(1) = "IL" Then
               
               arrNMLN(ctr) = arrSeg(3)
               arrNMFN(ctr) = arrSeg(4)
               arrMC(ctr) = arrSeg(9)
           ElseIf arrSeg(1) = "Y2" Then
           		 arrMCP(ctr) = arrSeg(3)
           End If
       ElseIf arrSeg(0) = "AAA" Then 'ERROR
           arrAAA(ctr) = arrSeg(3)
       ElseIf arrSeg(0) = "DMG" Then
           
           arrDMG(ctr) = FixDate(arrSeg(2)) 'DOB
       ElseIf arrSeg(0) = "EB" Then
           
           arrEB(ctr) = arrSeg(1)
           'If arrEB(ctr) = "R" Then
               
           'End If
       ElseIf arrSeg(0) = "DTP" Then
           If arrSeg(1) = "291" Then 'service date
               
               arrDTP(ctr) = FixDate(arrSeg(3))
           End If
       ElseIf arrSeg(0) = "REF" Then
           If arrSeg(1) = "6P" Then 'sec. insurance
              
               arrREF(ctr) = arrREF(ctr) & " " & arrSeg(3) & ","
           End If
       ElseIf arrSeg(0) = "SE" Then
           
           arrSE(ctr) = ctr
           ctr = ctr + 1
       End If
       ctrLine = ctrLine + 1
    Loop
	  x = 0
	  Do Until x = UBound(arrST) + 1
	  	kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
  		strBody = strBody & "<tr bgcolor='" & kulay & "' >" & _
				"<td class='tblgrn2'><nobr>" & arrNMLN(x) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & arrNMFN(x) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & arrDMG(x) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & arrDTP(x) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & ConvertEB(arrEB(x)) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & arrMCP(x) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & CleanREF(Trim(arrREF(x))) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & ConvertERROR(arrAAA(x)) & "</td></tr>" & vbCrLf
			csvBody = csvBody & """" & arrNMLN(x) & """,""" & arrNMFN(x) & """,""" & arrDMG(x) & """,""" & arrDTP(x) & """,""" & ConvertEB(arrEB(x)) & """,""" & _
				arrMCP(x) & """,""" & CleanREF(Trim(arrRef(x))) & """,""" & ConvertERROR(arrAAA(x)) & """" & vbCrLf
				
			'save in db where eb = yes if does not exist
			If Not IsMedApp(arrMC(x)) Then
				If (arrEB(x) = "R" And (arrMCP(x) = "New Hampshire Healthy Families" Or arrMCP(x) = "Well Sense Health Plan")) Or ConvertEB(arrEB(x)) = "YES" Then
					Set rsApp = Server.CreateObject("ADODB.RecordSet")
					
					rsApp.Open "SELECT * FROM medapprove_T WHERE medicaid = ''", g_strCONN, 1, 3'"INSERT INTO medapprove_T VALUES ('" & arrMC(x) & "','" & arrNMLN(x) & "','" & arrNMFN(x) & "','" & arrDMG(x) & "','0')", g_strCONN, 1, 3
					rsApp.AddNew
					rsApp("medicaid") = arrMC(x)
					rsApp("lname") = arrNMLN(x)
					rsApp("fname") = arrNMFN(x)
					rsApp("dob") = arrDMG(x)
					'rsApp("gender") = 0
					rsApp.Update
					rsApp.Close
					Set rsApp = Nothing
				End If
			End If
	     x = x + 1
	  Loop
	  RepCSV =  "Eligibility271.csv"
		'CONVERT TO CSV
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
		Prt.WriteLine "LANGUAGE BANK - REPORT"
		Prt.WriteLine csvHEAD
		Prt.WriteLine csvBody
		Prt.Close	
		Set Prt = Nothing
		fso.CopyFile RepPath & RepCSV, BackupStr
		Set fso = Nothing
		
		tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV)
	Else
    strBody = "<tr><td colspan='13' align='center'><i>&lt --- Invalid File --- &gt</i></td></tr>"
 	End If
Else
	strBody = "<tr><td colspan='13' align='center'><i>&lt --- Invalid File --- &gt</i></td></tr>"
End If
%>
<html>
	<head>
		<title>Language Bank - Report Result</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function exportMe()
		{
			document.frmResult.action = "printreport.asp?csv=1"
			document.frmResult.submit();
		}
		function PassMe(xxx)
		{
			window.opener.document.frmReport.hideID.value = xxx;
			window.opener.SubmitAko();
			self.close();
		}
		-->
		</script>
	</head>
	<body>
		<form method='post' name='frmResult'>
				<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>
					<tr>
						<td valign='top'>
							<table bgColor='white' border='0' cellSpacing='0' cellPadding='0' align='center'>
							<tr>
								<td>
									<img src='images/LBISLOGO.jpg' align='center'>
								</td>
							</tr>
							<tr>
								<td align='center'>
									261&nbsp;Sheep&nbsp;Davis&nbsp;Road,&nbsp;Concord,&nbsp;NH&nbsp;03301<br>
									Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
								</td>
							</tr>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top' >
							<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
								<tr bgcolor='#C2AB4B'>
									<td colspan='10' align='center'>
			
											<b><%=strMSG%></b>

									</td>
								</tr>
								<tr>
		
										<%=strHead%>
	
								</tr>
				
									<%=strBody%>
	
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
										<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring%>';">
									</td>
								</tr>
									<td colspan='10' align='center' height='100px' valign='bottom'>
										* If needed, please adjust the page orientation of your printer to landscape to view all columns in a single page   
									</td>
								<tr>
								</tr>
							</table>	
						</td>
					</tr>
				</table>
	
		</form>
	</body>
</html>
