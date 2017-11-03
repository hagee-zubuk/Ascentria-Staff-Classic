<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
tmpZip = Split(Request("zipus"), "|")
tmpInstZip = tmpZip(0)
tmpIntrZip = tmpZip(1)
const pi = 3.14159265358979323846
Function distance(lat1, lon1, lat2, lon2, unit)
  Dim theta, dist
  theta = lon1 - lon2
  dist = sin(deg2rad(lat1)) * sin(deg2rad(lat2)) + cos(deg2rad(lat1)) * cos(deg2rad(lat2)) * cos(deg2rad(theta))
  dist = acos(dist)
  dist = rad2deg(dist)
  distance = dist * 60 * 1.1515
  Select Case ucase(unit)
    Case "K"
      distance = distance * 1.609344
    Case "N"
      distance = distance * 0.8684
  End Select
End Function 
Function acos(rad)
  If Abs(rad) <> 1 Then
    acos = pi/2 - Atn(rad / Sqr(1 - rad * rad))
  ElseIf rad = -1 Then
    acos = pi
  End If
End function
Function deg2rad(Deg)
	deg2rad = cdbl(Deg * pi / 180)
End Function
Function rad2deg(Rad)
	rad2deg = cdbl(Rad * 180 / pi)
End Function

SET objSoapClient = Server.CreateObject("MSSOAP.SoapClient30")
	  objSoapClient.ClientProperty("ServerHTTPRequest") = True
		  
	  ' needs to be updated with the url of your Web Service WSDL and is
	  ' followed by the Web Service name
	  Call objSoapClient.mssoapinit("http://webservices.imacination.com/distance/Distance.jws?wsdl", "DistanceService", "Distance")
	  
' use the SOAP object to call the Web Method Required  
response.write "<!-- Z1: " & tmpInstZip & "<br>" & "Z2: " & tmpIntrZip & "-->"
ON ERROR RESUME NEXT
tmpdistance = objSoapClient.getDistance(CStr(Trim(tmpInstZip)),  CStr(Trim(tmpIntrZip)))
If err.Number <> 0 Then
	'strCalc = "The object for distance calculation returned an error: <br /><span style=""font-size: 9px;"">" & vbCRLF & err.Description & vbCrLf & "</span>"
	If left(tmpInstZip, 1) = "0" Then tmpInstZip = Mid(tmpInstZip,2)
	Set rsZip = Server.CreateObject("ADODB.RecordSet")
	sqlZip = "SELECT * FROM zip_T WHERE zip = '" & tmpInstZip & "'"
	rsZip.Open sqlZip, g_strCONNZIP, 3, 1
	If Not rsZip.EOF Then
		tmpLat = rsZip("lat")
		tmpLong = rsZip("long")
		tmpCity = rsZip("city")
	End If
	rsZip.Close
	Set rsZip = Nothing
	Set rsZip2 = Server.CreateObject("ADODB.RecordSet")
	If left(tmpIntrZip, 1) = "0" Then tmpIntrZip = Mid(tmpIntrZip,2)
	sqlZip2 = "SELECT * FROM zip_T WHERE zip = '" & tmpIntrZip & "'"
	rsZip2.Open sqlZip2, g_strCONNZIP, 3, 1
	If Not rsZip2.EOF Then
		tmpLat2 = rsZip2("lat")
		tmpLong2 = rsZip2("long")
		tmpCity2 = rsZip2("city")
	End If
	rsZip2.Close
	Set rsZip2 = Nothing
	tmpdistance = distance(tmpLat, tmpLong, tmpLat2, tmpLong2, "M")
	strCalc =  "<i>" & UCase(tmpCity2) & "</i> to <i>" & UCase(tmpCity) & "</i> is:<br><font size='3'>" & Z_formatNumber(tmpdistance, 2) & "</font> Miles."
Else
	zip1city = objSoapClient.getCity(Trim(tmpInstZip))
	zip2city = objSoapClient.getCity(Trim(tmpIntrZip))
	strCalc =  "<i>" & UCase(zip2city) & "</i> to <i>" & UCase(zip1city) & "</i> is:<br><font size='3'>" & Z_formatNumber(tmpdistance, 2) & "</font> Miles."
End If
err.reset
%>
<html>
	<head>
		<title>Language Bank - Find <%=tmpTitle%></title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmZip' action=''>
			<table cellpadding='0' cellspacing='0' border='0' align='left' width='100%'> 
				<tr>
					<td height='25px'>&nbsp;</td>
					<td class='header' colspan='6'>
						<nobr>Zip Code Calculator --&gt&gt
					</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='right'>
						Interpreter's Zip code:
					</td>
					<td>
						<input class='main' size='10' maxlength='10' name='txtIntrZip' readonly value='<%=tmpIntrZip%>'>
					</td>
					<td height='35px'>&nbsp;</td>
						<td align='right'>
						Institution's Zip code:
					</td>
					<td>
						<input class='main' size='10' maxlength='10' name='txtInstZip' readonly value='<%=tmpInstZip%>'>
					</td>
				</tr>
				<tr><td colspan='6'><hr align='center' width='75%'></td></tr>
				<tr>
					<td colspan='6' align='center'>
						<font size='2'><b><%=strCalc%></b></font>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>