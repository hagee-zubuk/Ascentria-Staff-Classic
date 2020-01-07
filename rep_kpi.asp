<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function Z_DispZero(zz)
	Z_DispZero = Z_CLng(zz)
	If ( Z_DispZero <= 0) Then
		Z_DispZero = ""
	End If
End Function

Function Z_MinRate()
	Z_MinRate = 0
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT MinWage FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		Z_MinRate = rsRate("MinWage")
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function

Function Z_InHouseRate()
	Z_InHouseRate = 0
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT inHouse FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONN, 1, 3
	If Not rsRate.EOF Then
		Z_InHouseRate = rsRate("inHouse")
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function

Function Z_GetRequestCount(sqlDt, sqlFilt)
	Z_GetRequestCount = 0
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT([appDate]) AS CTR " & _
			", SUM(CASE WHEN d.[State]='NH' THEN 1 ELSE 0 END) AS NHCtr " & _
			", SUM(CASE WHEN d.[State]='MA' THEN 1 ELSE 0 END) AS MACtr " & _
			", SUM(CASE WHEN (d.[State]<>'MA' AND d.[State]<>'NH') THEN 1 ELSE 0 END) AS O_Ctr " & _
			"FROM [request_T] INNER JOIN [dept_T] AS d ON [request_T].[DeptID]=d.[index] " & _
			"WHERE [request_T].[instID] <> 479 " & sqlDt & sqlFilt
On Error Resume Next				
	rsRef.Open sqlRef, g_strCONN, 1, 3
	Set Z_GetRequestCount = rsRef
	'rsRef.Close
On Error Goto 0	
	'Set rsRef = Nothing
End Function

DIM tmpIntr(), tmpTown(), tmpIntrName(), tmpLang(), tmpClass(), tmpBill(), tmpAhrs(), tmpApp(), tmpInst(), tmpDept(), tmpAmt(), tmpFac(), tmpMonthYr(), tmpCtr(), tmpMonthYr2(), tmpMonthYr3()
DIM tmpMonthYr4(), tmpHrs(), tmpHHrs(), tmpMile(), tmpToll(), arrTS(), arrAuthor(), arrPage(), tmpTrain(), tmpIHTrain(), tmpbhrs(), arrBody(), tmpHrs2(), tmpHrs3(), tmpHrs4() , tmpHrs5(), tmpZip()
DIM tmpHrsHP(), tmpHrsHP2()
server.scripttimeout = 360000

%>
<%
tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
tmpdate = replace(date, "/", "") 
tmpTime = replace(FormatDateTime(time, 3), ":", "")


RepCSV =  "KPI" & tmpdate & ".csv"
strMSG = "KPI report (rev. 2019-12-13)"
Set rsRep = Server.CreateObject("ADODB.RecordSet")
strIDT = ""

If tmpReport(1) <> "" Then strMSG = strMSG & " from " & tmpReport(1)
If tmpReport(2) <> "" Then strMSG = strMSG & " to " & tmpReport(2)

strHead = "<th class=""tblgrn"">Classification</td>" & vbCrlf & _
		"<th class=""tblgrn"">Status</td>" & vbCrlf & _
		"<th class=""tblgrn"" style=""width: 60px;"">Total</td>" & _
		"<th class=""tblgrn"" style=""width: 60px;"">NH</td>" & _
		"<th class=""tblgrn"" style=""width: 60px;"">MA</td>" & _
		"<th class=""tblgrn"" style=""width: 60px;"">Other</td></tr>"
		'O_Ctr'
		' & tmpReport(1) & " - " & tmpReport(2) & "</td>" & vbCrlf
CSVHead = "Classification,Status,Total,NH,MA,Other"
'' & tmpReport(1) & " - " & tmpReport(2)
tmpRef = 0
tmpRefMA = 0
tmpRefNH = 0
tmpRef_O = 0
tmpCan = 0
tmpCanMA = 0
tmpCanNH = 0
tmpCan_O = 0
tmpCanB = 0
tmpCanBMA = 0
tmpCanBNH = 0
tmpCanB_O = 0
tmpMis = 0
tmpMisMA = 0
tmpMisNH = 0
tmpMis_O = 0
tmpMis2 = 0
tmpMis2MA = 0
tmpMis2NH = 0
tmpMis2_O = 0
tmpPen = 0
tmpPenMA = 0
tmpPenNH = 0
tmpPen_O = 0
tmpCom = 0
tmpComMA = 0
tmpComNH = 0
tmpCom_O = 0
tmpEmer = 0
tmpEmerMA = 0
tmpEmerNH = 0
tmpEmer_O = 0

DIM strClasses(4), strSeq(4)
strSeq(0) = " AND Class = 3 "
strClasses(0) = "Court"
strSeq(1) = " AND Class = 5 "
strClasses(1) = "Legal"
strSeq(2) = " AND Class = 4 "
strClasses(2) = "Medical"
strSeq(3) = " AND (Class = 6) "
strClasses(3) = "Mental Health"
strSeq(4) = " AND (Class = 1 OR Class = 2) "
strClasses(4) = "Other"
strBody = ""
CSVBody = ""

' date clause
sqlDT = " " 'AND [request_T].[DeptID]=[dept_T].[index] "
If tmpReport(1) <> "" Then sqlDT = sqlDT & " AND appDate >= '" & tmpReport(1) & "'"
If tmpReport(2) <> "" Then sqlDT = sqlDT & " AND appDate <= '" & tmpReport(2) & "'"
sqlDT = sqlDT & " "
DIM lngReqs
For lngI = 0 To 4
	strBody = strBody & "<tr><td class='tblgrn2'><nobr>" & strClasses(lngI) & "</nobr></td>" & vbCrLf
	CSVBody = CSVBody & strClasses(lngI) & ","
	'REFERRALS
	strBody = strBody & "<td class='tblgrn3'><nobr># of Referrals</td>" & vbCrLf
	CSVBody = CSVBody & "# of Referrals,"
	Set rsReqs = Z_GetRequestCount(sqlDt, strSeq(lngI))
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsReqs("ctr") & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("NHCtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("MACtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("O_Ctr") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody & rsReqs("ctr") & "," & rsReqs("NHCtr") & "," & rsReqs("MACtr") & "," & rsReqs("O_Ctr") & vbCrLf
	tmpRef = tmpRef + rsReqs("ctr")
	tmpRefNH = tmpRefNH + Z_CLng( rsReqs("NHCtr") )
	tmpRefMA = tmpRefMA + Z_CLng( rsReqs("MACtr") )
	tmpRef_O = tmpRef_O + Z_CLng( rsReqs("O_Ctr") )
	rsReqs.Close
	Set rsReqs = Nothing
	'CANCELLED
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments,"
	Set rsReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=3 ")
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsReqs("ctr") & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("NHCtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("MACtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("O_Ctr") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody & rsReqs("ctr") & "," & rsReqs("NHCtr") & "," & rsReqs("MACtr") & "," & rsReqs("O_Ctr") & vbCrLf
	tmpCan = tmpCan + rsReqs("ctr")
	tmpCanNH = tmpCanNH + Z_CLng( rsReqs("NHCtr") )
	tmpCanMA = tmpCanMA + Z_CLng( rsReqs("MACtr") )
	tmpCan_O = tmpCan_O + Z_CLng( rsReqs("O_Ctr") )
	rsReqs.Close
	Set rsReqs = Nothing
	'CANCELLED BILLABLE
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments (Billable)</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Canceled Appointments (Billable),"
	Set rsReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=4 ")
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsReqs("ctr") & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("NHCtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("MACtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("O_Ctr") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody & rsReqs("ctr") & "," & rsReqs("NHCtr") & "," & rsReqs("MACtr") & "," & rsReqs("O_Ctr") & vbCrLf
	tmpCanB = tmpCanB + rsReqs("ctr")
	tmpCanBNH = tmpCanBNH + Z_CLng( rsReqs("NHCtr") )
	tmpCanBMA = tmpCanBMA + Z_CLng( rsReqs("MACtr") )
	tmpCanB_O = tmpCanB_O + Z_CLng( rsReqs("O_Ctr") )
	rsReqs.Close
	Set rsReqs = Nothing
	'MISSED
	strBody = strBody & "<tr><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments Missed by Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments Missed by Interpreters,"
	Set rsReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=2 AND [missed]<>1 ")
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsReqs("ctr") & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("NHCtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("MACtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("O_Ctr") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody & rsReqs("ctr") & "," & rsReqs("NHCtr") & "," & rsReqs("MACtr") & "," & rsReqs("O_Ctr") & vbCrLf
	tmpMis = tmpMis + Z_CLng(rsReqs("ctr"))
	tmpMisNH = tmpMisNH + Z_CLng(rsReqs("NHCtr"))
	tmpMisMA = tmpMisMA + Z_CLng(rsReqs("MACtr"))
	tmpMis_O = tmpMis_O + Z_CLng(rsReqs("O_Ctr"))
	rsReqs.Close
	Set rsReqs = Nothing
	'MISSED 2
	strBody = strBody & "<tr><td class='tblgrn2'>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments LB Unable to Send Interpreter</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Appointments LB Unable to Send Interpreter,"
	Set rsReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=2 AND [missed]=1 ")
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsReqs("ctr") & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("NHCtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("MACtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("O_Ctr") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody & rsReqs("ctr") & "," & rsReqs("NHCtr") & "," & rsReqs("MACtr") & "," & rsReqs("O_Ctr") & vbCrLf
	tmpMis2 = tmpMis2 + rsReqs("ctr")
	tmpMis2NH = tmpMis2NH + Z_CLng( rsReqs("NHCtr") )
	tmpMis2MA = tmpMis2MA + Z_CLng( rsReqs("MACtr") )
	tmpMis2_O = tmpMis2_O + Z_CLng( rsReqs("O_Ctr") )
	rsReqs.Close
	Set rsReqs = Nothing
	'PENDING
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Pending Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Pending Appointments,"
	Set rsReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=0 ")
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsReqs("ctr") & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("NHCtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("MACtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("O_Ctr") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody & rsReqs("ctr") & "," & rsReqs("NHCtr") & "," & rsReqs("MACtr") & "," & rsReqs("O_Ctr") & vbCrLf
	tmpPen = tmpPen + rsReqs("ctr")
	tmpPenNH = tmpPenNH + Z_CLng( rsReqs("NHCtr") )
	tmpPenMA = tmpPenMA + Z_CLng( rsReqs("MACtr") )
	tmpPen_O = tmpPen_O + Z_CLng( rsReqs("O_Ctr") )
	rsReqs.Close
	Set rsReqs = Nothing
	'EMERGENCY
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Emergency Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Emergency Appointments,"
	Set rsReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [Emergency]=1 ")
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsReqs("ctr") & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("NHCtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("MACtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("O_Ctr") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody & rsReqs("ctr") & "," & rsReqs("NHCtr") & "," & rsReqs("MACtr") & "," & rsReqs("O_Ctr") & vbCrLf
	tmpEmer = tmpEmer + rsReqs("ctr")
	tmpEmerNH = tmpEmerNH + Z_CLng( rsReqs("NHCtr") )
	tmpEmerMA = tmpEmerMA + Z_CLng( rsReqs("MACtr") )
	tmpEmer_O = tmpEmer_O + Z_CLng( rsReqs("O_Ctr") )
	rsReqs.Close
	Set rsReqs = Nothing
	'COMLPETED
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Completed Appointments</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Completed Appointments,"
	Set rsReqs = Z_GetRequestCount(sqlDt, strSeq(lngI) & " AND [status]=1 ")
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsReqs("ctr") & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("NHCtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("MACtr") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsReqs("O_Ctr") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody & rsReqs("ctr") & "," & rsReqs("NHCtr") & "," & rsReqs("MACtr") & "," & rsReqs("O_Ctr") & vbCrLf
	tmpCom = tmpCom + rsReqs("ctr")
	tmpComNH = tmpComNH + Z_CLng( rsReqs("NHCtr") )
	tmpComMA = tmpComMA + Z_CLng( rsReqs("MACtr") )
	tmpCom_O = tmpCom_O + Z_CLng( rsReqs("O_Ctr") )
	rsReqs.Close
	Set rsReqs = Nothing
	' NEW ROW! 171204: Facilities Clients requesting appointments'
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Facilities Clients</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Facilities Clients,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT([II]) AS InstCnt, COUNT([NH]) CntNH, COUNT([MA]) AS CntMA, COUNT([OO]) AS Cnt_O FROM ( " & _
			"SELECT COUNT(r1.[InstID]) AS II " & _
			", SUM (CASE WHEN d.[State]='NH' THEN 1 ELSE NULL END) AS NH " & _
			", SUM (CASE WHEN d.[State]='MA' THEN 1 ELSE NULL END) AS MA " & _
			", SUM(CASE WHEN (d.[State]<>'MA' AND d.[State]<>'NH') THEN 1 ELSE NULL END) AS OO " & _
			"FROM [request_T] AS r1 INNER JOIN [dept_T] AS d ON r1.[DeptID]=d.[index] " & _
			"WHERE r1.[instID] <> 479 " & sqlDT & strSeq(lngI) & _
			"GROUP BY r1.[instID] ) AS zz"
	'sqlRef = "SELECT COUNT(DISTINCT([request_T].[InstID])) AS instcnt " & _
	'		"FROM [request_T] INNER JOIN [dept_T] AS d ON [request_T].[DeptID]=d.[index] " & _
	'		" WHERE [request_T].[instID] <> 479 " & _
	'		sqlDT & _
	'		strSeq(lngI)
	rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsRef("instcnt") & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsRef("CntNH") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsRef("CntMA") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsRef("Cnt_O") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody &  rsRef("instcnt") & "," & rsRef("CntNH") & "," & rsRef("CntMA") & "," & rsRef("Cnt_O") & vbCrLf
	rsRef.Close
	Set rsRef = Nothing
	' NEW ROW! 180117: How many distinct interpreters took appointments at this time
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Interpreters Involved</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Interpreters Involved,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT([II]) AS IntrCnt" & _
			", SUM (CASE WHEN [OO]>0 THEN 1 ELSE 0 END) AS Cnt_O" & _
			", SUM (CASE WHEN [NH]>0 THEN 1 ELSE 0 END) AS CntNH" & _
			", SUM (CASE WHEN [MA]>0 THEN 1 ELSE 0 END) AS CntMA FROM (" & _
			"SELECT COUNT(	i.[index]) AS ii" & _
			", SUM (CASE WHEN i.[State]='NH' THEN 1 ELSE 0 END) AS NH" & _
			", SUM (CASE WHEN i.[State]='MA' THEN 1 ELSE 0 END) AS MA " & _
			", SUM(CASE WHEN (i.[State]<>'MA' AND i.[State]<>'NH') THEN 1 ELSE 0 END) AS OO " & _
			"FROM [request_T] AS r INNER JOIN [dept_T] AS d ON r.[DeptID]=d.[index] " & _
			"INNER JOIN [interpreter_T] AS i ON r.[IntrID]=i.[index] " & _
			"WHERE r.[instID] <> 479 AND [IntrID] > 0 " & sqlDT & strSeq(lngI) & _
			"GROUP BY i.[index], i.[State]) AS zz"
	rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsRef("intrcnt") & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsRef("CntNH") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsRef("CntMA") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsRef("Cnt_O") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody &  rsRef("intrcnt") & "," & rsRef("CntNH") & "," & rsRef("CntMA") & "," & rsRef("Cnt_O") & vbCrLf
	rsRef.Close
	Set rsRef = Nothing

	' NEW ROW! 180117: How many distinct languages provided at this time
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Languages Requested</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Languages Requested,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT([II]) AS LangCnt" & _
			", SUM (CASE WHEN [OO]>0 THEN 1 ELSE 0 END) AS Cnt_O" & _
			", SUM (CASE WHEN [NH]>0 THEN 1 ELSE 0 END) AS CntNH" & _
			", SUM (CASE WHEN [MA]>0 THEN 1 ELSE 0 END) AS CntMA FROM ( " & _
			"SELECT COUNT(r1.[LangID]) AS II" & _
			", SUM (CASE WHEN d.[State]='NH' THEN 1 ELSE 0 END) AS NH" & _
			", SUM (CASE WHEN d.[State]='MA' THEN 1 ELSE 0 END) AS MA " & _
			", SUM(CASE WHEN (d.[State]<>'MA' AND d.[State]<>'NH') THEN 1 ELSE 0 END) AS OO " & _
			"FROM [request_T] AS r1 INNER JOIN [dept_T] AS d ON r1.[DeptID]=d.[index] " & _
			"WHERE r1.[instID] <> 479 AND [IntrID] > 0 " & _
			sqlDt & strSeq(lngI) & " GROUP BY r1.[LangID] ) AS zz"
	'Response.Write sqlRef
	rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsRef("langcnt") & "</td>" & vbCrLf
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsRef("CntNH") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsRef("CntMA") ) & "</td>"
	strBody = strBody & "<td class='tblgrn4'>" & Z_DispZero( rsRef("Cnt_O") ) & "</td>"
	strBody = strBody & "</tr>" & vbCrLf
	CSVBody = CSVBody &  rsRef("langcnt") & "," & rsRef("CntNH") & "," & rsRef("CntMA") & "," & rsRef("Cnt_O") & vbCrLf
	rsRef.Close
	Set rsRef = Nothing

	strBody = strBody & "<tr><td colspan=""6"">&nbsp;</td></tr>"
	CSVBody = CSVBody &  vbCrLf
Next

'''''''''''TOTALS'''''''''''''''
strBody = strBody & "<tr><td class='tblgrn2'><nobr>TOTALS</td>" & vbCrLf
CSVBody = CSVBody &  "TOTALS,"
'REFERRALS
strBody = strBody & "<td class='tblgrn3'><nobr># of Referrals</td><td class='tblgrn4 tot'>" & tmpRef & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpRefNH ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpRefMA ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpRef_O ) & "</td>" & _
		"</tr>" & vbCrLf
CSVBody = CSVBody &  "# of Referrals," & tmpRef & "," & tmpRefNH & "," & tmpRefMA & "," & tmpRef_O & vbCrLf
'CANCELLED
strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments</td>" & _
		"<td class='tblgrn4 tot'>" & tmpCan & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpCanNH ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpCanMA ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpCan_O ) & "</td>" & _
		"</tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Canceled Appointments," & tmpCan & "," & tmpCanNH & "," & tmpCanMA & "," & tmpCan_O & vbCrLf
'CANCELLED BILLABLE
strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn3'><nobr># of Canceled Appointments (Billable)</td>" & _
		"<td class='tblgrn4 tot'>" & tmpCanB & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpCanBNH ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpCanBMA ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpCanB_O ) & "</td>" & _
		"</tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Canceled Appointments (Billable)," & tmpCanB & "," & tmpCanBNH & "," & tmpCanBMA & "," & tmpCanB_O & vbCrLf
'MISSED
strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments Missed by Interpreter</td>" & _
		"<td class='tblgrn4 tot'>" & tmpMis & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpMisNH ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpMisMA ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpMis_O ) & "</td>" & _
		"</tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Appointments Missed by Interpreter," & tmpMis & "," & tmpMisNH & "," & tmpMisMA & "," & tmpMis_O & vbCrLf
'MISSED 2
strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn3'><nobr># of Appointments LB Unable to Send Interpreter</td>" & _
		"<td class='tblgrn4 tot'>" & tmpMis2 & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpMis2NH ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpMis2MA ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpMis2_O ) & "</td>" & _
		"</tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Appointments LB Unable to Send Interpreter," & tmpMis2 & "," & tmpMis2NH & "," & tmpMis2MA  & "," & tmpMis2_O & vbCrLf
'PENDING
strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn3'><nobr># of Pending Appointments</td>" & _
		"<td class='tblgrn4 tot'>" & tmpPen & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpPenNH ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpPenMA ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpPen_O ) & "</td>" & _
		"</tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Pending Appointments," & tmpPen & "," & tmpPenNH & "," & tmpPenMA & "," & tmpPen_O & vbCrLf
'PENDING
strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn3'><nobr># of Emergency Appointments</td>" & _
		"<td class='tblgrn4 tot'>" & tmpEmer & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpEmerNH ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpEmerMA ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpEmer_O ) & "</td>" & _
		"</tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Emergency Appointments," & tmpEmer & "," & tmpEmerNH & "," & tmpEmerMA & "," & tmpEmer_O & vbCrLf
'COMLPETED
strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn3'><nobr># of Completed Appointments</td>" & _
		"<td class='tblgrn4 tot'>" & tmpCom & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpComNH ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpComMA ) & "</td>" & _
		"<td class='tblgrn4'>" & Z_DispZero( tmpCom_O ) & "</td>" & _
		"</tr>" & vbCrLf
CSVBody = CSVBody &  ",# of Completed Appointments," & tmpCom & "," & tmpComNH & "," & tmpComMA & "," & tmpCom_O & vbCrLf
' NEW ROW! 171204: Facilities Clients requesting appointments'
strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Facilities Clients</td>" & vbCrLf
CSVBody = CSVBody &  ",# of Facilities Clients,"
Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT([II]) AS InstCnt, COUNT([NH]) AS CntNH, COUNT([MA]) AS CntMA, COUNT([O]) AS Cnt_O FROM ( " & _
			"SELECT COUNT(r1.[InstID]) AS II " & _
			", SUM (CASE WHEN d.[State]='NH' THEN 1 ELSE NULL END) AS NH " & _
			", SUM (CASE WHEN d.[State]='MA' THEN 1 ELSE NULL END) AS MA " & _
			", SUM(CASE WHEN (d.[State]<>'MA' AND d.[State]<>'NH') THEN 1 ELSE NULL END) AS O " & _
			"FROM [request_T] AS r1 INNER JOIN [dept_T] AS d ON r1.[DeptID]=d.[index] " & _
			"WHERE r1.[instID] <> 479 " & sqlDT & "GROUP BY r1.[instID] ) AS zz"
rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsRef("InstCnt") & "</td>" & _ 
			"<td class='tblgrn4'>" & Z_DispZero( rsRef("CntNH") ) & "</td>" & _ 
			"<td class='tblgrn4'>" & Z_DispZero( rsRef("CntMA") ) & "</td>" & _ 
			"<td class='tblgrn4'>" & Z_DispZero( rsRef("Cnt_O") ) & "</td>" & _ 
			"</tr>" & vbCrLf
CSVBody = CSVBody &  rsRef("instcnt") & "," & rsRef("CntNH") & "," & rsRef("CntMA") & "," & rsRef("Cnt_O") & vbCrLf
tmpCom = tmpCom + rsRef("instcnt")
rsRef.Close
Set rsRef = Nothing
' NEW ROW! 180117: How many distinct interpreters took appointments at this time
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Interpreters Involved</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Interpreters Involved,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT([II]) AS IntrCnt" & _
			", SUM (CASE WHEN [NH]>0 THEN 1 ELSE 0 END) AS CntNH" & _
			", SUM (CASE WHEN [MA]>0 THEN 1 ELSE 0 END) AS CntMA" & _
			", SUM (CASE WHEN [OO]>0 THEN 1 ELSE 0 END) As Cnt_O  FROM (" & _
			"SELECT COUNT(	i.[index]) AS ii" & _
			", SUM (CASE WHEN i.[State]='NH' THEN 1 ELSE 0 END) AS NH" & _
			", SUM (CASE WHEN i.[State]='MA' THEN 1 ELSE 0 END) AS MA " & _
			", SUM (CASE WHEN (i.[State]<>'MA' AND i.[State]<>'NH') THEN 1 ELSE 0 END) AS OO " & _
			"FROM [request_T] AS r INNER JOIN [dept_T] AS d ON r.[DeptID]=d.[index] " & _
			"INNER JOIN [interpreter_T] AS i ON r.[IntrID]=i.[index] " & _
			"WHERE r.[instID] <> 479 AND [IntrID] > 0 " & sqlDT & "GROUP BY i.[index], i.[State]) AS zz"
	rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsRef("intrcnt") & "</td>" & _ 
			"<td class='tblgrn4'>" & Z_DispZero( rsRef("CntNH") ) & "</td>" & _ 
			"<td class='tblgrn4'>" & Z_DispZero( rsRef("CntMA") ) & "</td>" & _ 
			"<td class='tblgrn4'>" & Z_DispZero( rsRef("Cnt_O") ) & "</td>" & _ 
			"</tr>" & vbCrLf
	CSVBody = CSVBody &  rsRef("IntrCnt") & "," & rsRef("CntNH") & "," & rsRef("CntMA") & "," & rsRef("Cnt_O") & vbCrLf
	rsRef.Close
	Set rsRef = Nothing
' NEW ROW! 180117: How many distinct languages provided at this time
	strBody = strBody & "<tr><td class='tblgrn3'>&nbsp;</td><td class='tblgrn3'><nobr># of Languages Requested</td>" & vbCrLf
	CSVBody = CSVBody &  ",# of Languages Requested,"
	Set rsRef = Server.CreateObject("ADODB.RecordSet")
	sqlRef = "SELECT COUNT([II]) AS LangCnt, SUM (CASE WHEN [NH]>0 THEN 1 ELSE 0 END) AS CntNH" & _
			", SUM (CASE WHEN [MA]>0 THEN 1 ELSE 0 END) AS CntMA" & _
			", SUM (CASE WHEN [OO]>0 THEN 1 ELSE 0 END) AS Cnt_O FROM ( " & _
			"SELECT COUNT(r1.[LangID]) AS II" & _
			", SUM (CASE WHEN d.[State]='NH' THEN 1 ELSE 0 END) AS NH" & _
			", SUM (CASE WHEN d.[State]='MA' THEN 1 ELSE 0 END) AS MA " & _
			", SUM(CASE WHEN (d.[State]<>'MA' AND d.[State]<>'NH') THEN 1 ELSE 0 END) AS OO " & _
			"FROM [request_T] AS r1 INNER JOIN [dept_T] AS d ON r1.[DeptID]=d.[index] " & _
			"WHERE r1.[instID] <> 479 AND [IntrID] > 0 " & _
			sqlDt & " GROUP BY r1.[LangID] ) AS zz"
	rsRef.Open sqlRef, g_strCONN, 1, 3
	strBody = strBody & "<td class='tblgrn4 tot'>" & rsRef("langcnt") & "</td>" & _
			"<td class='tblgrn4'>" & Z_DispZero( rsRef("CntNH") ) & "</td>" & _
			"<td class='tblgrn4'>" & Z_DispZero( rsRef("CntMA") ) & "</td>" & _
			"<td class='tblgrn4'>" & Z_DispZero( rsRef("Cnt_O") ) & "</td>" & _
			"</tr>" & vbCrLf
	CSVBody = CSVBody &  rsRef("langcnt") & "," & rsRef("CntNH") & "," & rsRef("CntMA") & "," & rsRef("Cnt_O") & vbCrLf
	rsRef.Close
	Set rsRef = Nothing
CSVBody = CSVBody &  vbCrLf
If Request("csv") <> 1 Then
	'CONVERT TO CSV
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
	Prt.WriteLine "LANGUAGE BANK - REPORT"
	Prt.WriteLine strMSG
	Prt.WriteLine CSVHead
	Prt.WriteLine CSVBody
	Prt.Close	
	Set Prt = Nothing
	
	'COPY FILE TO BACKUP
	fso.CopyFile RepPath & RepCSV, BackupStr
	
	Set fso = Nothing
	'EXPORT CSV
	'If Request("bill") <> 1 Then
	tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV) & "&NF=" & Z_DoEncrypt("KPI_Report.csv")
End If
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
	<title>Language Bank - Report Result</title>
	<link href='style.css' type='text/css' rel='stylesheet'>
	<script language='JavaScript'>
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
	</script>
	<style>
td.tot { 
	background-color: #FAEEA7;
	font-weight: bolder;
	font-size: 110%;
}
td.tblgrn2 { font-weight: bold; font-size: 110%; }
th.tblgrn {
	background-color: white;
	border: 2px solid #01a1af;
	font-size: 9px;
	text-align: center;
	padding: 1px 0px 2px 0px;
}
tr:nth-child(even) {background-color: #efe;}
tr:nth-child(odd) {background-color: transparent;}
	</style>
</head>
<body>
<div style="width: 300px; text-align: center; margin: 0px auto 20px;">
	<img src='images/LBISLOGO.jpg' align='center' style="width: 287px; height: 67px;" />
	340 Granite Street 3<sup>rd</sup> Floor, Manchester, NH 03102<br />
	Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
</div>
		<form method='post' name='frmResult'>
			<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>
				<tr><td valign='top' >
							<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
								<tr>
									<td colspan="6" align="center" bgcolor='#f58426'>
<b><%=strMSG%></b>
									</td>
								</tr>
<tr><%=strHead%></tr>
<%=strBody%>
								<tr><td colspan="6">&nbsp;</td></tr>
								<tr><td colspan="6" align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
										<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'"
												onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring%>';">
										<br /><br />
										* If needed, please adjust the page orientation of your printer to landscape to view all columns in a single page   
									</td>
								</tr>
							</table>	
						</td>
					</tr>
				</table>
		</form>
	</body>
</html>
