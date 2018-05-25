<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function GetIntrAddr(xxx)
	GetIntrAddr = "N/A"
	If xxx < 1 Then Exit Function
	Set rsAdr = Server.CreateObject("ADODB.RecordSet")
	sqlAdr = "SELECT * FROM Interpreter_T WHERE [index] = " & xxx
	rsAdr.Open sqlAdr, g_strCONN, 3, 1
	If Not rsAdr.EOF Then
		GetIntrAddr = rsAdr("address1") & ", " & rsAdr("city") & ", " & rsAdr("state") & ", " & rsAdr("zip code") 
	End If
	rsAdr.Close
	Set rsAdr = Nothing
End Function

Server.ScriptTimeout = 360000

RepCSV =  "AnnReport.csv" 

CSVHead = "Date,Institution,Department,Appointment Address,Interpreter Name,Interpreter Address,Mileage"	

Set rsMile = Server.CreateObject("ADODB.RecordSet")

sqlMile = "SELECT * FROM Request_T " & _
	"WHERE NOT processed IS NULL AND " & _
	"status = 1 AND InstActMil > 0 AND " & _
	"appdate >= '1/1/2012' AND appDate < '7/1/2012' " & _
	"ORDER BY Appdate"
	
rsMile.Open sqlMile, g_strCONN, 3, 1
Do Until rsMile.EOF
	CSVBody = CSVBody & """" & rsMile("appDate") & """,""" & GetInst(rsMile("InstID")) & """,""" & _
		GetDept(rsMile("DeptID")) & ""","""
		
	If rsMile("CliAdd") Then
		CSVBody = CSVBody & "*" & rsMile("caddress") & ", " & rsMile("ccity") & ", " & rsMile("cstate") & ", " & rsMile("czip") & """,""" 
	Else
		CSVBody = CSVBody & GetDeptAdr(rsMile("deptID")) & ""","""
	End If
	
	CSVBody = CSVBody & GetIntr(rsMile("IntrID")) & """,""" & GetIntrAddr(rsMile("intrID")) & """,""" & rsMile("InstActMil") & """" & vbCrLf
	
	rsMile.MoveNext
Loop
rsMile.Close
Set rsMile = Nothing

Set fso = CreateObject("Scripting.FileSystemObject")
Set Prt = fso.CreateTextFile(RepPath &  RepCSV, True)
	Prt.WriteLine CSVHead
	Prt.WriteLine CSVBody
	Prt.Close	
Set Prt = Nothing
Set fso = Nothing

tmpstring = "CSV/" & repCSV
tmpstring = "dl_csv.asp?FN=" & Z_DoEncrypt(repCSV)

%>
<html>
	<head></head>
	<body>
		<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='<%=tmpstring%>';">
	</body>
</html>