<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilsReport.asp" -->
<%
CSVHEAD = "TIMESTAMP,USER,PAGE,,DATA" & vbCrLf & ",,,Timestamp,Requestor,Date,Start Time,End Time,Location," & _
	"Language,Last Name,First Name,Address,City,State,Zip,Phone,Alter. Phone, Directions,Special Cir," & _
	"DOB,Institution,Department,Doc Num,Courtroom Num,Interpreter,Actual Date,Actual Start Time,Actual End Time," & _
	"Sent to Requester,Sent to Interpreter,Printed,Inst. Rate, Status, Client,Paid,Billable,Verified,Comment,Processed," & _
	"Cancel Reason,Intr. Rate,Emergency,Missed Reason,Use Client Address,Travel Time Inst.,Travel Time Intr.," & _
	"Mileage Inst.,Mileage Intr.,HospitalPilot ID,Processed PR,Client Apartment/Suite,Override Travel Time Inst.," & _
	"Override Travel Time Intr.,Override Mileage Inst.,Override Mileage Intr.,Bill Inst., Travel Time Rate," & _
	"Mileage Rate,Intr. Comment,Bill Comment,Emergency Fee,LangBank Comment,Gender,Minor,Intr. Confirm Time," & _
	"Toll,Total Hrs.,Intr. Confirm Mileage,Actual Travel Time,Actual Mileage,Show Intr.,LangBank Confirm Time," & _
	"LangBank Confirm Mileage,Pay Intr.,Payable Hrs.,Override Pay Hrs.,Override Mileage,Approve Hrs.,Inst. Actual Mileage," & _
	"Inst. Actual Travel Time,Completed,Mileage Process,Late,Late Reason,Outpatient,Has Medicaid,Medicaid,Verified Medicaid," & _
	"Auto Accident,Worker Compensation,Secondary Insurance,Court Amount,Upload File,Approve File,Filename,Processed Medicaid," & _
	"Happen,Medicaid Denied,Changed to Inst.,Change to Inst. Reason, Sys. comment,Training Appt.,Billing Trail," & _
	"Raw Travel Time, Raw Mileage,Meridian,NH Health,WellSense,Acknowlege,Download Vform,MR No.,Assigned by,Block Sched," & _
	"No Reason,No Reason DateStamp,Phone Appt,Judge"

'dispalys history of request
Set rsApp = Server.CreateObject("ADODB.RecordSet")
sqlApp = "SELECT * FROM hist_T WHERE LBID = " & Request("ReqID") & " ORDER BY timestamp"
rsApp.Open sqlApp, g_strCONNHist2, 3, 1
Do Until rsApp.EOF
	CSVBody = CSVBody & """" & rsApp("timestamp") & """,""" & rsApp("author") & """,""" & rsApp("pageused") & """," & vbCrlf & ",,," & rsApp("Hist") & vbCrLf
	rsApp.MoveNext
Loop
rsApp.Close
Set rsApp = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
<%
HistID = BackupStr & "\Dhist" & Request("ReqID") & ".csv"
Set fso = CreateObject("Scripting.FileSystemObject")
Set Prt = fso.CreateTextFile(HistID, True)
Prt.WriteLine "LANGUAGE BANK - DETAILED HISTORY - " & Request("ReqID")
Prt.WriteLine CSVHead
Prt.WriteLine CSVBody
Prt.Close	
Set Prt = Nothing

Set dload = Server.CreateObject("SCUpload.Upload")
	dload.Download HistID
Set dload = Nothing
%>
