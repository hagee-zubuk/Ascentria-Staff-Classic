<%
'Paths
BackupStr 		= "C:\WORK\ascentria\Temp\staff\misc\"
RepPath 		= "C:\WORK\ascentria\Temp\staff\csv\"
RepPath2 		= BackupStr
RepCSV 			= "/misc/"
RepCSV2 		= "/misc/"
pdfStr 			= BackupStr
EmailLog 		= "C:\WORK\ascentria\Temp\staff\log\EmailLog.txt"
LoginLog 		= "C:\WORK\ascentria\Temp\staff\log\LoginLog.txt"
AdminLog 		= "C:\WORK\ascentria\Temp\staff\AdminLog.txt"
CalPath 		= "C:\WORK\ascentria\Temp\staff\cal\"
SurveyPath 		= BackupStr
DirectionPath 	= BackupStr
'PUBLIC DEFENDERS
F604AStr 		= "\\sqlsrv1\F604A\" '"\\webserv6\F604A\"
tmpF604AStr 	= BackupStr
'Secondary Insurance
secinsPath 		= BackupStr
'Client import
clientList		= BackupStr & "client.txt"
clientListDone	= BackupStr & "clientdone.txt"
'x12
x12path 		= BackupStr & "x12"
x12pathbackup	= BackupStr & "x12"
'271
f271Str			= BackupStr & "f271\"
tmpf271StrStr 	= BackupStr & "f271\"
'upload path chnage to sqlsrv path when live
uploadpath 		= BackupStr & "Upload\"

' Google
googlemapskey	= "AIzaSyAHcSoJYxk465hDVj1_wMXTAozARDkfFgo"
'DIM 	g_strCONN, g_strDBPath

'g_strCONN = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPath & ";"10.10.1.35  .\SQLEXPRESS
g_strCONNDB = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=langbank;Integrated Security=SSPI;"
'"Provider=SQLOLEDB;Data Source=192.168.111.25\SQLEXPRESS;Initial Catalog=langbank;User ID=testpatrick;Password=zubuk#zubuk;"
'Set g_strCONN = Server.CreateObject("ADODB.Connection")
'g_strCONN.Open g_strCONNDB
g_strCONN = g_strCONNDB

'HIST SQL
g_strCONNDB2 = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=histLB;Integrated Security=SSPI;"
'Set g_strCONNHIST2 = Server.CreateObject("ADODB.Connection")
'g_strCONNHIST2.Open g_strCONNDB2
g_strCONNHIST2 = g_strCONNDB2

'HIST Access
'HistoryDB = "C:\work\LSS-LBIS\db\HistLangBank.mdb"
'g_strCONNHist = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & HistoryDB & ";"
HistoryDB = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=HistLangBank;Integrated Security=SSPI;"'"Provider=SQLOLEDB;Data Source=192.168.111.25\SQLEXPRESS;Initial Catalog=langbank;User ID=testpatrick;Password=zubuk#zubuk;"
'Set g_strCONNHist = Server.CreateObject("ADODB.Connection")
'g_strCONNHist.Open HistoryDB
g_strCONNHist = HistoryDB

'FOR HOSPITALPILOT
'g_strDBPathHP = "C:\work\InterReq\db\interpreter.mdb"
'g_strCONNHP = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPathHP & ";"
g_strDBPathHP = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=interpretersql;Integrated Security=SSPI;"'"Provider=SQLOLEDB;Data Source=192.168.111.25\SQLEXPRESS;Initial Catalog=langbank;User ID=testpatrick;Password=zubuk#zubuk;"
'Set g_strCONNHP = Server.CreateObject("ADODB.Connection")
'g_strCONNHP.Open g_strDBPathHP
g_strCONNHP = g_strDBPathHP

'FOR WIZARD DB
g_strDBPathW = "C:\work\LSS-LBIS\db\LBWizard.mdb"
g_strCONNW = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPathW & ";"


'FOR UPLOAD
g_strCONNDBupload = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=langbankuploads;Integrated Security=SSPI;"
'Set g_strCONNupload = Server.CreateObject("ADODB.Connection")
'g_strCONNupload.Open g_strCONNDBupload
g_strCONNupload = g_strCONNDBupload

%>
<!-- #include file="_zEmail.asp" -->