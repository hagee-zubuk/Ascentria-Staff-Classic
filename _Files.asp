<%
'paths needed
DIM 	g_strCONN, g_strDBPath

'g_strCONN = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPath & ";"10.10.1.35  .\SQLEXPRESS
g_strCONNDB = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=langbank;Integrated Security=SSPI;"
'"Provider=SQLOLEDB;Data Source=192.168.111.25\SQLEXPRESS;Initial Catalog=langbank;User ID=testpatrick;Password=zubuk#zubuk;"
Set g_strCONN = Server.CreateObject("ADODB.Connection")
g_strCONN.Open g_strCONNDB

'HIST SQL
g_strCONNDB2 = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=histLB;Integrated Security=SSPI;"
Set g_strCONNHIST2 = Server.CreateObject("ADODB.Connection")
g_strCONNHIST2.Open g_strCONNDB2

'Paths
RepPath = "C:\WORK\ascentria\staff\misc\"
RepPath2 = "C:\WORK\ascentria\staff\misc\"
RepCSV = "/misc/"
RepCSV2 = "/misc/"
BackupStr = "C:\WORK\ascentria\staff\misc\"
pdfStr = "C:\WORK\ascentria\staff\misc\"
EmailLog = "C:\WORK\ascentria\staff\log\EmailLog.txt"
LoginLog = "C:\WORK\ascentria\staff\log\LoginLog.txt"
AdminLog = "C:\WORK\ascentria\staff\log\AdminLog.txt"
CalPath = "C:\WORK\ascentria\staff\cal\"
SurveyPath = "C:\WORK\ascentria\staff\misc\"
DirectionPath = "C:\WORK\ascentria\staff\misc\"

'HIST Access
'HistoryDB = "C:\work\LSS-LBIS\db\HistLangBank.mdb"
'g_strCONNHist = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & HistoryDB & ";"
HistoryDB = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=HistLangBank;Integrated Security=SSPI;"'"Provider=SQLOLEDB;Data Source=192.168.111.25\SQLEXPRESS;Initial Catalog=langbank;User ID=testpatrick;Password=zubuk#zubuk;"
Set g_strCONNHist = Server.CreateObject("ADODB.Connection")
g_strCONNHist.Open HistoryDB

'FOR HOSPITALPILOT
'g_strDBPathHP = "C:\work\InterReq\db\interpreter.mdb"
'g_strCONNHP = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPathHP & ";"
g_strDBPathHP = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=interpretersql;Integrated Security=SSPI;"'"Provider=SQLOLEDB;Data Source=192.168.111.25\SQLEXPRESS;Initial Catalog=langbank;User ID=testpatrick;Password=zubuk#zubuk;"
Set g_strCONNHP = Server.CreateObject("ADODB.Connection")
g_strCONNHP.Open g_strDBPathHP

'FOR WIZARD DB
g_strDBPathW = "C:\work\LSS-LBIS\db\LBWizard.mdb"
g_strCONNW = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPathW & ";"


'FOR INTERPRETER TRACKING
'g_strCONNDB3 = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=langbankappt;Integrated Security=SSPI;"
'Set g_strCONNIntr = Server.CreateObject("ADODB.Connection")
'g_strCONNIntr.Open g_strCONNDB3


'PUBLIC DEFENDERS
F604AStr = "\\sqlsrv1\F604A\" '"\\webserv6\F604A\"
tmpF604AStr = "C:\WORK\ascentria\staff\misc\"

'Secondary Insurance
secinsPath = "C:\WORK\ascentria\staff\misc\"

'Client import
clientList		= "C:\WORK\ascentria\staff\misc\client.txt"
clientListDone	= "C:\WORK\ascentria\staff\misc\clientdone.txt"

'x12
x12path 		= "C:\WORK\ascentria\staff\misc\x12"
x12pathbackup	= "C:\WORK\ascentria\staff\misc\x12"

'271
f271Str			= "C:\WORK\ascentria\staff\misc\f271\"
tmpf271StrStr 	= "C:\WORK\ascentria\staff\misc\f271\"

'upload path chnage to sqlsrv path when l;ive
uploadpath 		= "C:\WORK\ascentria\staff\misc\Upload\"

'FOR UPLOAD
g_strCONNDBupload = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=langbankuploads;Integrated Security=SSPI;"
Set g_strCONNupload = Server.CreateObject("ADODB.Connection")
g_strCONNupload.Open g_strCONNDBupload

googlemapskey = "AIzaSyAHcSoJYxk465hDVj1_wMXTAozARDkfFgo"
%>