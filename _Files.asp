<%
'paths needed
DIM 	g_strCONN, g_strDBPath

'g_strCONN = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPath & ";"10.10.1.35  .\SQLEXPRESS
g_strCONNDB = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=langbank;Integrated Security=SSPI;"'"Provider=SQLOLEDB;Data Source=192.168.111.25\SQLEXPRESS;Initial Catalog=langbank;User ID=testpatrick;Password=zubuk#zubuk;"
Set g_strCONN = Server.CreateObject("ADODB.Connection")
g_strCONN.Open g_strCONNDB

'HIST SQL
g_strCONNDB2 = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=histLB;Integrated Security=SSPI;"
Set g_strCONNHIST2 = Server.CreateObject("ADODB.Connection")
g_strCONNHIST2.Open g_strCONNDB2

'Paths
RepPath = "C:\work\LSS-LBIS\web\CSV\"
RepPath2 = "C:\work\LSS-LBIS\web\CSV\"
RepCSV = "/CSV/"
RepCSV2 = "/CSV/"
BackupStr = "C:\work\LSS-LBIS\CSV\"
pdfStr = "C:\work\LSS-LBIS\PDF\"
EmailLog = "c:\work\lss-lbis\log\EmailLog.txt"
LoginLog = "c:\work\lss-lbis\log\LoginLog.txt"
AdminLog = "c:\work\lss-lbis\log\AdminLog.txt"
CalPath = "C:\work\LSS-LBIS\Cal\"
SurveyPath = "C:\work\LSS-LBIS\DHHSsurvey\"
DirectionPath = "C:\work\LSS-LBIS\misc\"

'HIST Access
'HistoryDB = "C:\work\LSS-LBIS\db\HistLangBank.mdb"
'g_strCONNHist = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & HistoryDB & ";"
HistoryDB = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=HistLangBank;Integrated Security=SSPI;"'"Provider=SQLOLEDB;Data Source=192.168.111.25\SQLEXPRESS;Initial Catalog=langbank;User ID=testpatrick;Password=zubuk#zubuk;"
Set g_strCONNHist = Server.CreateObject("ADODB.Connection")
g_strCONNHist.Open HistoryDB

'FOR HOSPITALPILOT
'g_strDBPathHP = "C:\work\InterReq\db\interpreter.mdb"
'g_strCONNHP = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPathHP & ";"
g_strDBPathHP = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=interpretersql;Integrated Security=SSPI;"'"Provider=SQLOLEDB;Data Source=192.168.111.25\SQLEXPRESS;Initial Catalog=langbank;User ID=testpatrick;Password=zubuk#zubuk;"
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
tmpF604AStr = "C:\work\LSS-LBIS\web\PDF\"

'Secondary Insurance
secinsPath = "C:\work\LSS-LBIS\insurance\"

'Client import
clientList = "C:\work\LSS-LBIS\client\client.txt"
clientListDone = "C:\work\LSS-LBIS\client\clientdone.txt"

'x12
x12path = "C:\work\LSS-LBIS\web\x12\"
x12pathbackup = "C:\work\LSS-LBIS\x12\"

'271
f271Str = "C:\work\LSS-LBIS\271\"
tmpf271StrStr = "C:\work\LSS-LBIS\web\271\"

'upload path chnage to sqlsrv path when l;ive
uploadpath = "\\10.10.16.35\Interpreter_Upload\"

'FOR UPLOAD
g_strCONNDBupload = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=langbankuploads;Integrated Security=SSPI;"
Set g_strCONNupload = Server.CreateObject("ADODB.Connection")
g_strCONNupload.Open g_strCONNDBupload

googlemapskey = "AIzaSyAHcSoJYxk465hDVj1_wMXTAozARDkfFgo"
%>
<!-- #include file="_zEmail.asp" -->