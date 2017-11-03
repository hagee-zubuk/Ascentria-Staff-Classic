<%Language=VBScript%>
<%
Dim mlMail

Set mlMail = CreateObject("CDO.Message")
mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "10.10.1.3"
mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "concord0\support"
mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "zubuk#zubuk"
mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
mlMail.Configuration.Fields.Update
mlMail.To = "hagee@zubuk.com"
'mlMail.Cc = "language.services@thelanguagebank.org"
'mlMail.Bcc = "patrick@zubuk.com"
mlMail.From = "language.services@thelanguagebank.org"
mlMail.Subject = "TEST Cancellation "
strBody = "This is to let you know that appointment on suchandsuch is CANCELED.<br>" & _
	 "If you have any questions please contact the LanguageBank office immediately at 410-6183 or email us at " & _
	 "<a href='mailto:info@thelanguagebank.org'>info@thelanguagebank.org</a>.<br>" & _
	 "E-mail about this cancelation was initiated by someuser.<br><br>" & _
	 "Thanks,<br>" & _
	 "Language Bank"
mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
mlMail.Send
%>. done.