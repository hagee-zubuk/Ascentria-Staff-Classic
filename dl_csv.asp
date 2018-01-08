<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
strFN = Request("FN") 
'973CEF10E9090A841EA5A5320075BCF708E6278837AD01DDAA05A6D67CAAF4DBE7910FE0D101A27D
fname = Z_DoDecrypt(strFN)
strFN = RepPath & fname

Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Type = 1 'adTypeBinary
objStream.Open
objStream.LoadFromFile(strFN)
Response.ContentType = "text/csv"
Response.Addheader "Content-Disposition", "attachment; filename=""" & fname  & """"

Response.BinaryWrite objStream.Read
objStream.Close
Set objStream = Nothing
%>
