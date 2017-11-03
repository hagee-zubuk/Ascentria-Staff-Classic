<%
strAnn = ""
Set rsAnn = Server.CreateObject("ADODB.RecordSet")
sqlAnn = "SELECT LB FROM Announce_T"
rsAnn.Open SqlAnn, g_strCONN, 3, 1
If Not rsAnn.EOF Then
	strAnn = rsAnn("LB") '"ANNOUNCEMENT: " & rsAnn("LB")
End If
rsAnn.Close
Set rsAnn = Nothing
%>