<%@ Page Language="VB" AutoEventWireup="false" aspcompat="false" debug="true"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">
</script>
<%
    Dim strZIP(), strCity1, strCity2, strInstZIP, strIntrZIP, strTmp As String
    strTmp = ""
    btnSubmit.Visible = False
    If Request.QueryString("zipus") <> "" Then strTmp = Request.QueryString("zipus").Trim
    If strTmp = "" Then
        If Request.Form("txtIntrZip") <> "" Then
            strIntrZIP = Request.Form("txtIntrZip").Trim
            strInstZIP = Request.Form("txtInstZip").Trim
            If strInstZIP <> "" And strIntrZIP <> "" Then strTmp = "HERE"
        End If
    Else
        strZIP = Split(strTmp, "|")
        strInstZIP = strZIP(0)
        strIntrZIP = strZIP(1)
    End If
    If strTmp <> "" Then

        Dim objWs As New wsDistance.DistanceService
        Dim dblDistance As Double
        dblDistance = objWs.getDistance(strInstZIP, strIntrZIP)
    
        strCity1 = objWs.getCity(strIntrZIP)
        strCity2 = objWs.getCity(strInstZIP)

        strCalc.Text = "<i>" & UCase(strCity1) & "</i> to <i>" & UCase(strCity2) & "</i> is:<br /> " & _
                "<font size='3'>" & FormatNumber(dblDistance, 2) & "</font> Miles."
        txtInstZip.Value = strZIP(0)
        txtIntrZip.Value = strZIP(1)
    Else
        
    End If
%>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
		<link href='style.css' type='text/css' rel='stylesheet'>
</head>
<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
    <form id="form1" runat="server">
    <div>
			<table cellpadding='0' cellspacing='0' border='0' align='left' width='100%'> 
				<tr>
					<td height='25px'>&nbsp;</td>
					<td class='header' colspan='6'>
						Zip Code Calculator --&gt;&gt;
					</td>
				</tr>
				<tr>
					<td height='35px'>&nbsp;</td>
					<td align='right'>Interpreter's Zip code:</td>
					    <td><input class='main' size='10' runat="server"
					            maxlength='10' name='txtIntrZip' id='txtIntrZip' readonly="readonly">
					    </td>
					    <td height='35px'>&nbsp;</td>
						<td align='right'>Institution's Zip code:</td>
					    <td><input class='main' size='10' runat="server"
					            maxlength='10' name='txtInstZip' id='txtInstZip' readonly="readonly" >
					    </td>
				    </tr>
                <tr><td>&nbsp;</td><td colspan="5" align="center">
                        <input class="main" id="btnSubmit" name="btnSubmit" type="submit" runat="server" value="Look up"/></td></tr>
				<tr><td colspan='6'><hr align='center' width='75%'></td></tr>
				<tr><td colspan='6' align='center' style="height: 19px">
						<font size='2'><b><asp:Label name="strCalc" ID="strCalc" runat="server"></asp:Label></b></font>
					</td>
				</tr>
			</table>    
    </div>
    </form>
</body>
</html>
