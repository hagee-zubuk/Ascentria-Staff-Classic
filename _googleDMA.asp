<%
g_strDMAKey = "AIzaSyAchmuBCbxyH2PzSOXhhiiNNI_7RV3fOQw"	' Distance Matrix API key: 
														' https://console.cloud.google.com/google/maps-apis/apis/distance-matrix-backend.googleapis.com/credentials?project=ascentria-adhoc&duration=PT1H
' **************************************************************************** 
g_strURL = "https://maps.googleapis.com/maps/api/distancematrix/json?units=imperial&key=" & g_strDMAKey

Function Z_GetDMAData(ori, dst)
	Dim oXMLHTTP
	Dim strStatusTest, strURL
	ori = Replace(ori, "#", "")
	dst = Replace(dst, "#", "")
	Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	strURL = g_strURL & "&origins=" & ori & "&destinations=" & dst
	oXMLHTTP.Open "GET", strURL, False
	oXMLHTTP.Send

	If oXMLHTTP.Status = 200 Then
		Z_GetDMAData = oXMLHTTP.responseText

	Else
		Z_GetDMAData = ""
	End If
End Function

' **************************************************************************** 
' the JSON class!
Class acaDistanceMatrix
	Public DBCONN
	Public fltRealTT, fltRealM, fltActTT, fltActMil
	Public RawJSON, Status

	Public Sub FetchMileageV2(strReq, strItrAdr, strItrZip, blnForce)
		' *** TAKE ADDRESS INFORMATION FROM APPOINTMENT REQUEST RECORD'
		Set rsAddr = Server.CreateObject("ADODB.RecordSet")
		strSQL = "SELECT req.[index]" & _
				", CASE " & _
				"WHEN req.[CliAdd] = 1 THEN " & _
				"req.[CAddress] + ', ' + req.[CCity] + ', ' + req.[CState] + ' ' + req.[CZip] " & _
				"ELSE " & _
				"dep.[Address] + ', ' + dep.[City] + ', ' + dep.[State] + ' ' + dep.[Zip] " & _
				"END AS [dest_Address] " & _
				", CASE " & _
				"WHEN req.[CliAdd] = 1 THEN " & _
				"req.[CZip] " & _
				"ELSE " & _
				"dep.[Zip] " & _
				"END AS [dest_ZIP], req.[intrid] AS [intr_id], req.[instid] AS [inst_id]" & _
				", gdt.* " & _
				"FROM [Request_T] AS req " & _
				"INNER JOIN [institution_T] AS ins ON req.[instid] = ins.[index] " & _
				"INNER JOIN [dept_T] AS dep ON req.[deptid] = dep.[index] " & _
				"LEFT JOIN [tmpGoogleDist] AS gdt ON req.[index]=gdt.[reqid] AND req.[intrid]=gdt.[intrid] "  & _
				"WHERE req.[index] = " & strReq
		' Response.Write strSQL
		rsAddr.Open strSQL, DBCONN, 3, 1
		fltRealTT	= 0.0
		fltRealM	= 0.0
		fltActTT	= 0.0
		fltActMil	= 0.0
		RawJSON		= ""
		Status 		= ""
		If Not rsAddr.EOF Then
			' *** INTERROGATE GOOGLE MAPS FOR TRAVEL TIME, AND MILEAGE
			If Z_FixNull(rsAddr("reqid")) = "" Then blnForce = True

			If blnForce Then
				strDstAdr = rsAddr("dest_Address")
				strDstZIP = rsAddr("dest_ZIP")
				strIntrID = rsAddr("intr_id")
			
				Call GetMileage(strReq, strIntrID, strItrAdr, strDstAdr)
			Else
				' get it from the database
				fltRealTT = Round(Z_CDbl(rsAddr("durval")) / 30, 2)
				fltRealM  = Round(Z_CDbl(rsAddr("dstval")) *  2, 2)
				If fltRealM > 40 Then
					fltActMil = Round((fltRealM - 40), 3)
					fltActTT = Round( (fltActMil / (fltRealM/fltRealTT) ), 3)
				End If
			End If
		End If
		rsAddr.Close
		Set rsAddr = Nothing		
	End Sub

	Public Sub FetchMileageFromReqID(strReq, blnForce)
		' *** TAKE ADDRESS INFORMATION FROM APPOINTMENT REQUEST RECORD'
		Set rsAddr = Server.CreateObject("ADODB.RecordSet")
		strSQL = "SELECT req.[index]" & _
				", CASE " & _
				"WHEN req.[CliAdd] = 1 THEN " & _
				"req.[CAddress] + ', ' + req.[CCity] + ', ' + req.[CState] + ' ' + req.[CZip] " & _
				"ELSE " & _
				"dep.[Address] + ', ' + dep.[City] + ', ' + dep.[State] + ' ' + dep.[Zip] " & _
				"END AS [dest_Address] " & _
				", CASE " & _
				"WHEN req.[CliAdd] = 1 THEN " & _
				"req.[CZip] " & _
				"ELSE " & _
				"dep.[Zip] " & _
				"END AS [dest_ZIP] " & _
				", itr.[address1] + ', ' + itr.[City] + ', ' + itr.[State] + ' ' + itr.[Zip Code] AS [orig_Address]" & _
				", itr.[Zip Code] AS [orig_ZIP], req.[intrid] AS [intr_id], req.[instid] AS [inst_id]" & _
				", gdt.* " & _
				"FROM [Request_T] AS req " & _
				"INNER JOIN [institution_T] AS ins ON req.[instid] = ins.[index] " & _
				"INNER JOIN [dept_T] AS dep ON req.[deptid] = dep.[index] " & _
				"INNER JOIN [interpreter_T] AS itr ON req.[intrid]=itr.[index] " & _
				"LEFT JOIN [tmpGoogleDist] AS gdt ON req.[index]=gdt.[reqid] AND req.[intrid]=gdt.[intrid] "  & _
				"WHERE req.[index] = " & strReq
		'Response.Write sqlConfirm
		rsAddr.Open strSQL, DBCONN, 3, 1
		fltRealTT	= 0.0
		fltRealM	= 0.0
		fltActTT	= 0.0
		fltActMil	= 0.0
		RawJSON		= ""
		Status 		= ""
		If Not rsAddr.EOF Then
			' *** INTERROGATE GOOGLE MAPS FOR TRAVEL TIME, AND MILEAGE
			If Z_FixNull(rsAddr("reqid")) = "" Then blnForce = True

			If blnForce Then
				strDstAdr = rsAddr("dest_Address")
				strDstZIP = rsAddr("dest_ZIP")
				strItrAdr = rsAddr("orig_Address")
				strItrZip = rsAddr("orig_ZIP")
				strIntrID = rsAddr("intr_id")
			
				Call GetMileage(strReq, strIntrID, strItrAdr, strDstAdr)
			Else
				' get it from the database
				fltRealTT = Round(Z_CDbl(rsAddr("durval")) / 30, 2)
				fltRealM  = Round(Z_CDbl(rsAddr("dstval")) *  2, 2)
				If fltRealM > 40 Then
					fltActMil = Round((fltRealM - 40), 3)
					If tmpRealTT > 0 And tmpRealM > 0 Then fltActTT = Round( (fltActMil / (tmpRealM/tmpRealTT) ), 3)
				End If
			End If
		End If
		rsAddr.Close
		Set rsAddr = Nothing
	End Sub

	Private Sub GetMileage(strReq, strIntr, strItrAdr, strDstAdr)
		strJSON = Z_GetDMAData(strItrAdr, strDstAdr)
		RawJSON = strJSON
			
		Set oJSON = New aspJSON
		oJSON.loadJSON(strJSON)
		Status = Z_FixNull( oJSON.data("status") )
		
		If Status <> "OK" Then
			strJSON = Z_GetDMAData(strItrZip, strDstZIP)
			Set oJSON = New aspJSON
			oJSON.loadJSON(strJSON)
			Status = Z_FixNull( oJSON.data("status") )
		End If
		
		If Status = "OK" Then
			tmpRealTT = Z_CDbl(oJSON.data("rows")(0)("elements")(0)("duration")("value")) / 3600
			fltRealTT = Round((2 * tmpRealTT), 2) 
			tmpRealM = Z_CDbl(oJSON.data("rows")(0)("elements")(0)("distance")("value")) / 1609.34
			fltRealM = Round((2 * tmpRealM), 2)
			If fltRealM > 40 Then
				fltActMil = Round((fltRealM - 40), 3)
				If tmpRealTT > 0 And tmpRealM > 0 Then
					fltActTT = Round( (fltActMil / (tmpRealM/tmpRealTT) ), 3)
				Else
					fltActTT = 0
				End If
			' Else
			End If

			Set rsGoog = Server.CreateObject("ADODB.RecordSet")
			strSQL = "SELECT * FROM [tmpGoogleDist] WHERE [reqid]=" & strReq & " AND [intrid]=" & strIntr
			rsGoog.Open strSQL, DBCONN, 1, 3
			If rsGoog.EOF Then
				rsGoog.AddNew
				rsGoog("reqid") = strReq
			End If
			rsGoog("intrid")	= strIntr
			rsGoog("distance")	= oJSON.data("rows")(0)("elements")(0)("distance")("text")
			rsGoog("duration")	= oJSON.data("rows")(0)("elements")(0)("duration")("text")
			rsGoog("dstval")	= Z_CDbl(oJSON.data("rows")(0)("elements")(0)("distance")("value")) / 1609.34
			rsGoog("durval")	= Z_CDbl(oJSON.data("rows")(0)("elements")(0)("duration")("value")) / 60
			If (tmpRealTT > 0 ) Then
				rsGoog("rate")	= tmpRealM/tmpRealTT
			Else
				rsGoog("rate")	= 0
			End If
			rsGoog("raw") 		= strItrAdr & " || " & strDstAdr
			rsGoog("fetch") 	= Now
			rsGoog.Update
			rsGoog.Close
			Set rsGoog = Nothing

		End If
	End Sub
End Class


' **************************************************************************** 
' the JSON class!
Class aspJSON
	Public data
	Private p_JSONstring
	private aj_in_string, aj_in_escape, aj_i_tmp, aj_char_tmp, aj_s_tmp, aj_line_tmp, aj_line, aj_lines, aj_currentlevel, aj_currentkey, aj_currentvalue, aj_newlabel, aj_XmlHttp, aj_RegExp, aj_colonfound

	Private Sub Class_Initialize()
		Set data = Collection()

	    Set aj_RegExp = new regexp
	    aj_RegExp.Pattern = "\s{0,}(\S{1}[\s,\S]*\S{1})\s{0,}"
	    aj_RegExp.Global = False
	    aj_RegExp.IgnoreCase = True
	    aj_RegExp.Multiline = True
	End Sub

	Private Sub Class_Terminate()
		Set data = Nothing
	    Set aj_RegExp = Nothing
	End Sub

	Public Sub loadJSON(inputsource)
		inputsource = aj_MultilineTrim(inputsource)
		If Len(inputsource) = 0 Then Err.Raise 1, "loadJSON Error", "No data to load."
		
		select case Left(inputsource, 1)
			case "{", "["
			case else
				Set aj_XmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
				aj_XmlHttp.open "GET", inputsource, False
				aj_XmlHttp.setRequestHeader "Content-Type", "text/json"
				aj_XmlHttp.setRequestHeader "CharSet", "UTF-8"
				aj_XmlHttp.Send
				inputsource = aj_XmlHttp.responseText
				set aj_XmlHttp = Nothing
		end select

		p_JSONstring = CleanUpJSONstring(inputsource)
		aj_lines = Split(p_JSONstring, Chr(13) & Chr(10))

		Dim level(99)
		aj_currentlevel = 1
		Set level(aj_currentlevel) = data
		For Each aj_line In aj_lines
			aj_currentkey = ""
			aj_currentvalue = ""
			If Instr(aj_line, ":") > 0 Then
				aj_in_string = False
				aj_in_escape = False
				aj_colonfound = False
				For aj_i_tmp = 1 To Len(aj_line)
					If aj_in_escape Then
						aj_in_escape = False
					Else
						Select Case Mid(aj_line, aj_i_tmp, 1)
							Case """"
								aj_in_string = Not aj_in_string
							Case ":"
								If Not aj_in_escape And Not aj_in_string Then
									aj_currentkey = Left(aj_line, aj_i_tmp - 1)
									aj_currentvalue = Mid(aj_line, aj_i_tmp + 1)
									aj_colonfound = True
									Exit For
								End If
							Case "\"
								aj_in_escape = True
						End Select
					End If
				Next
				if aj_colonfound then
					aj_currentkey = aj_Strip(aj_JSONDecode(aj_currentkey), """")
					If Not level(aj_currentlevel).exists(aj_currentkey) Then level(aj_currentlevel).Add aj_currentkey, ""
				end if
			End If
			If right(aj_line,1) = "{" Or right(aj_line,1) = "[" Then
				If Len(aj_currentkey) = 0 Then aj_currentkey = level(aj_currentlevel).Count
				Set level(aj_currentlevel).Item(aj_currentkey) = Collection()
				Set level(aj_currentlevel + 1) = level(aj_currentlevel).Item(aj_currentkey)
				aj_currentlevel = aj_currentlevel + 1
				aj_currentkey = ""
			ElseIf right(aj_line,1) = "}" Or right(aj_line,1) = "]" or right(aj_line,2) = "}," Or right(aj_line,2) = "]," Then
				aj_currentlevel = aj_currentlevel - 1
			ElseIf Len(Trim(aj_line)) > 0 Then
				if Len(aj_currentvalue) = 0 Then aj_currentvalue = aj_line
				aj_currentvalue = getJSONValue(aj_currentvalue)

				If Len(aj_currentkey) = 0 Then aj_currentkey = level(aj_currentlevel).Count
				level(aj_currentlevel).Item(aj_currentkey) = aj_currentvalue
			End If
		Next
	End Sub

	Public Function Collection()
		set Collection = CreateObject("Scripting.Dictionary")
	End Function

	Public Function AddToCollection(dictobj)
		if TypeName(dictobj) <> "Dictionary" then Err.Raise 1, "AddToCollection Error", "Not a collection."
		aj_newlabel = dictobj.Count
		dictobj.Add aj_newlabel, Collection()
		set AddToCollection = dictobj.item(aj_newlabel)
	end function

	Private Function CleanUpJSONstring(aj_originalstring)
		aj_originalstring = Replace(aj_originalstring, Chr(13) & Chr(10), "")
		aj_originalstring = Mid(aj_originalstring, 2, Len(aj_originalstring) - 2)
		aj_in_string = False : aj_in_escape = False : aj_s_tmp = ""
		For aj_i_tmp = 1 To Len(aj_originalstring)
			aj_char_tmp = Mid(aj_originalstring, aj_i_tmp, 1)
			If aj_in_escape Then
				aj_in_escape = False
				aj_s_tmp = aj_s_tmp & aj_char_tmp
			Else
				Select Case aj_char_tmp
					Case "\" : aj_s_tmp = aj_s_tmp & aj_char_tmp : aj_in_escape = True
					Case """" : aj_s_tmp = aj_s_tmp & aj_char_tmp : aj_in_string = Not aj_in_string
					Case "{", "["
						aj_s_tmp = aj_s_tmp & aj_char_tmp & aj_InlineIf(aj_in_string, "", Chr(13) & Chr(10))
					Case "}", "]"
						aj_s_tmp = aj_s_tmp & aj_InlineIf(aj_in_string, "", Chr(13) & Chr(10)) & aj_char_tmp
					Case "," : aj_s_tmp = aj_s_tmp & aj_char_tmp & aj_InlineIf(aj_in_string, "", Chr(13) & Chr(10))
					Case Else : aj_s_tmp = aj_s_tmp & aj_char_tmp
				End Select
			End If
		Next

		CleanUpJSONstring = ""
		aj_s_tmp = split(aj_s_tmp, Chr(13) & Chr(10))
		For Each aj_line_tmp In aj_s_tmp
			aj_line_tmp = replace(replace(aj_line_tmp, chr(10), ""), chr(13), "")
			CleanUpJSONstring = CleanUpJSONstring & aj_Trim(aj_line_tmp) & Chr(13) & Chr(10)
		Next
	End Function

	Private Function getJSONValue(ByVal val)
		val = Trim(val)
		If Left(val,1) = ":"  Then val = Mid(val, 2)
		If Right(val,1) = "," Then val = Left(val, Len(val) - 1)
		val = Trim(val)

		Select Case val
			Case "true"  : getJSONValue = True
			Case "false" : getJSONValue = False
			Case "null" : getJSONValue = Null
			Case Else
				If (Instr(val, """") = 0) Then
					If IsNumeric(val) Then
						getJSONValue = CDbl(val)
					Else
						getJSONValue = val
					End If
				Else
					If Left(val,1) = """" Then val = Mid(val, 2)
					If Right(val,1) = """" Then val = Left(val, Len(val) - 1)
					getJSONValue = aj_JSONDecode(Trim(val))
				End If
		End Select
	End Function

	Private JSONoutput_level
	Public Function JSONoutput()
		dim wrap_dicttype, aj_label
		JSONoutput_level = 1
		wrap_dicttype = "[]"
		For Each aj_label In data
			 If Not aj_IsInt(aj_label) Then wrap_dicttype = "{}"
		Next
		JSONoutput = Left(wrap_dicttype, 1) & Chr(13) & Chr(10) & GetDict(data) & Right(wrap_dicttype, 1)
	End Function

	Private Function GetDict(objDict)
		dim aj_item, aj_keyvals, aj_label, aj_dicttype
		For Each aj_item In objDict
			Select Case TypeName(objDict.Item(aj_item))
				Case "Dictionary"
					GetDict = GetDict & Space(JSONoutput_level * 4)
					
					aj_dicttype = "[]"
					For Each aj_label In objDict.Item(aj_item).Keys
						 If Not aj_IsInt(aj_label) Then aj_dicttype = "{}"
					Next
					If aj_IsInt(aj_item) Then
						GetDict = GetDict & (Left(aj_dicttype,1) & Chr(13) & Chr(10))
					Else
						GetDict = GetDict & ("""" & aj_JSONEncode(aj_item) & """" & ": " & Left(aj_dicttype,1) & Chr(13) & Chr(10))
					End If
					JSONoutput_level = JSONoutput_level + 1
					
					aj_keyvals = objDict.Keys
					GetDict = GetDict & (GetSubDict(objDict.Item(aj_item)) & Space(JSONoutput_level * 4) & Right(aj_dicttype,1) & aj_InlineIf(aj_item = aj_keyvals(objDict.Count - 1),"" , ",") & Chr(13) & Chr(10))
				Case Else
					aj_keyvals =  objDict.Keys
					GetDict = GetDict & (Space(JSONoutput_level * 4) & aj_InlineIf(aj_IsInt(aj_item), "", """" & aj_JSONEncode(aj_item) & """: ") & WriteValue(objDict.Item(aj_item)) & aj_InlineIf(aj_item = aj_keyvals(objDict.Count - 1),"" , ",") & Chr(13) & Chr(10))
			End Select
		Next
	End Function

	Private Function aj_IsInt(val)
		aj_IsInt = (TypeName(val) = "Integer" Or TypeName(val) = "Long")
	End Function

	Private Function GetSubDict(objSubDict)
		GetSubDict = GetDict(objSubDict)
		JSONoutput_level= JSONoutput_level -1
	End Function

	Private Function WriteValue(ByVal val)
		Select Case TypeName(val)
			Case "Double", "Integer", "Long": WriteValue = val
			Case "Null"						: WriteValue = "null"
			Case "Boolean"					: WriteValue = aj_InlineIf(val, "true", "false")
			Case Else						: WriteValue = """" & aj_JSONEncode(val) & """"
		End Select
	End Function

	Private Function aj_JSONEncode(ByVal val)
		val = Replace(val, "\", "\\")
		val = Replace(val, """", "\""")
		'val = Replace(val, "/", "\/")
		val = Replace(val, Chr(8), "\b")
		val = Replace(val, Chr(12), "\f")
		val = Replace(val, Chr(10), "\n")
		val = Replace(val, Chr(13), "\r")
		val = Replace(val, Chr(9), "\t")
		aj_JSONEncode = Trim(val)
	End Function

	Private Function aj_JSONDecode(ByVal val)
		val = Replace(val, "\""", """")
		val = Replace(val, "\\", "\")
		val = Replace(val, "\/", "/")
		val = Replace(val, "\b", Chr(8))
		val = Replace(val, "\f", Chr(12))
		val = Replace(val, "\n", Chr(10))
		val = Replace(val, "\r", Chr(13))
		val = Replace(val, "\t", Chr(9))
		aj_JSONDecode = Trim(val)
	End Function

	Private Function aj_InlineIf(condition, returntrue, returnfalse)
		If condition Then aj_InlineIf = returntrue Else aj_InlineIf = returnfalse
	End Function

	Private Function aj_Strip(ByVal val, stripper)
		If Left(val, 1) = stripper Then val = Mid(val, 2)
		If Right(val, 1) = stripper Then val = Left(val, Len(val) - 1)
		aj_Strip = val
	End Function

	Private Function aj_MultilineTrim(TextData)
		aj_MultilineTrim = aj_RegExp.Replace(TextData, "$1")
	End Function

	private function aj_Trim(val)
		aj_Trim = Trim(val)
		Do While Left(aj_Trim, 1) = Chr(9) : aj_Trim = Mid(aj_Trim, 2) : Loop
		Do While Right(aj_Trim, 1) = Chr(9) : aj_Trim = Left(aj_Trim, Len(aj_Trim) - 1) : Loop
		aj_Trim = Trim(aj_Trim)
	end function
End Class
%>