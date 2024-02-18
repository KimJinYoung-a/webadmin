<%
Function GetCheckStatus(byVal sellsite, byRef LastCheckDate, byRef isSuccess)
	dim strSql

    strSql = " IF NOT Exists("
    strSql = strSql + " 	select LastcheckDate"
    strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp]"
    strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "'"
	strSql = strSql + " )"
	strSql = strSql + " BEGIN"
	strSql = strSql + "		insert into db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp](sellsite, lastcheckdate, issuccess) "
	strSql = strSql + "		values('" & sellsite & "', '" & Left(DateAdd("d", -1, Now()), 10) & "', 'N') "
	strSql = strSql + " END"
	dbget.Execute strSql

	strSql = " select convert(varchar(10), LastCheckDate, 121) as LastCheckDate, isSuccess from db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' "

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		LastCheckDate = rsget("LastCheckDate")
		isSuccess = rsget("isSuccess")
	rsget.Close
End Function

function SetCheckStatus(sellsite, LastCheckDate, isSuccess)
	dim strSql

	strSql = " update db_temp.[dbo].[tbl_xSite_TMPOrder_timestamp] "
	strSql = strSql + " set lastcheckdate = '" & LastCheckDate & "', issuccess = '" & isSuccess & "' "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' "
	''response.write strSql
	dbget.Execute strSql
end function

Function GetOrderFromExtSite(sellsite, selldate, gubunCode, resultMessage)
	Select Case sellsite
		Case "ezwel"
			Call GetOrderFrom_ezwel(selldate, resultMessage)
		Case "boribori1010"
			Call GetOrderFrom_boribori(selldate, gubunCode, resultMessage)
		Case "wconcept1010"
			Call GetOrderFrom_wconcept(selldate, gubunCode, resultMessage)
		Case "benepia1010"
			Call GetOrderFrom_benepia(selldate, gubunCode, resultMessage)
		Case Else
			response.write "잘못된 접근입니다."
			dbget.close : response.end
	End Select
End Function

function GetCSCheckStatus(byVal sellsite, byVal csGubun, byRef LastCheckDate, byRef isSuccess)
	dim strSql

    strSql = " IF NOT Exists("
    strSql = strSql + " 	select LastcheckDate"
    strSql = strSql + " 	from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp]"
    strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "'"
	strSql = strSql + " )"
	strSql = strSql + " BEGIN"
	strSql = strSql + "		insert into db_temp.[dbo].[tbl_xSite_TMPCS_timestamp](sellsite, csGubun, lastcheckdate, issuccess, LastUpdate) "
	strSql = strSql + "		values('" & sellsite & "', '" & csGubun & "', '" & Left(DateAdd("d", -1, Now()), 10) & "', 'N', getdate()) "
	strSql = strSql + " END"
	dbget.Execute strSql

	strSql = " select convert(varchar(10), LastCheckDate, 121) as LastCheckDate, isSuccess from db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	strSql = strSql + " 	where sellsite='" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "'"

	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		LastCheckDate = rsget("LastCheckDate")
		isSuccess = rsget("isSuccess")
	rsget.Close
end function

function SetCSCheckStatus(sellsite, csGubun, LastCheckDate, isSuccess)
	dim strSql

	strSql = " update db_temp.[dbo].[tbl_xSite_TMPCS_timestamp] "
	strSql = strSql + " set lastcheckdate = '" & LastCheckDate & "', issuccess = '" & isSuccess & "', LastUpdate = getdate() "
	strSql = strSql + " where sellsite = '" + CStr(sellsite) + "' and csGubun='" + CStr(csGubun) + "' "
	''response.write strSql
	dbget.Execute strSql
end function

Function GetOrderFrom_ezwel(selldate, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId=ezwel&sellDate="&selldate&"&code=1001"

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSOrder_ezwel(selldate, csGubun, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	Select Case csGubun
		Case "CANCLEREQ"
			reqCode = "1013"
		Case "CANCLEDONE"
			reqCode = "1005"
		Case "RETURNREQ"
			reqCode = "1007"
		Case "RETURNDONE"
			reqCode = "1008"
		Case "CHANGEREQ"
			reqCode = "1011"
		Case "CHANGEDONE"
			reqCode = "1012"
	End Select

	getParam = ""
	getParam = getParam & "mallId=ezwel&sellDate="&selldate&"&code="&reqCode
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order/cs?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/cs?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetOrderFrom_boribori(selldate, gubunCode, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId=boribori1010&sellDate="&selldate&"&code="&gubunCode

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetOrderFrom_wconcept(selldate, gubunCode, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId=wconcept1010&sellDate="&selldate&"&code=03"

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr:80/external/apis/order?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.result.result.returnMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				resultMessage	= strObj.message
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetOrder_WconceptDetail(outMallOrderSerial, orgDetailKey, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode
	getParam = ""
	getParam = getParam & "sellSite=wconcept1010&outmallOrderSerial="&outMallOrderSerial&"&originDetailKey="&orgDetailKey
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr:80/external/apis/order/detail?"& getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/detail?"& getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				resultMessage	= strObj.message
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
End Function

Function GetOrder_WconceptConfirm(outMallOrderSerial, orgDetailKey, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode
	getParam = ""
	getParam = getParam & "sellSite=wconcept1010&outmallOrderSerial="&outMallOrderSerial&"&originDetailKey="&orgDetailKey
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr:80/external/apis/order/confirm?"& getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/confirm?"& getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
' rw iRbody
' response.end
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				resultMessage	= strObj.message
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
End Function

Function GetOrderFrom_benepia(selldate, gubunCode, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId=benepia1010&sellDate="&selldate&"&code=1&page=1&size=1000"

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr:80/external/apis/order?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				status			= strObj.status
				code			= strObj.code
				resultMessage	= strObj.message
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetOrder_benepiaDetail(outMallOrderSerial, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode
	getParam = ""
	getParam = getParam & "sellSite=benepia1010&outmallOrderSerial="&outMallOrderSerial
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr:80/external/apis/order/detail?"& getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/detail?"& getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
'				resultMessage	= strObj.result.resultMessage
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				resultMessage	= strObj.message
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
End Function

Function GetOrder_benepiaConfirm(outMallOrderSerial, orgDetailKey, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode
	getParam = ""
	getParam = getParam & "sellSite=benepia1010&outmallOrderSerial="&outMallOrderSerial&"&originDetailKey="&orgDetailKey
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr:80/external/apis/order/confirm?"& getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/confirm?"& getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
' rw iRbody
' response.end
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				resultMessage	= strObj.message
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
End Function

Function GetCSOrder_boribori(selldate, csGubun, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	Select Case csGubun
		Case "ordercancel"
			reqCode = "r"
		Case "ordersoldout"
			reqCode = "f"
		Case "return"
			reqCode = "refund"
		Case "exchange"
			reqCode = "exchange"
	End Select

	getParam = ""
	getParam = getParam & "mallId=boribori1010&sellDate="&selldate&"&code="&reqCode
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order/cs?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/cs?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSOrderCancel_benepia(selldate, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	getParam = ""
	getParam = getParam & "mallId=benepia1010&page=1&size=100&sellDate="&selldate
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order/cs/cancel?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/cs/cancel?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSOrderExchange_benepia(selldate, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	getParam = ""
	getParam = getParam & "mallId=benepia1010&page=1&size=100&sellDate="&selldate
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order/cs/return?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/cs/return?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSOrderCancel_ebay(selldate, resultMessage, sellsite)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	getParam = ""
	getParam = getParam & "mallId="&sellsite&"&sellDate="&selldate
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order/cs/cancel?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/cs/cancel?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSOrderReturn_ebay(selldate, resultMessage, sellsite, icode)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	getParam = ""
	getParam = getParam & "mallId="&sellsite&"&sellDate="&selldate&"&code="&icode
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order/cs/return?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/cs/return?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSOrderExchange_ebay(selldate, resultMessage, sellsite, icode)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	getParam = ""
	getParam = getParam & "mallId="&sellsite&"&sellDate="&selldate&"&code="&icode
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order/cs/change?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/cs/change?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function


Function GetCSOrder_wconcept(selldate, csGubun, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	Select Case csGubun
		Case "ordercancelRequest"
			reqCode = "32"
		Case "ordercancelComplete"
			reqCode = "38"
		Case "return"
			reqCode = "22"
		Case "exchange"
			reqCode = "43"
	End Select

	getParam = ""
	getParam = getParam & "mallId=wconcept1010&sellDate="&selldate&"&code="&reqCode
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order/cs?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/cs?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = strObj.message
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function


Function GetOrder_kakaostore(sellsite, selldate, hasMoreData, page, gubunCode)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode
	getParam = ""
	getParam = getParam & "mallId="&sellsite&"&sellDate="&selldate&"&page="&page&"&code="&gubunCode
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr/apis/order?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				If strObj.result.last = true Then
					hasMoreData = "N"
				Else
					hasMoreData = "Y"
				End If
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function setOrder_kakaostoreConfirm(gubunCode)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode, arrRows, strParam

	Dim sqlStr, i
	sqlStr = ""
	sqlStr = sqlStr & " SELECT orderId FROM db_temp.[dbo].[tbl_xSite_TMPOrder_kakaostore] WHERE orderStatus = '"&gubunCode&"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) then
		arrRows = rsget.getRows()
	End If
	rsget.Close

	If IsArray(arrRows) Then
		Set obj = jsObject()
			Set obj("orderIds")= jsArray()									'내부처리완료목록
			For i = 0 To Ubound(arrRows, 2)
				obj("orderIds")(i) = arrRows(0, i)
			Next
			strParam = obj.jsString
		Set obj = nothing

		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			If application("Svr_Info")="Dev" Then
				objXML.open "POST", "http://external-dev.10x10.co.kr/apis/order/ready", false
			Else
				objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/order/ready", false
			End If
			objXML.setRequestHeader "Content-Type", "application/json"
			objXML.setTimeouts 5000,80000,80000,80000
			objXML.Send(strParam)
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				If (session("ssBctID")="kjy8517") Then
					rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
				End If
			Set strObj = nothing
		Set objXML= nothing
	End If
End Function

Function GetOrder_kakaostoreDetail(orderId, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr/apis/order/"& orderId, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/"& orderId, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage = strObj.result.resultMessage
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetInvoiceList(sellsite, returnJsonList)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	getParam = ""
	getParam = getParam & "?mallId="&sellsite
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://temporary-dev.10x10.co.kr:80/internal/temporder/invoice/list"&getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/temporary/internal/temporder/invoice/list"&getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		returnJsonList = iRbody
		If (session("ssBctID")="kjy8517") Then
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set objXML= nothing
End Function

Function GetCSOrder_kakaostore(currdate, csGubun, orderStatus)
	Call deleteKakaoStoreOrder(orderStatus)
	Dim isOrderComplete, page, hasMoreData, sqlStr, arrRows
	isOrderComplete = "N"
	page = 1
	Do Until isOrderComplete = "Y"
		Call GetOrder_kakaostore("kakaostore", currdate, hasMoreData, page, orderStatus)
		If hasMoreData = "N" Then
			isOrderComplete = "Y"
		Else 
			page = page + 1
		End If
		response.flush
	Loop

	sqlStr = ""
	sqlStr = sqlStr & " SELECT orderId "
	sqlStr = sqlStr & " FROM db_temp.[dbo].[tbl_xSite_TMPOrder_kakaostore] "
	sqlStr = sqlStr & " WHERE 1=1 "
	sqlStr = sqlStr & " AND orderStatus = '"& orderStatus &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open sqlStr, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		arrRows = rsget.getRows()
	End If
	rsget.Close

	If IsArray(arrRows) Then
		For i = 0 To Ubound(arrRows, 2)
			Call GetCSOrder_kakaostoreDetail(arrRows(0, i), csGubun, resultMessage)
			rw resultMessage
			If (i mod 10) = 9 Then response.flush
		Next
	End If
End Function

Function deleteKakaoStoreOrder(orderStatus)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://temporary-dev.10x10.co.kr:80/internal/temporder/kakaostore/delete/"&orderStatus, false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/temporary/internal/temporder/kakaostore/delete/"&orderStatus, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		If (session("ssBctID")="kjy8517") Then
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set objXML= nothing
End Function

Function GetCSOrder_kakaostoreDetail(orderId, csGubun, resultMessage)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam, reqCode

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr/apis/order/cs/"&csGubun&"/"& orderId, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/cs/"&csGubun&"/"& orderId, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage = strObj.result.resultMessage
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function
'http://gateway.10x10.co.kr/temporary/internal/temporder/invoice/list?mallId=ezwel
'http://temporary-dev.10x10.co.kr:80/internal/temporder/invoice/list?mallId=ezwel
%>