<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'이미지 등록
Public Function fnkakaostoreItemImageReg(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message
	Set obj = jsObject()
		obj("action") = iaction
		obj("adminId") = session("ssBctID")
		obj("itemId") = iitemid
		obj("mallId") = CMALLNAME
		obj("user") = "admin"
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr/apis/product/create/image", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/create/image", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				On Error Resume Next
					resultMessage	= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = replace(strObj.message, "'", "")
					End If
				On Error Goto 0
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = replace(strObj.message, "'", "")
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

'상품 등록
Public Function fnkakaostoreItemReg(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message
	Set obj = jsObject()
		obj("action") = iaction
		obj("adminId") = session("ssBctID")
		obj("itemId") = iitemid
		obj("mallId") = CMALLNAME
		obj("user") = "admin"
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr/apis/product/create", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/create", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				On Error Resume Next
					resultMessage	= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = replace(strObj.message, "'", "")
					End If
				On Error Goto 0
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = replace(strObj.message, "'", "")
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

'상품 수정
Public Function fnkakaostoreItemEdit(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message, strSQL
	Set obj = jsObject()
		obj("action") = iaction
		obj("adminId") = session("ssBctID")
		obj("itemId") = iitemid
		obj("mallId") = CMALLNAME
		obj("user") = "admin"
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr/apis/product/update", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
				strSQL = ""
				strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_outmall_API_Que " & VBCRLF
				strSQL = strSQL & " SET readdate=getdate() " & VBCRLF
				strSQL = strSQL & " ,findate=getdate() " & VBCRLF
				strSQL = strSQL & " ,resultCode='DUPP' " & VBCRLF
				strSQL = strSQL & " ,lastErrMsg='' " & VBCRLF
				strSQL = strSQL & " WHERE mallid = 'kakaostore' " & VBCRLF
				strSQL = strSQL & " and itemid = '"&iitemid&"' " & VBCRLF
				strSQL = strSQL & " and apiAction in ('EDIT', 'PRICE', 'SOLDOUT', 'EDITBATCH') " & VBCRLF
				strSQL = strSQL & " and readdate is null " & VBCRLF
				strSQL = strSQL & " and lastUserid = 'system' "
				dbget.Execute strSQL
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
				On Error Resume Next
					resultMessage			= strObj.result.resultMessage
					If Err.number <> 0 Then
						resultMessage = replace(strObj.message, "'", "")
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

'상품 상태 수정
Public Function fnkakaostoreSellYN(iitemid, iaction, ichgSellYn, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message
	Set obj = jsObject()
		obj("action") = iaction
		obj("adminId") = session("ssBctID")
		obj("itemId") = iitemid
		obj("mallId") = CMALLNAME
		obj("status") = ichgSellYn
		obj("user") = "admin"
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr/apis/product/update/status", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/status", false
		End If
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(istrParam)
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
						resultMessage = replace(strObj.message, "'", "")
					End If
				On Error Goto 0
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Public Function fnkakaostoreChkstat(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message
	Set obj = jsObject()
		obj("action") = iaction
		obj("adminId") = session("ssBctID")
		obj("itemId") = iitemid
		obj("mallId") = CMALLNAME
		obj("user") = "admin"
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr/apis/product/view", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/view", false
		End If
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage	= strObj.result.resultMessage
			ElseIf objXML.Status = "504" Then
				resultMessage = iRbody
			Else
				status			= strObj.status
				code			= strObj.code
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Public Function fnkakaostoreCategory(iaction)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message, strSql
	Set obj = jsObject()
		obj("action") = iaction
		obj("mallId") = CMALLNAME
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr/apis/product/code", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/code", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			status			= strObj.status
			code			= strObj.code
			If status = "200" Then
				strSql = ""
				strSql = strSql & " EXEC [db_etcmall].[dbo].[usp_Ten_API_kakaostore_Category_Make] "
				dbget.Execute strSql
				rw "OK"
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

''이하는 주문 관련
Public Function fnKakaoStoreSongjangUpload(outmallorderserial, orgDetailKey, parcelCode, invoiceNumber, byRef resultMessage, byRef failCount, orgOutmallorderserial)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam
	getParam = ""
	getParam = getParam & "detailKey=" & orgDetailKey
	getParam = getParam & "&invoiceNumber=" & invoiceNumber
	getParam = getParam & "&mallId=" & CMALLNAME 
	getParam = getParam & "&orgOutmallOrderSerial=" & orgOutmallorderserial
	getParam = getParam & "&outmallOrderSerial=" & outmallOrderSerial
	getParam = getParam & "&parcelCode=" & parcelCode

	'http://localhost:11117/admin/etc/kakaostore/kakaostore_SongjangProc.asp?mallId=kakaostore&ord_no=1708467089&ord_dtl_sn=1986019494&hdc_cd=33&inv_no=2323
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/order/invoice?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/order/invoice?" & getParam, false
		End If
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultMessage		= strObj.result.message
				failCount			= strObj.result.failCount
			End If
		Set strObj = nothing

		If (session("ssBctID")="kjy8517") Then
			rw "REQ : <textarea cols=40 rows=10>"&getParam&"</textarea>"
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set objXML= nothing
End Function

Public Function fnkakaostoreSongjangUploadByManager(isellsite, outmallorderserial, originDetailKey, sendState)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, result

	Set obj = jsObject()
		obj("originDetailKey") = originDetailKey
		obj("mallId") = isellsite
		obj("outMallOrderSerial") = outmallorderserial
		obj("sendState") = sendState
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://temporary-dev.10x10.co.kr:80/internal/temporder/invoice/manager", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/temporary/internal/temporder/invoice/manager", false
		End If
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				fnkakaostoreSongjangUploadByManager = strObj.result
			End If
		Set strObj = nothing

		If (session("ssBctID")="kjy8517") Then
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set objXML= nothing
End Function

Public Function getOutmallRefOrgOrderNO(iorderno, iorderdtlsn, isellsite)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, result

	Set obj = jsObject()
		obj("originDetailKey") = iorderdtlsn
		obj("mallId") = isellsite
		obj("outMallOrderSerial") = iorderno
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://temporary-dev.10x10.co.kr:80/internal/temporder/orgordernum", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/temporary/internal/temporder/orgordernum", false
		End If
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				result			= strObj.result
				getOutmallRefOrgOrderNO = result
			End If
		Set strObj = nothing

		If (session("ssBctID")="kjy8517") Then
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set objXML= nothing
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
%>