<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 등록
Public Function fnEzwelItemReg(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://external-dev.10x10.co.kr:80/apis/product/create", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/create", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
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

'상품 수정
Public Function fnEzwelItemEdit(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://external-dev.10x10.co.kr:80/apis/product/update", false
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

'상품 상태 수정
Public Function fnEzwelSellYN(iitemid, iaction, ichgSellYn, byRef resultMessage)
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
			objXML.open "POST", "http://external-dev.10x10.co.kr:80/apis/product/update/status", false
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

'상품 조회
Public Function fnEzwelChkstat(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://external-dev.10x10.co.kr:80/apis/product/view", false
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

'정보고시 조회
Public Function fnEzwelLayout(igoodsGrpCd, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message
	Set obj = jsObject()
		obj("action") = iaction
		obj("goodsGrpCd") = igoodsGrpCd		'1001 / 1002 / 1003 ~~~~
		obj("mallId") = CMALLNAME
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr:80/apis/product/code", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/code", false
		End If
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			'If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			'End If
		Set strObj = nothing
	Set objXML= nothing
End Function

''이하는 주문 관련
Public Function fnEzwelSongjangUpload(outmallorderserial, orgDetailKey, parcelCode, invoiceNumber, byRef resultCode, byRef resultMessage, byRef failCount, orgOutmallorderserial)
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

	'http://localhost:11117/admin/etc/ezwel/Ezwel_SongjangProc.asp?mallId=ezwel&ord_no=9000018375564&ord_dtl_sn=19&hdc_cd=33&inv_no=2323
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
				resultCode			= strObj.result.ezwelResponseModel.resultCode
				resultMessage		= strObj.result.ezwelResponseModel.resultMsg
				failCount			= strObj.result.failCount
			ElseIf objXML.Status = "504" Then
				resultCode			= strObj.code
				resultMessage 		= iRbody
			Else
				status				= strObj.status
				resultCode			= strObj.code
			End If
		Set strObj = nothing

		If (session("ssBctID")="kjy8517") Then
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set objXML= nothing
End Function

Public Function fnEzwelDlvFinish(outmallorderserial, orgDetailKey)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message

	Set obj = jsObject()
		obj("detailKey") = orgDetailKey
		obj("mallId") = CMALLNAME
		obj("outmallOrderSerial") = outmallorderserial
		obj("statusKey") = "1004"
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr:80/apis/order/update/status", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/order/update/status", false
		End If
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				resultCode			= strObj.result.ezwelResponseModel.resultCode
				resultMessage		= strObj.result.ezwelResponseModel.resultMsg
				failCount			= strObj.result.failCount
				response.write "OK"
			ElseIf objXML.Status = "504" Then
				resultCode			= strObj.code
				resultMessage 		= iRbody
			Else
				status				= strObj.status
				resultCode			= strObj.code
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

Public Function fnEzwelSongjangUploadByManager(isellsite, outmallorderserial, originDetailKey, sendState)
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
				fnEzwelSongjangUploadByManager = strObj.result
			End If
		Set strObj = nothing

		If (session("ssBctID")="kjy8517") Then
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set objXML= nothing
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
%>