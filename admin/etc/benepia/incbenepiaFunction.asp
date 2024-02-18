<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 + 이미지 등록
Public Function fnbenepiaItemReg(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/create", false
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
				On Error Goto 0
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

'상품만 등록
Public Function fnbenepiaOnlyItemReg(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/create/only", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/create/only", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				On Error Resume Next
					resultMessage	= strObj.result.resultMessage
				On Error Goto 0
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

'상품이미지 등록
Public Function fnbenepiaImageReg(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/create/image", false
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
				On Error Goto 0
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
Public Function fnbenepiaSellYN(iitemid, iaction, ichgSellYn, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/status", false
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

'상품 수정
Public Function fnbenepiaItemEdit(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update", false
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
				strSQL = strSQL & " WHERE mallid = 'benepia1010' " & VBCRLF
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

'상품 카테고리 수정
Public Function fnbenepiaCategoryEdit(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/category", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/category", false
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

'가격 수정
Public Function fnbenepiaPriceEdit(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/price", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/price", false
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

'재고 수정
Public Function fnbenepiaQuantityEdit(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/option/qty", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/option/qty", false
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

'상품 정보 수정
Public Function fnbenepiaItemInfoEdit(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/only", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/only", false
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

'상품 배송 정보 수정
Public Function fnbenepiaItemDeliveryEdit(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/delivery", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/delivery", false
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

'상품 이미지 수정
Public Function fnbenepiaItemImageEdit(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj, imgType
	Dim objXML, status, code, message, strSQL
	imgType = "1"		'기본 PC이미지
	If iaction = "EDITIMAGEMOB" Then
		imgType = "2"
	End If

	Set obj = jsObject()
		obj("action") = iaction
		obj("adminId") = session("ssBctID")
		obj("itemId") = iitemid
		obj("mallId") = CMALLNAME
		obj("status") = imgType
		obj("user") = "admin"
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/image", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/image", false
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

'상품 설명 수정
Public Function fnbenepiaItemContentEdit(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj, imgType
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/content", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/content", false
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

'상품 안전인증 수정
Public Function fnbenepiaItemSafeInfoEdit(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj, imgType
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/safecert", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/safecert", false
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

'상품 정보고시 수정
Public Function fnbenepiaItemInfoCodeEdit(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj, imgType
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/infocode", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/infocode", false
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

'옵션 수정
Public Function fnbenepiaItemEditOption(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj, imgType
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/option", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/option", false
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

'상품 + 옵션 조회
Public Function fnwbenepiaChkstat(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj, imgType
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/view", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/view", false
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

'상품 조회
Public Function fnwbenepiaChkItem(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj, imgType
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/view/item", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/view/item", false
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

'옵션 조회
Public Function fnwbenepiaChkOpt(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj, imgType
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/view/option", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/view/option", false
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

Public Function fnbenepiaCategory(idepth, icateCode)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message, strSql, datalist, i
	Dim CateKey, name, depth, lastLevel, ctryChidFg, parentCateKey
	Set obj = jsObject()
		obj("categoryLevel") = idepth
		obj("mallId") = CMALLNAME
		obj("parentCategoryId") = icateCode
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/code/category", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/code/category", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)

		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			status			= strObj.status
			code			= strObj.code
			If strObj.result.responseCode = "S000" Then
				Set datalist = strObj.result.responseEncData
					For i=0 to datalist.length-1
						CateKey = datalist.get(i).ctryCode
						name = datalist.get(i).ctryNm
						depth = datalist.get(i).ctryLvl
						lastLevel = datalist.get(i).ctryChidFg
						parentCateKey = datalist.get(i).parCtryCode

						strSql = ""
						strSql = strSql & " INSERT INTO db_etcmall.[dbo].[tbl_benepia_category] (CateKey, name, depth, lastLevel, parentCateKey, regdate) VALUES "
						strSql = strSql & " ('"& CateKey &"', '"& name &"', '"& depth &"', '"& lastLevel &"', '"& parentCateKey &"', GETDATE()) "
						dbget.execute strSql
					Next
			End If
			' If (session("ssBctID")="kjy8517") Then
			' 	rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			' End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Public Function fnbenepiaCommonCode(iccd, infocodedtlCode)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message, strSql, datalist, i
	Dim action
	Dim stdClsId, itmCnts, itmExpl, itmSeq

	Select Case iccd
		Case "md"				action = "MD"
		Case "area"				action = "AREA"
		Case "safe"				action = "SAFE_CODE"
		Case "infocode"			action = "PRODUCT_NOTICE_INFO"
		Case "infocodedtl"		action = "PRODUCT_NOTICE_INFO_DETAIL"
		Case "casedelivery"		action = "CASE_DELIVERY"
		Case "parcel"			action = "PARCELCODE"
		Case "brand"			action = "BRAND"
		Case "locaddress"		action = "LOC_ADDRESS"
	End Select

	Set obj = jsObject()
		obj("action") = action
		obj("mallId") = CMALLNAME
		If action = "PRODUCT_NOTICE_INFO_DETAIL" Then
			obj("goodsGrpCd") = infocodedtlCode
		End If
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/code", false
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

			If action = "PRODUCT_NOTICE_INFO_DETAIL" Then
				Set datalist = strObj.result.responseEncData
					For i=0 to datalist.length-1
						stdClsId = datalist.get(i).stdClsId
						itmCnts = datalist.get(i).itmCnts
						itmExpl = datalist.get(i).itmExpl
						itmSeq = datalist.get(i).itmSeq

						strSql = ""
						strSql = strSql & " INSERT INTO db_item.dbo.tbl_benepia_infoCode (stdClsId, itmCnts, itmExpl, itmSeq) VALUES "
						strSql = strSql & " ('"& stdClsId &"', '"& itmCnts &"', '"& itmExpl &"', '"& itmSeq &"') "
						dbget.execute strSql
					Next
				Set datalist = nothing
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

''이하는 주문 관련
Public Function fnbenepiaSongjangUpload(outmallorderserial, orgDetailKey, parcelCode, invoiceNumber, byRef resultMessage, byRef failCount, orgOutmallorderserial)
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
	'http://localhost:11117/admin/etc/benepia/benepia_SongjangProc.asp?mallId=benepia1010&ord_no=229694620&ord_dtl_sn=214154419&hdc_cd=33&inv_no=2323

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
				resultMessage		= strObj.result.resultMessage
				failCount			= strObj.result.failCount
			End If
		Set strObj = nothing

		If (session("ssBctID")="kjy8517") Then
			rw "REQ : <textarea cols=40 rows=10>"&getParam&"</textarea>"
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set objXML= nothing
End Function

Public Function fnbenepiaSongjangUploadByManager(isellsite, outmallorderserial, originDetailKey, sendState)
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
				fnbenepiaSongjangUploadByManager = strObj.result
			End If
		Set strObj = nothing

		If (session("ssBctID")="kjy8517") Then
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set objXML= nothing
End Function
%>