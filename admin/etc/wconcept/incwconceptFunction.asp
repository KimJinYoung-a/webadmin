<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 등록
Public Function fnwconceptItemReg(iitemid, iaction, byRef resultMessage)
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

'상품 등록
Public Function fnwconceptItemRegStep(iitemid, iaction, byRef resultMessage)
	Dim obj, istrParam, iRbody, strObj
	Dim objXML, status, code, message, strStep
	Set obj = jsObject()
		obj("action") = iaction
		obj("adminId") = session("ssBctID")
		obj("itemId") = iitemid
		obj("mallId") = CMALLNAME
		obj("user") = "admin"
		istrParam = obj.jsString
	Set obj = nothing

	Select Case iaction
		Case "REGSTEP1"		strStep = "step1"
		Case "REGSTEP2"		strStep = "step2"
		Case "REGSTEP3"		strStep = "step3"
		Case "REGSTEP4"		strStep = "step4"
		Case "REGSTEP5"		strStep = "step5"
		Case "REGSTEP6"		strStep = "step6"
	End Select

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/create/"&strStep, false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/create/"&strStep, false
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
						resultMessage = strObj.message
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
Public Function fnwconceptItemEdit(iitemid, iaction, byRef resultMessage)
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
				strSQL = strSQL & " WHERE mallid = 'wconcept1010' " & VBCRLF
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

'가격 수정
Public Function fnwconceptItemEditPrice(iitemid, iaction, byRef resultMessage)
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

'컨텐츠 수정
Public Function fnwconceptItemEditContent(iitemid, iaction, byRef resultMessage)
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

'이미지 수정
Public Function fnwconceptItemEditIMAGE(iitemid, iaction, byRef resultMessage)
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

'추가이미지 수정
Public Function fnwconceptItemEditAddIMAGE(iitemid, iaction, byRef resultMessage)
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
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/update/image/add", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/update/image/add", false
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
Public Function fnwconceptItemEditOption(iitemid, iaction, byRef resultMessage)
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

'정보고시 수정
Public Function fnwconceptInfoCode(iitemid, iaction, byRef resultMessage)
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

'상품 상태 수정
Public Function fnwconceptSellYN(iitemid, iaction, ichgSellYn, byRef resultMessage)
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

Public Function fnwconceptChkstat(iitemid, iaction, byRef resultMessage)
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
		 	objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/view", false
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

Public Function fnwconceptChkItem(iitemid, iaction, byRef resultMessage)
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
		 	objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/view/item", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/view/item", false
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

Public Function fnwconceptChkOpt(iitemid, iaction, byRef resultMessage)
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
		 	objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/product/view/option", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/product/view/option", false
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

Public Function fnwconceptCommonCode(ccd, goodsGrpCd)
	Dim obj, istrParam, iRbody, strObj, i
	Dim objXML, status, code, message, strSql, cCode, cCodeName, MediumCode, CategoryCode, productNoticeInfoList
	Dim icode, icodeName, defineTxt
	Set obj = jsObject()
		obj("action") = ccd
		obj("mallId") = CMALLNAME
		If ccd = "PRODUCT_NOTICE_INFO" Then
			obj("goodsGrpCd") = goodsGrpCd
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
			If status = "200" Then
				If ccd = "PRODUCT_NOTICE_INFO" Then
					cCode			= strObj.result.result.prodNoticeInfo.ccode
					cCodeName		= strObj.result.result.prodNoticeInfo.ccodeName

					MediumCode		= Split(goodsGrpCd, ",")(0)
					CategoryCode	= Split(goodsGrpCd, ",")(1)

					SET productNoticeInfoList = strObj.result.result.productNoticeInfoList
					If productNoticeInfoList.length > 0 Then
						For i=0 to productNoticeInfoList.length-1
							icode		= productNoticeInfoList.get(i).icode
							icodeName	= productNoticeInfoList.get(i).icodeName
							defineTxt	= productNoticeInfoList.get(i).defineTxt

							strSql = ""
							strSql = strSql & " INSERT INTO db_etcmall.dbo.[tbl_wconcept_infoCode] ([MediumCode], [CategoryCode], [cCode], [cCodeName], [icode], [icodeName], [defineTxt]) VALUES "
							strSql = strSql & " ('"& MediumCode &"', '"& CategoryCode &"', '"& cCode &"', '"& cCodeName &"', '"& icode &"', '"& icodeName &"', '"& defineTxt &"')  "
							dbget.Execute strSql
						Next
					End If
					SET productNoticeInfoList = nothing

					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.[tbl_wconcept_category] "
					strSql = strSql & " SET cCode = '"& cCode &"' "
					strSql = strSql & " , cCodeName = '"& cCodeName &"'"
					strSql = strSql & " WHERE MediumCode = '"& MediumCode &"' "
					strSql = strSql & " AND CategoryCode = '"& CategoryCode &"' "
					dbget.Execute strSql
					rw "OK"
				Else
					rw "OK"
				End If
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

''이하는 주문 관련
Public Function fnwconceptSongjangUpload(outmallorderserial, orgDetailKey, parcelCode, invoiceNumber, byRef resultMessage, byRef failCount, orgOutmallorderserial)
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

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr/external/apis/order/invoice?" & getParam, false
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
				If isNull(resultMessage) OR resultMessage = "" Then
					resultMessage = "FAIL"
				End If
				failCount			= strObj.result.failCount
			ElseIf objXML.Status = "504" Then
				resultMessage 		= iRbody
			Else
				status				= strObj.status
				resultCode			= strObj.code
			End If
		Set strObj = nothing

		If (session("ssBctID")="kjy8517") Then
			rw "REQ : <textarea cols=40 rows=10>"&getParam&"</textarea>"
			rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
		End If
	Set objXML= nothing
End Function

Public Function fnwconceptSongjangUploadByManager(isellsite, outmallorderserial, originDetailKey, sendState)
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
				fnwconceptSongjangUploadByManager = strObj.result
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