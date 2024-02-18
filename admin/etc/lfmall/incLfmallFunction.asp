<%
'############################################## 실제 수행하는 API 함수 모음 시작 ############################################
'상품 등록
Public Function fnLfmallItemReg(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	istrParam = "?itemid="&iitemid&"&scmid="&session("ssBctID")
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://110.93.128.100:8090/lfmall/product/productreg" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품등록] " & html2db(Err.Description)
			Exit Function
		End If

		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success
				iMessage		= strObj.message
				If isSuccess = true Then
					iErrStr = "OK||"&iitemid&"||성공[상품등록]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품등록] " & iMessage
					If (session("ssBctID")="kjy8517") Then
						response.write BinaryToText(objXML.ResponseBody,"utf-8")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품등록] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품등록_NO] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품 수정
Public Function fnLfmallItemEdit(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccess
	istrParam = "?itemid="&iitemid&"&scmid="&session("ssBctID")
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://110.93.128.100:8090/lfmall/product/update" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상품수정] " & html2db(Err.Description)
			Exit Function
		End If
		'rw objXML.Status
		'rw BinaryToText(objXML.ResponseBody,"utf-8")
		'response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success
				iMessage		= strObj.message
				If isSuccess = true Then
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||성공[상품수정]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품수정] " & iMessage
					If (session("ssBctID")="kjy8517") Then
						response.write BinaryToText(objXML.ResponseBody,"utf-8")
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상품수정] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상품수정] "& html2db(replace(iRbody, """", ""))
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품 상태 수정
Public Function fnLfmallSellYN(iitemid, ichgSellYn, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, itemstat, isSuccess

	Select Case ichgSellYn
		Case "Y"	itemstat = "7"		'상품판매시작 MD 승인 필요
		Case "X"	itemstat = "5"		'상품정지(삭제같음)
		Case Else	itemstat = "6"		'일시품절
	End Select

	istrParam = "?itemid="&iitemid&"&stat="&itemstat&"&scmid="&session("ssBctID")
	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://110.93.128.100:8090/lfmall/product/manage" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[상태수정] " & html2db(Err.Description)
			Exit Function
		End If
'		rw objXML.Status
'		rw BinaryToText(objXML.ResponseBody,"utf-8")
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success
				iMessage		= strObj.message
				If isSuccess = true Then
					If ichgSellyn = "Y" Then
						iErrStr =  "OK||"&iitemid&"||판매승인대기[상태수정]"
					ElseIf ichgSellyn = "N" Then
						iErrStr =  "OK||"&iitemid&"||품절처리[상태수정]"
					End If
				Else
					If Instr(iMessage, "상품 코드의 판매 상태 및 변경 상태 코드 확인 바랍니다") > 0 Then
						If ichgSellyn = "Y" Then
							iErrStr =  "OK||"&iitemid&"||판매승인대기[상태수정]"
						ElseIf ichgSellyn = "N" Then
							iErrStr =  "OK||"&iitemid&"||품절처리[상태수정]"
						End If
					Else
						iErrStr = "ERR||"&iitemid&"||실패[상태수정] "& html2db(iMessage)
					End If
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[상태수정] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[상태수정] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function

'상품 조회
Public Function fnLfmallStatChk(iitemid, byRef iErrStr)
    Dim objXML, strSql, i, iRbody, iMessage, istrParam, strObj, isSuccessCode, isSuccess
	Dim slitmAprvlGbcd, productStatusName, itemSellGbcd, ihmallSellYn, ihmallPrice, ostkYn, itemAthzGbcd, itemAthzGbcdNm
	istrParam = "?itemid="&iitemid&"&scmid="&session("ssBctID")

	On Error Resume Next
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.open "GET", "http://110.93.128.100:8090/lfmall/product/info" & istrParam, false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()

		If Err.number <> 0 Then
			iErrStr = "ERR||"&iitemid&"||실패[조회] " & html2db(Err.Description)
			Exit Function
		End If
'		rw objXML.Status
'		response.end
		If objXML.Status = "200" OR objXML.Status = "201" Then
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
'			response.write iRbody
'			response.end
			Set strObj = JSON.parse(iRbody)
				isSuccess		= strObj.success
				iMessage		= strObj.message
				If isSuccess = true Then
					productStatusName = strObj.outValue.body.product.productStatusName
					strSql = ""
					strSql = strSql & " UPDATE db_etcmall.dbo.tbl_lfmall_regitem " & VbCRLF
					strSql = strSql & " SET lastConfirmdate = getdate() " & VbCRLF
					strSql = strSql & " WHERE itemid='" & iitemid & "'"
					dbget.Execute(strSql)
					iErrStr = "OK||"&iitemid&"||성공[조회("&html2db(iMessage)&")]"
				Else
					iErrStr = "ERR||"&iitemid&"||실패[조회] "& html2db(iMessage)
				End If
			Set strObj = nothing
		Else
			iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
			Set strObj = JSON.parse(iRbody)
				iMessage			= strObj.message
				'rw iRbody
				If iMessage <> "" Then
					iErrStr = "ERR||"&iitemid&"||실패[조회] "& html2db(iMessage)
				Else
					iErrStr = "ERR||"&iitemid&"||실패[조회] 통신오류"
				End If
			Set strObj = nothing
		End If
	Set objXML= nothing
End Function
'############################################## 실제 수행하는 API 함수 모음 끝 ############################################
%>