<%
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

Function GetCSQnA_kakaostore()
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId=kakaostore"

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr/apis/qna?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/qna?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				rw "Q&A 미답변 건수 : " & strObj.result.totalCount
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function resCSQnA_kakaostore(iNum, iRply)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim istrParam
	Set obj = jsObject()
		obj("mallId") = "kakaostore"
		obj("qnaId") = iNum
		obj("replyContent") = iRply
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr/apis/qna", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/qna", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				rw iNum & " 답변완료"
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSQnA_wconcept(selldate)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId=wconcept1010&sellDate="&selldate

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr/external/apis/qna?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/qna?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				rw "Q&A 미답변 건수 : " & strObj.result.result.result.contentsCount
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function resCSQnA_wconcept(iNum, iRply, outmallGoodNo)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim istrParam, strSql
	Set obj = jsObject()
		obj("mallId") = "wconcept1010"
		obj("qnaId") = iNum
		obj("outmallGoodNo") = outmallGoodNo
		obj("replyContent") = iRply
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/qna", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/qna", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				rw iNum & " 답변완료"
			End If

			If strObj.result.result.result.errorInfo.detail = "이미 답변 완료된 상품 문의입니다." Then
				strSql = ""
				strSql = strSql & " UPDATE db_temp.dbo.tbl_Sabannet_Detail "
				strSql = strSql & " SET CS_STATUS = '003' "
				strSql = strSql & " ,TenStatus = 'C' "
				strSql = strSql & " WHERE SabanetNum = '"& iNum &"' "
				strSql = strSql & " and SellSite = 'wconcept1010' "
				dbget.Execute strSql
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSQnA_ebay(mallid, selldate, code)
	Dim obj, iRbody, strObj
	Dim objXML, status, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId="&mallid&"&sellDate="&selldate&"&code="&code

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr/external/apis/qna?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/qna?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				If strObj.result.resultMessage = "OK" Then
					If code = "2" Then
						rw "Q&A[쪽지] 미답변 건수 : " & strObj.result.ebayResponse.length
					Else
						rw "Q&A 미답변 건수 : " & strObj.result.ebayResponse.length
					End If
				Else
					rw "Q&A 미답변 건수 : 0 "
				End If
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function resCSQnA_ebay(iNum, iRply, token, sellsite)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim istrParam, strSql
	Set obj = jsObject()
		obj("mallId") = sellsite
		obj("qnaId") = iNum
		obj("replyContent") = iRply
		obj("outmallGoodNo") = token
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/qna", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/qna", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				If strObj.result.resultMessage = "OK" Then
					rw iNum & " 답변완료"
				Else
					rw iNum & " 실패"
				End If
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSQnA_ebay_refer(mallid, selldate)
	Dim obj, iRbody, strObj
	Dim objXML, status, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId="&mallid&"&sellDate="&selldate

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr/external/apis/qna/cs?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/qna/cs?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				If strObj.result.resultMessage = "OK" Then
					If isNull(strObj.result.ebayResponse.Data) Then
						rw "Refer Q&A 미답변 건수 : 0"
					Else
						rw "Refer Q&A 미답변 건수 : " & strObj.result.ebayResponse.Data.length
					End If
				End If
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function resCSQnA_ebay_refer(iNum, iRply, sellsite)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim istrParam, strSql
	Set obj = jsObject()
		obj("mallId") = sellsite
		obj("qnaId") = iNum
		obj("replyContent") = iRply
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/qna/cs", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/qna/cs", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				If strObj.result.resultMessage = "OK" Then
					rw iNum & " 답변완료"
				Else
					rw iNum & " 실패"
				End If
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSQnA_benepia(selldate)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId=benepia1010&sellDate="&selldate

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr/external/apis/qna?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/qna?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				If strObj.result.benepiaResponse.responseCode = "S000" Then
					rw "Q&A 미답변 건수 : " & strObj.result.benepiaResponse.responseEncData.length
				ElseIf strObj.result.benepiaResponse.responseCode = "S001" Then
					rw "Q&A 미답변 건수 : 0 "
				End If
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function resCSQnA_benepia(iNum, iRply)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim istrParam, strSql
	Set obj = jsObject()
		obj("mallId") = "benepia1010"
		obj("qnaId") = iNum
		obj("replyContent") = iRply
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://gateway-dev.10x10.co.kr/external/apis/qna", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/qna", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				If strObj.result.benepiaResponse.responseCode = "S000" Then
					rw iNum & " 답변완료"
				Else
					rw iNum & " 실패"
				End If
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSQnA_boribori(currDate)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId=boribori1010&sellDate="&currDate

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr/apis/qna?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/qna?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				rw "Q&A 미답변 건수 : " & strObj.result.data.result.length
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function resCSQnA_boribori(iNum, iRply)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim istrParam
	Set obj = jsObject()
		obj("mallId") = "boribori1010"
		obj("qnaId") = iNum
		obj("replyContent") = iRply
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr/apis/qna", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/qna", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				rw iNum & " 답변완료"
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetCSQnA_boribori_Refer(currDate)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim getParam
	getParam = ""
	getParam = getParam & "mallId=boribori1010&sellDate="&currDate

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr/apis/qna/cs?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/qna/cs?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				rw "Q&A Refer 건수 : " & strObj.result.data.result.length
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function resCSQnA_boribori_Refer(iNum, iRply)
	Dim obj, iRbody, strObj
	Dim objXML, status, code, message
	Dim istrParam
	Set obj = jsObject()
		obj("mallId") = "boribori1010"
		obj("qnaId") = iNum
		obj("replyContent") = iRply
		istrParam = obj.jsString
	Set obj = nothing

	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "POST", "http://external-dev.10x10.co.kr/apis/qna/cs", false
		Else
			objXML.open "POST", "http://gateway.10x10.co.kr/external/apis/qna/cs", false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send(istrParam)
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")
		Set strObj = JSON.parse(iRbody)
			If objXML.Status = "200" OR objXML.Status = "201" Then
				rw iNum & " 답변완료"
			End If

			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function getCSAnswerComplete(imallid)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT SabanetNum, isNull(RPLY_CNTS, '') as RPLY_CNTS, PRODUCT_ID, COMPAYNY_GOODS_CD "
	strSql = strSql & " FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) "
	strSql = strSql & " WHERE SellSite = '"& imallid &"' "
	strSql = strSql & " AND TenStatus = 'S' "
	strSql = strSql & " and CS_STATUS = '001' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		getCSAnswerComplete = rsget.getRows
	End If
	rsget.close
End Function

Function getCSAnswerCompleteMallId(imallid, v)
	Dim strSql
	strSql = ""
	strSql = strSql & " SELECT SabanetNum, isNull(RPLY_CNTS, '') as RPLY_CNTS "
	strSql = strSql & " FROM db_temp.dbo.tbl_Sabannet_Detail with (nolock) "
	strSql = strSql & " WHERE SellSite = '"& imallid &"' "
	strSql = strSql & " AND TenStatus = 'S' "
	strSql = strSql & " and CS_STATUS = '001' "
	If imallid = "boribori1010" Then
		If v = "1" Then
			strSql = strSql & " and MALL_ID = '보리보리(Q&A)' "
		Else
			strSql = strSql & " and MALL_ID = '보리보리(CS)' "
		End If
	ElseIf imallid = "auction1010" Then
		If v = "1" Then
			strSql = strSql & " and MALL_ID = '옥션(Q&A)' "
		Else
			strSql = strSql & " and MALL_ID = '옥션(CS_Q&A)' "
		End If
	ElseIf imallid = "gmarket1010" Then
		If v = "1" Then
			strSql = strSql & " and MALL_ID = '지마켓(Q&A)' "
		Else
			strSql = strSql & " and MALL_ID = '지마켓(CS_Q&A)' "
		End If
	End If
	rsget.CursorLocation = adUseClient
	rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	If Not(rsget.EOF or rsget.BOF) Then
		getCSAnswerCompleteMallId = rsget.getRows
	End If
	rsget.close
End Function
%>