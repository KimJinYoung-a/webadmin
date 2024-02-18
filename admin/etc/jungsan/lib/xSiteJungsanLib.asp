<%
Function GetJungsan_ezwel(reqDate, hasNext, vpage, vTotalPage)
	Dim sellsite : sellsite = "ezwel"
	Dim objXML, xmlDOM, sqlStr, resultMessage
	Dim accountCnt, strObj, page
	Dim getParam, totalPage, iRbody
	getParam = ""
	getParam = getParam & "mallId=ezwel&page="&vpage&"&requestDate="&reqDate
	GetJungsan_ezwel = False
	accountCnt = 0

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://external-dev.10x10.co.kr:80/apis/settlement?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/settlement?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

		Set strObj = JSON.parse(iRbody)
			resultMessage	= iRbody

'			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
'			End If

			If objXML.Status = "200" OR objXML.Status = "201" Then
				accountCnt		= strObj.result.ezwelResponseModel.accountCnt
				hasnext			= strObj.result.resultSettlement.hasNext
				vpage			= strObj.result.resultSettlement.page
				page			= strObj.result.ezwelResponseModel.page
				totalPage		= strObj.result.ezwelResponseModel.totalPage
				vTotalPage		= totalPage
			Else
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
				response.end
			End If

			If accountCnt > 0 Then
				If CSTR(page) = CSTR(totalPage) Then
					sqlStr = " EXEC db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_ezwel] "
					dbget.Execute sqlStr
				End If
			Else
				response.write reqDate & " JungsanData Not Exists"
				response.end
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function

Function GetJungsan_wconcept1010(reqDate, vpage)
	Dim sellsite : sellsite = "wconcept1010"
	Dim objXML, xmlDOM, sqlStr, resultMessage
	Dim totalCount, strObj, page
	Dim getParam, totalPage, iRbody
	getParam = ""
	getParam = getParam & "mallId=wconcept1010&page="&vpage&"&requestDate="&reqDate
	GetJungsan_wconcept1010 = False
	totalCount = 0

	Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If application("Svr_Info")="Dev" Then
			objXML.open "GET", "http://gateway-dev.10x10.co.kr/external/apis/settlement?" & getParam, false
		Else
			objXML.open "GET", "http://gateway.10x10.co.kr/external/apis/settlement?" & getParam, false
		End If
		objXML.setRequestHeader "Content-Type", "application/json"
		objXML.setTimeouts 5000,80000,80000,80000
		objXML.Send()
		iRbody = BinaryToText(objXML.ResponseBody,"utf-8")

		Set strObj = JSON.parse(iRbody)
			resultMessage	= iRbody
			If (session("ssBctID")="kjy8517") Then
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
			End If

			If objXML.Status = "200" OR objXML.Status = "201" Then
				totalCount = strObj.result.result.result.totalCount
				totalPage = strObj.result.result.result.totalPage
			Else
				rw "RES : <textarea cols=40 rows=10>"&iRbody&"</textarea>"
				response.end
			End If

			If totalCount > 0 Then
				If CSTR(vpage) = CSTR(totalPage) Then
					sqlStr = " EXEC db_jungsan.[dbo].[usp_Ten_OUTAMLL_Jungsan_MappingTmp_wconcept] "
					dbget.Execute sqlStr
				End If
			Else
				response.write reqDate & " JungsanData Not Exists"
				response.end
			End If
		Set strObj = nothing
	Set objXML= nothing
End Function
%>