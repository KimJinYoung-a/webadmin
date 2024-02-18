<%@ language=vbscript %>
<% option explicit %>
<% response.charset = "euc-kr" %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim vCateCode, vQuery, vResultArr, vResult, vResultMsg, i
	vCateCode = Trim(requestCheckVar(request("catecode"),300))
	
	If vCateCode <> "" Then
		If Right(vCateCode,1) = "," Then
			vCateCode = Left(vCateCode, Len(vCateCode)-1)
		End If
		
		vQuery = "select safetyinfotype from db_item.dbo.tbl_display_cate where catecode in(" & vCateCode & ") and safetyinfotype in('necessary','choice')"
		rsget.CursorLocation = adUseClient
		rsget.Open vQuery,dbget,adOpenForwardOnly,adLockReadOnly
		
		If not rsget.eof Then
			Do Until rsget.Eof
				vResultArr = vResultArr & rsget(0) & ","
				rsget.movenext
			Loop
		End If
		rsget.close
		
		If InStr(vResultArr, "necessary") > 0 Then
			vResult = "necessary"
		End If
		
		If vResult <> "necessary" AND InStr(vResultArr, "choice") > 0 Then
				vResult = "choice"
		End If
	End If
	
	If vResult = "necessary" Then
		vResultMsg = "선택된 카테고리에 안전인증 정보가 필수 입력 정보입니다.br정보 입력을 안 한 경우br판매정지 또는 판매상 불이익을 당할 수 있습니다."
	ElseIf vResult = "choice" Then
		vResultMsg = "선택된 카테고리에 품목 중br안전인증 정보인 품목이 일부 존재합니다.br안전인증 필수 품목을 다시 한번 확인해주세요."
	End If
	
	Response.Write vResultMsg
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->