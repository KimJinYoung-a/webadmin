<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  텐바이텐 메일진
' History : 2018.04.27 이상구 생성(메일러 연동 생성 메일러로 발송 내역 전송. 메일 가져오기 생성.)
'			2019.06.24 정태훈 수정(템플릿 기능 신규 추가)
'			2020.05.28 한용민 수정(TMS 메일러 추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<%
Dim idx, isusing, chisusing, reservationOK, saveHtml, fixedHTML, fixedHTMLNoMember
Dim strSql, i, lineSel, mode

Function GetTextFromUrl(url)
  Dim oXMLHTTP

  Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")

  oXMLHTTP.Open "GET", url, False
  oXMLHTTP.Send

  If oXMLHTTP.Status = 200 Then
    GetTextFromUrl = oXMLHTTP.responseText
  End If
End Function

idx 			= requestCheckVar(request("idx"), 32)
mode 			= requestCheckVar(request("mode"), 8)
isusing 		=  requestCheckVar(request("isusing"), 32)
reservationOK 	= requestCheckVar(request("reservationOK"),32)
saveHtml 		= requestCheckVar(request("saveHtml"),32)
lineSel 		=  request("lineSel")

if lineSel <> "" then
	if checkNotValidHTML(lineSel) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
	response.write "</script>"
	response.End
	end if
end If

if mode="delete" then
	strSql = "Update [db_sitemaster].[dbo].tbl_mailzine set " & vbcrlf
	strSql = strSql & " deleteyn='Y'" & vbcrlf
	strSql = strSql & " where idx in (" & lineSel & ")"
	dbget.execute strSql
	Response.Write "<script>parent.location.reload();</script>"
	dbget.close()
	response.end
else
	If (reservationOK = "OK") Then
		''if (saveHtml = "Y") then
		''	fixedHTML = GetTextFromUrl("http://webadmin.10x10.co.kr/admin/mailzine/mailzine_display_new.asp?idx=" & idx & "&member=member&type=view")
		''	fixedHTMLNoMember = GetTextFromUrl("http://webadmin.10x10.co.kr/admin/mailzine/mailzine_display_new.asp?idx=" & idx & "&member=notmember&type=view")

		''	response.write fixedHTML
		''	dbget.close() : response.end
		''end if

		strSql = ""
		strSql = strSql & " Update [db_sitemaster].[dbo].tbl_mailzine set " & vbcrlf
		strSql = strSql & " reservationDATE = getdate() " & vbcrlf
		strSql = strSql & " where idx = " & idx
		dbget.execute strSql

		Response.Write "<script>parent.location.reload();</script>"
		dbget.close()
		response.end
	Else
		If isusing = "Y" Then
			chisusing = "N"
		Else
			chisusing = "Y"
		End If
			strSql = ""
			strSql = strSql & " Update [db_sitemaster].[dbo].tbl_mailzine set " & vbcrlf
			strSql = strSql & " isusing = '" & chisusing & "' " & vbcrlf
			strSql = strSql & " where idx = " & idx
			dbget.execute strSql
			Response.Write "<script>parent.location.reload();</script>"
			dbget.close()
			response.end
	End If
End If
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
