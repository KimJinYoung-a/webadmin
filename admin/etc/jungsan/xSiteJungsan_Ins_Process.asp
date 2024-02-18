<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
Server.ScriptTimeOut = 1200 ''초단위
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/admin/etc/jungsan/lib/xSiteJungsanLib.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim sellsite : sellsite	= requestCheckVar(html2db(request("sellsite")),32)
Dim reqDate : reqDate = request("reqDate")
Dim vPage, hasnext, isJungsanComplete, vTotalPage
vPage = request("page")
If vPage = "" Then
	vPage = 1
End If

If sellsite = "ezwel" Then
	If (reqDate = "") Then
		reqDate = Replace(Left(DateAdd("m", -1, NOW()), 7), "-", "")
	End If
	If Len(reqDate) <> 6 Then
		rw "날짜 형식이 잘 못 되었습니다."
		response.end
	End If
ElseIf sellsite = "wconcept1010" Then
	reqDate = Left(reqDate,4)&"-"&Mid(reqDate,5,2)&"-"&Mid(reqDate,7,2)
	If Len(reqDate) <> 10 Then
		rw "날짜 형식이 잘 못 되었습니다."
		response.end
	End If
Else
	response.write "잘못된 접근입니다."
	dbget.close : response.end
End If

isJungsanComplete = "N"
Select Case sellsite
	Case "ezwel"
		rw "호출월 : " & reqDate
		Do Until isJungsanComplete = "Y"
			Call GetJungsan_ezwel(reqDate, hasnext, vPage, vTotalPage)
			If hasnext = "N" Then
				isJungsanComplete = "Y"
				rw "완료"
			Else
				rw "API 호출 중 입니다. ("& vPage - 1 & "/" & vTotalPage & ")"
				rw "-------------------------"
			End If
			response.flush
		Loop
	Case "wconcept1010"
		rw "호출일 : " & reqDate
		Call GetJungsan_wconcept1010(reqDate, vPage)
		response.flush
		rw "완료"
End Select
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
