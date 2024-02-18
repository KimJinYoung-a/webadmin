<%@ language=vbscript %>
<% option Explicit %>
<% Response.CharSet = "euc-kr" %>
<%
'####################################################
' Description : 돌아온 크리스박스의 기적!
' History : 2015.12.07 유태욱 생성
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim mode, evt_code, winnumber, sqlStr
	mode = requestcheckvar(request("mode"),32)
	evt_code = getNumeric(requestcheckvar(request("evt_code"),32))
	winnumber = getNumeric(requestcheckvar(request("winnumber"),4))

If session("ssBctId")="winnie" Or session("ssBctId")="gawisonten10" Or session("ssBctId") ="greenteenz" Or session("ssBctId") = "edojun" Or session("ssBctId") = "tozzinet" Or session("ssBctId") = "thensi7" Then
Else
	response.write "<script>alert('관계자만 볼 수 있는 페이지 입니다.');window.close();</script>"
	response.End
End If

dim refer
	refer = request.ServerVariables("HTTP_REFERER")
if InStr(refer,"10x10.co.kr")<1 then
	Response.Write "<script type='text/javascript'>alert('잘못된 접속입니다.');</script>"
	dbget.close() : Response.End
end If

If mode = "winnumber" Then
	if evt_code="" then
		Response.Write "<script type='text/javascript'>alert('이벤트코드가 없습니다.');</script>"
		dbget.close() : Response.End
	end If
	if winnumber="" then
		Response.Write "<script type='text/javascript'>alert('확률이 없습니다.');</script>"
		dbget.close() : Response.End
	end If

	sqlStr = "update db_temp.dbo.tbl_event_etc_yongman" + vbcrlf
	sqlStr = sqlStr & " set bigo='"& winnumber &"' where" + vbcrlf
	sqlStr = sqlStr & " isusing='Y' and event_code='"& evt_code &"'"

	'response.write sqlStr & "<br>"
	dbget.execute sqlStr
	
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('OK');"
	Response.Write "	parent.top.location.replace('/admin/datamart/mkt/event67929_manage.asp');"
	Response.Write "</script>"
	dbget.close() : Response.End

else
	Response.Write "<script type='text/javascript'>alert('구분자가 없습니다.');</script>"
	dbget.close() : Response.End
end if

%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
