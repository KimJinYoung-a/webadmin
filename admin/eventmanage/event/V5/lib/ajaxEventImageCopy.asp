<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "Cache-Control","no-cache,must-revalidate"

'###############################################
' PageName : ajaxEventImageCopy.asp
' Discription : I형(통합형) 이벤트 이미지 복사
' History : 2019.09.11 정태훈
'###############################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
dim eCode : eCode = requestCheckVar(Request("eC"),10)

if eCode="" then
    Response.Write "1"
	dbget.close()	:	response.End
end if

dim strSql
    strSql = "exec db_event.dbo.usp_SCM_EventImageCopy_Set " & Cstr(eCode) & ", '" & Cstr(session("ssBctId")) &"'"
    dbget.execute strSql

if Err.Number <> 0 then
    Response.Write "2"
    dbget.close()	:	response.End
else
    response.write "0"
    dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->