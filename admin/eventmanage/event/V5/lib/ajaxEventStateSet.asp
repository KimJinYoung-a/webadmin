<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "Cache-Control","no-cache,must-revalidate"

'###############################################
' PageName : ajaxEventStateSet.asp
' Discription : I형(통합형) 이벤트 상태 설정
' History : 2019.01.24 정태훈
'###############################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
dim eState : eState = requestCheckVar(Request("eState"),2)
dim eCode : eCode = requestCheckVar(Request("eC"),10)


if eCode="" or eState="" then
    Response.Write "1"
	dbget.close()	:	response.End
end if

dim strSql
strSql = "UPDATE [db_event].[dbo].[tbl_event]" & vbCrlf
strSql = strSql + " SET evt_state=" & eState & vbCrlf
strSql = strSql + ", evt_lastupdate=getdate()" & vbCrlf
strSql = strSql + ", adminid='" & session("ssBctId") & "'" & vbCrlf
strSql = strSql + " WHERE evt_code=" & eCode
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