<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0   
Response.AddHeader "Pragma","no-cache"   
Response.AddHeader "Cache-Control","no-cache,must-revalidate"   

'###############################################
' PageName : appDedicatedItem_process.asp
' Discription : 앱전용 응모템 설정 프로세스
' History : 2023.02.07 정태훈
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
dim eCode, mode, strSql, itemarr
eCode = requestCheckVar(Request.Form("evt_code"),10)
mode = requestCheckVar(Request.Form("mode"),10)
itemarr = Trim(Request.Form("itemarr"))

if eCode="" then
    response.write "<script type='text/javascript'>"
    response.write "	alert('유효하지 않은 데이터 입니다. 다시 시도해 주세요.');history.back();"
    response.write "</script>"
    response.End
end if

if mode="Add" then
    strSql = "INSERT INTO [db_event].[dbo].[tbl_event_secret_shop_item](evt_code,itemidarr)" & vbCrlf
    strSql = strSql + " VALUES(" & eCode & ",'" & itemarr & "')"
    dbget.execute strSql
elseif mode="Modify" then
    strSql = "UPDATE [db_event].[dbo].[tbl_event_secret_shop_item]" & vbCrlf
    strSql = strSql + " SET itemidarr='" & itemarr & "'" & vbCrlf
    strSql = strSql + " where evt_code=" & eCode
    dbget.execute strSql
end if

	response.write "<script type='text/javascript'>"
	response.write "	location.replace('pop_secret_shop_setting.asp?evt_code="&eCode&"');"
	response.write "</script>"
	dbget.close()	:	response.End
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->