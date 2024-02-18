<%@ language=vbscript %>
<% option explicit %>
<%
Response.Expires = 0
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "Cache-Control","no-cache,must-revalidate"

'###############################################
' PageName : ajaxEventImageLinkSet.asp
' Discription : I형(통합형) 이벤트 이미지맵 저장
' History : 2019.12.16 정태훈
'###############################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->

<%
dim strSql
Dim mode : mode = requestCheckvar(request("mode"),1)
dim masterIdx : masterIdx = requestCheckVar(Request("masterIdx"),10)
dim x1 : x1 = requestCheckVar(Request("x1"),4)
dim y1 : y1 = requestCheckVar(Request("y1"),4)
dim x2 : x2 = requestCheckVar(Request("x2"),4)
dim y2 : y2 = requestCheckVar(Request("y2"),4)
dim linkurl : linkurl = requestCheckVar(Request("linkurl"),128)
dim didx : didx = requestCheckVar(Request("didx"),10)

if masterIdx="" then
    Response.Write "1"
	dbget.close()	:	response.End
end if

if mode="W" then
    if linkurl <> "" then
        if checkNotValidHTML(linkurl) then
            Response.Write "3"
            dbget.close()	:	response.End
        end if
    end If
    strSql = "EXEC [db_event].[dbo].[usp_SCM_Event_ImageMap_Set] " & Cstr(masterIdx) & ", " & Cstr(x1) & ", " & Cstr(y1)  & ", " & Cstr(x2)  & ", " & Cstr(y2) & ", '" & Cstr(linkurl) & "', '" & Cstr(Session("ssBctId")) & "'"
    dbget.execute strSql
elseif mode="E" then
    if linkurl <> "" then
        if checkNotValidHTML(linkurl) then
            Response.Write "3"
            dbget.close()	:	response.End
        end if
    end If
    strSql = "EXEC [db_event].[dbo].[usp_SCM_Event_ImageMap_Upd] " & Cstr(didx) & ", " & Cstr(x1) & ", " & Cstr(y1)  & ", " & Cstr(x2)  & ", " & Cstr(y2) & ", '" & Cstr(linkurl) & "', '" & Cstr(Session("ssBctId")) & "'"
    dbget.execute strSql
else
    strSql = "EXEC [db_event].[dbo].[usp_SCM_Event_ImageMap_Del] " & Cstr(didx)
    dbget.execute strSql
end if

if Err.Number <> 0 then
    Response.Write "2"
    dbget.close()	:	response.End
else
    response.write "0"
    dbget.close()	:	response.End
end if
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->