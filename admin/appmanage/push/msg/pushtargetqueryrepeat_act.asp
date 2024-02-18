<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 예약 푸시 메시지 작성
' Hieditor : 2022.03.17 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->
<%
dim targetkey, mode, strSql, replacetagcode
    targetkey = requestCheckVar(trim(Request("targetkey")),10)
    mode = requestCheckVar(trim(Request("mode")),32)
%>
<%
if mode="replacetagcode" then
    if targetkey="" or isnull(targetkey) then
        session.codePage = 949
        response.end
    end if

    strSql = "SELECT" & vbcrlf
    strSql = strSql & " q.targetKey,q.targetName,q.targetQuery,q.isusing,q.repeatpushyn, q.target_procedureyn, q.replacetagcode" & vbcrlf
    strSql = strSql & " From db_contents.[dbo].[tbl_app_targetQuery] q with (readuncommitted)" & vbcrlf
    strSql = strSql & " WHERE q.isusing=N'Y'" & vbcrlf

    if targetkey<>"" THEN
        strSql = strSql & " and q.targetkey=N'"& targetkey &"'" & vbcrlf
    end if

    'response.write strSql &"<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    IF not rsget.EOF THEN
        replacetagcode 	= db2html(rsget("replacetagcode"))
    End IF			
    rsget.Close	
%>
<br><br>※ 실제 고객 데이터로 치환되는코드 (제목,내용,링크)
<br><font color="red"><%= replacetagcode %></font>
<%
end if
session.codePage = 949
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->