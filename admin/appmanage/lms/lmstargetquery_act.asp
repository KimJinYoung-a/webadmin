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
' Description : LMS발송관리
' Hieditor : 2020.03.19 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->
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
    strSql = strSql & " q.targetkey,q.targetName,q.targetQuery,q.isusing,q.repeatlmsyn,q.target_procedureyn,q.replacetagcode" & vbcrlf
    strSql = strSql & " From db_contents.dbo.tbl_lms_targetQuery q with (readuncommitted)" & vbcrlf
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
<br><br>※ 실제 고객 데이터로 치환되는코드 (제목,내용,실패시문자제목,실패시문자내용)
<br><font color="red"><%= replacetagcode %></font>
<%
end if
session.codePage = 949
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->