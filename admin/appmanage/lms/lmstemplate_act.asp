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
dim sendmethod, template_code, mode, strSql
dim contents, button_name, button_url_mobile, button_name2, button_url_mobile2, failed_type , failed_subject, failed_msg
    sendmethod = requestCheckVar(trim(Request("sendmethod")),16)
    template_code = requestCheckVar(trim(Request("template_code")),32)
    mode = requestCheckVar(trim(Request("mode")),32)
%>
<%
if mode="templateajax" then
    if sendmethod="" or isnull(sendmethod) then
        session.codePage = 949
        response.end
    end if
%>
템플릿 : <% drawSelectBoxtemplate "template_code", template_code, " onchange='calltemplatecontentsajax(inputfrm.sendmethod.value,this.value);'", sendmethod %><br><br>
<%
elseif mode="templatecontentsajax" then
    if sendmethod="" or isnull(sendmethod) then
        session.codePage = 949
        response.end
    end if
    if template_code="" or isnull(template_code) then
        session.codePage = 949
        response.end
    end if

    strSql = "SELECT" & vbcrlf
    strSql = strSql & " t.contents, t.button_name, t.button_url_mobile, t.button_name2, t.button_url_mobile2, t.failed_type, t.failed_subject" & vbcrlf
    strSql = strSql & " , t.failed_msg" & vbcrlf
    strSql = strSql & " From db_contents.dbo.tbl_lms_template t with (readuncommitted)" & vbcrlf
    strSql = strSql & " WHERE t.isusing=N'Y'" & vbcrlf

    if sendmethod<>"" THEN
        strSql = strSql & " and t.sendmethod=N'"& sendmethod &"'" & vbcrlf
    end if
    if template_code<>"" THEN
        strSql = strSql & " and t.template_code=N'"& html2db(template_code) &"'" & vbcrlf
    end if

    'response.write strSql &"<br>"
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
    IF not rsget.EOF THEN
        contents 	= db2html(rsget("contents"))
        button_name 	= db2html(rsget("button_name"))
        button_url_mobile 	= db2html(rsget("button_url_mobile"))
        button_name2 	= db2html(rsget("button_name2"))
        button_url_mobile2 	= db2html(rsget("button_url_mobile2"))
        failed_type 	= rsget("failed_type")
        failed_subject 	= db2html(rsget("failed_subject"))
        failed_msg 	= db2html(rsget("failed_msg"))
    End IF			
    rsget.Close	

    response.write "{""resultcode"":""00"",""contents"":"""&replace(contents,vbcrlf,"!@#")&""",""button_name"":"""&button_name&""",""button_url_mobile"":"""&button_url_mobile&""",""button_name2"":"""&button_name2&""",""button_url_mobile2"":"""&button_url_mobile2&""",""failed_type"":"""&failed_type&""",""failed_subject"":"""&failed_subject&""",""failed_msg"":"""&replace(failed_msg,vbcrlf,"!@#")&"""}"
    session.codePage = 949
    response.end
end if

session.codePage = 949
response.end
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->