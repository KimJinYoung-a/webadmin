<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : LMS발송관리
' Hieditor : 2020.03.20 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->
<%
dim iResult, sMode, strSql, refer, menupos
dim targetkey,targetName,targetQuery,isusing,repeatlmsyn, target_procedureyn, replacetagcode
	sMode		= requestCheckVar(trim(Request("sM")),1)
    targetkey		= requestCheckVar(trim(Request("targetkey")),10)
    targetName		= requestCheckVar(trim(Request("targetName")),100)
    targetQuery		= trim(Request("targetQuery"))
    isusing		= requestCheckVar(trim(Request("isusing")),1)
    repeatlmsyn		= requestCheckVar(trim(Request("repeatlmsyn")),1)
    target_procedureyn		= requestCheckVar(trim(Request("target_procedureyn")),1)
    menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)
    replacetagcode		= requestCheckVar(trim(Request("replacetagcode")),256)

refer = request.ServerVariables("HTTP_REFERER")

IF sMode = "I" THEN
	targetQuery = replace(targetQuery,"'", "''")

    strSql = "SELECT targetkey FROM db_contents.dbo.tbl_lms_targetQuery with (readuncommitted) Where targetkey="& targetkey

    'response.write strSql &"<br>"		
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	IF not (rsget.eof or rsget.bof) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('이미 존재하는 코드값입니다.다시 등록해주세요.');"
        session.codePage = 949
        response.write "	history.back();"
        response.write "</script>"
        rsget.close	: dbget.close()	: response.End
	end if
	rsget.close	
	
	strSql = " INSERT INTO db_contents.dbo.tbl_lms_targetQuery (" & vbcrlf
    strSql = strSql & " targetkey, targetName, targetQuery, isusing, repeatlmsyn, target_procedureyn, replacetagcode" & vbcrlf
    strSql = strSql & " ) Values(" & vbcrlf
	strSql = strSql & " "& targetkey &",N'"& html2db(targetName) &"',N'"& html2db(targetQuery) &"',N'"& isusing &"'" & vbcrlf
    strSql = strSql & " ,N'"& repeatlmsyn &"',N'"& target_procedureyn &"',N'"& replacetagcode &"'" & vbcrlf
    strSql = strSql & " )" & vbcrlf

    'response.write strSql &"<br>"
	dbget.execute strSql
				
ELSEIF sMode="U" THEN	
	targetQuery = replace(targetQuery,"'", "''")

    strSql =" UPDATE db_contents.dbo.tbl_lms_targetQuery" & vbcrlf
    strSql = strSql & " Set targetName = N'"& html2db(targetName) &"'" & vbcrlf
    strSql = strSql & " , targetQuery = N'"& html2db(targetQuery) &"'" & vbcrlf
    strSql = strSql & " , isusing =N'"& isusing &"'" & vbcrlf
    strSql = strSql & " , repeatlmsyn =N'"& repeatlmsyn &"'" & vbcrlf
    strSql = strSql & " , target_procedureyn =N'"& target_procedureyn &"'" & vbcrlf
    strSql = strSql & " , replacetagcode=N'"& replacetagcode &"' WHERE" & vbcrlf
	strSql = strSql & " targetkey="& targetkey &"" & vbcrlf

    'response.write strSql &"<br>"
	dbget.execute strSql
END IF	

response.write "<script type='text/javascript'>alert('ok');</script>"
session.codePage = 949
Response.write "<script type='text/javascript'>location.replace('"& refer &"');</script>"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->