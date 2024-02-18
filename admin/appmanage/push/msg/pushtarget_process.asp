<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : 푸시 타켓 관리
' Hieditor : 2019.06.17 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/push/apppush_msg_cls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<%
dim iResult, sMode, strSql, refer, replacetagcode
dim targetKey,targetName,targetQuery,isusing,repeatpushyn
	sMode		= requestCheckVar(Request("sM"),1)
    targetKey		= requestCheckVar(Request("targetKey"),10)
    targetName		= requestCheckVar(Request("targetName"),100)
    targetQuery		= Request("targetQuery")
    isusing		= requestCheckVar(Request("isusing"),1)
    repeatpushyn		= requestCheckVar(Request("repeatpushyn"),1)
    replacetagcode		= requestCheckVar(trim(Request("replacetagcode")),256)

refer = request.ServerVariables("HTTP_REFERER")

IF sMode = "I" THEN
	targetQuery = replace(targetQuery,"'", "''")

    strSql = "SELECT targetKey FROM db_contents.[dbo].[tbl_app_targetQuery] with (nolock) Where targetKey="& targetKey

    'response.write strSql &"<br>"		
    rsget.CursorLocation = adUseClient
    rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
	IF not (rsget.eof or rsget.bof) then
		Call sbAlertMsg ("이미 존재하는 코드값입니다.다시 등록해주세요", "back", "") 
		dbget.close()	:	response.End
	end if
	rsget.close	
	
	strSql = " INSERT INTO db_contents.[dbo].[tbl_app_targetQuery] (targetKey, targetName, targetQuery, isusing, repeatpushyn, target_procedureyn"
    strSql = strSql & " , replacetagcode" & vbcrlf
    strSql = strSql & " ) Values(" & vbcrlf
	strSql = strSql & " "& targetKey &",'"& targetName &"','"& targetQuery &"','"& isusing &"','"& repeatpushyn &"',N'N'" & vbcrlf
    strSql = strSql & " ,N'"& replacetagcode &"') " & vbcrlf

    'response.write strSql &"<br>"
	dbget.execute strSql
				
ELSEIF sMode="U" THEN	
	targetQuery = replace(targetQuery,"'", "''")

    strSql =" UPDATE db_contents.[dbo].[tbl_app_targetQuery]" & vbcrlf
    strSql = strSql & " Set targetName = '"& targetName &"', targetQuery = '"& targetQuery &"' " & vbcrlf
    strSql = strSql & " , isusing ='"& isusing &"', repeatpushyn ='"& repeatpushyn &"'" & vbcrlf
    strSql = strSql & " , target_procedureyn =N'N'" & vbcrlf
    strSql = strSql & " , replacetagcode=N'"& replacetagcode &"' WHERE" & vbcrlf
	strSql = strSql & " targetKey="& targetKey &"" & vbcrlf

    'response.write strSql &"<br>"
	dbget.execute strSql
END IF	
	
response.write "<script type='text/javascript'>alert('ok'); location.replace('"& refer &"');</script>"	
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->