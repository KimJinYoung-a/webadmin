<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : LMS발송관리
' Hieditor : 2020.03.19 한용민 생성
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
dim tidx, sendmethod, template_code, template_name, contents, button_name, button_url_mobile, button_name2, button_url_mobile2
dim failed_type, failed_subject, failed_msg, isusing, adminid, sortno
    tidx = requestCheckVar(getNumeric(trim(Request("tidx"))),10)
    sendmethod = requestCheckVar(trim(Request("sendmethod")),16)
    template_code = requestCheckVar(trim(Request("template_code")),32)
    template_name = requestCheckVar(trim(Request("template_name")),64)
    contents = requestCheckVar(trim(Request("contents")),1000)
    button_name = requestCheckVar(trim(Request("button_name")),64)
    button_url_mobile = requestCheckVar(trim(Request("button_url_mobile")),256)
    button_name2 = requestCheckVar(trim(Request("button_name2")),64)
    button_url_mobile2 = requestCheckVar(trim(Request("button_url_mobile2")),256)
    failed_type = requestCheckVar(trim(Request("failed_type")),3)
    failed_subject = requestCheckVar(trim(Request("failed_subject")),50)
    failed_msg = requestCheckVar(trim(Request("failed_msg")),1000)
    isusing = requestCheckVar(trim(Request("isusing")),1)
    sortno = requestCheckVar(getNumeric(trim(Request("sortno"))),10)
    adminid = session("ssBctId")
	sMode		= requestCheckVar(trim(Request("sM")),1)
    menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)

if sortno="" or isnull(sortno) then sortno=0

refer = request.ServerVariables("HTTP_REFERER")

if sendmethod="" or isnull(sendmethod) then
    response.write "<script type='text/javascript'>"
    response.write "	alert('발송방법을 입력해 주세요.');"
    session.codePage = 949
    response.write "	history.back();"
    response.write "</script>"
    dbget.close()	:	response.End
end if
sendmethod = replace(sendmethod,vbcrlf,"")
if template_code="" or isnull(template_code) then
    response.write "<script type='text/javascript'>"
    response.write "	alert('템플릿코드를 입력해 주세요.');"
    session.codePage = 949
    response.write "	history.back();"
    response.write "</script>"
    dbget.close()	:	response.End
end if
template_code = replace(template_code,vbcrlf,"")
if template_name="" or isnull(template_name) then
    response.write "<script type='text/javascript'>"
    response.write "	alert('템플릿명을 입력해 주세요.');"
    session.codePage = 949
    response.write "	history.back();"
    response.write "</script>"
    dbget.close()	:	response.End
end if
template_name = replace(template_name,vbcrlf,"")
if contents="" or isnull(contents) then
    response.write "<script type='text/javascript'>"
    response.write "	alert('내용을 입력해 주세요.');"
    session.codePage = 949
    response.write "	history.back();"
    response.write "</script>"
    dbget.close()	:	response.End
end if
'contents = replace(contents,vbcrlf,"\n")

if checkNotValidHTML(contents) then
    response.write "<script type='text/javascript'>"
    response.write "	alert('내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
    session.codePage = 949
    response.write "	history.back();"
    response.write "</script>"
    dbget.close()	:	response.End
end if
if button_name="" or isnull(button_name) then
    ' response.write "<script type='text/javascript'>"
    ' response.write "	alert('카카오톡 버튼 이름을 입력해 주세요.');"
    ' session.codePage = 949
    ' response.write "	history.back();"
    ' response.write "</script>"
    ' dbget.close()	:	response.End
else
    button_name = replace(button_name,vbcrlf,"")
end if
if button_url_mobile="" or isnull(button_url_mobile) then
    ' response.write "<script type='text/javascript'>"
    ' response.write "	alert('카카오톡 버튼 모바일 주소를 입력해 주세요.');"
    ' session.codePage = 949
    ' response.write "	history.back();"
    ' response.write "</script>"
    ' dbget.close()	:	response.End
else
    button_url_mobile = replace(button_url_mobile,vbcrlf,"")
end if
if failed_subject<>"" and not(isnull(failed_subject)) then
    if len(failed_subject)>50 then
        response.write "<script type='text/javascript'>"
        response.write "	alert('카카오톡 실패시 문자제목이 제한길이를 초과하였습니다. 50자 까지 작성 가능합니다.');"
        session.codePage = 949
        response.write "	history.back();"
        response.write "</script>"
        dbget.close()	:	response.End
    end if
    failed_subject = replace(failed_subject,vbcrlf,"")
end if
if failed_msg<>"" and not(isnull(failed_msg)) then
    if checkNotValidHTML(failed_msg) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('카카오톡 실패시 문자내용에 유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
        session.codePage = 949
        response.write "	history.back();"
        response.write "</script>"
        dbget.close()	:	response.End
    end if
end if

IF sMode = "I" THEN
    strSql = "SELECT sendmethod, template_code FROM db_contents.dbo.tbl_lms_template with (readuncommitted) Where sendmethod=N'"& sendmethod &"' and template_code=N'"& html2db(template_code) &"'"

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

	strSql = " INSERT INTO db_contents.dbo.tbl_lms_template (" & vbcrlf
	strSql = strSql & " sendmethod, template_code, template_name, contents" & vbcrlf
    strSql = strSql & " , button_name, button_url_mobile, button_name2, button_url_mobile2, failed_type, failed_subject" & vbcrlf
    strSql = strSql & " , failed_msg, isusing, regadminid, lastadminid, regdate, lastupdate, sortno" & vbcrlf
    strSql = strSql & " ) Values(" & vbcrlf
    strSql = strSql & " N'"& sendmethod &"',N'"& html2db(template_code) &"',N'"& html2db(template_name) &"',N'"& html2db(contents) &"'" & vbcrlf
    strSql = strSql & " ,N'"& html2db(button_name) &"',N'"& html2db(button_url_mobile) &"',N'"& html2db(button_name2) &"',N'"& html2db(button_url_mobile2) &"'" & vbcrlf
    strSql = strSql & " ,N'"& failed_type &"',N'"& html2db(failed_subject) &"',N'"& html2db(failed_msg) &"',N'"& isusing &"',N'"& adminid &"'" & vbcrlf
    strSql = strSql & " ,N'"& adminid &"',getdate(),getdate(),"& sortno &"" & vbcrlf
    strSql = strSql & " ) " & vbcrlf

    'response.write strSql &"<br>"
	dbget.execute strSql
				
ELSEIF sMode="U" THEN
    if tidx="" or isnull(tidx) then
        response.write "<script type='text/javascript'>"
        response.write "	alert('템플릿번호가 없습니다.');"
        session.codePage = 949
        response.write "	history.back();"
        response.write "</script>"
        dbget.close()	:	response.End
    end if

    strSql =" UPDATE db_contents.dbo.tbl_lms_template" & vbcrlf
    strSql = strSql & " Set sendmethod = N'"& sendmethod &"'" & vbcrlf
    strSql = strSql & " , template_code = N'"& html2db(template_code) &"'" & vbcrlf
    strSql = strSql & " , template_name = N'"& html2db(template_name) &"'" & vbcrlf
    strSql = strSql & " , contents = N'"& html2db(contents) &"'" & vbcrlf
    strSql = strSql & " , button_name = N'"& html2db(button_name) &"'" & vbcrlf
    strSql = strSql & " , button_url_mobile = N'"& html2db(button_url_mobile) &"'" & vbcrlf
    strSql = strSql & " , button_name2 = N'"& html2db(button_name2) &"'" & vbcrlf
    strSql = strSql & " , button_url_mobile2 = N'"& html2db(button_url_mobile2) &"'" & vbcrlf
    strSql = strSql & " , failed_type = N'"& failed_type &"'" & vbcrlf
    strSql = strSql & " , failed_subject = N'"& html2db(failed_subject) &"'" & vbcrlf
    strSql = strSql & " , failed_msg = N'"& html2db(failed_msg) &"'" & vbcrlf
    strSql = strSql & " , isusing = N'"& isusing &"'" & vbcrlf
    strSql = strSql & " , lastadminid = N'"& adminid &"'" & vbcrlf
    strSql = strSql & " , lastupdate =getdate()" & vbcrlf
    strSql = strSql & " , sortno ="& sortno &" WHERE " & vbcrlf
	strSql = strSql & " tidx="& tidx &"" & vbcrlf

    'response.write strSql &"<br>"
	dbget.execute strSql
ELSE
    response.write "<script type='text/javascript'>"
    response.write "	alert('정상적인 경로가 아닙니다.');"
    session.codePage = 949
    response.write "	history.back();"
    response.write "</script>"
    dbget.close()	:	response.End
END IF

response.write "<script type='text/javascript'>alert('ok');</script>"
session.codePage = 949
Response.write "<script type='text/javascript'>location.replace('"& refer &"');</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->