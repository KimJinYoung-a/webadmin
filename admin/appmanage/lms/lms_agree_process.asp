<%@ codepage="65001" language="VBScript" %>
<% option Explicit %>
<%
session.codepage = 65001
response.Charset="UTF-8"
%>
<%
'###########################################################
' Description : LMS/친구톡/알림톡 수신동의 관리
' Hieditor : 2021.08.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib_utf8.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheaderUTF8.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/offshop_function_utf8.asp"-->
<!-- #include virtual="/lib/classes/appmanage/lms/lms_msg_cls.asp" -->
<%
dim userid, regdate, lastupdate, reguserid, lastuserid, kakaoalrimyn
dim adminuserid, i, menupos, sql, mode
	userid = requestcheckvar(trim(request("userid")),32)
	kakaoalrimyn = requestcheckvar(trim(request("kakaoalrimyn")),1)
	menupos = requestcheckvar(getNumeric(trim(request("menupos"))),10)
	mode = requestcheckvar(trim(request("mode")),32)

adminuserid=session("ssBctId")									
    
if mode = "lms_agree_edit" then
	if userid="" or isnull(userid) then
		Response.Write "<script type='text/javascript'>"
        Response.Write "    alert('아이디가 없습니다.');"
        session.codePage = 949
        Response.Write "    history.go(-1);"
        Response.Write "</script>"
		dbget.close() : Response.End
	end if
	if kakaoalrimyn="" or isnull(kakaoalrimyn) then
		Response.Write "<script type='text/javascript'>"
        Response.Write "    alert('알림톡 수신여부를 선택해 주세요.');"
        session.codePage = 949
        Response.Write "    history.go(-1);"
        Response.Write "</script>"
		dbget.close() : Response.End
	end if

    sql = "if exists(select top 1 userid from db_contents.[dbo].[tbl_lms_agree] with (nolock) where userid=N'"& userid &"')"
    sql = sql & " begin"
    sql = sql & "   update db_contents.[dbo].[tbl_lms_agree]" + vbcrlf
    sql = sql & "   set lastupdate=getdate()" + vbcrlf
    sql = sql & "   , lastuserid=N'"& adminuserid &"'" + vbcrlf
    sql = sql & "   , kakaoalrimyn=N'"& kakaoalrimyn &"' where" + vbcrlf
    sql = sql & "   userid=N'"& userid &"'"
    sql = sql & " end"
    sql = sql & " else"
    sql = sql & " begin"
    sql = sql & "   insert into db_contents.[dbo].[tbl_lms_agree](" + vbcrlf
    sql = sql & "   userid, regdate, lastupdate, reguserid, lastuserid, kakaoalrimyn)" + vbcrlf
    sql = sql & "       select u.userid, getdate(), getdate(), N'"& adminuserid &"', N'"& adminuserid &"', N'"& kakaoalrimyn &"'" + vbcrlf
    sql = sql & "       from db_user.dbo.tbl_user_n u with (nolock)"
    sql = sql & "       where userid=N'"& userid &"'"
    sql = sql & " end"

    'response.write sql & "<Br>"
    dbget.execute sql
else
    Response.Write "<script type='text/javascript'>"
    Response.Write "    alert('구분자가 없습니다.');"
    session.codePage = 949
    Response.Write "    history.go(-1);"
    Response.Write "</script>"
    dbget.close() : Response.End
end if
%>

<!-- #include virtual="/admin/lib/poptail.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<%
session.codePage = 949
%>
<script type='text/javascript'>
	opener.location.reload();
	self.close();
</script>
