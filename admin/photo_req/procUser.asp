<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/photo_req/requestCls.asp"-->
<%
Dim sMode
Dim uid, seltype, uname, isusing
Dim strSql

sMode		= requestCheckVar(Request("sM"),1)
uid 	= requestCheckVar(Request("uid"),10)
seltype 	= requestCheckVar(Request("seltype"),20)
uname 	= requestCheckVar(Request("uname"),20)
isusing 	= requestCheckVar(Request("isusing"),1)

IF sMode = "I" THEN
	strSql = "SELECT  user_id FROM [db_partner].[dbo].[tbl_photo_user] Where user_type='"&seltype&"' and user_ID='"&uid&"'"
	rsget.Open strSql,dbget
	IF not (rsget.eof or rsget.bof) then
		Response.Write "<script>alert('이미 작업자 리스트에 포함되어있습니다.');history.back();</script>"
		dbget.close()
		response.End
	end if
	rsget.close

	strSql = " INSERT INTO [db_partner].[dbo].[tbl_photo_user] (user_type, user_id, user_name, user_useyn, user_regdate)"&_
			" Values('"&seltype&"','"&uid&"','"&uname&"','"&isusing&"',getdate()) "
	dbget.execute strSql

ELSEIF sMode="U" THEN
	strSql =" UPDATE [db_partner].[dbo].[tbl_photo_user] Set user_name = '"&uname&"', user_useyn ='"&isusing&"'"&_
			" WHERE user_type ='"&seltype&"' and user_id='"&uid&"'"
	dbget.execute strSql
END IF

	Response.Write "<script>alert('처리되었습니다.');location.href='/admin/photo_req/popUserList.asp?selCT="&seltype&"';</script>"
	dbget.close()
	Response.End
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->