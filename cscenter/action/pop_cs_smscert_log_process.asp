<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode, idx

mode = requestCheckVar(request("mode"), 32)
idx = requestCheckVar(request("idx"), 32)


dim sqlStr

if (mode<>"reInput") then
	sqlStr = " insert into db_log.dbo.tbl_userConfirm(userid, confDiv, usermail, usercell, smsCD, isConfirm, pFlag, evtFlag) "
	sqlStr = sqlStr + " select userid, confDiv, usermail, usercell, smsCD, isConfirm, pFlag, evtFlag "
	sqlStr = sqlStr + " from db_log.dbo.tbl_userConfirm "
	sqlStr = sqlStr + " where idx = " & idx & " and isConfirm = 'N' and confDate is NULL "
	dbget.Execute sqlStr

	response.write "<script>alert('재전송 되었습니다.');</script>"
    response.write "<script>history.back();</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
