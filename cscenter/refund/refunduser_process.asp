<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim mode
dim menupos
dim strSql
dim idx, userid, defaultCSRefundLimit, useyn, reguserid


mode = requestCheckVar(request("mode"), 32)
menupos = requestCheckVar(request("menupos"), 32)

idx = requestCheckVar(request("idx"), 32)
userid = requestCheckVar(request("userid"), 32)
defaultCSRefundLimit = requestCheckVar(request("defaultCSRefundLimit"), 32)
useyn = requestCheckVar(request("useyn"), 32)

reguserid = session("ssBctId")

if (mode = "modify") then
	strSql = " insert into db_cs.dbo.tbl_cs_refund_user_history(orgidx, userid, defaultCSRefundLimit, useyn, reguserid, regdate) "
	strSql = strSql + " values('" + CStr(idx) + "', '" + CStr(userid) + "', '" + CStr(defaultCSRefundLimit) + "', '" + CStr(useyn) + "', '" + CStr(reguserid) + "', getdate()) "
	rsget.Open strSql, dbget, 1

	strSql = " update db_cs.dbo.tbl_cs_refund_user "
	strSql = strSql + " set userid = '" + CStr(userid) + "' "
	strSql = strSql + " 	, defaultCSRefundLimit = '" + CStr(defaultCSRefundLimit) + "' "
	strSql = strSql + " 	, useyn = '" + CStr(useyn) + "' "
	strSql = strSql + " 	, lastupdate=getdate() "
	strSql = strSql + " where idx = " + CStr(idx) + " "
	''response.write strSql
	rsget.Open strSql, dbget, 1

end if

'==============================================================================
response.write	"<script language='javascript'>" &_
				"	alert('수정되었습니다.'); location.href = 'refunduserlist.asp?menupos=" & menupos & "'; " &_
				"</script>"

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
