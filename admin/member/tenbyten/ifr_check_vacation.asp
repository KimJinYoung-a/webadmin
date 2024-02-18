<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<%

dim empno

dim strSql
dim mode, startday, endday, totWorkDay


mode=requestCheckVar(Request("mode"),32)
empno=requestCheckVar(Request("empno"),32)
startday=requestCheckVar(Request("startday"),15)
endday=requestCheckVar(Request("endday"),15)


if (mode = "checkparthour") then
	totWorkDay = 0
	strSql = " exec [db_partner].[dbo].[usp_Ten_user_tenbyten_GetVacationHour] '" + CStr(empno) + "', '" + CStr(startday) + "', '" + CStr(endday) + "' "
	rsget.Open strSql,dbget,1
		totWorkDay = rsget(0)
	rsget.Close

	Response.Write "<script language=javascript>" &_
			"	parent.jsReActFromIframe('" + CStr(totWorkDay) + "');" &_
			"</script>"

else
	'
end if

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
