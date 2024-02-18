<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim strSql
dim mode
dim inoutidx, matchMemo, matchstate

mode = requestCheckVar(request("mode"), 32)
inoutidx = requestCheckVar(request("inoutidx"), 32)
matchMemo = requestCheckVar(request("matchMemo"), 100)
matchstate = requestCheckVar(request("matchstate"), 100)


if (mode = "insMatchMemo") then
	'==============================================================================
	strSql = " update db_log.dbo.tbl_IBK_ISS_ACCT_INOUT "
	strSql = strSql + " set matchmemo = '" + html2db(CStr(matchMemo)) + "' "
	strSql = strSql + " where inoutidx = " + CStr(inoutidx) + " and matchmemo is NULL "
	rsget.Open strSql, dbget, 1

	if (matchstate <> "") then
		if (matchstate = "X") then
			strSql = " update db_log.dbo.tbl_IBK_ISS_ACCT_INOUT "
			strSql = strSql + " set matchstate = 'X' "
			strSql = strSql + " where inoutidx = " + CStr(inoutidx) + " and matchstate is NULL "
			rsget.Open strSql, dbget, 1
		elseif (matchstate = "D") then
			strSql = " update db_log.dbo.tbl_IBK_ISS_ACCT_INOUT "
			strSql = strSql + " set matchstate = NULL "
			strSql = strSql + " where inoutidx = " + CStr(inoutidx) + " and IsNull(matchstate, 'N') = 'X' "
			rsget.Open strSql, dbget, 1
		end if
	end if

	response.write	"<script language='javascript'>" &_
					"	alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close(); " &_
					"</script>"

elseif (mode = "modMatchMemo") then
	'==============================================================================
	strSql = " update db_log.dbo.tbl_IBK_ISS_ACCT_INOUT "
	strSql = strSql + " set matchmemo = '" + html2db(CStr(matchMemo)) + "' "
	strSql = strSql + " where inoutidx = " + CStr(inoutidx) + " and matchmemo is not NULL "
	rsget.Open strSql, dbget, 1

	if (matchstate <> "") then
		if (matchstate = "X") then
			strSql = " update db_log.dbo.tbl_IBK_ISS_ACCT_INOUT "
			strSql = strSql + " set matchstate = 'X' "
			strSql = strSql + " where inoutidx = " + CStr(inoutidx) + " and matchstate is NULL "
			rsget.Open strSql, dbget, 1
		elseif (matchstate = "D") then
			strSql = " update db_log.dbo.tbl_IBK_ISS_ACCT_INOUT "
			strSql = strSql + " set matchstate = NULL "
			strSql = strSql + " where inoutidx = " + CStr(inoutidx) + " and IsNull(matchstate, 'N') = 'X' "
			rsget.Open strSql, dbget, 1
		end if
	end if

	response.write	"<script language='javascript'>" &_
					"	alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close(); " &_
					"</script>"

elseif (mode = "delMatchMemo") then
	'==============================================================================
	strSql = " update db_log.dbo.tbl_IBK_ISS_ACCT_INOUT "
	strSql = strSql + " set matchmemo = NULL "
	strSql = strSql + " where inoutidx = " + CStr(inoutidx) + " and matchmemo is not NULL "
	rsget.Open strSql, dbget, 1

	if (matchstate <> "") then
		if (matchstate = "X") then
			strSql = " update db_log.dbo.tbl_IBK_ISS_ACCT_INOUT "
			strSql = strSql + " set matchstate = 'X' "
			strSql = strSql + " where inoutidx = " + CStr(inoutidx) + " and matchstate is NULL "
			rsget.Open strSql, dbget, 1
		elseif (matchstate = "D") then
			strSql = " update db_log.dbo.tbl_IBK_ISS_ACCT_INOUT "
			strSql = strSql + " set matchstate = NULL "
			strSql = strSql + " where inoutidx = " + CStr(inoutidx) + " and IsNull(matchstate, 'N') = 'X' "
			rsget.Open strSql, dbget, 1
		end if
	end if

	response.write	"<script language='javascript'>" &_
					"	alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close(); " &_
					"</script>"

else
	response.write "잘못된 접근입니다."
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
