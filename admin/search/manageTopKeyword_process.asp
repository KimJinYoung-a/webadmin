<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim topKeyword, modiType, reguserid, useYN, searchCount, idx

topKeyword = requestCheckVar(request("topKeyword"), 32)
modiType = requestCheckVar(request("modiType"), 32)
reguserid = session("ssBctId")
useYN = requestCheckVar(request("useYN"), 32)
searchCount = requestCheckVar(request("searchCount"), 32)
idx = requestCheckVar(request("idx"), 32)



dim SQL
dim mode

mode = requestCheckVar(request("mode"), 32)

if (mode = "add") then

	SQL = " update db_log.dbo.tbl_keyword_top_modi set useYN = 'N' where topKeyword = '" + html2db(CStr(topKeyword)) + "' and useYN = 'Y' "
	rsget.Open SQL, dbget, 1

	SQL = " insert into db_log.dbo.tbl_keyword_top_modi(topKeyword, modiType, reguserid, useYN, regdate, searchCount) "
	SQL = SQL + " values('" + html2db(CStr(topKeyword)) + "', '" + html2db(CStr(modiType)) + "', '" + html2db(CStr(reguserid)) + "', '" + html2db(CStr(useYN)) + "', getdate(), " + CStr(searchCount) + ") "
	rsget.Open SQL, dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('등록되었습니다.'); opener.location.reload(); opener.focus(); window.close(); " &_
					"</script>"

elseif (mode = "del") then

	SQL = " update db_log.dbo.tbl_keyword_top_modi set useYN = 'N' where idx = " + CStr(idx) + " "
	rsget.Open SQL, dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('삭제되었습니다.'); location.href = '" + CStr(refer) + "' " &_
					"</script>"

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
