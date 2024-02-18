<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim orgKeyword, relatedKeyword, modiType, reguserid, useYN, searchCount, idx
dim prect, rect, UserAddCNT

orgKeyword = requestCheckVar(request("orgKeyword"), 32)
relatedKeyword = requestCheckVar(request("relatedKeyword"), 32)
modiType = requestCheckVar(request("modiType"), 32)
reguserid = session("ssBctId")
useYN = requestCheckVar(request("useYN"), 32)
searchCount = requestCheckVar(request("searchCount"), 32)
idx = requestCheckVar(request("idx"), 32)
prect = requestCheckVar(request("prect"), 32)
rect = requestCheckVar(request("rect"), 32)
UserAddCNT = requestCheckVar(request("UserAddCNT"), 32)



dim SQL
dim mode

mode = requestCheckVar(request("mode"), 32)

if (mode = "add") then

	SQL = " update db_log.dbo.tbl_keyword_related_modi set useYN = 'N' where orgKeyword = '" + html2db(CStr(orgKeyword)) + "' and relatedKeyword = '" + html2db(CStr(relatedKeyword)) + "' and useYN = 'Y' "
	rsget.Open SQL, dbget, 1

	SQL = " insert into db_log.dbo.tbl_keyword_related_modi(orgKeyword, relatedKeyword, modiType, reguserid, useYN, regdate, searchCount) "
	SQL = SQL + " values('" + html2db(CStr(orgKeyword)) + "', '" + html2db(CStr(relatedKeyword)) + "', '" + html2db(CStr(modiType)) + "', '" + html2db(CStr(reguserid)) + "', '" + html2db(CStr(useYN)) + "', getdate(), " + CStr(searchCount) + ") "
	rsget.Open SQL, dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('등록되었습니다.'); opener.location.reload(); opener.focus(); window.close(); " &_
					"</script>"

elseif (mode = "del") then

	SQL = " update db_log.dbo.tbl_keyword_related_modi set useYN = 'N' where idx = " + CStr(idx) + " "
	rsget.Open SQL, dbget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('삭제되었습니다.'); location.href = '" + CStr(refer) + "' " &_
					"</script>"

elseif (mode = "delevt") then

	SQL = " update [db_EVT].[dbo].[tbl_keywords_Relate] set isUsingType = 0, isAutoType = 2, lastupdate = getdate() where prect = '" & prect & "' and rect = '" & rect & "' "
	rsEVTget.Open SQL, dbEVTget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('삭제되었습니다.'); location.href = '" + CStr(refer) + "' " &_
					"</script>"

elseif (mode = "useevt") then

	SQL = " update [db_EVT].[dbo].[tbl_keywords_Relate] set isUsingType = 1, lastupdate = getdate() where prect = '" & prect & "' and rect = '" & rect & "' "
	rsEVTget.Open SQL, dbEVTget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('사용전환 되었습니다.'); location.href = '" + CStr(refer) + "' " &_
					"</script>"

elseif (mode = "modievtUserAddCNT") then

	SQL = " update [db_EVT].[dbo].[tbl_keywords_Relate] set UserAddCNT = " & UserAddCNT & ", lastupdate = getdate() where prect = '" & prect & "' and rect = '" & rect & "' "
	rsEVTget.Open SQL, dbEVTget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('수정 되었습니다.'); opener.focus(); opener.location.reload(); window.close() " &_
					"</script>"

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
