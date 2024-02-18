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

if (mode = "delevt") then

	SQL = " update [db_EVT].[dbo].[tbl_keywords_autoComplete] set isUsingType = 0, isAutoType = 2, lastupdate = getdate() where rect = '" & rect & "' "
	rsEVTget.Open SQL, dbEVTget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('삭제되었습니다.'); location.href = '" + CStr(refer) + "' " &_
					"</script>"

elseif (mode = "useevt") then

	SQL = " update [db_EVT].[dbo].[tbl_keywords_autoComplete] set isUsingType = 1, lastupdate = getdate() where rect = '" & rect & "' "
	rsEVTget.Open SQL, dbEVTget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('사용전환 되었습니다.'); location.href = '" + CStr(refer) + "' " &_
					"</script>"

elseif (mode = "modievtUserAddCNT") then

	SQL = " update [db_EVT].[dbo].[tbl_keywords_autoComplete] set UserAddCNT = " & UserAddCNT & ", lastupdate = getdate() where rect = '" & rect & "' "
	rsEVTget.Open SQL, dbEVTget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('수정 되었습니다.'); opener.focus(); opener.location.reload(); window.close() " &_
					"</script>"

elseif (mode = "add") then

	SQL = " IF NOT EXISTS (select top 1 rect from [db_EVT].[dbo].[tbl_keywords_autoComplete] where rect = '" & html2db(rect) & "') "
	SQL = SQL + " BEGIN "
	SQL = SQL + " 	insert into [db_EVT].[dbo].[tbl_keywords_autoComplete](rect) "
	SQL = SQL + " 	select '" & html2db(rect) & "' "
	SQL = SQL + " END "
	rsEVTget.Open SQL, dbEVTget, 1

	response.write	"<script language='javascript'>" &_
					"	alert('등록 되었습니다.'); opener.focus(); opener.location.reload(); window.close() " &_
					"</script>"

end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
