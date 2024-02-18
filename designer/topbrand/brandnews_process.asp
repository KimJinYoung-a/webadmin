<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim sqlStr, i
dim mode, idx, title, contents, userid, menupos

mode = html2db(request("mode"))
idx = html2db(request("idx"))
title = html2db(request("title"))
contents = html2db(request("contents"))
userid = session("ssBctId")
menupos = html2db(request("menupos"))


if (mode = "write") then
    sqlStr = " insert into [db_brand].[dbo].tbl_topbrand_news(makerid, title, contents, regdate, isusing) "
    sqlStr = sqlStr + " values('" + userid + "', '" + title + "', '" + contents + "', getdate(), 'Y') "
    rsget.Open sqlStr, dbget, 1
elseif (mode = "modify") then
    sqlStr = " update [db_brand].[dbo].tbl_topbrand_news set title = '" + title + "', contents = '" + contents + "' "
    sqlStr = sqlStr + " where idx = " + CStr(idx) + " and makerid = '" + userid + "' "
    rsget.Open sqlStr, dbget, 1
elseif (mode = "delete") then
    sqlStr = " update [db_brand].[dbo].tbl_topbrand_news set isusing = 'N' "
    sqlStr = sqlStr + " where idx = " + CStr(idx) + " and makerid = '" + userid + "' "
    rsget.Open sqlStr, dbget, 1
else
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if

response.write "<script>alert('저장되었습니다.'); location.href = 'brandnews_list.asp?menupos=" + CStr(menupos) + "';</script>"
dbget.close()	:	response.End

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->