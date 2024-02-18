<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/xmlhttpUtil.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode, idx, xSiteId, gubun, makerid, comment, reguserid

mode = requestCheckvar(request("mode"),32)
idx = requestCheckvar(request("idx"),32)
xSiteId = requestCheckvar(request("xSiteId"),32)
gubun = requestCheckvar(request("gubun"),32)
makerid = requestCheckvar(request("makerid"),32)
comment = html2db(requestCheckvar(request("comment"),200))

reguserid = session("ssBctId")


dim sqlStr

if (mode="ins") then

	sqlStr = " update db_partner.dbo.tbl_xSite_BrandInfo "
	sqlStr = sqlStr + " set useyn = 'N' where useyn = 'Y' and xSiteId = '" + CStr(xSiteId) + "' and gubun = '" + CStr(gubun) + "' and makerid = '" + CStr(makerid) + "' "
	dbget.Execute sqlStr

	sqlStr = " insert into db_partner.dbo.tbl_xSite_BrandInfo(xSiteId, makerid, gubun, comment, reguserid) "
	sqlStr = sqlStr + " values('" + CStr(xSiteId) + "', '" + CStr(makerid) + "', '" + CStr(gubun) + "', '" + CStr(comment) + "', '" + CStr(reguserid) + "') "
	dbget.Execute sqlStr

	response.write "<script>alert('저장되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

elseif (mode="del") then

	sqlStr = " update db_partner.dbo.tbl_xSite_BrandInfo "
	sqlStr = sqlStr + " set useyn = 'N' where useyn = 'Y' and idx = " + CStr(idx) + " "
	dbget.Execute sqlStr

	response.write "<script>alert('삭제되었습니다.'); opener.location.reload(); opener.focus(); window.close();</script>"
	dbget.close()
	response.end

end if

%>
<script language='javascript'>
alert('저장되었습니다.');
location.replace('<%= refer %>');
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
