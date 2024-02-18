<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%

dim mode
dim itemgubun, itemid, itemoption
dim yyyymm, lastmwdiv
dim strSql

mode = request("mode")
itemgubun = request("itemgubun")
itemid = request("itemid")
itemoption = request("itemoption")
yyyymm = request("yyyymm")
lastmwdiv = request("lastmwdiv")


if (mode = "updatelastmwdiv") then

	strSql = " update "
	strSql = strSql + " [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary "
	strSql = strSql + " set lastmwdiv = '" + CStr(lastmwdiv) + "' "
	strSql = strSql + " where Itemgubun = '" + CStr(Itemgubun) + "' and  itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' and lastmwdiv is NULL "

	if (yyyymm <> "all") then
		strSql = strSql + " and yyyymm = '" + CStr(yyyymm) + "' "
	end if
	'Response.Write strSql

	rsget.Open strSql,dbget,1

	Response.Write "<script>alert('저장되었습니다.'); opener.reload(); opener.focus(); window.close();</script>"
elseif (mode="updatelastipgo") then

	'// 사용안함, 2015-06-15, skyer9
	Response.Write "<script>alert('\n\n!!! 폐기 메뉴 !!!\n\n시스템팀 문의 : 폐기된 기능입니다.');</script>"
	Response.Write "<br><br>!!! 폐기 메뉴 !!!<br><br>시스템팀 문의 : 폐기된 기능입니다."
	Response.end

    strSql = " update "
	strSql = strSql + " [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary "
	strSql = strSql + " set lastipgodate = '" + CStr(yyyymm) + "' "
	strSql = strSql + " where Itemgubun = '" + CStr(Itemgubun) + "' and  itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' and lastipgodate is NULL "
    strSql = strSql + " and yyyymm>='"&yyyymm&"'"
    'Response.Write strSql

    dbget.Execute strSql
    Response.Write "<script>alert('저장되었습니다.'); window.close();</script>"
else
	Response.Write "<script>alert('잘못된 접근입니다..');</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
