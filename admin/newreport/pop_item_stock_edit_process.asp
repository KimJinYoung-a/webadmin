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

	Response.Write "<script>alert('����Ǿ����ϴ�.'); opener.reload(); opener.focus(); window.close();</script>"
elseif (mode="updatelastipgo") then

	'// ������, 2015-06-15, skyer9
	Response.Write "<script>alert('\n\n!!! ��� �޴� !!!\n\n�ý����� ���� : ���� ����Դϴ�.');</script>"
	Response.Write "<br><br>!!! ��� �޴� !!!<br><br>�ý����� ���� : ���� ����Դϴ�."
	Response.end

    strSql = " update "
	strSql = strSql + " [db_summary].[dbo].tbl_monthly_accumulated_logisstock_summary "
	strSql = strSql + " set lastipgodate = '" + CStr(yyyymm) + "' "
	strSql = strSql + " where Itemgubun = '" + CStr(Itemgubun) + "' and  itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "' and lastipgodate is NULL "
    strSql = strSql + " and yyyymm>='"&yyyymm&"'"
    'Response.Write strSql

    dbget.Execute strSql
    Response.Write "<script>alert('����Ǿ����ϴ�.'); window.close();</script>"
else
	Response.Write "<script>alert('�߸��� �����Դϴ�..');</script>"
end if

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
