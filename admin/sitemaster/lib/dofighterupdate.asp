<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim refer,mode,idx
refer = request.ServerVariables("HTTP_REFERER")
mode = request("mode")
idx = request("idx")


%>


<%
dim sqlstr

	 sqlstr = "update [db_sitemaster].[dbo].tbl_design_fighter" + vbcrlf
	 sqlstr = sqlstr + " set sellcnt1 = T.cnt" + vbcrlf
	 sqlstr = sqlstr + " from (" + vbcrlf
	 sqlstr = sqlstr + " select d.itemid, sum(d.itemno) as cnt" + vbcrlf
	 sqlstr = sqlstr + " from [db_order].[dbo].tbl_order_master m," + vbcrlf
	 sqlstr = sqlstr + " [db_order].[dbo].tbl_order_detail d" + vbcrlf
	 sqlstr = sqlstr + " where m.orderserial = d.orderserial" + vbcrlf
	 sqlstr = sqlstr + " and m.ipkumdiv > 3" + vbcrlf
	 sqlstr = sqlstr + " and m.cancelyn = 'N'" + vbcrlf
	 sqlstr = sqlstr + " and d.cancelyn <> 'Y'" + vbcrlf
	 sqlstr = sqlstr + " and datediff(ww,m.regdate,getdate()) <= 1" + vbcrlf
	 sqlstr = sqlstr + " and d.itemid in (select top 1 itemid1 from [db_sitemaster].[dbo].tbl_design_fighter where idx='" &idx & "'" + vbcrlf
	 sqlstr = sqlstr + " order by idx desc)" + vbcrlf
	 sqlstr = sqlstr + " group by d.itemid" + vbcrlf
	 sqlstr = sqlstr + " ) as T" + vbcrlf
	 sqlstr = sqlstr + " where [db_sitemaster].[dbo].tbl_design_fighter.itemid1 = T.itemid" + vbcrlf

	 sqlstr = sqlstr + "update [db_sitemaster].[dbo].tbl_design_fighter" + vbcrlf
	 sqlstr = sqlstr + " set sellcnt2 = T.cnt" + vbcrlf
	 sqlstr = sqlstr + " from (" + vbcrlf
	 sqlstr = sqlstr + " select d.itemid, sum(d.itemno) as cnt" + vbcrlf
	 sqlstr = sqlstr + " from [db_order].[dbo].tbl_order_master m," + vbcrlf
	 sqlstr = sqlstr + " [db_order].[dbo].tbl_order_detail d" + vbcrlf
	 sqlstr = sqlstr + " where m.orderserial = d.orderserial" + vbcrlf
	 sqlstr = sqlstr + " and m.ipkumdiv > 3" + vbcrlf
	 sqlstr = sqlstr + " and m.cancelyn = 'N'" + vbcrlf
	 sqlstr = sqlstr + " and d.cancelyn <> 'Y'" + vbcrlf
	 sqlstr = sqlstr + " and datediff(ww,m.regdate,getdate()) <= 1" + vbcrlf
	 sqlstr = sqlstr + " and d.itemid in (select top 1 itemid2 from [db_sitemaster].[dbo].tbl_design_fighter where idx='" &idx & "'" + vbcrlf
	 sqlstr = sqlstr + " order by idx desc)" + vbcrlf
	 sqlstr = sqlstr + " group by d.itemid" + vbcrlf
	 sqlstr = sqlstr + " ) as T" + vbcrlf
	 sqlstr = sqlstr + " where [db_sitemaster].[dbo].tbl_design_fighter.itemid2 = T.itemid" + vbcrlf

	 sqlstr = sqlstr + "update [db_sitemaster].[dbo].tbl_design_fighter" + vbcrlf
	 sqlstr = sqlstr + " set wishcnt1 = T.cnt" + vbcrlf
	 sqlstr = sqlstr + " from (" + vbcrlf
	 sqlstr = sqlstr + " select itemid, count(itemid) as cnt" + vbcrlf
	 sqlstr = sqlstr + " from [db_my10x10].[dbo].tbl_myfavorite" + vbcrlf
	 sqlstr = sqlstr + " where datediff(ww,regdate,getdate()) <= 1" + vbcrlf
	 sqlstr = sqlstr + " and itemid in (select top 1 itemid1 from [db_sitemaster].[dbo].tbl_design_fighter where idx='" &idx & "'" + vbcrlf
	 sqlstr = sqlstr + " order by idx desc)" + vbcrlf
	 sqlstr = sqlstr + " group by itemid" + vbcrlf
	 sqlstr = sqlstr + " ) as T" + vbcrlf
	 sqlstr = sqlstr + " where [db_sitemaster].[dbo].tbl_design_fighter.itemid1 = T.itemid" + vbcrlf

	 sqlstr = sqlstr + "update [db_sitemaster].[dbo].tbl_design_fighter" + vbcrlf
	 sqlstr = sqlstr + " set wishcnt2 = T.cnt" + vbcrlf
	 sqlstr = sqlstr + " from (" + vbcrlf
	 sqlstr = sqlstr + " select itemid, count(itemid) as cnt" + vbcrlf
	 sqlstr = sqlstr + " from [db_my10x10].[dbo].tbl_myfavorite" + vbcrlf
	 sqlstr = sqlstr + " where datediff(ww,regdate,getdate()) <= 1" + vbcrlf
	 sqlstr = sqlstr + " and itemid in (select top 1 itemid2 from [db_sitemaster].[dbo].tbl_design_fighter where idx='" &idx & "'" + vbcrlf
	 sqlstr = sqlstr + " order by idx desc)" + vbcrlf
	 sqlstr = sqlstr + " group by itemid" + vbcrlf
	 sqlstr = sqlstr + " ) as T" + vbcrlf
	 sqlstr = sqlstr + " where [db_sitemaster].[dbo].tbl_design_fighter.itemid2 = T.itemid" + vbcrlf
'	 response.write sqlstr
'	 dbget.close()	:	response.End
	rsget.Open sqlStr,dbget,1

%>

<script language="javascript">
alert('반영 되었습니다.');
opener.location.reload();
window.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->