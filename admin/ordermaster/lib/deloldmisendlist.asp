<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
dim mode
mode = request("mode")
dim sqlStr

1
'// 에러낸다. 2016-12-15, skyer9
'// 폐기예정

if mode="deloldmisendlist" then

	'// 폐기기능, 2016-12-15, skyer9
	response.end

	sqlStr = "delete from [db_temp].[dbo].tbl_mibeasong_list"
	sqlStr = sqlStr + " where idx in ("
	sqlStr = sqlStr + "     select l.idx from [db_temp].[dbo].tbl_mibeasong_list l"
	sqlStr = sqlStr + "     left join [db_order].[dbo].tbl_order_master m"
	sqlStr = sqlStr + "     on l.orderserial=m.orderserial"
	sqlStr = sqlStr + "     where ((m.orderserial is null))"
	sqlStr = sqlStr + " )"

	dbget.Execute sqlStr


	sqlStr = "delete from [db_temp].[dbo].tbl_mibeasong_list"
	sqlStr = sqlStr + " where orderserial in ("
	sqlStr = sqlStr + "     select l.orderserial from [db_temp].[dbo].tbl_mibeasong_list l"
    sqlStr = sqlStr + "     ,[db_order].[dbo].tbl_order_master m"
    sqlStr = sqlStr + "     where l.orderserial=m.orderserial"
    sqlStr = sqlStr + "     and m.ipkumdiv>7"
    sqlStr = sqlStr + "     and datediff(d,l.regdate,getdate())>45"
    sqlStr = sqlStr + " )"

	dbget.Execute sqlStr

	sqlStr = "delete from [db_temp].[dbo].tbl_mibeasong_list"
	sqlStr = sqlStr + " where orderserial in ("
	sqlStr = sqlStr + "     select l.orderserial from [db_temp].[dbo].tbl_mibeasong_list l"
    sqlStr = sqlStr + "     ,[db_order].[dbo].tbl_order_master m"
    sqlStr = sqlStr + "     where l.orderserial=m.orderserial"
    sqlStr = sqlStr + "     and m.cancelyn='Y'"
    sqlStr = sqlStr + "     and datediff(d,l.regdate,getdate())>45"
    sqlStr = sqlStr + " )"

	dbget.Execute sqlStr
end if
%>

<script language="javascript">
alert('삭제 되었습니다.');
opener.location.reload();
window.close();
</script>

<!-- #include virtual="/lib/db/dbclose.asp" -->
