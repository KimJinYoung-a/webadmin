<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim refer
refer = request.ServerVariables("HTTP_REFERER")

dim mode
dim itemArr, oneItem
dim itemgubun, itemid, itemoption, dayforsellcount, dayforsafestock, dayforleadtime, dayformaxstock

mode	= request.form("mode")
itemArr	= request.form("itemArr")

dim sqlStr, i

if mode="saveestockbaseday" then
	itemArr = Split(itemArr, "|")

	for i = 0 to UBound(itemArr)
		oneItem = itemArr(i)

		if (Trim(oneItem) <> "") then
			oneItem = Split(oneItem, ",")

			itemgubun = oneItem(0)
			itemid = oneItem(1)
			itemoption = oneItem(2)
			dayforsellcount = oneItem(3)
			dayforsafestock = oneItem(4)
			dayforleadtime = oneItem(5)
			dayformaxstock = oneItem(6)

			sqlStr = " IF NOT EXISTS(SELECT itemid FROM [db_item].[dbo].[tbl_item_option_stock] WHERE itemgubun = '" + CStr(itemgubun) + "' and itemid = " + CStr(itemid) + " and itemoption = '" + CStr(itemoption) + "') "
			sqlStr = sqlStr + " BEGIN "
			sqlStr = sqlStr + " 	insert into [db_item].[dbo].[tbl_item_option_stock](itemgubun,itemid,itemoption,DayForSellCount,DayForSafeStock,DayForLeadTime,DayForMaxStock) "
			sqlStr = sqlStr + " 	values('" + CStr(itemgubun) + "'," + CStr(itemid) + ",'" + CStr(itemoption) + "', " + CStr(dayforsellcount) + ", " + CStr(DayForSafeStock) + ", " + CStr(DayForLeadTime) + ", " + CStr(DayForMaxStock) + ") "
			sqlStr = sqlStr + " END "
			sqlStr = sqlStr + " ELSE "
			sqlStr = sqlStr + " BEGIN "
			sqlStr = sqlStr + " 	update [db_item].[dbo].tbl_item_option_Stock "
			sqlStr = sqlStr + " 	set DayForSellCount = " + CStr(dayforsellcount) + " "
			sqlStr = sqlStr + " 	, DayForSafeStock = " + CStr(DayForSafeStock) + " "
			sqlStr = sqlStr + " 	, DayForLeadTime = " + CStr(DayForLeadTime) + " "
			sqlStr = sqlStr + " 	, DayForMaxStock = " + CStr(DayForMaxStock) + " "
			sqlStr = sqlStr + " 	where 1 = 1 "
			sqlStr = sqlStr + " 	and itemgubun = '" + CStr(itemgubun) + "' "
			sqlStr = sqlStr + " 	and itemid = " + CStr(itemid) + " "
			sqlStr = sqlStr + " 	and itemoption = '" + CStr(itemoption) + "' "
			sqlStr = sqlStr + " END "

			' sqlStr = " update [db_item].[dbo].tbl_item_option_Stock "
			' sqlStr = sqlStr + " set DayForSellCount = " + CStr(dayforsellcount) + " "
			' sqlStr = sqlStr + " , DayForSafeStock = " + CStr(dayforsafestock) + " "
			' sqlStr = sqlStr + " , DayForLeadTime = " + CStr(dayforleadtime) + " "
			' sqlStr = sqlStr + " , DayForMaxStock = " + CStr(dayformaxstock) + " "
			' sqlStr = sqlStr + " where 1 = 1 "
			' sqlStr = sqlStr + " and itemgubun='" + CStr(itemgubun) + "'"
			' sqlStr = sqlStr + " and itemid=" + CStr(itemid) + ""
			' sqlStr = sqlStr + " and itemoption='" + CStr(itemoption) + "'"

			''response.write sqlStr
			rsget.Open sqlStr,dbget,1

			'// 2018-01-03, skyer9
			'// 재고기준일수 별도 등록 상품만
			sqlStr = " exec [db_summary].[dbo].[usp_Ten_Refresh_MakeItem_RequireNO] '" + CStr(itemgubun) + "', " + CStr(itemid) + ", '" + CStr(itemoption) + "' "
			rsget.Open sqlStr,dbget,1

		end if
	next

end if
%>


<script language="javascript">
<% if mode="xxx" or mode="xxxx" then %>
alert('저장 되었습니다.');
opener.location.reload();
window.close();
<% else %>
alert('저장 되었습니다.');
location.replace('<%= refer %>');
<% end if %>
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
