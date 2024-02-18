<%@ language=vbscript %>
<% option explicit %>
<% response.Charset="euc-kr" %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<%

dim orderserial
dim sqlStr, i
dim toItemId, toItemOption

toItemId = requestCheckVar(request("toItemId"),32)
toItemOption = requestCheckVar(request("toItemOption"),32)



'==============================================================================
dim isaddok, itemid, itemoption, makerid, itemname, itemoptionname, orgitemcostprice, imagesmall
dim sellcash, optaddprice, buycash, optaddbuyprice
dim issaleitem, ismileageshopitem, isspacialuseritem


sqlStr = " select top 1 "
sqlStr = sqlStr + " 	i.itemid "
sqlStr = sqlStr + " 	, IsNull(o.itemoption, '0000') as itemoption "
sqlStr = sqlStr + " 	, i.itemname "
sqlStr = sqlStr + " 	, IsNull(o.optionname, '') as optionname "
sqlStr = sqlStr + " 	, i.makerid "
sqlStr = sqlStr + " 	, i.sellcash "
sqlStr = sqlStr + " 	, i.orgprice "
sqlStr = sqlStr + " 	, i.mileage "
sqlStr = sqlStr + " 	, i.listimage "
sqlStr = sqlStr + " 	, i.buycash "
sqlStr = sqlStr + " 	, IsNull(o.optaddprice, 0) as optaddprice "
sqlStr = sqlStr + " 	, IsNull(o.optaddbuyprice, 0) as optaddbuyprice "
sqlStr = sqlStr + " 	, i.sailyn "
sqlStr = sqlStr + " 	, i.ItemDiv "
sqlStr = sqlStr + " 	, i.specialuseritem "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + " 	db_item.dbo.tbl_item i "
sqlStr = sqlStr + " 	left join dbo.tbl_item_option o "
sqlStr = sqlStr + " on "
sqlStr = sqlStr + " 	i.itemid = o.itemid "
sqlStr = sqlStr + " where "
sqlStr = sqlStr + " 	1 = 1 "
sqlStr = sqlStr + " 	and i.itemid = " & CStr(toItemId) & " "
sqlStr = sqlStr + " 	and IsNull(o.itemoption, '0000') = '" & CStr(toItemOption) & "' "

rsget.Open sqlStr,dbget,1

isaddok 			= "false"
orgitemcostprice	= 0

if not rsget.Eof then
	makerid 			= rsget("makerid")
	itemname 			= db2Html(rsget("itemname"))
	itemoptionname 		= db2Html(rsget("optionname"))

	orgitemcostprice	= rsget("orgprice") + rsget("optaddprice")

	sellcash			= rsget("sellcash")
	optaddprice			= rsget("optaddprice")

	issaleitem			= rsget("sailyn")
	if (rsget("ItemDiv") = "82") then
		ismileageshopitem	= "Y"
	else
		ismileageshopitem	= "N"
	end if
	isspacialuseritem	= rsget("specialuseritem")

	buycash 			= rsget("buycash")
	optaddbuyprice		= rsget("optaddbuyprice")

	imagesmall 			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemID(toItemId) + "/" + rsget("listimage")

	isaddok 			= "true"
end if
rsget.Close



'==============================================================================
dim ItemCouponYn, CurrItemCouponIdx, ItemCouponType, ItemCouponValue, couponbuyprice

'사용가능한 상품쿠폰
sqlStr = " select top 1 "
sqlStr = sqlStr + " 	IsNull(i.ItemCouponYn, 'N') as ItemCouponYn "
sqlStr = sqlStr + " 	, IsNull(i.CurrItemCouponIdx, '0') as CurrItemCouponIdx "
sqlStr = sqlStr + " 	, IsNull(i.ItemCouponType, '1') as ItemCouponType "
sqlStr = sqlStr + " 	, IsNull(i.ItemCouponValue, 0) as ItemCouponValue "
sqlStr = sqlStr + " 	, IsNull(d.couponbuyprice, 0) as couponbuyprice "
sqlStr = sqlStr + " 	, IsNull(m.itemcouponidx, 0) as validcouponidx "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + " 	db_item.dbo.tbl_item i "
sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item_coupon_master m "
sqlStr = sqlStr + " 	on "
sqlStr = sqlStr + " 		1 = 1 "
sqlStr = sqlStr + " 		and m.itemcouponidx = i.CurrItemCouponIdx "
sqlStr = sqlStr + " 		and m.itemcouponstartdate <= getdate() "
sqlStr = sqlStr + " 		and m.itemcouponexpiredate > getdate() "
sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item_coupon_detail d "
sqlStr = sqlStr + " 	on "
sqlStr = sqlStr + " 		1 = 1 "
sqlStr = sqlStr + " 		and m.itemcouponidx = d.itemcouponidx "
sqlStr = sqlStr + " 		and d.itemid = i.itemid "
sqlStr = sqlStr + " where "
sqlStr = sqlStr + " 	i.itemid = " & CStr(toItemId) & " "

rsget.Open sqlStr,dbget,1

ItemCouponYn = "N"
CurrItemCouponIdx = 0
ItemCouponType = "1"
ItemCouponValue = "0"
couponbuyprice = "0"

if not rsget.Eof then
	ItemCouponYn		= rsget("ItemCouponYn")
	CurrItemCouponIdx	= rsget("CurrItemCouponIdx")
	ItemCouponType		= rsget("ItemCouponType")
	ItemCouponValue		= rsget("ItemCouponValue")
	couponbuyprice		= rsget("couponbuyprice")
end if
rsget.Close



response.write "<script>parent.ReActItemAdd(true, " & CStr(toItemId) & ", '" & CStr(toItemOption) & "', '" & CStr(makerid) & "', '" & CStr(Replace(Replace(itemname, "'", "\'"), Chr(34), "&quot;")) & "', '" & CStr(Replace(Replace(itemoptionname, "'", "\'"), Chr(34), "&quot;")) & "', '" & orgitemcostprice & "', '" & CStr(sellcash) & "', '" & CStr(optaddprice) & "', '" & CStr(buycash) & "', '" & CStr(optaddbuyprice) & "', '" & imagesmall & "', '" & issaleitem & "', '" & ismileageshopitem & "', '" & isspacialuseritem & "', '" & CStr(ItemCouponYn) & "', '" & CStr(CurrItemCouponIdx) & "', '" & CStr(ItemCouponType) & "', '" & CStr(ItemCouponValue) & "', '" & CStr(couponbuyprice) & "')</script>"

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->