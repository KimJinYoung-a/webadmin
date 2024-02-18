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
dim detailidx, toItemId, toItemOption

detailidx = requestCheckVar(request("detailidx"),32)
toItemId = requestCheckVar(request("toItemId"),32)
toItemOption = requestCheckVar(request("toItemOption"),32)



'==============================================================================
dim ojumunDetail
set ojumunDetail = new CJumunMaster
ojumunDetail.SearchOneJumunDetail detailidx

orderserial = ojumunDetail.FJumunDetail.FOrderSerial



'==============================================================================
dim ojumun
set ojumun = new COrderMaster

if (orderserial <> "") then
    ojumun.FRectOrderSerial = orderserial
    ojumun.QuickSearchOrderMaster
end if

if (ojumun.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    ojumun.FRectOldOrder = "on"
    ojumun.QuickSearchOrderMaster
end if

if (ojumun.FResultCount<1) then
	response.write "<script>alert('잘못된 접속입니다.');</script>"
	response.end
end if



'==============================================================================
dim oordermaster, oorderdetail, selecteditemindex

set oordermaster = new COrderMaster
oordermaster.FRectOrderSerial = orderserial
oordermaster.QuickSearchOrderMaster

if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if


set oorderdetail = new COrderMaster
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.QuickSearchOrderDetail

if (oorderdetail.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oorderdetail.FRectOldOrder = "on"
    oorderdetail.QuickSearchOrderDetail
end if


selecteditemindex = 0
for i = 0 to oorderdetail.FResultCount - 1
	if (CStr(oorderdetail.FItemList(i).Fidx) = CStr(ojumunDetail.FJumunDetail.Fdetailidx)) then
		selecteditemindex = i
	end if
next



'==============================================================================
'상품쿠폰/보너스쿠폰/기타할인 동일적용 가능여부 확인
'상품쿠폰 적용가가 동일한가 확인

dim IsItemCouponDiscountItem, IsBonusCouponDiscountItem

IsItemCouponDiscountItem = oorderdetail.FItemList(selecteditemindex).IsItemCouponDiscountItem
IsBonusCouponDiscountItem = oorderdetail.FItemList(selecteditemindex).IsBonusCouponDiscountItem



'oorderdetail.FItemList(selecteditemindex).GetItemCouponPrice



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

'response.write orgitemcostprice

itemid 				= toItemId
itemoption 			= toItemOption



'==============================================================================
dim ItemCouponYn, CurrItemCouponIdx, ItemCouponType, ItemCouponValue, couponbuyprice

'현재 사용가능한 상품쿠폰
sqlStr = " select top 1 "
sqlStr = sqlStr + " 	IsNull(i.ItemCouponYn, 'N') as ItemCouponYn, IsNull(i.CurrItemCouponIdx, '0') as CurrItemCouponIdx, IsNull(i.ItemCouponType, '1') as ItemCouponType, IsNull(i.ItemCouponValue, 0) as ItemCouponValue, IsNull(d.couponbuyprice, 0) as couponbuyprice, IsNull(m.itemcouponidx, 0) as validcouponidx "
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



'==============================================================================
dim ItemCouponYnOrg, CurrItemCouponIdxOrg, ItemCouponTypeOrg, ItemCouponValueOrg, couponbuypriceOrg

if IsNull(oorderdetail.FItemList(selecteditemindex).Fitemcouponidx) then
	oorderdetail.FItemList(selecteditemindex).Fitemcouponidx = 0
end if

'원주문 상품쿠폰
sqlStr = " select top 1 "
sqlStr = sqlStr + " 	IsNull(m.itemcoupontype, '1') as ItemCouponType, IsNull(m.itemcouponvalue, 0) as ItemCouponValue, IsNull(d.couponbuyprice, 0) as couponbuyprice "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + " 	db_item.dbo.tbl_item_coupon_master m "
sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_coupon_detail d "
sqlStr = sqlStr + " 	on "
sqlStr = sqlStr + " 		m.itemcouponidx = d.itemcouponidx "
sqlStr = sqlStr + " where "
sqlStr = sqlStr + " 	1 = 1 "
sqlStr = sqlStr + " 	and m.itemcouponidx = " & CStr(oorderdetail.FItemList(selecteditemindex).Fitemcouponidx) & " "
sqlStr = sqlStr + " 	and d.itemid = " & CStr(toItemId) & " "

rsget.Open sqlStr,dbget,1

ItemCouponYnOrg = "N"
CurrItemCouponIdxOrg = 0
ItemCouponTypeOrg = "1"
ItemCouponValueOrg = "0"
couponbuypriceOrg = "0"

if not rsget.Eof then
	ItemCouponYnOrg			= "Y"
	CurrItemCouponIdxOrg	= oorderdetail.FItemList(selecteditemindex).Fitemcouponidx
	ItemCouponTypeOrg		= rsget("ItemCouponType")
	ItemCouponValueOrg		= rsget("ItemCouponValue")
	couponbuypriceOrg		= rsget("couponbuyprice")
end if
rsget.Close

''dim TT : TT = "<script>parent.ReActItemAdd(true, " & CStr(itemid) & ", '" & CStr(itemoption) & "', '" & CStr(makerid) & "', '" & CStr(Replace(Replace(itemname, "'", "\'"), Chr(34), "&quot;")) & "', '" & CStr(Replace(Replace(itemoptionname, "'", "\'"), Chr(34), "&quot;")) & "', '" & orgitemcostprice & "', '" & CStr(sellcash) & "', '" & CStr(optaddprice) & "', '" & CStr(buycash) & "', '" & CStr(optaddbuyprice) & "', '" & imagesmall & "', '" & issaleitem & "', '" & ismileageshopitem & "', '" & isspacialuseritem & "', '" & CStr(ItemCouponYn) & "', '" & CStr(CurrItemCouponIdx) & "', '" & CStr(ItemCouponType) & "', '" & CStr(ItemCouponValue) & "', '" & CStr(couponbuyprice) & "', '" & CStr(ItemCouponYnOrg) & "', '" & CStr(CurrItemCouponIdxOrg) & "', '" & CStr(ItemCouponTypeOrg) & "', '" & CStr(ItemCouponValueOrg) & "', '" & CStr(couponbuypriceOrg) & "')</script>"
''response.write "<script>alert('"&Replace(TT,"'","")&"')</script>"
''response.write TT
response.write "<script>parent.ReActItemAdd(true, " & CStr(itemid) & ", '" & CStr(itemoption) & "', '" & CStr(makerid) & "', '" & CStr(Replace(Replace(itemname, "'", "\'"), Chr(34), "&quot;")) & "', '" & CStr(Replace(Replace(itemoptionname, "'", "\'"), Chr(34), "&quot;")) & "', '" & orgitemcostprice & "', '" & CStr(sellcash) & "', '" & CStr(optaddprice) & "', '" & CStr(buycash) & "', '" & CStr(optaddbuyprice) & "', '" & imagesmall & "', '" & issaleitem & "', '" & ismileageshopitem & "', '" & isspacialuseritem & "', '" & CStr(ItemCouponYn) & "', '" & CStr(CurrItemCouponIdx) & "', '" & CStr(ItemCouponType) & "', '" & CStr(ItemCouponValue) & "', '" & CStr(couponbuyprice) & "', '" & CStr(ItemCouponYnOrg) & "', '" & CStr(CurrItemCouponIdxOrg) & "', '" & CStr(ItemCouponTypeOrg) & "', '" & CStr(ItemCouponValueOrg) & "', '" & CStr(couponbuypriceOrg) & "')</script>"

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->