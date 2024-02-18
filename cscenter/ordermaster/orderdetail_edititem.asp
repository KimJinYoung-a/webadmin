<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs센터 상품변경
' History : 이상구 생성
'			2023.06.12 한용민 수정(표준코딩으로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/jumuncls.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_couponcls.asp" -->
<!-- #include virtual="/lib/classes/cscenter/sp_itemcouponcls.asp" -->
<!-- #include virtual="/cscenter/lib/csOrderFunction.asp" -->
<!-- #include virtual="/lib/classes/items/newcouponcls.asp" -->
<%
dim i, idx, orderserial, result
	idx = requestCheckVar(request("idx"),10)

dim toItemId, toItemOption
toItemId = requestCheckVar(request("toItemId"),12)
toItemOption = requestCheckVar(request("toItemOption"),4)

dim errMsg


'==============================================================================
dim ojumunDetail
set ojumunDetail = new CJumunMaster
ojumunDetail.SearchOneJumunDetail idx

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
	response.write "잘못된 접속입니다."
	response.end
end if

if (IsNull(ojumunDetail.FJumunDetail.Fcurrstate) = True) then
	ojumunDetail.FJumunDetail.Fcurrstate = ""
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


dim orgItemIsSaleDiscountItem			: orgItemIsSaleDiscountItem = oorderdetail.FItemList(selecteditemindex).IsSaleDiscountItem
dim orgItemIsItemCouponDiscountItem		: orgItemIsItemCouponDiscountItem = oorderdetail.FItemList(selecteditemindex).IsItemCouponDiscountItem
dim orgItemIsBonusCouponDiscountItem	: orgItemIsBonusCouponDiscountItem = oorderdetail.FItemList(selecteditemindex).IsBonusCouponDiscountItem

dim orgItemIsBuyCashSaleApplied			: orgItemIsBuyCashSaleApplied = oorderdetail.FItemList(selecteditemindex).IsBuyCashSaleApplied
dim orgItemIsBuyCashCouponApplied		: orgItemIsBuyCashCouponApplied = oorderdetail.FItemList(selecteditemindex).IsBuyCashItemCouponApplied

dim orgItemIsPlusSaleItem				: orgItemIsPlusSaleItem = oorderdetail.FItemList(selecteditemindex).IsPlusSaleItem
dim orgItemIsMileageShopItem			: orgItemIsMileageShopItem = oorderdetail.FItemList(selecteditemindex).IsMileageShopItem
dim orgItemIsSpecialShopDiscountItem	: orgItemIsSpecialShopDiscountItem = oorderdetail.FItemList(selecteditemindex).IsSpecialShopDiscountItem

dim BonusCouponValue, BonusCouponType
BonusCouponValue = 0
BonusCouponType = -1

dim itemChangeAvail : itemChangeAvail = True
if (orgItemIsPlusSaleItem or orgItemIsMileageShopItem or orgItemIsSpecialShopDiscountItem) then
	if False and (orgItemIsPlusSaleItem = True) then
        '// 플러스세일 상품도 변경허용, 2020-02-03
		errMsg = "상품변경 불가 : 취소대상상품이 플러스세일상품입니다."
	elseif (orgItemIsMileageShopItem = True) then
        itemChangeAvail = False
		errMsg = "상품변경 불가 : 취소대상상품이 마일리지샵상품입니다."
	elseif (orgItemIsSpecialShopDiscountItem = True) then
        itemChangeAvail = False
		errMsg = "상품변경 불가 : 취소대상상품이 우수회원샵상품입니다."
	end if
end if


'==============================================================================
dim sqlStr

dim addItemName
dim addItemOptionName
dim addMakerid
dim addOrgprice
dim addOrgsuplycash
dim addSellcash
dim addListimage
dim addBuycash
dim addSailyn
dim addItemIsMileageItem
dim addIsSpecialShopitem
dim addIsUpchebeasong

dim IsMaySameItem : IsMaySameItem = False
dim IsMaySameItemName : IsMaySameItemName = False
dim IsMaySameBrand : IsMaySameBrand = True

dim IsSameSaleStatus : IsSameSaleStatus = False

'// 상품쿠폰
dim IsSameCouponStatus : IsSameCouponStatus = True
dim IsExpireCouponStatus : IsExpireCouponStatus = False

'// 보너스쿠폰
dim IsSameBonusCouponStatus : IsSameBonusCouponStatus = False
dim IsExpireBonusCouponStatus : IsExpireBonusCouponStatus = True
dim IsBonusCouponExists : IsBonusCouponExists = False


if (itemChangeAvail = True) and (toItemId <> "") and (toItemOption <> "") then
	sqlStr = " select top 1 "
	sqlStr = sqlStr + " 	i.itemname "
	sqlStr = sqlStr + " 	, IsNull(o.optionname, '') as optionname "
	sqlStr = sqlStr + " 	, i.makerid "
	sqlStr = sqlStr + " 	, (i.orgprice + IsNull(o.optaddprice, 0)) as orgprice "
	sqlStr = sqlStr + " 	, (i.orgsuplycash + IsNull(o.optaddbuyprice, 0)) as orgsuplycash "
	sqlStr = sqlStr + " 	, (i.sellcash + IsNull(o.optaddprice, 0)) as sellcash "
	sqlStr = sqlStr + " 	, i.listimage "
	sqlStr = sqlStr + " 	, (i.buycash + IsNull(o.optaddbuyprice, 0)) as buycash "
	sqlStr = sqlStr + " 	, i.sailyn "
	sqlStr = sqlStr + " 	, i.ItemDiv "
	sqlStr = sqlStr + " 	, i.specialuseritem "
    sqlStr = sqlStr + " 	, i.mwdiv "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_item.dbo.tbl_item i "
	sqlStr = sqlStr + " 	left join dbo.tbl_item_option o "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	i.itemid = o.itemid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and i.itemid = " & CStr(toItemId) & " "
	sqlStr = sqlStr + " 	and IsNull(o.itemoption, '0000') = '" & CStr(toItemOption) & "' "
    ''response.write sqlStr

	rsget.Open sqlStr,dbget,1

	if not rsget.Eof then
		addMakerid 			= rsget("makerid")
		addItemname 		= db2Html(rsget("itemname"))
		addItemoptionname 	= db2Html(rsget("optionname"))

		addOrgprice			= rsget("orgprice")
		addOrgsuplycash		= rsget("orgsuplycash")
		addSellcash			= rsget("sellcash")
		addBuycash 			= rsget("buycash")

		addSailyn				= rsget("sailyn")
		addItemIsMileageItem 	= (rsget("ItemDiv") = "82")
		addIsSpecialShopitem	= rsget("specialuseritem")

		addListimage 			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemID(toItemId) + "/" + rsget("listimage")

        addIsUpchebeasong	= CHKIIF(rsget("mwdiv")="U", "Y", "N")
	else
		if itemChangeAvail then
			itemChangeAvail = False
			errMsg = "추가대상 상품이 없습니다."
		end if
	end if
	rsget.Close

	if itemChangeAvail then
		if (addItemIsMileageItem = True) then
			itemChangeAvail = False
			errMsg = "상품변경 불가 : 추가대상상품이 마일리지샵상품입니다."
		elseif (addIsSpecialShopitem = "Y") then
			itemChangeAvail = False
			errMsg = "상품변경 불가 : 추가대상상품이 우수회원샵상품입니다."
		end if
	end if

	if itemChangeAvail and (oorderdetail.FItemList(selecteditemindex).Fitemid = CLng(toItemId)) then
		'' 2015-09-23, skyer9
		''itemChangeAvail = False
		''errMsg = "상품변경 불가 : 동일 상품코드입니다.<br>옵션변경을 이용하세요."
	end if

	if itemChangeAvail and (LCase(oorderdetail.FItemList(selecteditemindex).Fmakerid) <> LCase(addMakerid)) and Not (oorderdetail.FItemList(selecteditemindex).Fisupchebeasong="N" and addIsUpchebeasong="N") then
		''itemChangeAvail = False
		''errMsg = "상품변경 불가 : 동일 브랜드만 상품변경 가능합니다."
		IsMaySameBrand = False
	end if

	if (itemChangeAvail = True) and (oorderdetail.FItemList(selecteditemindex).GetOrgItemCostPrice = addOrgprice) and (oorderdetail.FItemList(selecteditemindex).Forgsuplycash = addOrgsuplycash) then
		'// 같은 상품 기준 : 동일 소비자가, 동일 기본 매입가(판매시 할인 매입가 아님)
		'// 옵션체크는 안하도록 변경
		IsMaySameItem = True

		if (Left(Replace(addItemName, " ", ""), 4) = Left(Replace(ojumunDetail.FJumunDetail.Fitemname, " ", ""), 4)) then
			'// 상품명 앞부분 동일
			IsMaySameItemName = True
		end if

		'// 상품명 비교 안함(2014-05-20)
		IsMaySameItemName = True
	end if

	if (((orgItemIsSaleDiscountItem = True) and (addSailyn = "Y")) or ((orgItemIsSaleDiscountItem <> True) and (addSailyn <> "Y"))) and (addSellcash = oorderdetail.FItemList(selecteditemindex).GetSalePrice) and (oorderdetail.FItemList(selecteditemindex).Forgsuplycash = addOrgsuplycash) then
		'// 동일 세일상태
		IsSameSaleStatus = True
	end if

	if orgItemIsItemCouponDiscountItem and (oorderdetail.FItemList(selecteditemindex).Forgsuplycash <> oorderdetail.FItemList(selecteditemindex).Fbuycash) then
		'// 상품쿠폰이 적용되었고, 상품쿠폰 매입가가 설정되어 있는 경우

		sqlStr = " select top 1 "
		sqlStr = sqlStr + " (case when u.itemcouponexpiredate >= getdate() then 'N' else 'Y' end) as expireYN "
		sqlStr = sqlStr + " , (case when IsNull(d.itemid, 0) <> 0 and IsNull(d.couponbuyprice, 0) <> 0 then 'Y' else 'N' end) as availYN "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_item.dbo.tbl_user_item_coupon u "
		sqlStr = sqlStr + " 	join db_item.dbo.tbl_item_coupon_master m "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and u.orderserial = '" + CStr(orderserial) + "' "
		sqlStr = sqlStr + " 		and u.userid = '" + CStr(oordermaster.FOneItem.Fuserid) + "' "
		sqlStr = sqlStr + " 		and u.itemcouponidx = m.itemcouponidx "
		sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item_coupon_detail d "
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and m.itemcouponidx = d.itemcouponidx "
		sqlStr = sqlStr + " 		and d.itemid = " + CStr(toItemId) + " "

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1

		if not rsget.Eof then
			IsExpireCouponStatus 	= (rsget("expireYN") = "Y")
			IsSameCouponStatus		= (rsget("availYN") = "Y")
		else
			IsExpireCouponStatus	= True
			IsSameCouponStatus		= False
		end if
		rsget.Close

	end if

	if orgItemIsBonusCouponDiscountItem then
		''response.write "aaa"

		sqlStr = " select top 1 couponvalue, coupontype "
		sqlStr = sqlStr + " from [db_user].[dbo].tbl_user_coupon "
		sqlStr = sqlStr + " where 1 = 1 and userid='" + CStr(oordermaster.FOneItem.Fuserid) + "' "
		sqlStr = sqlStr + " and orderserial = '" + CStr(orderserial) + "' "

		''response.write sqlStr
		rsget.Open sqlStr,dbget,1

		if not rsget.Eof then
			IsSameBonusCouponStatus 		= True
			IsExpireBonusCouponStatus		= False
			BonusCouponValue 				= rsget("couponvalue")
			BonusCouponType 				= rsget("coupontype")
		else
			IsSameBonusCouponStatus 		= False
			IsExpireBonusCouponStatus		= True
		end if
		rsget.Close

	end if

end if

dim isupchebeasong, requiremakerid
isupchebeasong = ojumunDetail.FJumunDetail.Fisupchebeasong

if (isupchebeasong = "Y") then
	requiremakerid = ojumunDetail.FJumunDetail.Fmakerid
end if

dim title, contents_jupsu

title = "상품변경(동일소비자가)"
if ojumunDetail.FJumunDetail.FcurrState = 7 then
	title = "교환출고(상품변경)"
end if

'// 기본문구 설정
if Not IsNull(session("ssBctCname")) then
	contents_jupsu = "텐바이텐 고객센터 " + CStr(session("ssBctCname")) + " 입니다"
end if

dim posit_sn
	posit_sn = getposit_sn(session("ssBctSn"),session("ssBctId"))	' 직위 받아옴
%>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language='javascript' SRC="/js/ajax.js"></script>
<script language='javascript' SRC="/cscenter/js/newcsas.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
window.resizeTo(1400,800);

// 사유구분(ajax) 를 사용하기 위해 필요
var IsPossibleModifyCSMaster = true;
var IsPossibleModifyItemList = true;

function SearchItemByItemname(makerid, itemname) {
	var popwin = window.open('pop_item_search_one.asp?itemname=' + itemname + '&makerid=' + makerid + '&onlineonly=Y','SearchItemByItemname','width=1000,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function SearchItemByMakerid(makerid) {
	var popwin = window.open('pop_item_search_one.asp?makerid=' + makerid + '&onlineonly=Y','SearchItemByMakerid','width=1000,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function SearchItemByPrice(makerid, saleprice) {
	var popwin = window.open('pop_item_search_one.asp?makerid=' + makerid + '&saleprice=' + saleprice + '&onlineonly=Y','SearchItemByPrice','width=1000,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function SearchItemAll() {
	var popwin = window.open('pop_item_search_one.asp' + '?onlineonly=Y','SearchItemAll','width=1000,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ReActItemOne(toItemId, toItemOption) {
	document.location.href = "/cscenter/ordermaster/orderdetail_edititem.asp?idx=<%= idx %>&toItemId=" + toItemId + "&toItemOption=" + toItemOption;
}

function applyPriceToAddItem(applyPriceType) {
	var frm = document.frm;

	var orgprice = "0";
	var orgbuyprice = "0";

	var currsellprice = "0";
	var currbuyprice = "0";

	var saleprice = "0";
	var itemcouponprice = "0";
	var bonuscouponprice = "0";
	var etcdiscountprice = "0";
	var buyprice = "0";

	jsChangeTitle();

	orgprice = frm.toOrgItemCostPrice.value * 1;
	orgbuyprice = frm.toOrgbuycash.value * 1;

	currsellprice = frm.toSellcash.value * 1;
	currbuyprice = frm.toBuycash.value * 1;

	if (applyPriceType == "Y") {
		saleprice 			= frm.fromSalePrice.value * 1;
		itemcouponprice 	= frm.fromItemCouponPrice.value * 1;
		bonuscouponprice 	= frm.fromBonusCouponPrice.value * 1;
		etcdiscountprice 	= frm.fromEtcDiscountPrice.value * 1;
		buyprice 			= frm.fromBuycash.value * 1;
	} else if (applyPriceType == "N") {
		saleprice 			= frm.toOrgItemCostPrice.value * 1;
		itemcouponprice 	= frm.toOrgItemCostPrice.value * 1;
		bonuscouponprice 	= frm.toOrgItemCostPrice.value * 1;
		etcdiscountprice 	= frm.toOrgItemCostPrice.value * 1;
		buyprice 			= frm.toOrgbuycash.value * 1;
	} else if (applyPriceType == "S") {
		saleprice 			= frm.toSellcash.value * 1;
		itemcouponprice 	= frm.toSellcash.value * 1;
		bonuscouponprice 	= frm.toSellcash.value * 1;
		etcdiscountprice 	= frm.toSellcash.value * 1;
		buyprice 			= frm.toOrgbuycash.value * 1;
	} else if (applyPriceType == "D") {
		// 다른상품(할인/쿠폰 적용)
		saleprice 			= frm.toOrgItemCostPrice.value * 1;
		itemcouponprice 	= frm.toOrgItemCostPrice.value * 1;
		bonuscouponprice 	= frm.toOrgItemCostPrice.value * 1;
		etcdiscountprice 	= frm.toOrgItemCostPrice.value * 1;
		buyprice 			= frm.toOrgbuycash.value * 1;

		if (saleprice > currsellprice) {
			// 추가 상품 할인중이면 할인가 입력
			saleprice = currsellprice
			itemcouponprice = currsellprice
			bonuscouponprice = currsellprice
			etcdiscountprice = currsellprice
			buyprice = currbuyprice
		}

		if (IsSameBonusCouponStatus == true) {
			if (BonusCouponType == 2) {
				// 일단 정액쿠폰만 작업함.
				bonuscouponprice 	= itemcouponprice - (frm.fromSalePrice.value*1 - frm.fromBonusCouponPrice.value*1);
				etcdiscountprice 	= bonuscouponprice - (frm.fromBonusCouponPrice.value*1 - frm.fromEtcDiscountPrice.value*1);
			} else {
				alert("\n\n적용불가!! 작업안되어 있음!!\n\n");
				frm.applyToAddItem[0].checked = true;
				applyPriceToAddItem('N');
				return;
			}
		}
	} else {
		// CS할인
		saleprice 			= frm.toOrgItemCostPrice.value * 1;
		itemcouponprice 	= frm.toOrgItemCostPrice.value * 1;
		bonuscouponprice 	= frm.toOrgItemCostPrice.value * 1;
		etcdiscountprice 	= frm.toOrgItemCostPrice.value * 1;
		buyprice 			= frm.toOrgbuycash.value * 1;

		if (saleprice > currsellprice) {
			// 추가 상품 할인중이면 할인가 입력
			saleprice = currsellprice
			itemcouponprice = currsellprice
			bonuscouponprice = currsellprice
			etcdiscountprice = currsellprice
			buyprice = currbuyprice
		}

		if (saleprice > frm.fromEtcDiscountPrice.value*1) {
			saleprice 			= frm.fromEtcDiscountPrice.value*1;
			itemcouponprice 	= frm.fromEtcDiscountPrice.value*1;
			bonuscouponprice 	= frm.fromEtcDiscountPrice.value*1;
			etcdiscountprice 	= frm.fromEtcDiscountPrice.value*1;
		}
	}

	frm.toSalePrice.value = saleprice;
	frm.toItemCouponPrice.value = itemcouponprice;
	frm.toBonusCouponPrice.value = bonuscouponprice;
	frm.toEtcDiscountPrice.value = etcdiscountprice;
	frm.toAddBuycash.value = buyprice;

	if (orgprice > saleprice) {
		$("#toItem_sellcash").html("<font color='red'>" + FormatNumber(saleprice) + "</font>");
	} else {
		$("#toItem_sellcash").html(FormatNumber(saleprice));
	}

	if (saleprice > itemcouponprice) {
		$("#toItem_itemcoupon_cash").html("<font color='green'>" + FormatNumber(itemcouponprice) + "</font>");
	} else {
		$("#toItem_itemcoupon_cash").html(FormatNumber(itemcouponprice));
	}

	if (itemcouponprice > bonuscouponprice) {
		$("#toItem_bonuscoupon_cash").html("<font color='purple'><b>" + FormatNumber(bonuscouponprice) + "</b></font>");
	} else {
		$("#toItem_bonuscoupon_cash").html("<b>" + FormatNumber(bonuscouponprice) + "</b>");
	}

	if (bonuscouponprice > etcdiscountprice) {
		$("#toItem_etcdiscount_cash").html("<font color='red'><b>" + FormatNumber(etcdiscountprice) + "</b></font>");
	} else {
		$("#toItem_etcdiscount_cash").html("<b>" + FormatNumber(etcdiscountprice) + "</b>");
	}

	if (orgbuyprice > buyprice) {
		$("#toItem_buycash").html("<font color='red'><b>" + FormatNumber(buyprice) + "</b></font>");
	} else {
		$("#toItem_buycash").html("<b>" + FormatNumber(buyprice) + "</b>");
	}
}

function jsChangeTitle() {
	var frm = document.frm;

	if (!frm.title) {
		return;
	}

	<% if ojumunDetail.FJumunDetail.FcurrState = 7 then %>
		frm.title.value = "교환출고(상품변경, 할인쿠폰정보 적용)";
	<% else %>
		if (frm.applyToAddItem[2].checked == true) {
			frm.title.value = "상품변경(CS할인)";
		} else if (frm.applyToAddItem[1].checked == true) {
			frm.title.value = "상품변경(옵션변경, 할인쿠폰정보 적용)";
		} else if (frm.applyToAddItem[0].checked == true) {
			frm.title.value = "상품변경";
		} else if (frm.applyToAddItem[3].checked == true) {
			frm.title.value = "상품변경";
		} else {
			alert("ERROR");
			return;
		}
	<% end if %>
}

function FormatNumber(nStr) {
	var radixdivided, integernumber, primenumber;

	nStr += '';

	radixdivided = nStr.split('.');
	integernumber = radixdivided[0];

	if (radixdivided.length > 1) {
		primenumber = "." + radixdivided[1];
	} else {
		primenumber = "";
	}

	var regex = /(\d+)(\d{3})/;
	while (regex.test(integernumber)) {
		integernumber = integernumber.replace(regex, '$1' + ',' + '$2');
	}
	return integernumber + primenumber;
}

$(document).ready(function() {
	<% if (IsMaySameItem and IsSameSaleStatus and IsSameCouponStatus and Not IsExpireCouponStatus) then %>
    applyPriceToAddItem("Y");
	<% else %>
	applyPriceToAddItem("N");
	<% end if %>
});

function CheckItemNo() {
    var frm = document.frm;

    var itemnoorg = frm.itemnoorg.value*1;
    var itemnocancel = frm.itemnocancel.value;

    if (itemnocancel*0 != 0) {
		alert('수량에 숫자를 입력하세요.');
		frm.itemnocancel.value = itemnoorg;
		frm.itemnoadd.value = itemnoorg;

		return;
    }

    if (itemnocancel*1 <= 0) {
		alert('수량에 0 또는 마이너스를 넣을 수 없습니다.');
		frm.itemnocancel.value = itemnoorg;
		frm.itemnoadd.value = itemnoorg;

		return;
    }

    if (itemnocancel*1 > itemnoorg) {
		alert('원주문 수량을 초과하여 변경할 수 없습니다.');
		frm.itemnocancel.value = itemnoorg;
		frm.itemnoadd.value = itemnoorg;

		return;
    }

    frm.itemnoadd.value = itemnocancel;
}

var isupchebeasong 		= "<%= ojumunDetail.FJumunDetail.Fisupchebeasong %>";
var detailstate 		= "<%= ojumunDetail.FJumunDetail.FcurrState %>";
var ipkumdiv 			= "<%= ojumun.FOneItem.Fipkumdiv %>";
var subtotalprice		= "<%= ojumun.FOneItem.Fsubtotalprice %>";

var IsMaySameItem 		= <%= LCase(IsMaySameItem) %>;
var IsMaySameItemName 	= <%= LCase(IsMaySameItemName) %>;
var IsSameSaleStatus 	= <%= LCase(IsSameSaleStatus) %>;
var IsSameCouponStatus 	= <%= LCase(IsSameCouponStatus) %>;
var IsExpireCouponStatus = <%= LCase(IsExpireCouponStatus) %>;
var IsMaySameBrand = <%= LCase(IsMaySameBrand) %>;

var IsSameBonusCouponStatus = <%= LCase(IsSameBonusCouponStatus) %>;
var BonusCouponValue = <%= (BonusCouponValue) %>;
var BonusCouponType = <%= LCase(BonusCouponType) %>;


function SaveChangeItem(isadmin) {
    var frm = document.frm;
	var IsSameItemApply = frm.applyToAddItem[1].checked;
    //var itemno;
	var itemnocancel, itemnoadd;

	if (isadmin != true) {
		frm.itemnoadd.value = frm.itemnocancel.value;
	}

    itemnocancel = frm.itemnocancel.value*1;
	itemnoadd = frm.itemnoadd.value*1;

	if (frm.toItemId.value == "") {
		alert('추가할 상품을 선택하세요.');

		return;
	}

	if (detailstate == "") {
		detailstate = "0";
	}

	if (frm.gubun01.value == "") {
		alert("사유구분을 선택하세요.");
		return;
	}

	if ((frm.fromItemId.value == frm.toItemId.value) && (frm.fromItemOption.value == frm.toItemOption.value)) {
		alert('[변경불가] 같은 상품입니다.');
		return;
	}

	if (detailstate >= "7") {
		alert('!!! 이미 출고된 상품입니다. !!!\n\n교환출고 기능을 이용하세요.');
		return;
	}

	if (isadmin != true) {
		if (IsMaySameBrand != true) {
			alert('[관리자 문의] <동일 브랜드 상품> 만 선택가능합니다.');
			return;
		}

		if (IsSameItemApply != true) {
			alert('[관리자 문의] <동일상품(할인/쿠폰 적용)> 만 선택가능합니다.'+IsSameItemApply);
			return;
		}

		if ((IsMaySameItem != true) || (IsSameSaleStatus != true) || (IsSameCouponStatus != true) || (IsExpireCouponStatus == true)) {
			alert('[관리자 문의] 동일종류 상품이 아니거나, 할인/쿠폰 상태가 다릅니다.'+IsMaySameItem+','+IsSameSaleStatus+','+IsSameCouponStatus+','+IsExpireCouponStatus);
			return;
		}
	} else {
		if (IsMaySameItem != true) {
			if (confirm("상품이 다를 수 있습니다.\n\n계속 진행하시겠습니까?") != true) {
				return;
			}
		}

		if (((IsSameCouponStatus != true)) && (IsSameItemApply == true)) {
			if (confirm("[관리자권한]\n\n소비자가 또는 기본매입가가 다르거나 옵션이 있는 상품입니다.\n\n<동일상품(할인/쿠폰 적용)> 계속 진행하시겠습니까?") != true) {
				return;
			}

			// 2014-10-23, skyer9
			//alert('소비자가 또는 기본매입가가 다르거나 옵션이 있는 상품입니다.\n\n<동일상품(할인/쿠폰 적용)> 을 선택할 수 없습니다.');
			//return;
		}

		if (frm.fromBuycash.value*1 < frm.toAddBuycash.value*1) {
			if (confirm("====== [[ 역 마 진]] =================================\n\n\n\n계속 진행하시겠습니까?") != true) {
				return;
			}
		}
	}

	if (frm.fromEtcDiscountPrice.value*1 < frm.toEtcDiscountPrice.value*1) {
		alert('[상품변경 불가]\n\n추가상품의 기타할인 적용가가 더 큽니다.\n할인 적용방식을 선택하세요.');
		return;
	} else if (frm.fromEtcDiscountPrice.value*1 > frm.toEtcDiscountPrice.value*1) {
		if (ipkumdiv >= 4) {
			if (confirm("차액에 대한 환불이 접수됩니다.\n\n진행하시겠습니까?") != true) {
				return;
			}
		} else {
			if (confirm("추가되는 상품의 금액이 취소되는 상품의 금액보다 작습니다\n[결제완료이전 : 환불없음]\n\n진행하시겠습니까?") != true) {
				return;
			}
		}
	}

	// =========================================================================
	frm.refundrequire.value 	= 0;
	frm.canceltotal.value 		= 0;
	frm.refunditemcostsum.value = 0;
	frm.refundcouponsum.value 	= 0;
	frm.allatsubtractsum.value 	= 0;

	frm.canceltotal.value 		= ((frm.fromSalePrice.value * itemnocancel) - (frm.toSalePrice.value * itemnoadd));
	frm.refunditemcostsum.value	= ((frm.fromSalePrice.value * itemnocancel) - (frm.toSalePrice.value * itemnoadd));

	frm.refundcouponsum.value = ((frm.fromSalePrice.value*itemnocancel - frm.fromBonusCouponPrice.value*itemnocancel) - (frm.toSalePrice.value*itemnoadd - frm.toBonusCouponPrice.value*itemnoadd));
	frm.allatsubtractsum.value = ((frm.fromBonusCouponPrice.value*itemnocancel - frm.fromEtcDiscountPrice.value*itemnocancel) - (frm.toBonusCouponPrice.value*itemnoadd - frm.toEtcDiscountPrice.value*itemnoadd));
	frm.refundrequire.value 	= frm.refunditemcostsum.value*1 - frm.refundcouponsum.value*1 - frm.allatsubtractsum.value*1;

	if (frm.refundrequire.value*1 < 0) {
		alert("취소금액보다 추가금액이 큽니다.\n\n상품변경 할 수 없습니다.");
		return;
	}

	if (subtotalprice*1 < frm.refundrequire.value*1) {
		alert("실 결제액보다 환불액이 큽니다.\n\n상품변경 할 수 없습니다.");
		return;
	}

	var msg = "상품변경 하시겠습니까?";

	if ((isupchebeasong == "Y") && (detailstate >= "3") && (detailstate < "7")) {
		msg = "업체배송이면서 상품준비 이후입니다\n\n" + msg;
	} else if (detailstate >= "7") {
		msg = "이미 출고된 상품입니다. 정산이 이루어진 상품인 경우 변경할 수 없습니다.\n\n" + msg;
	}

	if (isadmin == true) {
		msg = "[파트장권한] " + msg;
	}

	if ((frm.refundrequire.value*1 > 0) && (ipkumdiv >= 4)) {
		frm.title.value = frm.title.value + " + 차액환불";
	}

	msg = frm.title.value + "\n\n" + msg;

	if (confirm(msg) == true) {
		if (isadmin == true) {
			frm.forceedit.value = "Y";
		}

		frm.submit();
	}
}

// 교환출고(상품변경)
function SaveChangeOrder(isadmin) {
    var frm = document.frm;
	var IsSameItemApply = frm.applyToAddItem[1].checked;
    // var itemno;
	var itemnocancel, itemnoadd;

	if (isadmin != true) {
		frm.itemnoadd.value = frm.itemnocancel.value;
	}

	itemnocancel = frm.itemnocancel.value*1;
	itemnoadd = frm.itemnoadd.value*1;

	if (frm.toItemId.value == "") {
		alert('추가할 상품을 선택하세요.');
		return;
	}

	if ((frm.fromItemId.value == frm.toItemId.value) && (frm.fromItemOption.value == frm.toItemOption.value)) {
		alert('[변경불가] 동일 상품입니다.');
		return;
	}

	if (detailstate != "7") {
		alert('[접수불가] 출고완료 상품만 교환출고 가능합니다.');
		return;
	}

	/*
	if (frm.toOrgItemCostPrice.value*1 < frm.toSalePrice.value*1) {
		alert('[접수불가] 소비자가보다 판매가가 더 높습니다.');
		return;
	}
	*/

	// =========================================================================
	frm.refundrequire.value 	= 0;
	frm.canceltotal.value 		= 0;
	frm.refunditemcostsum.value = 0;
	frm.refundcouponsum.value 	= 0;
	frm.allatsubtractsum.value 	= 0;

	frm.canceltotal.value 		= ((frm.fromSalePrice.value * itemnocancel) - (frm.toSalePrice.value * itemnoadd));
	frm.refunditemcostsum.value	= ((frm.fromSalePrice.value * itemnocancel) - (frm.toSalePrice.value * itemnoadd));

	frm.refundcouponsum.value = ((frm.fromSalePrice.value*itemnocancel - frm.fromBonusCouponPrice.value*itemnocancel) - (frm.toSalePrice.value*itemnoadd - frm.toBonusCouponPrice.value*itemnoadd));
	frm.allatsubtractsum.value = ((frm.fromBonusCouponPrice.value*itemnocancel - frm.fromEtcDiscountPrice.value*itemnocancel) - (frm.toBonusCouponPrice.value*itemnoadd - frm.toEtcDiscountPrice.value*itemnoadd));
	frm.refundrequire.value 	= frm.refunditemcostsum.value*1 - frm.refundcouponsum.value*1 - frm.allatsubtractsum.value*1;

	if (frm.itemnoadd.value !== frm.itemnocancel.value) {
		<% if not(C_ADMIN_AUTH) then %>
		alert('[시스템팀 문의] 상품수량이 다릅니다.');
		return;
		<% end if %>
	}

	if (IsSameItemApply != true) {
		if (isadmin != true) {
			alert('[관리자 문의] <동일상품(할인/쿠폰 적용)> 만 선택가능합니다.'+IsSameItemApply);
			return;
		} else {
			if ((frm.refundcouponsum.value*1 == 0) && (frm.allatsubtractsum.value*1 == 0)) {
				// 쿠폰,기타할인 없는 경우
			} else {
				<% if C_ADMIN_AUTH then %>
				<% 'if C_ADMIN_AUTH or C_CSPowerUser then %>
				alert('[관리자]!!!! 강제진행 : 쿠폰관련 작업안되어 있음. !!!!');
				<% else %>
				alert('[시스템팀 문의] 쿠폰 또는 기타할인 적용 상품입니다.');
				return;
				<% end if %>
			}
		}
	}

	alert("환불금액 : " + frm.refundrequire.value);

	if (frm.refundrequire.value*1 !== 0 ) {
		if (isadmin != true) {
			alert('[관리자 문의] 추가/취소 상품의 판매가가 다릅니다.');
			return;
		}
	}

	if (frm.refundrequire.value*1 < 0) {
		alert("취소금액보다 추가금액이 큽니다.\n\n상품변경 할 수 없습니다.");
		return;
	}

	if (subtotalprice*1 < frm.refundrequire.value*1) {
		alert("실 결제액보다 환불액이 큽니다.\n\n상품변경 할 수 없습니다.");
		return;
	}

	if (frm.gubun01.value == "") {
		alert("사유구분을 선택하세요.");
		return;
	}

	if ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD01")) {
		// 단순변심 접수불가 => 배송비 처리안함
		alert("[접수불가] 단순변심일 경우 교환출고 불가합니다.");
		return;
	}

	if (IsMaySameBrand != true) {
		alert('[접수불가] <동일 브랜드 상품> 만 선택가능합니다.');
		return;
	}

	if ((IsMaySameItem != true) || (IsSameSaleStatus != true) || (IsSameCouponStatus != true) || (IsExpireCouponStatus == true)) {
		if (isadmin) {
			if (confirm("[관리자권한]\n\n동일종류 상품이 아니거나, 할인/쿠폰 상태가 다릅니다.\n\n계속 진행하시겠습니까?") != true) {
				return;
			}
		} else {
			alert('[관리자 문의] 동일종류 상품이 아니거나, 할인/쿠폰 상태가 다릅니다.');
			return;
		}
	}

	var msg = "교환출고(상품변경) 하시겠습니까?";

	if (confirm(msg) != true) {
		return;
	}

	// frm.title.value = "교환출고(상품변경)";
	frm.mode.value = "orderChange";

	frm.submit();
}

</script>

<form name="frm" method="post" action="/cscenter/ordermaster/orderdetail_process.asp" style="margin:0px;">
<input type="hidden" name="detailidx" value="<%= ojumunDetail.FJumunDetail.Fdetailidx %>">
<input type="hidden" name="orderserial" value="<%= ojumunDetail.FJumunDetail.FOrderSerial %>">
<input type="hidden" name="mode" value="itemChange">
<input type="hidden" name="forceedit" value="N">

<input type="hidden" name="isupchebeasong" value="<%= isupchebeasong %>">
<input type="hidden" name="requiremakerid" value="<%= requiremakerid %>">

<input type="hidden" name="itemnoorg" value="<%= ojumunDetail.FJumunDetail.Fitemno %>">

<input type="hidden" name="fromItemId" value="<%= ojumunDetail.FJumunDetail.Fitemid %>">
<input type="hidden" name="fromItemOption" value="<%= ojumunDetail.FJumunDetail.FitemOption %>">
<input type="hidden" name="toItemId" value="<%= toItemId %>">
<input type="hidden" name="toItemOption" value="<%= toItemOption %>">

<input type="hidden" name="fromOrgItemCostPrice" value="<%= oorderdetail.FItemList(selecteditemindex).GetOrgItemCostPrice %>">
<input type="hidden" name="fromOrgbuycash" value="<%= oorderdetail.FItemList(selecteditemindex).Forgsuplycash %>">
<input type="hidden" name="fromSalePrice" value="<%= oorderdetail.FItemList(selecteditemindex).GetSalePrice %>">
<input type="hidden" name="fromItemCouponPrice" value="<%= oorderdetail.FItemList(selecteditemindex).GetItemCouponPrice %>">
<input type="hidden" name="fromBonusCouponPrice" value="<%= oorderdetail.FItemList(selecteditemindex).GetBonusCouponPrice %>">
<input type="hidden" name="fromEtcDiscountPrice" value="<%= oorderdetail.FItemList(selecteditemindex).GetEtcDiscountPrice %>">
<input type="hidden" name="fromBuycash" value="<%= oorderdetail.FItemList(selecteditemindex).Fbuycash %>">

<input type="hidden" name="fromBonusCouponIdx" value="<%= oorderdetail.FItemList(selecteditemindex).Fbonuscouponidx %>">
<input type="hidden" name="fromItemCouponIdx" value="<%= oorderdetail.FItemList(selecteditemindex).Fitemcouponidx %>">

<input type="hidden" name="toOrgItemCostPrice" value="<%= addOrgprice %>">
<input type="hidden" name="toOrgbuycash" value="<%= addOrgsuplycash %>">
<input type="hidden" name="toSellcash" value="<%= addSellcash %>">
<input type="hidden" name="toBuycash" value="<%= addBuycash %>">
<input type="hidden" name="toSalePrice" value="">
<input type="hidden" name="toItemCouponPrice" value="">
<input type="hidden" name="toBonusCouponPrice" value="">
<input type="hidden" name="toEtcDiscountPrice" value="">
<input type="hidden" name="toAddBuycash" value="">

<input type="hidden" name="contents_finish" value="상품변경이 정상적으로 처리되었습니다.">

<input type="hidden" name="refundrequire" value="">
<input type="hidden" name="canceltotal" value="">
<input type="hidden" name="refunditemcostsum" value="">
<input type="hidden" name="refundcouponsum" value="">
<input type="hidden" name="allatsubtractsum" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="3">
			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>주문아이템정보 수정</b>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">주문번호</td>
		<td colspan="2"><%= oordermaster.FOneItem.Forderserial %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">결재방법</td>
		<td colspan="2"><%= oordermaster.FOneItem.JumunMethodName %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">거래상태</td>
		<td colspan="2"><%= oordermaster.FOneItem.IpkumDivName %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>" colspan="3"><b>취소 대상상품</b></td>
	</tr>

	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
		<td colspan="2"><%= ojumunDetail.FJumunDetail.Fitemname %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td width="75" bgcolor="<%= adminColor("tabletop") %>">브랜드 ID</td>
		<td width="445"><%= ojumunDetail.FJumunDetail.Fmakerid %></td>
		<td rowspan="5" align="center"><img src="<%= ojumunDetail.FJumunDetail.FImageList %>"></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemid %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">기존옵션</td>
		<td>[<%= ojumunDetail.FJumunDetail.Fitemoption %>] <%= ojumunDetail.FJumunDetail.Fitemoptionname %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">취소수량</td>
		<td>
			<input type="text" class="text" name="itemnocancel" value="<%= ojumunDetail.FJumunDetail.Fitemno %>" size="3" maxlength="9" onFocusOut="CheckItemNo()"> / <%= ojumunDetail.FJumunDetail.Fitemno %>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">취소상태</td>
		<td><%= ojumunDetail.FJumunDetail.Fcancelyn %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">가격정보</td>
		<td colspan="2">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">소비자가<br>(+옵션가)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">판매가<br>(할인가)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">구매가<br>(상품쿠폰)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">보너스쿠폰<br>적용가</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">기타할인<br>적용가</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">기본매입가</td>
				<td bgcolor="<%= adminColor("tabletop") %>" align="center">매입가<br>(판매시)</td>
			</tr>
			<tr height="25" bgcolor="FFFFFF">
				<td align="right" style="padding-right:20px">
                	<font color="<%= oorderdetail.FItemList(selecteditemindex).GetOrgItemCostColor %>">
                		<%= FormatNumber(oorderdetail.FItemList(selecteditemindex).GetOrgItemCostPrice,0) %>
                	</font>
				</td>
				<td align="right" style="padding-right:20px">
                	<span title="<%= oorderdetail.FItemList(selecteditemindex).GetSaleText %>" style="cursor:hand">
                		<font color="<%= oorderdetail.FItemList(selecteditemindex).GetSaleColor %>">
                			<%= FormatNumber(oorderdetail.FItemList(selecteditemindex).GetSalePrice,0) %>
                		</font>
                	</span>
				</td>
				<td align="right" style="padding-right:20px">
                	<span title="<%= oorderdetail.FItemList(selecteditemindex).GetItemCouponText %>" style="cursor:hand">
                		<font color="<%= oorderdetail.FItemList(selecteditemindex).GetItemCouponColor %>">
                			<%= FormatNumber(oorderdetail.FItemList(selecteditemindex).GetItemCouponPrice,0) %>
                		</font>
                	</span>
				</td>
				<td align="right" style="padding-right:20px">
                    <span title="<%= oorderdetail.FItemList(selecteditemindex).GetBonusCouponText %>" style="cursor:hand">
                    	<font color="<%= oorderdetail.FItemList(selecteditemindex).GetBonusCouponColor %>">
                    		<b><%= FormatNumber(oorderdetail.FItemList(selecteditemindex).GetBonusCouponPrice,0) %></b>
                    	</font>
                    </span>
				</td>
				<td align="right" style="padding-right:20px">
                    <span title="<%= oorderdetail.FItemList(selecteditemindex).GetEtcDiscountText %>" style="cursor:hand">
                    	<font color="<%= oorderdetail.FItemList(selecteditemindex).GetEtcDiscountColor %>">
                    		<b><%= FormatNumber(oorderdetail.FItemList(selecteditemindex).GetEtcDiscountPrice,0) %></b>
                    	</font>
                    </span>
				</td>
				<td align="right" style="padding-right:25px">
                    <%= FormatNumber(oorderdetail.FItemList(selecteditemindex).Forgsuplycash,0) %>
				</td>
				<td align="right" style="padding-right:25px">
					<% if (oorderdetail.FItemList(selecteditemindex).IsBuyCashSaleApplied) then %>
						<span title="<%= oorderdetail.FItemList(selecteditemindex).GetSaleBuycashText %>" style="cursor:hand">
							<font color="<%= oorderdetail.FItemList(selecteditemindex).GetSaleBuycashColor %>">
								<b><%= FormatNumber(oorderdetail.FItemList(selecteditemindex).Fbuycash,0) %></b>
							</font>
						</span>
					<% elseif (oorderdetail.FItemList(selecteditemindex).IsBuyCashItemCouponApplied) then %>
						<span title="<%= oorderdetail.FItemList(selecteditemindex).GetItemCouponBuycashText %>" style="cursor:hand">
							<font color="<%= oorderdetail.FItemList(selecteditemindex).GetItemCouponBuycashColor %>">
								<b><%= FormatNumber(oorderdetail.FItemList(selecteditemindex).Fbuycash,0) %></b>
							</font>
						</span>
					<% else %>
						<b><%= FormatNumber(oorderdetail.FItemList(selecteditemindex).Fbuycash,0) %></b>
					<% end if %>
				</td>
			</tr>
			</table>
		</td>
	</tr>

	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>" colspan="3">
			<b><font color="red">추가 대상상품</font></b>
			<input type="button" class="button" value="검색하기" style="width:100px; height:22px;" onClick="javascript:SearchItemAll()">
			<input type="button" class="button" value="동일브랜드" style="width:100px; height:22px;" onClick="javascript:SearchItemByMakerid('<%= Replace(ojumunDetail.FJumunDetail.Fmakerid, "'", "\'") %>')">
			<input type="button" class="button" value="동일판매가" style="width:100px; height:22px;" onClick="javascript:SearchItemByPrice('<%= Replace(ojumunDetail.FJumunDetail.Fmakerid, "'", "\'") %>', '<%= oorderdetail.FItemList(selecteditemindex).GetSalePrice %>')">
			<input type="button" class="button" value="동일상품명" style="width:100px; height:22px;" onClick="javascript:SearchItemByItemname('<%= Replace(ojumunDetail.FJumunDetail.Fmakerid, "'", "\'") %>', '<%= Server.URLencode(ojumunDetail.FJumunDetail.Fitemname) %>')">
		</td>
	</tr>
	<% if (toItemId <> "") and (toItemOption <> "") then %>
	<tr height="25" bgcolor="#FFFFFF" id="tradd01">
		<td bgcolor="<%= adminColor("tabletop") %>">상품명</td>
		<td colspan="2">
			<%= addItemName %>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd02">
		<td bgcolor="<%= adminColor("tabletop") %>">브랜드 ID</td>
		<td>
			<%= addMakerid %>
		</td>
		<td rowspan="4" align="center">
			<img src="<%= addListimage %>">
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd03">
		<td bgcolor="<%= adminColor("tabletop") %>">상품코드</td>
		<td>
			<%= toItemId %>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd04">
		<td bgcolor="<%= adminColor("tabletop") %>">추가옵션</td>
		<td>
			[<%= toItemOption %>] <%= addItemOptionName %>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd05">
		<td bgcolor="<%= adminColor("tabletop") %>">추가수량</td>
		<td>
			<input type="text" class="text_ro" name="itemnoadd" value="<%= ojumunDetail.FJumunDetail.Fitemno %>" size="3" maxlength="9">
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd06">
		<td bgcolor="<%= adminColor("tabletop") %>">가격정보</td>
		<td colspan="2">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">소비자가<br>(+옵션가)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">판매가<br>(할인가)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">구매가<br>(상품쿠폰)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">보너스쿠폰<br>적용가</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">기타할인<br>적용가</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">기본매입가</td>
				<td bgcolor="<%= adminColor("tabletop") %>" align="center">매입가<br>(현재)</td>
			</tr>
			<tr height="25" bgcolor="FFFFFF">
				<td align="right" style="padding-right:20px">
					<%= FormatNumber(addOrgprice,0) %>
				</td>
				<td align="right" style="padding-right:20px" id="toItem_sellcash">
					<% if (addOrgprice > addSellcash) then %>
					<font color="red"><%= FormatNumber(addSellcash,0) %></font>
					<% else %>
					<%= FormatNumber(addSellcash,0) %>
					<% end if %>
				</td>
				<td align="right" style="padding-right:20px" id="toItem_itemcoupon_cash">
					<%= FormatNumber(addSellcash,0) %>
				</td>
				<td align="right" style="padding-right:20px" id="toItem_bonuscoupon_cash">
					<b><%= FormatNumber(addSellcash,0) %></b>
				</td>
				<td align="right" style="padding-right:20px" id="toItem_etcdiscount_cash">
					<b><%= FormatNumber(addSellcash,0) %></b>
				</td>
				<td align="right" style="padding-right:25px">
					<%= FormatNumber(addOrgsuplycash,0) %>
				</td>
				<td align="right" style="padding-right:25px" id="toItem_buycash">
					<b>
					<% if (addOrgsuplycash > addBuycash) then %>
					<font color="red"><%= FormatNumber(addBuycash,0) %></font>
					<% else %>
					<%= FormatNumber(addBuycash,0) %>
					<% end if %>
					</b>
				</td>
			</tr>
			<tr height="25" bgcolor="FFFFFF">
				<td align="left" colspan="7">
					* 할인상태 :
					<% if (IsSameSaleStatus) then %>
					동일
					<% elseif (addSailyn = "Y") then %>
					<font color="red">현재 할인중</font>
					<% else %>
					<font color="red">서로다름</font>
					<% end if %>
					<% if (orgItemIsBonusCouponDiscountItem or orgItemIsItemCouponDiscountItem) then %>
						* 쿠폰상태 :
						<% if (IsSameCouponStatus and Not IsExpireCouponStatus) then %>
						적용가능
						<% elseif (IsSameCouponStatus and IsExpireCouponStatus) then %>
						<font color="red">적용가능 기간경과</font>
						<% else %>
						<font color="red">적용불가 상품</font>
						<% end if %>
					<% end if %>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd07">
		<td bgcolor="<%= adminColor("tabletop") %>">사유구분</td>
		<td colspan="2">
                <input type="hidden" name="gubun01" value="">
                <input type="hidden" name="gubun02" value="">
                <input class="text_ro" type="text" name="gubun01name" value="" size="16" Readonly >
                &gt;
                <input class="text_ro" type="text" name="gubun02name" value="" size="16" Readonly >
                <input class="csbutton" type="button" value="선택" onClick="divCsAsGubunSelect(frm.gubun01.value, frm.gubun02.value, frm.gubun01.name, frm.gubun02.name, frm.gubun01name.name, frm.gubun02name.name,'frm','causepop');">
                <div id="causepop" style="position:absolute;"></div>

                <!-- 일부 사유 미리 표시 -->
                <%
                '참조쿼리
				'select top 100 m.comm_cd, m.comm_name, d.comm_cd, d.comm_name
				'from
				'	db_cs.dbo.tbl_cs_comm_code m
				'	left join db_cs.dbo.tbl_cs_comm_code d
				'	on
				'		m.comm_cd = d.comm_group
				'where
				'	1 = 1
				'	and m.comm_group = 'Z020'
				'	and m.comm_isdel <> 'Y'
				'	and d.comm_isdel <> 'Y'
				'order by m.comm_cd, d.comm_cd
                %>
                [<a href="javascript:selectGubun('C004','CD01','공통','단순변심','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">단순변심</a>]
                [<a href="javascript:selectGubun('C004','CD05','공통','품절','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">품절</a>]
                [<a href="javascript:selectGubun('C005','CE01','상품관련','상품불량','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">상품불량</a>]
                [<a href="javascript:selectGubun('C004','CD99','공통','기타','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">기타</a>]
            	<br>
            	<div id="chkmodifyitemstockoutyn" style="display: inline;"><input type="checkbox" name="modifyitemstockoutyn" value="Y" checked> 품절정보 저장(업배상품)</div>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd07">
		<td bgcolor="<%= adminColor("tabletop") %>">할인적용</td>
		<td colspan="2">
			<input type="radio" name="applyToAddItem" value="N" <% if Not (IsMaySameItem and IsSameSaleStatus and IsSameCouponStatus and Not IsExpireCouponStatus) then %>checked<% end if %> onClick="applyPriceToAddItem('N')" > 할인안함
			<input type="radio" name="applyToAddItem" value="Y" <% if (IsMaySameItem and IsSameSaleStatus and IsSameCouponStatus and Not IsExpireCouponStatus) then %>checked<% end if %> onClick="applyPriceToAddItem('Y')" > 동일상품(할인/쿠폰 적용)
			<% if C_ADMIN_AUTH then %>
			<input type="radio" name="applyToAddItem" value="D" onClick="applyPriceToAddItem('D')" > 다른상품(할인/쿠폰 적용)
			<% end if %>
			<input type="radio" name="applyToAddItem" value="C" onClick="applyPriceToAddItem('C')" > CS할인
			<input type="radio" name="applyToAddItem" value="S" onClick="applyPriceToAddItem('S')" > 현재판매가
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
			접수제목
		</td>
		<td colspan=2>
                <input class='text' type="text" name="title" value="<%= title %>" size="56" maxlength="56">
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
			접수내용
		</td>
		<td colspan=2>
                <textarea class='textarea' name="contents_jupsu" cols="68" rows="6"><%= contents_jupsu %></textarea>
		</td>
	</tr>

	<% end if %>
	<tr bgcolor="#FFFFFF" height="40">
		<td colspan="3" align="center">
<% if (ojumunDetail.FJumunDetail.Fcancelyn <> "Y" and ojumun.FOneItem.Fcancelyn <> "Y") and itemChangeAvail = True and ojumunDetail.FJumunDetail.Fitemno > 0 then %>
			<input type="button" class="button" value="상품변경" style="width:100px; height:22px;" onclick="javascript:SaveChangeItem(false)">
			<% if C_ADMIN_AUTH or (C_CSPowerUser) or (posit_sn<=11 and C_CSUser) then %>
		    <input type="button" class="button" value="강제변경" style="width:100px; height:22px;" onclick="javascript:SaveChangeItem(true)">
			<% end if %>
			&nbsp;
			<input type="button" class="button" value="교환출고(상품변경)" style="width:130px; height:22px;" onclick="javascript:SaveChangeOrder(false)">
			<% if C_ADMIN_AUTH or (C_CSPowerUser) or (posit_sn<=11 and C_CSUser) then %>
			<input type="button" class="button" value="교환출고(강제변경)" style="width:130px; height:22px;" onclick="javascript:SaveChangeOrder(true)">
			<% end if %>
<% elseif (ojumunDetail.FJumunDetail.Fcancelyn = "Y" or ojumun.FOneItem.Fcancelyn = "Y") then %>
			<b>취소된 상품은 상품변경 불가</b>
<% elseif ojumunDetail.FJumunDetail.Fitemno < 0 then %>
			<b>마이너스 주문 상품변경 불가</b>
<% else %>
			<b><%= errMsg %></b>
<% end if %>
		</td>
	</tr>
</table>
</form>

<%
set ojumun       = Nothing
set ojumunDetail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
