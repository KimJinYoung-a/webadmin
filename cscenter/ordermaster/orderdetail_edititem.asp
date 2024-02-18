<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� ��ǰ����
' History : �̻� ����
'			2023.06.12 �ѿ�� ����(ǥ���ڵ����� ����)
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
	response.write "<script>alert('�߸��� �����Դϴ�.');</script>"
	response.write "�߸��� �����Դϴ�."
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
        '// �÷������� ��ǰ�� �������, 2020-02-03
		errMsg = "��ǰ���� �Ұ� : ��Ҵ���ǰ�� �÷������ϻ�ǰ�Դϴ�."
	elseif (orgItemIsMileageShopItem = True) then
        itemChangeAvail = False
		errMsg = "��ǰ���� �Ұ� : ��Ҵ���ǰ�� ���ϸ�������ǰ�Դϴ�."
	elseif (orgItemIsSpecialShopDiscountItem = True) then
        itemChangeAvail = False
		errMsg = "��ǰ���� �Ұ� : ��Ҵ���ǰ�� ���ȸ������ǰ�Դϴ�."
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

'// ��ǰ����
dim IsSameCouponStatus : IsSameCouponStatus = True
dim IsExpireCouponStatus : IsExpireCouponStatus = False

'// ���ʽ�����
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
			errMsg = "�߰���� ��ǰ�� �����ϴ�."
		end if
	end if
	rsget.Close

	if itemChangeAvail then
		if (addItemIsMileageItem = True) then
			itemChangeAvail = False
			errMsg = "��ǰ���� �Ұ� : �߰�����ǰ�� ���ϸ�������ǰ�Դϴ�."
		elseif (addIsSpecialShopitem = "Y") then
			itemChangeAvail = False
			errMsg = "��ǰ���� �Ұ� : �߰�����ǰ�� ���ȸ������ǰ�Դϴ�."
		end if
	end if

	if itemChangeAvail and (oorderdetail.FItemList(selecteditemindex).Fitemid = CLng(toItemId)) then
		'' 2015-09-23, skyer9
		''itemChangeAvail = False
		''errMsg = "��ǰ���� �Ұ� : ���� ��ǰ�ڵ��Դϴ�.<br>�ɼǺ����� �̿��ϼ���."
	end if

	if itemChangeAvail and (LCase(oorderdetail.FItemList(selecteditemindex).Fmakerid) <> LCase(addMakerid)) and Not (oorderdetail.FItemList(selecteditemindex).Fisupchebeasong="N" and addIsUpchebeasong="N") then
		''itemChangeAvail = False
		''errMsg = "��ǰ���� �Ұ� : ���� �귣�常 ��ǰ���� �����մϴ�."
		IsMaySameBrand = False
	end if

	if (itemChangeAvail = True) and (oorderdetail.FItemList(selecteditemindex).GetOrgItemCostPrice = addOrgprice) and (oorderdetail.FItemList(selecteditemindex).Forgsuplycash = addOrgsuplycash) then
		'// ���� ��ǰ ���� : ���� �Һ��ڰ�, ���� �⺻ ���԰�(�ǸŽ� ���� ���԰� �ƴ�)
		'// �ɼ�üũ�� ���ϵ��� ����
		IsMaySameItem = True

		if (Left(Replace(addItemName, " ", ""), 4) = Left(Replace(ojumunDetail.FJumunDetail.Fitemname, " ", ""), 4)) then
			'// ��ǰ�� �պκ� ����
			IsMaySameItemName = True
		end if

		'// ��ǰ�� �� ����(2014-05-20)
		IsMaySameItemName = True
	end if

	if (((orgItemIsSaleDiscountItem = True) and (addSailyn = "Y")) or ((orgItemIsSaleDiscountItem <> True) and (addSailyn <> "Y"))) and (addSellcash = oorderdetail.FItemList(selecteditemindex).GetSalePrice) and (oorderdetail.FItemList(selecteditemindex).Forgsuplycash = addOrgsuplycash) then
		'// ���� ���ϻ���
		IsSameSaleStatus = True
	end if

	if orgItemIsItemCouponDiscountItem and (oorderdetail.FItemList(selecteditemindex).Forgsuplycash <> oorderdetail.FItemList(selecteditemindex).Fbuycash) then
		'// ��ǰ������ ����Ǿ���, ��ǰ���� ���԰��� �����Ǿ� �ִ� ���

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

title = "��ǰ����(���ϼҺ��ڰ�)"
if ojumunDetail.FJumunDetail.FcurrState = 7 then
	title = "��ȯ���(��ǰ����)"
end if

'// �⺻���� ����
if Not IsNull(session("ssBctCname")) then
	contents_jupsu = "�ٹ����� ������ " + CStr(session("ssBctCname")) + " �Դϴ�"
end if

dim posit_sn
	posit_sn = getposit_sn(session("ssBctSn"),session("ssBctId"))	' ���� �޾ƿ�
%>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language='javascript' SRC="/js/ajax.js"></script>
<script language='javascript' SRC="/cscenter/js/newcsas.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
window.resizeTo(1400,800);

// ��������(ajax) �� ����ϱ� ���� �ʿ�
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
		// �ٸ���ǰ(����/���� ����)
		saleprice 			= frm.toOrgItemCostPrice.value * 1;
		itemcouponprice 	= frm.toOrgItemCostPrice.value * 1;
		bonuscouponprice 	= frm.toOrgItemCostPrice.value * 1;
		etcdiscountprice 	= frm.toOrgItemCostPrice.value * 1;
		buyprice 			= frm.toOrgbuycash.value * 1;

		if (saleprice > currsellprice) {
			// �߰� ��ǰ �������̸� ���ΰ� �Է�
			saleprice = currsellprice
			itemcouponprice = currsellprice
			bonuscouponprice = currsellprice
			etcdiscountprice = currsellprice
			buyprice = currbuyprice
		}

		if (IsSameBonusCouponStatus == true) {
			if (BonusCouponType == 2) {
				// �ϴ� ���������� �۾���.
				bonuscouponprice 	= itemcouponprice - (frm.fromSalePrice.value*1 - frm.fromBonusCouponPrice.value*1);
				etcdiscountprice 	= bonuscouponprice - (frm.fromBonusCouponPrice.value*1 - frm.fromEtcDiscountPrice.value*1);
			} else {
				alert("\n\n����Ұ�!! �۾��ȵǾ� ����!!\n\n");
				frm.applyToAddItem[0].checked = true;
				applyPriceToAddItem('N');
				return;
			}
		}
	} else {
		// CS����
		saleprice 			= frm.toOrgItemCostPrice.value * 1;
		itemcouponprice 	= frm.toOrgItemCostPrice.value * 1;
		bonuscouponprice 	= frm.toOrgItemCostPrice.value * 1;
		etcdiscountprice 	= frm.toOrgItemCostPrice.value * 1;
		buyprice 			= frm.toOrgbuycash.value * 1;

		if (saleprice > currsellprice) {
			// �߰� ��ǰ �������̸� ���ΰ� �Է�
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
		frm.title.value = "��ȯ���(��ǰ����, ������������ ����)";
	<% else %>
		if (frm.applyToAddItem[2].checked == true) {
			frm.title.value = "��ǰ����(CS����)";
		} else if (frm.applyToAddItem[1].checked == true) {
			frm.title.value = "��ǰ����(�ɼǺ���, ������������ ����)";
		} else if (frm.applyToAddItem[0].checked == true) {
			frm.title.value = "��ǰ����";
		} else if (frm.applyToAddItem[3].checked == true) {
			frm.title.value = "��ǰ����";
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
		alert('������ ���ڸ� �Է��ϼ���.');
		frm.itemnocancel.value = itemnoorg;
		frm.itemnoadd.value = itemnoorg;

		return;
    }

    if (itemnocancel*1 <= 0) {
		alert('������ 0 �Ǵ� ���̳ʽ��� ���� �� �����ϴ�.');
		frm.itemnocancel.value = itemnoorg;
		frm.itemnoadd.value = itemnoorg;

		return;
    }

    if (itemnocancel*1 > itemnoorg) {
		alert('���ֹ� ������ �ʰ��Ͽ� ������ �� �����ϴ�.');
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
		alert('�߰��� ��ǰ�� �����ϼ���.');

		return;
	}

	if (detailstate == "") {
		detailstate = "0";
	}

	if (frm.gubun01.value == "") {
		alert("���������� �����ϼ���.");
		return;
	}

	if ((frm.fromItemId.value == frm.toItemId.value) && (frm.fromItemOption.value == frm.toItemOption.value)) {
		alert('[����Ұ�] ���� ��ǰ�Դϴ�.');
		return;
	}

	if (detailstate >= "7") {
		alert('!!! �̹� ���� ��ǰ�Դϴ�. !!!\n\n��ȯ��� ����� �̿��ϼ���.');
		return;
	}

	if (isadmin != true) {
		if (IsMaySameBrand != true) {
			alert('[������ ����] <���� �귣�� ��ǰ> �� ���ð����մϴ�.');
			return;
		}

		if (IsSameItemApply != true) {
			alert('[������ ����] <���ϻ�ǰ(����/���� ����)> �� ���ð����մϴ�.'+IsSameItemApply);
			return;
		}

		if ((IsMaySameItem != true) || (IsSameSaleStatus != true) || (IsSameCouponStatus != true) || (IsExpireCouponStatus == true)) {
			alert('[������ ����] �������� ��ǰ�� �ƴϰų�, ����/���� ���°� �ٸ��ϴ�.'+IsMaySameItem+','+IsSameSaleStatus+','+IsSameCouponStatus+','+IsExpireCouponStatus);
			return;
		}
	} else {
		if (IsMaySameItem != true) {
			if (confirm("��ǰ�� �ٸ� �� �ֽ��ϴ�.\n\n��� �����Ͻðڽ��ϱ�?") != true) {
				return;
			}
		}

		if (((IsSameCouponStatus != true)) && (IsSameItemApply == true)) {
			if (confirm("[�����ڱ���]\n\n�Һ��ڰ� �Ǵ� �⺻���԰��� �ٸ��ų� �ɼ��� �ִ� ��ǰ�Դϴ�.\n\n<���ϻ�ǰ(����/���� ����)> ��� �����Ͻðڽ��ϱ�?") != true) {
				return;
			}

			// 2014-10-23, skyer9
			//alert('�Һ��ڰ� �Ǵ� �⺻���԰��� �ٸ��ų� �ɼ��� �ִ� ��ǰ�Դϴ�.\n\n<���ϻ�ǰ(����/���� ����)> �� ������ �� �����ϴ�.');
			//return;
		}

		if (frm.fromBuycash.value*1 < frm.toAddBuycash.value*1) {
			if (confirm("====== [[ �� �� ��]] =================================\n\n\n\n��� �����Ͻðڽ��ϱ�?") != true) {
				return;
			}
		}
	}

	if (frm.fromEtcDiscountPrice.value*1 < frm.toEtcDiscountPrice.value*1) {
		alert('[��ǰ���� �Ұ�]\n\n�߰���ǰ�� ��Ÿ���� ���밡�� �� Ů�ϴ�.\n���� �������� �����ϼ���.');
		return;
	} else if (frm.fromEtcDiscountPrice.value*1 > frm.toEtcDiscountPrice.value*1) {
		if (ipkumdiv >= 4) {
			if (confirm("���׿� ���� ȯ���� �����˴ϴ�.\n\n�����Ͻðڽ��ϱ�?") != true) {
				return;
			}
		} else {
			if (confirm("�߰��Ǵ� ��ǰ�� �ݾ��� ��ҵǴ� ��ǰ�� �ݾ׺��� �۽��ϴ�\n[�����Ϸ����� : ȯ�Ҿ���]\n\n�����Ͻðڽ��ϱ�?") != true) {
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
		alert("��ұݾ׺��� �߰��ݾ��� Ů�ϴ�.\n\n��ǰ���� �� �� �����ϴ�.");
		return;
	}

	if (subtotalprice*1 < frm.refundrequire.value*1) {
		alert("�� �����׺��� ȯ�Ҿ��� Ů�ϴ�.\n\n��ǰ���� �� �� �����ϴ�.");
		return;
	}

	var msg = "��ǰ���� �Ͻðڽ��ϱ�?";

	if ((isupchebeasong == "Y") && (detailstate >= "3") && (detailstate < "7")) {
		msg = "��ü����̸鼭 ��ǰ�غ� �����Դϴ�\n\n" + msg;
	} else if (detailstate >= "7") {
		msg = "�̹� ���� ��ǰ�Դϴ�. ������ �̷���� ��ǰ�� ��� ������ �� �����ϴ�.\n\n" + msg;
	}

	if (isadmin == true) {
		msg = "[��Ʈ�����] " + msg;
	}

	if ((frm.refundrequire.value*1 > 0) && (ipkumdiv >= 4)) {
		frm.title.value = frm.title.value + " + ����ȯ��";
	}

	msg = frm.title.value + "\n\n" + msg;

	if (confirm(msg) == true) {
		if (isadmin == true) {
			frm.forceedit.value = "Y";
		}

		frm.submit();
	}
}

// ��ȯ���(��ǰ����)
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
		alert('�߰��� ��ǰ�� �����ϼ���.');
		return;
	}

	if ((frm.fromItemId.value == frm.toItemId.value) && (frm.fromItemOption.value == frm.toItemOption.value)) {
		alert('[����Ұ�] ���� ��ǰ�Դϴ�.');
		return;
	}

	if (detailstate != "7") {
		alert('[�����Ұ�] ���Ϸ� ��ǰ�� ��ȯ��� �����մϴ�.');
		return;
	}

	/*
	if (frm.toOrgItemCostPrice.value*1 < frm.toSalePrice.value*1) {
		alert('[�����Ұ�] �Һ��ڰ����� �ǸŰ��� �� �����ϴ�.');
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
		alert('[�ý����� ����] ��ǰ������ �ٸ��ϴ�.');
		return;
		<% end if %>
	}

	if (IsSameItemApply != true) {
		if (isadmin != true) {
			alert('[������ ����] <���ϻ�ǰ(����/���� ����)> �� ���ð����մϴ�.'+IsSameItemApply);
			return;
		} else {
			if ((frm.refundcouponsum.value*1 == 0) && (frm.allatsubtractsum.value*1 == 0)) {
				// ����,��Ÿ���� ���� ���
			} else {
				<% if C_ADMIN_AUTH then %>
				<% 'if C_ADMIN_AUTH or C_CSPowerUser then %>
				alert('[������]!!!! �������� : �������� �۾��ȵǾ� ����. !!!!');
				<% else %>
				alert('[�ý����� ����] ���� �Ǵ� ��Ÿ���� ���� ��ǰ�Դϴ�.');
				return;
				<% end if %>
			}
		}
	}

	alert("ȯ�ұݾ� : " + frm.refundrequire.value);

	if (frm.refundrequire.value*1 !== 0 ) {
		if (isadmin != true) {
			alert('[������ ����] �߰�/��� ��ǰ�� �ǸŰ��� �ٸ��ϴ�.');
			return;
		}
	}

	if (frm.refundrequire.value*1 < 0) {
		alert("��ұݾ׺��� �߰��ݾ��� Ů�ϴ�.\n\n��ǰ���� �� �� �����ϴ�.");
		return;
	}

	if (subtotalprice*1 < frm.refundrequire.value*1) {
		alert("�� �����׺��� ȯ�Ҿ��� Ů�ϴ�.\n\n��ǰ���� �� �� �����ϴ�.");
		return;
	}

	if (frm.gubun01.value == "") {
		alert("���������� �����ϼ���.");
		return;
	}

	if ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD01")) {
		// �ܼ����� �����Ұ� => ��ۺ� ó������
		alert("[�����Ұ�] �ܼ������� ��� ��ȯ��� �Ұ��մϴ�.");
		return;
	}

	if (IsMaySameBrand != true) {
		alert('[�����Ұ�] <���� �귣�� ��ǰ> �� ���ð����մϴ�.');
		return;
	}

	if ((IsMaySameItem != true) || (IsSameSaleStatus != true) || (IsSameCouponStatus != true) || (IsExpireCouponStatus == true)) {
		if (isadmin) {
			if (confirm("[�����ڱ���]\n\n�������� ��ǰ�� �ƴϰų�, ����/���� ���°� �ٸ��ϴ�.\n\n��� �����Ͻðڽ��ϱ�?") != true) {
				return;
			}
		} else {
			alert('[������ ����] �������� ��ǰ�� �ƴϰų�, ����/���� ���°� �ٸ��ϴ�.');
			return;
		}
	}

	var msg = "��ȯ���(��ǰ����) �Ͻðڽ��ϱ�?";

	if (confirm(msg) != true) {
		return;
	}

	// frm.title.value = "��ȯ���(��ǰ����)";
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

<input type="hidden" name="contents_finish" value="��ǰ������ ���������� ó���Ǿ����ϴ�.">

<input type="hidden" name="refundrequire" value="">
<input type="hidden" name="canceltotal" value="">
<input type="hidden" name="refunditemcostsum" value="">
<input type="hidden" name="refundcouponsum" value="">
<input type="hidden" name="allatsubtractsum" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="3">
			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ����������� ����</b>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
		<td colspan="2"><%= oordermaster.FOneItem.Forderserial %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td colspan="2"><%= oordermaster.FOneItem.JumunMethodName %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
		<td colspan="2"><%= oordermaster.FOneItem.IpkumDivName %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>" colspan="3"><b>��� ����ǰ</b></td>
	</tr>

	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
		<td colspan="2"><%= ojumunDetail.FJumunDetail.Fitemname %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td width="75" bgcolor="<%= adminColor("tabletop") %>">�귣�� ID</td>
		<td width="445"><%= ojumunDetail.FJumunDetail.Fmakerid %></td>
		<td rowspan="5" align="center"><img src="<%= ojumunDetail.FJumunDetail.FImageList %>"></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
		<td><%= ojumunDetail.FJumunDetail.Fitemid %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�����ɼ�</td>
		<td>[<%= ojumunDetail.FJumunDetail.Fitemoption %>] <%= ojumunDetail.FJumunDetail.Fitemoptionname %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">��Ҽ���</td>
		<td>
			<input type="text" class="text" name="itemnocancel" value="<%= ojumunDetail.FJumunDetail.Fitemno %>" size="3" maxlength="9" onFocusOut="CheckItemNo()"> / <%= ojumunDetail.FJumunDetail.Fitemno %>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">��һ���</td>
		<td><%= ojumunDetail.FJumunDetail.Fcancelyn %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td colspan="2">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">�Һ��ڰ�<br>(+�ɼǰ�)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">�ǸŰ�<br>(���ΰ�)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">���Ű�<br>(��ǰ����)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">���ʽ�����<br>���밡</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">��Ÿ����<br>���밡</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">�⺻���԰�</td>
				<td bgcolor="<%= adminColor("tabletop") %>" align="center">���԰�<br>(�ǸŽ�)</td>
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
			<b><font color="red">�߰� ����ǰ</font></b>
			<input type="button" class="button" value="�˻��ϱ�" style="width:100px; height:22px;" onClick="javascript:SearchItemAll()">
			<input type="button" class="button" value="���Ϻ귣��" style="width:100px; height:22px;" onClick="javascript:SearchItemByMakerid('<%= Replace(ojumunDetail.FJumunDetail.Fmakerid, "'", "\'") %>')">
			<input type="button" class="button" value="�����ǸŰ�" style="width:100px; height:22px;" onClick="javascript:SearchItemByPrice('<%= Replace(ojumunDetail.FJumunDetail.Fmakerid, "'", "\'") %>', '<%= oorderdetail.FItemList(selecteditemindex).GetSalePrice %>')">
			<input type="button" class="button" value="���ϻ�ǰ��" style="width:100px; height:22px;" onClick="javascript:SearchItemByItemname('<%= Replace(ojumunDetail.FJumunDetail.Fmakerid, "'", "\'") %>', '<%= Server.URLencode(ojumunDetail.FJumunDetail.Fitemname) %>')">
		</td>
	</tr>
	<% if (toItemId <> "") and (toItemOption <> "") then %>
	<tr height="25" bgcolor="#FFFFFF" id="tradd01">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ��</td>
		<td colspan="2">
			<%= addItemName %>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd02">
		<td bgcolor="<%= adminColor("tabletop") %>">�귣�� ID</td>
		<td>
			<%= addMakerid %>
		</td>
		<td rowspan="4" align="center">
			<img src="<%= addListimage %>">
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd03">
		<td bgcolor="<%= adminColor("tabletop") %>">��ǰ�ڵ�</td>
		<td>
			<%= toItemId %>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd04">
		<td bgcolor="<%= adminColor("tabletop") %>">�߰��ɼ�</td>
		<td>
			[<%= toItemOption %>] <%= addItemOptionName %>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd05">
		<td bgcolor="<%= adminColor("tabletop") %>">�߰�����</td>
		<td>
			<input type="text" class="text_ro" name="itemnoadd" value="<%= ojumunDetail.FJumunDetail.Fitemno %>" size="3" maxlength="9">
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd06">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td colspan="2">
			<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr height="25" bgcolor="FFFFFF">
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">�Һ��ڰ�<br>(+�ɼǰ�)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">�ǸŰ�<br>(���ΰ�)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">���Ű�<br>(��ǰ����)</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">���ʽ�����<br>���밡</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">��Ÿ����<br>���밡</td>
				<td width="14%" bgcolor="<%= adminColor("tabletop") %>" align="center">�⺻���԰�</td>
				<td bgcolor="<%= adminColor("tabletop") %>" align="center">���԰�<br>(����)</td>
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
					* ���λ��� :
					<% if (IsSameSaleStatus) then %>
					����
					<% elseif (addSailyn = "Y") then %>
					<font color="red">���� ������</font>
					<% else %>
					<font color="red">���δٸ�</font>
					<% end if %>
					<% if (orgItemIsBonusCouponDiscountItem or orgItemIsItemCouponDiscountItem) then %>
						* �������� :
						<% if (IsSameCouponStatus and Not IsExpireCouponStatus) then %>
						���밡��
						<% elseif (IsSameCouponStatus and IsExpireCouponStatus) then %>
						<font color="red">���밡�� �Ⱓ���</font>
						<% else %>
						<font color="red">����Ұ� ��ǰ</font>
						<% end if %>
					<% end if %>
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd07">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td colspan="2">
                <input type="hidden" name="gubun01" value="">
                <input type="hidden" name="gubun02" value="">
                <input class="text_ro" type="text" name="gubun01name" value="" size="16" Readonly >
                &gt;
                <input class="text_ro" type="text" name="gubun02name" value="" size="16" Readonly >
                <input class="csbutton" type="button" value="����" onClick="divCsAsGubunSelect(frm.gubun01.value, frm.gubun02.value, frm.gubun01.name, frm.gubun02.name, frm.gubun01name.name, frm.gubun02name.name,'frm','causepop');">
                <div id="causepop" style="position:absolute;"></div>

                <!-- �Ϻ� ���� �̸� ǥ�� -->
                <%
                '��������
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
                [<a href="javascript:selectGubun('C004','CD01','����','�ܼ�����','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">�ܼ�����</a>]
                [<a href="javascript:selectGubun('C004','CD05','����','ǰ��','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">ǰ��</a>]
                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">��ǰ�ҷ�</a>]
                [<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frm','causepop');">��Ÿ</a>]
            	<br>
            	<div id="chkmodifyitemstockoutyn" style="display: inline;"><input type="checkbox" name="modifyitemstockoutyn" value="Y" checked> ǰ������ ����(�����ǰ)</div>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" id="tradd07">
		<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
		<td colspan="2">
			<input type="radio" name="applyToAddItem" value="N" <% if Not (IsMaySameItem and IsSameSaleStatus and IsSameCouponStatus and Not IsExpireCouponStatus) then %>checked<% end if %> onClick="applyPriceToAddItem('N')" > ���ξ���
			<input type="radio" name="applyToAddItem" value="Y" <% if (IsMaySameItem and IsSameSaleStatus and IsSameCouponStatus and Not IsExpireCouponStatus) then %>checked<% end if %> onClick="applyPriceToAddItem('Y')" > ���ϻ�ǰ(����/���� ����)
			<% if C_ADMIN_AUTH then %>
			<input type="radio" name="applyToAddItem" value="D" onClick="applyPriceToAddItem('D')" > �ٸ���ǰ(����/���� ����)
			<% end if %>
			<input type="radio" name="applyToAddItem" value="C" onClick="applyPriceToAddItem('C')" > CS����
			<input type="radio" name="applyToAddItem" value="S" onClick="applyPriceToAddItem('S')" > �����ǸŰ�
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
			��������
		</td>
		<td colspan=2>
                <input class='text' type="text" name="title" value="<%= title %>" size="56" maxlength="56">
		</td>
	</tr>

	<tr bgcolor="#FFFFFF">
		<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
			��������
		</td>
		<td colspan=2>
                <textarea class='textarea' name="contents_jupsu" cols="68" rows="6"><%= contents_jupsu %></textarea>
		</td>
	</tr>

	<% end if %>
	<tr bgcolor="#FFFFFF" height="40">
		<td colspan="3" align="center">
<% if (ojumunDetail.FJumunDetail.Fcancelyn <> "Y" and ojumun.FOneItem.Fcancelyn <> "Y") and itemChangeAvail = True and ojumunDetail.FJumunDetail.Fitemno > 0 then %>
			<input type="button" class="button" value="��ǰ����" style="width:100px; height:22px;" onclick="javascript:SaveChangeItem(false)">
			<% if C_ADMIN_AUTH or (C_CSPowerUser) or (posit_sn<=11 and C_CSUser) then %>
		    <input type="button" class="button" value="��������" style="width:100px; height:22px;" onclick="javascript:SaveChangeItem(true)">
			<% end if %>
			&nbsp;
			<input type="button" class="button" value="��ȯ���(��ǰ����)" style="width:130px; height:22px;" onclick="javascript:SaveChangeOrder(false)">
			<% if C_ADMIN_AUTH or (C_CSPowerUser) or (posit_sn<=11 and C_CSUser) then %>
			<input type="button" class="button" value="��ȯ���(��������)" style="width:130px; height:22px;" onclick="javascript:SaveChangeOrder(true)">
			<% end if %>
<% elseif (ojumunDetail.FJumunDetail.Fcancelyn = "Y" or ojumun.FOneItem.Fcancelyn = "Y") then %>
			<b>��ҵ� ��ǰ�� ��ǰ���� �Ұ�</b>
<% elseif ojumunDetail.FJumunDetail.Fitemno < 0 then %>
			<b>���̳ʽ� �ֹ� ��ǰ���� �Ұ�</b>
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
