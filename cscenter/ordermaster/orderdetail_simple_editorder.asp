<%@ language=vbscript %>
<% option explicit %>
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
dim i, sqlStr
dim idx, orderserial, targetdetailidx, targetregitemno, toItemId, toItemOption
dim result
dim salemethod

orderserial = requestCheckVar(request("orderserial"),32)
targetdetailidx = requestCheckVar(request("targetdetailidx"),32)
targetregitemno = requestCheckVar(request("targetregitemno"),32)
toItemId = requestCheckVar(request("toItemId"),32)
toItemOption = requestCheckVar(request("toItemOption"),32)
salemethod = requestCheckVar(request("salemethod"),32)


'==============================================================================
''1. �ֹ� ����Ÿ
'==============================================================================
dim oordermaster, IsOrderCanceled, IsChangeOrder

set oordermaster = new COrderMaster

oordermaster.FRectOrderSerial = orderserial

if Left(orderserial,1)="A" then
    set oordermaster.FOneItem = new COrderMasterItem
else
    oordermaster.QuickSearchOrderMaster
end if

'' ���� 6���� ���� ���� �˻�
if (oordermaster.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oordermaster.FRectOldOrder = "on"
    oordermaster.QuickSearchOrderMaster
end if

IsOrderCanceled = (oordermaster.FOneItem.Fcancelyn = "Y")
IsChangeOrder   = (oordermaster.FOneItem.FjumunDiv="6")


'==============================================================================
dim oorderdetail

set oorderdetail = new COrderMaster
oorderdetail.FRectOrderSerial = orderserial
oorderdetail.QuickSearchOrderDetail

'' ���� 6���� ���� ���� �˻�
if (oorderdetail.FResultCount<1) and (Len(orderserial)=11) and (IsNumeric(orderserial)) then
    oorderdetail.FRectOldOrder = "on"
    oorderdetail.QuickSearchOrderDetail
end if


'==============================================================================
dim divcd
dim title
dim contents_jupsu
dim prevregno

dim oupchebeasongpay
dim upchebeasongpay
dim isupchebeasong, requiremakerid, detailstate


'==============================================================================
dim isitemselected

'==============================================================================
'2. �߰���ǰ
'==============================================================================
dim itemid, itemoption, makerid, itemname, itemoptionname, orgitemcostprice, imagesmall
dim sellcash, optaddprice, buycash, optaddbuyprice
dim issaleitem
dim mwdiv

dim fromItemId, fromItemOption
dim issameitemcost, issamemakerid, isitemcouponapplied, ispercentcouponapplied, itemcouponidxapplied, bonuscouponidxapplied
dim ItemCouponType, ItemCouponValue, couponbuyprice, iscouponapplyOK
dim addorgitemcost, additemcost, addpercentBonusCouponDiscount
dim fromitemcost

dim isregchangeorderOK
dim itemstate

dim add_SalePrice, add_ItemCouponPrice, add_BonusCouponPrice, add_buycash


toItemId = requestCheckVar(request("toItemId"),32)
toItemOption = requestCheckVar(request("toItemOption"),32)


issamemakerid = False
issameitemcost = False
isitemcouponapplied = False
ispercentcouponapplied = False
iscouponapplyOK = False
isregchangeorderOK = False
itemstate = ""

if (IsChangeOrder) then
	isregchangeorderOK = True
end if

if (toItemId <> "") and (toItemOption <> "") then

	'==============================================================================
	'// a. ���(ȸ��) ��ǰ����
	'==============================================================================
	for i = 0 to oorderdetail.FResultCount - 1
		if (CStr(targetdetailidx) = CStr(oorderdetail.FItemList(i).Fidx)) then

			fromItemId = oorderdetail.FItemList(i).Fitemid
			fromItemOption = oorderdetail.FItemList(i).Fitemoption

			isitemcouponapplied = Not IsNull(oorderdetail.FItemList(i).Fitemcouponidx)
			itemcouponidxapplied = oorderdetail.FItemList(i).Fitemcouponidx
			ispercentcouponapplied = Not IsNull(oorderdetail.FItemList(i).Fbonuscouponidx)
			bonuscouponidxapplied = oorderdetail.FItemList(i).Fbonuscouponidx

			fromitemcost = oorderdetail.FItemList(i).Fitemcost - (oorderdetail.FItemList(i).getAllAtDiscountedPrice + oorderdetail.FItemList(i).getPercentBonusCouponDiscountedPrice)

			isupchebeasong = oorderdetail.FItemList(i).Fisupchebeasong
			requiremakerid = oorderdetail.FItemList(i).Fmakerid
			detailstate = oorderdetail.FItemList(i).Fcurrstate

			'// TODO : ��ǰ ��ۻ��°� �ٸ� ��� üũ(����ǰ�� ������� ��ǰ�� ���ÿ� ����� �� ����.)
			itemstate = oorderdetail.FItemList(i).Fcurrstate

		end if
	next


	'==============================================================================
	'// b. ���� ��������
	'==============================================================================
	if (fromItemId <> "") and (isitemcouponapplied or ispercentcouponapplied) then

		if (isitemcouponapplied) then
			'��ǰ����
			sqlStr = " select top 1 "
			sqlStr = sqlStr + " 	IsNull(i.ItemCouponType, '1') as ItemCouponType "
			sqlStr = sqlStr + " 	, IsNull(i.ItemCouponValue, 0) as ItemCouponValue "
			sqlStr = sqlStr + " 	, IsNull(d.couponbuyprice, 0) as couponbuyprice "
			sqlStr = sqlStr + " from "

			sqlStr = sqlStr + " 	db_item.dbo.tbl_item i "
			sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item_coupon_master m "

			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		m.itemcouponidx = i.CurrItemCouponIdx "
			sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item_coupon_detail d "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and m.itemcouponidx = d.itemcouponidx "
			sqlStr = sqlStr + " 		and d.itemid = i.itemid "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and d.itemid = " + CStr(fromItemId) + " "

			if (isitemcouponapplied) then
				sqlStr = sqlStr + " 	and m.itemcouponidx = " + CStr(itemcouponidxapplied) + " "
			else
				sqlStr = sqlStr + " 	and m.itemcouponidx = " + CStr(bonuscouponidxapplied) + " "
			end if

			'response.write sqlStr
			rsget.Open sqlStr,dbget,1

			if not rsget.Eof then
				ItemCouponType 				= rsget("ItemCouponType")
				ItemCouponValue 			= rsget("ItemCouponValue")
				couponbuyprice 				= rsget("couponbuyprice")
			end if
			rsget.Close
		end if

		if (ispercentcouponapplied) then
			'��������
			sqlStr = " select top 1 "
			sqlStr = sqlStr + " 	IsNull(c.CouponType, '1') as ItemCouponType "
			sqlStr = sqlStr + " 	, IsNull(c.CouponValue, 0) as ItemCouponValue "
			sqlStr = sqlStr + " 	, IsNull(c.minbuyprice, 0) as couponbuyprice "
			sqlStr = sqlStr + " from "

			sqlStr = sqlStr + " 	db_user.dbo.tbl_user_coupon c "
			sqlStr = sqlStr + " where idx = " + CStr(bonuscouponidxapplied) + " and coupontype = 1 "

			'response.write sqlStr
			rsget.Open sqlStr,dbget,1

			if not rsget.Eof then
				ItemCouponType 				= rsget("ItemCouponType")
				ItemCouponValue 			= rsget("ItemCouponValue")
				couponbuyprice 				= rsget("couponbuyprice")
			end if
			rsget.Close
		end if

	end if


	'==============================================================================
	'// c. �߰��Ǵ� ��ǰ����
	'==============================================================================
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
	rsget.Open sqlStr,dbget,1

	if not rsget.Eof then
		itemid 				= rsget("itemid")
		itemoption 			= rsget("itemoption")

		makerid 			= rsget("makerid")
		itemname 			= db2Html(rsget("itemname"))
		itemoptionname 		= db2Html(rsget("optionname"))

		orgitemcostprice	= rsget("orgprice") + rsget("optaddprice")

		sellcash			= rsget("sellcash")
		optaddprice			= rsget("optaddprice")

		addorgitemcost		= sellcash + optaddprice
		additemcost			= sellcash + optaddprice

		add_SalePrice			= sellcash + optaddprice
		add_ItemCouponPrice		= sellcash + optaddprice
		add_BonusCouponPrice	= sellcash + optaddprice
		add_buycash				= rsget("buycash")

		if (salemethod = "C") then
			addorgitemcost = fromitemcost
			additemcost = fromitemcost

			add_SalePrice = fromitemcost
			add_ItemCouponPrice = fromitemcost
			add_BonusCouponPrice = fromitemcost
		end if

		issaleitem			= rsget("sailyn")

		buycash 			= rsget("buycash")
		optaddbuyprice		= rsget("optaddbuyprice")

		imagesmall 			= "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemID(toItemId) + "/" + rsget("listimage")

		mwdiv				= rsget("mwdiv")
	end if
	rsget.Close

	'==============================================================================
	'// d. ��ǰ��
	'==============================================================================
	sqlStr = " select top 1 "
	sqlStr = sqlStr + " 	i.makerid "
	sqlStr = sqlStr + " 	, i.sellcash "
	sqlStr = sqlStr + " 	, i.orgprice "
	sqlStr = sqlStr + " 	, i.buycash "
	sqlStr = sqlStr + " 	, IsNull(o.optaddprice, 0) as optaddprice "
	sqlStr = sqlStr + " 	, IsNull(o.optaddbuyprice, 0) as optaddbuyprice "
	sqlStr = sqlStr + " from "
	sqlStr = sqlStr + " 	db_item.dbo.tbl_item i "
	sqlStr = sqlStr + " 	left join dbo.tbl_item_option o "
	sqlStr = sqlStr + " on "
	sqlStr = sqlStr + " 	i.itemid = o.itemid "
	sqlStr = sqlStr + " where "
	sqlStr = sqlStr + " 	1 = 1 "
	sqlStr = sqlStr + " 	and i.itemid = " & CStr(fromItemId) & " "
	sqlStr = sqlStr + " 	and IsNull(o.itemoption, '0000') = '" & CStr(fromItemOption) & "' "
	'response.write sqlStr
	rsget.Open sqlStr,dbget,1

	if not rsget.Eof then

		'// ���� �ǸŰ�/���԰� ����
		issameitemcost = ((rsget("sellcash") + rsget("optaddprice")) = (sellcash + optaddprice)) and ((rsget("buycash") + rsget("optaddbuyprice")) = (buycash + optaddbuyprice))

		issamemakerid = (rsget("makerid") = makerid)

		if (detailstate = "7") then
			if issameitemcost and issamemakerid then
				isregchangeorderOK = True
			end if
		end if

	end if
	rsget.Close
'rw issameitemcost
'rw issamemakerid
'rw salemethod
	'==============================================================================
	'// e. ���� ���밡������
	'==============================================================================
	iscouponapplyOK = False
	addpercentBonusCouponDiscount = 0
	if (fromItemId <> "") and (isitemcouponapplied or ispercentcouponapplied) then

		if (isitemcouponapplied) then
			'��ǰ����
			sqlStr = " select top 1 "
			sqlStr = sqlStr + " 	IsNull(i.ItemCouponType, '1') "
			sqlStr = sqlStr + " from "

			sqlStr = sqlStr + " 	db_item.dbo.tbl_item i "
			sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item_coupon_master m "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		m.itemcouponidx = i.CurrItemCouponIdx "
			sqlStr = sqlStr + " 	left join db_item.dbo.tbl_item_coupon_detail d "
			sqlStr = sqlStr + " 	on "
			sqlStr = sqlStr + " 		1 = 1 "
			sqlStr = sqlStr + " 		and m.itemcouponidx = d.itemcouponidx "
			sqlStr = sqlStr + " 		and d.itemid = i.itemid "
			sqlStr = sqlStr + " where "
			sqlStr = sqlStr + " 	1 = 1 "
			sqlStr = sqlStr + " 	and d.itemid = " + CStr(toItemId) + " "
			sqlStr = sqlStr + " 	and m.itemcouponidx = " + CStr(itemcouponidxapplied) + " "
			sqlStr = sqlStr + " 	and m.itemcouponidx = i.CurrItemCouponIdx "
			sqlStr = sqlStr + " 	and i.ItemCouponYn = 'Y' "
			sqlStr = sqlStr + " 	and m.itemcouponstartdate <= getdate() "
			sqlStr = sqlStr + " 	and m.itemcouponexpiredate > getdate() "

			'response.write sqlStr
			rsget.Open sqlStr,dbget,1

			if not rsget.Eof then
				iscouponapplyOK = True
			end if
			rsget.Close

			if (iscouponapplyOK) and (salemethod <> "C") then

				if (CStr(ItemCouponType) = "1") then
					'����
					additemcost = (sellcash - (sellcash * ItemCouponValue) / 100) + optaddprice
				elseif (CStr(ItemCouponType) = "2") then
					'����
					additemcost = additemcost - ItemCouponValue
				end if

				if (CStr(couponbuyprice) <> "0") then
					buycash = couponbuyprice
				end if

				add_ItemCouponPrice = additemcost
				add_BonusCouponPrice = add_ItemCouponPrice
				add_buycash = buycash

			end if

		end if

		if (ispercentcouponapplied) then
			'��������(��������, ��ۺ������� ��ǰ���ݿ� ������ ���� �����Ƿ� ������� �ʴ´�.)
			'���������� ��ȿ�Ⱓ�̳� �귣�带 ������� �ʰ� ������ �������ش�.(���� �ǸŰ� �� ���԰� ������ ���)

			iscouponapplyOK = issameitemcost

			if (iscouponapplyOK) and (salemethod <> "C") then
				addpercentBonusCouponDiscount = (sellcash * ItemCouponValue) / 100
				additemcost = (sellcash - (sellcash * ItemCouponValue) / 100) + optaddprice

				add_BonusCouponPrice = additemcost
			end if

		end if

	end if


	if Not isregchangeorderOK then
		'// �ֹ���������
		divcd = "A900"
		if (salemethod = "R") then
			title = "�ֹ�����(����ȯ��)"
		elseif (salemethod = "C") then
			title = "�ֹ�����(CS����)"
		else
			if (oordermaster.FOneItem.FIpkumDiv < "4") then
				title = "�ֹ�����(�����Ϸ�����)"
			else
				title = "�ֹ�����(�����ǸŰ�)"
			end if
		end if
	end if

	set oupchebeasongpay = new COrderMaster
	upchebeasongpay = 2000

	if (orderserial <> "") and (isupchebeasong = "Y") then
		oupchebeasongpay.FRectOrderSerial = orderserial
		oupchebeasongpay.getUpcheBeasongPayList

		for i = 0 to oupchebeasongpay.FResultCount - 1
			if (oupchebeasongpay.FItemList(i).Fmakerid = requiremakerid) then
				'// ��ü����̸� ��ü �⺻��ۺ� ��������
				upchebeasongpay = oupchebeasongpay.FItemList(i).Fdefaultdeliverpay
			end if
		next

		if (upchebeasongpay = 0) then
			'// XXXX ��ü�������̸� ���ٹ�ۺ�� ����
			'�⺻��ۺ� ���� �ʵǾ� ������ 2500��(since 2012-06-18)
			upchebeasongpay = 2500
		end if
	end if

	'// �⺻���� ����
	if Not IsNull(session("ssBctCname")) then
		contents_jupsu = "�ٹ����� ������ " + CStr(session("ssBctCname")) + " �Դϴ�"
	end if

end if

%>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language='javascript'>

// ��������(ajax) �� ����ϱ� ���� �ʿ�
var IsPossibleModifyCSMaster = true;
var IsPossibleModifyItemList = true;
var IsCSReturnProcess = false;
var IsCSCancelProcess = false;

var IsChangeOrder = <%= LCase(IsChangeOrder) %>;

var arrdivcd, arrorderdetailidx, arrdetailstate;
var arritemid, arritemcost, arrallatitemdiscount, arrpercentBonusCouponDiscount, arrregitemno, arritemcostCouponNotApplied;
var arrisupchebeasong, arrmakerid;
var arrregitemno;

var arradd_orderdetailidx, arradd_itemcost, arradd_regitemno, arradd_percentBonusCouponDiscount;

function getOnload() {
	arrdivcd 				= document.getElementsByName("divcd");
	arrorderdetailidx 		= document.getElementsByName("orderdetailidx");
	arrdetailstate 			= document.getElementsByName("detailstate");

	arritemid 						= document.getElementsByName("itemid");
	arritemcost 					= document.getElementsByName("itemcost");
	arritemcostCouponNotApplied 	= document.getElementsByName("itemcostCouponNotApplied");
	arrallatitemdiscount 			= document.getElementsByName("allatitemdiscount");
	arrpercentBonusCouponDiscount 	= document.getElementsByName("percentBonusCouponDiscount");
	arrregitemno 					= document.getElementsByName("regitemno");

	arrisupchebeasong 	= document.getElementsByName("isupchebeasong");
	arrmakerid 			= document.getElementsByName("makerid");

	arrregitemno = document.getElementsByName("regitemno");

	arradd_orderdetailidx 				= document.getElementsByName("add_orderdetailidx");
	arradd_itemcost 					= document.getElementsByName("add_itemcost");
	arradd_regitemno 					= document.getElementsByName("add_regitemno");
	arradd_percentBonusCouponDiscount 	= document.getElementsByName("add_percentBonusCouponDiscount");

	<% if (targetdetailidx <> "") then %>
		OneItemSelected(frm);
		CalcAddedItemCost();
	<% end if %>

}
window.onload = getOnload;


/* ============================================================================
��ǰ����

 - 1. �ϳ��� ��ǰ�� ���ð���(������ ������ ����)
========================================================================== */
function OneItemSelected(frm) {
	var checkeditemexist = false;
	var selectedindex = -1;

	frm.targetisupchebeasong.value = "";
	frm.targetmakerid.value = "";
	frm.targetitemid.value = "";
	frm.targetitemcost.value = 0;

	frm.targetdetailidx.value = "";
	frm.targetregitemno.value = "";

	for (var i = 0; i < arrorderdetailidx.length; i++) {
		if (arrorderdetailidx[i].checked == true) {

			checkeditemexist = true;
			selectedindex = i;
			frm.targetisupchebeasong.value 	= arrisupchebeasong[i].value;
			frm.targetmakerid.value 		= arrmakerid[i].value;
			frm.targetitemcost.value 		= arritemcostCouponNotApplied[i].value;
			frm.targetitemid.value 			= arritemid[i].value;

			frm.targetdetailidx.value 		= arrorderdetailidx[i].value;
			frm.targetregitemno.value 		= arrregitemno[i].value;

			break;
		}
	}

	CalcSelectedItemCost();

	if (checkeditemexist == true) {
		// �ϳ��� ��ǰ�� ���ð���(������ ������ ����)
		__ShowOnlySelectedItem();
	} else {
		__ShowAllItem();

		for (var i = 0; i < arrdivcd.length; i++) {
			arrdivcd[i].checked = false;
		}

		frm.title.value = "";

		return;
	}
}

// newcsas.js �� �Լ��� �ߺ��� �߻��� ���� �߰����ش�.
function __CheckMaxItemNo(obj, maxno) {
	if (obj.value*0 != 0) {
		return;
	}

    if (obj.value*1 > maxno*1) {
        alert("�ֹ����� �̻����� ��ǰ������ �����Ҽ� �����ϴ�.");
        obj.value = maxno;
    }

	if (frm.targetdetailidx.value != "") {
		frm.targetregitemno.value = obj.value;
	}

	CalcSelectedItemCost();
	if (frm.add_totalselecteditemcost) {
		frm.add_regitemno.value = obj.value;
		CalcAddedItemCost();
	}
}

function CancelSelectItem() {
	var frm = document.frm;

	document.location.href = "orderdetail_simple_editorder.asp?orderserial=" + frm.orderserial.value
}

function SearchItemByMakerid() {
	var frm = document.frm;
	var isupchebeasong;
	var makerid;
	var excludeupbae;

	isupchebeasong = frm.targetisupchebeasong.value;
	makerid = frm.targetmakerid.value;

	if (makerid == "") {
		alert("���� ���(ȸ��)�� ��ǰ�� �����ϼ���.");
		return;
	}

	if (isupchebeasong == "N") {
		excludeupbae = "on";
	} else {
		excludeupbae = "";
	}

	var popwin = window.open('pop_item_search_one.asp?makerid=' + makerid + '&onlineonly=Y&nubeasong=' + excludeupbae,'SearchItemByMakerid','width=1000,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function SearchItemByPrice() {
	var frm = document.frm;
	var isupchebeasong, makerid, itemcost;
	var excludeupbae;

	isupchebeasong = frm.targetisupchebeasong.value;
	makerid = frm.targetmakerid.value;
	itemcost = frm.targetitemcost.value;

	if (isupchebeasong == "") {
		alert("���� ���(ȸ��)�� ��ǰ�� �����ϼ���.");
		return;
	}

	if (isupchebeasong == "N") {
		excludeupbae = "on";
	} else {
		excludeupbae = "";
	}

	var popwin = window.open('pop_item_search_one.asp?makerid=' + makerid + '&saleprice=' + itemcost + '&onlineonly=Y&nubeasong=' + excludeupbae,'SearchItemByPrice','width=1000,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ReActItemOne(toItemId, toItemOption) {
	var frm = document.frm;

	if (IsSameItemExist(toItemId, toItemOption) == true) {
		return;
	}

	document.location.href = "orderdetail_simple_editorder.asp?orderserial=" + frm.orderserial.value + "&targetdetailidx=" + frm.targetdetailidx.value + "&targetregitemno=" + frm.targetregitemno.value + "&toItemId=" + toItemId + "&toItemOption=" + toItemOption;

	return;
}

function ChangeSaleMethod(salemethod) {
	var frm = document.frm;

	document.location.href = "orderdetail_simple_editorder.asp?orderserial=" + frm.orderserial.value + "&targetdetailidx=" + frm.targetdetailidx.value + "&targetregitemno=" + frm.targetregitemno.value + "&toItemId=<%= toItemId %>&toItemOption=<%= toItemOption %>" + "&salemethod=" + salemethod;

	return;
}

function IsSameItemExist(ItemId, ItemOption) {
	var frm = document.frm;

	if (frm.targetitemid.value*1 == ItemId*1) {
		alert("���ϻ�ǰ�� �ֽ��ϴ�. �ɼǺ����� �̿��ϼ���.");
		return true;
	}

	return false;
}

function CalcSelectedItemCost() {
	var frm = document.frm;

	var cancelitemcostsum = 0;
	var cancelcouponsum = 0;
	var cancelallatsubtractsum = 0;

	for (var i = 0; i < arrorderdetailidx.length; i++) {
		if (arrorderdetailidx[i].checked == true) {
			cancelitemcostsum 		= cancelitemcostsum + (arritemcost[i].value*1 * arrregitemno[i].value*1)
			cancelcouponsum 		= cancelcouponsum + (arrpercentBonusCouponDiscount[i].value*1 * arrregitemno[i].value*1)
			cancelallatsubtractsum 	= cancelallatsubtractsum + (arrallatitemdiscount[i].value*1 * arrregitemno[i].value*1)
		}
	}

	frm.cancelitemcostsum.value = cancelitemcostsum;
	frm.cancelcouponsum.value = cancelcouponsum;
	frm.cancelallatsubtractsum.value = cancelallatsubtractsum;

	frm.totalselecteditemcost.value = cancelitemcostsum - (cancelcouponsum + cancelallatsubtractsum);
}

function CalcAddedItemCost() {
	var frm = document.frm;

	var additemcostsum = 0;
	var addcouponsum = 0;
	var addallatsubtractsum = 0;

	for (var i = 0; i < arradd_orderdetailidx.length; i++) {
		if (arradd_orderdetailidx[i].checked == true) {
			additemcostsum 		= additemcostsum + (arradd_itemcost[i].value*1 * arradd_regitemno[i].value*1)
			addcouponsum 		= addcouponsum + (arradd_percentBonusCouponDiscount[i].value*1 * arradd_regitemno[i].value*1)

			// �������
			addallatsubtractsum	= 0
		}
	}

	frm.additemcostsum.value = additemcostsum + addcouponsum;
	frm.addcouponsum.value = addcouponsum;
	frm.addallatsubtractsum.value = addallatsubtractsum;

	frm.add_totalselecteditemcost.value = additemcostsum;

	// ����
	frm.totaldiffitemcost.value = frm.totalselecteditemcost.value*1 - frm.add_totalselecteditemcost.value*1;

	frm.refunditemcostsum.value = frm.cancelitemcostsum.value*1 - frm.additemcostsum.value*1;
	frm.refundcouponsum.value = frm.cancelcouponsum.value*1 - frm.addcouponsum.value*1;
	frm.refundallatsubtractsum.value = frm.cancelallatsubtractsum.value*1 - frm.addallatsubtractsum.value*1;

	frm.canceltotal.value = frm.refunditemcostsum.value*1 - frm.refundcouponsum.value*1 - frm.refundallatsubtractsum.value*1;
	frm.refundrequire.value = frm.canceltotal.value;
}

// newcsas.js �� �Լ��� �ߺ��� �߻��� ���� �߰����ش�.
function __ShowOnlySelectedItem() {
    var e, t;

    for (var i = 0; i < arrorderdetailidx.length; i++) {
        e = arrorderdetailidx[i];
        t = arrorderdetailidx[i];

        if (e.type == "checkbox") {
			while (t.tagName != "TR") {
				t = t.parentElement;
			}

			if (e.checked == true) {
				t.style.display = '';
			} else {
				t.style.display = 'none';
			}
        }
    }
}

// newcsas.js �� �Լ��� �ߺ��� �߻��� ���� �߰����ش�.
function __ShowAllItem() {
    var e, t;

    for (var i = 0; i < arrorderdetailidx.length; i++) {
        e = arrorderdetailidx[i];
        t = arrorderdetailidx[i];

        if (e.type == "checkbox") {
			while (t.tagName != "TR") {
				t = t.parentElement;
			}

			t.style.display = '';
        }
    }
}

function ChangeDivCD(frm) {
	var divcd;
	var salemethod = "<%= salemethod %>";
	var ipkumdiv = "<%= oordermaster.FOneItem.FIpkumDiv %>";
	var title;

	for (var i = 0; i < arrdivcd.length; i++) {
		if (arrdivcd[i].checked == true) {
			divcd = arrdivcd[i].value;
			break;
		}
	}

	SetAddBeasongPay()

	if (divcd == "A900") {
		title = "�ֹ�����";

		if (salemethod == "R") {
			frm.title.value = title + "(����ȯ��)";
		} else if (salemethod == "C") {
			frm.title.value = title + "(CS����)";
		} else {
			if (ipkumdiv < "4") {
				frm.title.value = title + "(�����Ϸ�����)";
			} else {
				frm.title.value = title + "(�����ǸŰ�)";
			}
		}
	} else if (divcd == "A100") {
		title = "��ǰ���� ��ȯ���";

		if (salemethod == "R") {
			frm.title.value = title + "(ERROR)";
		// } else if (salemethod == "C") {
		// 	frm.title.value = title + "(ERROR)";
		} else {
			if ((ipkumdiv < "7") && (IsChangeOrder != true)) {
				frm.title.value = title + "(ERROR)";
			} else {
				frm.title.value = title + "(�����ǸŰ�)";
			}
		}
	}
}

function SetAddBeasongPay() {
	var frm = document.frm;
	var divcd;

	for (var i = 0; i < arrdivcd.length; i++) {
		if (arrdivcd[i].checked == true) {
			divcd = arrdivcd[i].value;
			break;
		}
	}

	if (divcd != "A100") {
		frm.add_customeraddbeasongpay.value = 0;
		frm.add_customeraddmethod.value = "";
		return;
	}

	if (!frm.isupchebeasong) {
		return;
	}

	if ((frm.gubun01.value == "C004") && (frm.gubun02.value == "CD01")) {
		// �ܼ�����
		frm.add_customeraddbeasongpay.value = frm.upchebeasongpay.value*2;
		frm.add_customeraddmethod.value = "1";
	} else {
		frm.add_customeraddbeasongpay.value = 0;
		frm.add_customeraddmethod.value = "";
	}
}

function SaveChangeItem(isadmin) {
    var frm = document.frm;

	var divcd = "";
	var salemethod = "<%= salemethod %>";
	var ipkumdiv = "<%= oordermaster.FOneItem.FIpkumDiv %>";
	var itemstate = "<%= itemstate %>";



	var issamemakerid = <%= LCase(issamemakerid) %>;
	var issameitemcost = <%= LCase(issameitemcost) %>;

	var isitemcouponapplied = <%= LCase(isitemcouponapplied) %>;
	var ispercentcouponapplied = <%= LCase(ispercentcouponapplied) %>;
	var iscouponapplyOK = <%= LCase(iscouponapplyOK) %>;

	var errorMSG, adminErrorMSG;

	if (issameitemcost != true) {
		alert("�ǸŰ�,���԰��� �ٸ���� ����� �� �����ϴ�.\n\n[������]�� ��� ��ǰ���� �����ϰ� ��ǰ���� �ϼ���.");
		return;
	}

	for (var i = 0; i < arrdivcd.length; i++) {
		if (arrdivcd[i].checked == true) {
			divcd = arrdivcd[i].value;
			break;
		}
	}

	if (divcd == "") {
		alert("���������� �����ϼ���.");
		return;
	}

	if (frm.gubun01.value == "") {
		alert("���������� �����ϼ���.");
		return;
	}

    errorMSG = "";
    adminErrorMSG = "";

	if (divcd == "A100") {
		if ((itemstate < "7") && (IsChangeOrder != true)) {
			adminErrorMSG = adminErrorMSG + "\n - ����� ��ǰ�Դϴ�..[��ϺҰ�]";
		}

		if (issamemakerid != true) {
			adminErrorMSG = adminErrorMSG + "\n - �귣�尡 �ٸ��ϴ�.[��ϺҰ�]";
		}

		if (issameitemcost != true) {
			adminErrorMSG = adminErrorMSG + "\n - ���� �ǸŰ�(���ΰ�) �Ǵ� ���԰��� �ٸ��ϴ�.[��ϺҰ�]";
		}

		if ((isitemcouponapplied == true) || (ispercentcouponapplied == true)) {
			if (iscouponapplyOK != true) {
				adminErrorMSG = adminErrorMSG + "\n - �߰���ǰ�� ������ ������ �� �����ϴ�.[��ϺҰ�]";
			}
		}

		if (salemethod != "") {
			errorMSG = errorMSG + "\n - ����ȯ�� �Ǵ� CS������ �����ڸ� �����մϴ�.";
		}
	} else {
		if (<%= LCase(IsChangeOrder) %> == true) {
			adminErrorMSG = adminErrorMSG + "\n - ��ȯ�ֹ��� ��ǰ���� �� �� �����ϴ�.[��ϺҰ�]";
		}

		if (ipkumdiv < "4") {
			errorMSG = errorMSG + "\n - �����Ϸ� ���� �ֹ��Դϴ�.";
		}

		if (itemstate == "7") {
			errorMSG = errorMSG + "\n - �̹� ���� ��ǰ�Դϴ�.";
		}

		if (issamemakerid != true) {
			errorMSG = errorMSG + "\n - �귣�尡 �ٸ��ϴ�.";
		}

		if (issameitemcost != true) {
			errorMSG = errorMSG + "\n - ���� �ǸŰ�(���ΰ�) �Ǵ� ���԰��� �ٸ��ϴ�.";
		}

		if ((isitemcouponapplied == true) || (ispercentcouponapplied == true)) {
			if (iscouponapplyOK != true) {
				errorMSG = errorMSG + "\n - �߰���ǰ�� ������ ������ �� �����ϴ�.";
			}
		}

		if (salemethod != "") {
			errorMSG = errorMSG + "\n - ����ȯ�� �Ǵ� CS������ �����ڸ� �����մϴ�.";
		}
	}

	if (adminErrorMSG != "") {
		alert("��ϺҰ�!!\n" + adminErrorMSG);
		return;
	}

	if (errorMSG != "") {
		if (isadmin == true) {
			if (confirm("������ ���� : \n" + errorMSG + "\n\n�����Ͻðڽ��ϱ�?") != true) {
				return;
			}
		} else {
			alert("�����ڹ���!!\n" + errorMSG);
			return;
		}
	}


	if ((isadmin == true) || (confirm("����Ͻðڽ��ϱ�?") == true)) {
		if (divcd == "A100") {
			frm.mode.value = "regchangeorder"
		} else {
			frm.mode.value = "regmodifyorder"
		}

		frm.submit();
	}
}

</script>
<script language='javascript' SRC="/js/ajax.js"></script>
<script language='javascript' SRC="/cscenter/js/newcsas.js"></script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ����������� ����</b>
		</td>
	</tr>

	<form name="frm" method="post" action="orderdetail_simple_editorder_process.asp">
	<input type="hidden" name="orderserial" value="<%= orderserial %>">
	<input type="hidden" name="mode" value="">

	<input type="hidden" name="targetisupchebeasong" value="">
	<input type="hidden" name="targetmakerid" value="">
	<input type="hidden" name="targetitemid" value="">
	<input type="hidden" name="targetitemcost" value=""><!-- ����������ݾ�, �߰���ǰ �˻��� -->

	<input type="hidden" name="targetdetailidx" value="">
	<input type="hidden" name="targetregitemno" value="">

	<input type="hidden" name="cancelitemcostsum" value="0">
	<input type="hidden" name="cancelcouponsum" value="0">
	<input type="hidden" name="cancelallatsubtractsum" value="0">

	<input type="hidden" name="additemcostsum" value="0">
	<input type="hidden" name="addcouponsum" value="0">
	<input type="hidden" name="addallatsubtractsum" value="0">

	<input type="hidden" name="refunditemcostsum" value="0">
	<input type="hidden" name="refundcouponsum" value="0">
	<input type="hidden" name="refundallatsubtractsum" value="0">

	<input type="hidden" name="canceltotal" value="0">
	<input type="hidden" name="refundrequire" value="0">

	<tr height="25" bgcolor="#FFFFFF" >
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
		<td><%= orderserial %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">������</td>
		<td><%= oordermaster.FOneItem.JumunMethodName %></td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">�ŷ�����</td>
		<td><%= oordermaster.FOneItem.IpkumDivName %></td>
	</tr>
</table>

<p>

<!-- ====================================================================== -->



<!-- ====================================================================== -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>">
			<b>1. ���(ȸ��) ����ǰ ����</b>
			<% if (toItemId <> "") and (toItemOption <> "") then %>
				<input type="button" class="button" value="��ǰ��������" onClick="javascript:CancelSelectItem()">
			<% end if %>
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF" >
		<td>
            <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
			<tr height="20" align="center" bgcolor="#F4F4F4">
				<td width="30">����</td>
				<td width="50">�̹���</td>
				<td width="30">����</td>
				<td width="50">������</td>
				<td width="50">��ǰ�ڵ�</td>
				<td width="90">�귣��ID</td>
				<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
				<td width="80">����/���ֹ�</td>
				<td width="60">�ǸŰ�<br>(���ΰ�)</td>
				<td width="60">������</td>
				<td width="60">���԰�</td>
				<td width="100">���</td>
			</tr>
<% for i = 0 to oorderdetail.FResultCount - 1 %>
	<% if (oorderdetail.FItemList(i).Fitemid <> 0) then %>
		<%
		isitemselected = "N"
		if (CStr(targetdetailidx) = CStr(oorderdetail.FItemList(i).Fidx)) then
			isitemselected = "Y"
		end if

		%>
		<% if (oorderdetail.FItemList(i).FCancelyn = "Y") or (oorderdetail.FItemList(i).Fitemno < 1) then %>
			<tr align="center" bgcolor='#DDDDFF' class='gray'>
		<% else %>
			<tr align="center" bgcolor='#FFFFFF' >
		<% end if %>

				<input type="hidden" name="detailstate" value="<%= oorderdetail.FItemList(i).Fcurrstate %>">
				<input type="hidden" name="isupchebeasong" value="<%= oorderdetail.FItemList(i).Fisupchebeasong %>">
				<input type="hidden" name="makerid" value="<%= oorderdetail.FItemList(i).Fmakerid %>">
				<input type="hidden" name="itemid" value="<%= oorderdetail.FItemList(i).Fitemid %>">
				<input type="hidden" name="itemcost" value="<%= oorderdetail.FItemList(i).Fitemcost %>">
				<input type="hidden" name="allatitemdiscount" value="<%= oorderdetail.FItemList(i).getAllAtDiscountedPrice %>">
				<input type="hidden" name="percentBonusCouponDiscount" value="<%= oorderdetail.FItemList(i).getPercentBonusCouponDiscountedPrice %>">
				<input type="hidden" name="itemcostCouponNotApplied" value="<%= oorderdetail.FItemList(i).FitemcostCouponNotApplied %>">

				<td height="25">
					<input type="checkbox" name="orderdetailidx" onClick="OneItemSelected(frm)" value="<%= oorderdetail.FItemList(i).Fidx %>" <% if (oorderdetail.FItemList(i).FCancelyn = "Y") or (oorderdetail.FItemList(i).Fitemno < 1) then %>disabled<% end if %> <% if isitemselected = "Y" then %>checked disabled<% end if %>>
				</td>
				<td width="50"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oorderdetail.FItemList(i).Fitemid %>" target="_blank"><img src="<%= oorderdetail.FItemList(i).FSmallImage %>" width="50" border="0"></a></td>
				<td><font color="<%= oorderdetail.FItemList(i).CancelStateColor %>"><%= oorderdetail.FItemList(i).CancelStateStr %></font></td>
				<td>
					<font color="<%= oorderdetail.FItemList(i).GetStateColor %>"><%= oorderdetail.FItemList(i).GetStateName %></font>
				</td>
				<td>
		<% if oorderdetail.FItemList(i).Fisupchebeasong="Y" then %>
					<font color="red"><%= oorderdetail.FItemList(i).Fitemid %><br>(��ü)</font>
		<% else %>
						<%= oorderdetail.FItemList(i).Fitemid %>
		<% end if %>
				</td>
				<td width="90">
					<acronym title="<%= oorderdetail.FItemList(i).Fmakerid %>">
					<%= Left(oorderdetail.FItemList(i).Fmakerid,32) %>
					</acronym>
				</td>
				<td align="left">
					<acronym title="<%= oorderdetail.FItemList(i).FItemName %>"><%= DDotFormat(oorderdetail.FItemList(i).FItemName,64) %></acronym>
		<% if (oorderdetail.FItemList(i).FItemoptionName <> "") then %>
					<br>
					<font color="blue">[<%= oorderdetail.FItemList(i).FItemoptionName %>]</font><br>
		<% end if %>
				</td>
				<td>

					<input type="text" name="regitemno" value="<% if (targetregitemno <> "") then %><%= targetregitemno %><% else %><%= oorderdetail.FItemList(i).Fitemno %><% end if %>" size="2" style="text-align:center" onKeyUp="__CheckMaxItemNo(this, <%= oorderdetail.FItemList(i).FItemNo %>);">
					/
					<input type="text" name="itemno" value="<%= oorderdetail.FItemList(i).Fitemno %>" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>
				</td>
				<td align="right">
					<% if (Not oorderdetail.FItemList(i).IsOldJumun) then %>
                    	<span title="<%= oorderdetail.FItemList(i).GetSaleText %>" style="cursor:hand">
                    	<font color="<%= oorderdetail.FItemList(i).GetSaleColor %>">
                    		<%= FormatNumber(oorderdetail.FItemList(i).GetSalePrice,0) %>
                    	</font>
                    	</span>
                	<% else %>
                		----
                	<% end if %>
				</td>

				<td align="right">
                	<span title="<%= oorderdetail.FItemList(i).GetItemCouponText %>" style="cursor:hand">
                	<font color="<%= oorderdetail.FItemList(i).GetItemCouponColor %>">
                		<b><%= FormatNumber(oorderdetail.FItemList(i).GetItemCouponPrice,0) %></b>
                	</font>
                	</span>

					<% if (oorderdetail.FItemList(i).getAllAtDiscountedPrice > 0) or (oorderdetail.FItemList(i).getPercentBonusCouponDiscountedPrice > 0) then %>
                	<span title="<%= oorderdetail.FItemList(i).GetBonusCouponText %>" style="cursor:hand">
                	<font color="<%= oorderdetail.FItemList(i).GetBonusCouponColor %>">
                		<br><b>(<%= FormatNumber(oorderdetail.FItemList(i).GetBonusCouponPrice,0) %>)</b>
                	</font>
                	</span>
                	<% end if %>
				</td>
				<td align="right">
					<%= FormatNumber(oorderdetail.FItemList(i).Fbuycash,0) %>
				</td>
				<td align="right"></td>
			</tr>
	<% end if %>
<% next %>
        	<tr bgcolor="FFFFFF" height="20">
        	    <td colspan="7">
        	        &nbsp;
        	    </td>
        	    <td align="right" colspan="3">
        	        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
        	        <tr>
        	            <td>���û�ǰ�հ�</td>
        	            <td align="right"><input type="text" name="totalselecteditemcost" size="7" value="0" readonly style="text-align:right;border: 1px solid #333333;" ></td>
        	        </tr>
        	        </table>
        	    </td>
        	    <td colspan="2"></td>
        	</tr>
			</table>
		</td>
	</tr>
</table>

<p>

<!-- ====================================================================== -->



<!-- ====================================================================== -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>" colspan="2">
			<b>2. �߰� ����ǰ ����</b>
			<input type="button" class="button" value="���Ϻ귣��" onClick="javascript:SearchItemByMakerid()">
			<input type="button" class="button" value="�����ǸŰ�" onClick="javascript:SearchItemByPrice()">
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
		<td colspan="2">


	<% if (toItemId <> "") and (toItemOption <> "") then %>
            <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="#BABABA">
			<tr height="20" align="center" bgcolor="#F4F4F4">
				<td width="30">����</td>
				<td width="50">�̹���</td>
				<td width="30">����</td>
				<td width="50">������</td>
				<td width="50">��ǰ�ڵ�</td>
				<td width="90">�귣��ID</td>
				<td>��ǰ��<font color="blue">[�ɼǸ�]</font></td>
				<td width="80">����/���ֹ�</td>
				<td width="60">�ǸŰ�<br>(���ΰ�)</td>
				<td width="60">������</td>
				<td width="60">���԰�</td>
				<td width="100">���</td>
			</tr>

			<tr align="center" bgcolor='#FFFFFF' >
				<input type="hidden" name="add_detailstate" value="">
				<input type="hidden" name="add_makerid" value="<%= makerid %>">
				<input type="hidden" name="add_itemid" value="<%= itemid %>">
				<input type="hidden" name="add_itemoption" value="<%= itemoption %>">

				<input type="hidden" name="add_SalePrice" value="<%= add_SalePrice %>">
				<input type="hidden" name="add_ItemCouponPrice" value="<%= add_ItemCouponPrice %>">
				<input type="hidden" name="add_BonusCouponPrice" value="<%= add_BonusCouponPrice %>">
				<input type="hidden" name="add_buycash" value="<%= add_buycash %>">

				<input type="hidden" name="add_itemcost" value="<%= additemcost %>">
				<input type="hidden" name="add_allatitemdiscount" value="0">
				<input type="hidden" name="add_percentBonusCouponDiscount" value="<%= addpercentBonusCouponDiscount %>">

		<%if (iscouponapplyOK) and (salemethod <> "C") then %>
				<input type="hidden" name="iscouponapplied" value="Y">
				<input type="hidden" name="itemcouponidxapplied" value="<%= itemcouponidxapplied %>">
				<input type="hidden" name="bonuscouponidxapplied" value="<%= bonuscouponidxapplied %>">
		<% else %>
				<input type="hidden" name="iscouponapplied" value="N">
				<input type="hidden" name="itemcouponidxapplied" value="">
				<input type="hidden" name="bonuscouponidxapplied" value="">
		<% end if %>

				<td height="25">
					<input type="checkbox" name="add_orderdetailidx" value="" checked disabled>
				</td>
				<td width="50"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= itemid %>" target="_blank"><img src="<%= imagesmall %>" width="50" border="0"></a></td>
				<td>����</td>
				<td></td>
				<td>
		<% if mwdiv = "U" then %>
					<font color="red"><%= itemid %><br>(��ü)</font>
		<% else %>
						<%= itemid %>
		<% end if %>
				</td>
				<td width="90">
					<acronym title="<%= makerid %>">
					<%= Left(makerid,32) %>
					</acronym>
				</td>
				<td align="left">
					<acronym title="<%= itemname %>"><%= DDotFormat(itemname,64) %></acronym>
		<% if (itemoptionname <> "") then %>
					<br>
					<font color="blue">[<%= itemoptionname %>]</font><br>
		<% end if %>
				</td>
				<td>
					<input type="text" name="add_regitemno" value="<%= targetregitemno %>" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>
					/
					<input type="text" name="add_itemno" value="0" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>
				</td>
				<td align="right">
                	<font color="<% if (issaleitem = "Y") then %>red<% else %>black<% end if %>">
                		<%= FormatNumber((addorgitemcost),0) %>
                	</font>
				</td>

				<td align="right">
					<% if (addorgitemcost) <> additemcost then %>
						<% if isitemcouponapplied then %>
							<font color="green">
						<% else %>
							<font color="purple">
						<% end if %>
					<% end if %>
					<b><%= FormatNumber((additemcost),0) %></b>
				</td>
				<td align="right">
					<%= FormatNumber((buycash + optaddbuyprice),0) %>
				</td>
				<td align="right"></td>
			</tr>
        	<tr bgcolor="FFFFFF" height="20">
        	    <td colspan="7">
        	        &nbsp;
        	    </td>
        	    <td align="right" colspan="3">
        	        <table width="100%" border="0" cellspacing="0" cellpadding="0" class="a">
        	        <tr>
        	            <td>�߰���ǰ�հ�</td>
        	            <td align="right"><input type="text" name="add_totalselecteditemcost" size="7" value="0" readonly style="text-align:right;border: 1px solid #333333;" ></td>
						<input type="hidden" name="add_totalselectedcoupon" value="">
						<input type="hidden" name="add_totalselectedallatsubtract" value="">
        	        </tr>
        	        <tr>
        	            <td><font color="red">����</font></td>
        	            <td align="right"><input type="text" name="totaldiffitemcost" size="7" value="0" readonly style="text-align:right;border: 1px solid #333333;" ></td>
        	        </tr>
        	        </table>
        	    </td>
        	    <td colspan="2"></td>
        	</tr>
			</table>
	<% end if %>

		</td>
	</tr>
</table>

<p>

<% if (toItemId <> "") and (toItemOption <> "") then %>

	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr height="25" bgcolor="#FFFFFF" >
			<td bgcolor="<%= adminColor("tabletop") %>" colspan="3">
				<b>3. CS����</b>
			</td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
				��ǰ��
			</td>
			<td>
				<% if (oordermaster.FOneItem.FIpkumDiv >= "4" and Not IsChangeOrder) or IsChangeOrder then %>
					<%= oordermaster.FOneItem.IpkumDivName %>
				<% else %>
					<font color="red"><b><%= oordermaster.FOneItem.IpkumDivName %></b></font>
				<% end if %>
				|
				<% if issamemakerid then %>
					�귣�� ����
				<% else %>
					<font color="red"><b>�귣�� �ٸ�</b></font>
				<% end if %>
				|
				<% if issameitemcost then %>
					���� �ǸŰ�,���԰� ����
				<% else %>
					<font color="red"><b>���� �ǸŰ�,���԰� �ٸ�</b></font>
				<% end if %>

				<% if isitemcouponapplied then %>
					|
					<font color="green">��ǰ����</font>
					<% if iscouponapplyOK <> True then %>
						<font color="red"><b>(�߰���ǰ�� ����Ұ�)</b></font>
					<% else %>
						(�߰���ǰ�� ����)
					<% end if %>
				<% end if %>
				<% if ispercentcouponapplied then %>
					|
					<font color="purple">��������</font>
					<% if iscouponapplyOK <> True then %>
						<font color="red"><b>(�߰���ǰ�� ����Ұ�)</b></font>
					<% else %>
						(�߰���ǰ�� ����)
					<% end if %>
				<% end if %>
			</td>
		</tr>

		<tr bgcolor="#FFFFFF">
			<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
				����ó��
			</td>
			<td>
				<% if (fromitemcost > additemcost) and (oordermaster.FOneItem.FIpkumDiv >= "4") then %>
					<input type="radio" name="salemethod" value="R" onClick="ChangeSaleMethod('R')" <% if (salemethod = "R") then %>checked<% end if %>> ȯ��
				<%elseif (fromitemcost < additemcost) or (salemethod = "C") then %>
					<input type="radio" name="salemethod" value="C" onClick="ChangeSaleMethod('C')" <% if (salemethod = "C") then %>checked<% end if %>> CS����
				<% end if %>
				<input type="radio" name="salemethod" value="" onClick="ChangeSaleMethod('')" <% if (salemethod = "") then %>checked<% end if %>> ����
			</td>
		</tr>

	<% if (fromitemcost = additemcost) or ((fromitemcost > additemcost) and (salemethod = "R" or oordermaster.FOneItem.FIpkumDiv < "4")) then %>

			<tr bgcolor="#FFFFFF">
				<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
					��������
				</td>
				<td colspan=2>
					<input type="radio" name="divcd" value="A900" <% if (divcd = "A900") then %>checked<% end if %> onClick="ChangeDivCD(frm)"> �ֹ�����
					<input type="radio" name="divcd" value="A100" <% if (Not isregchangeorderOK) then %>disabled<% end if %>  onClick="ChangeDivCD(frm)"> ��ǰ���� ��ȯ���
					&nbsp;
					<% if (detailstate = "7") then %>
						<% if Not isregchangeorderOK then %>
						* ��ȯ��� ��ϺҰ�
						: ���� �ǸŰ� ���԰� �ƴѰ�� �����ڵ� ��ϺҰ�
						<% end if %>
					<% end if %>
				</td>
			</tr>

			<tr bgcolor="#FFFFFF">
				<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
					��������
				</td>
				<td colspan=2>
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
		                [<a href="javascript:selectGubun('C004','CD01','����','�ܼ�����','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">�ܼ�����</a>]
		                [<a href="javascript:selectGubun('C004','CD05','����','ǰ��','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">ǰ��</a>]
		                [<a href="javascript:selectGubun('C005','CE01','��ǰ����','��ǰ�ҷ�','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">��ǰ�ҷ�</a>]
		                [<a href="javascript:selectGubun('C004','CD99','����','��Ÿ','gubun01','gubun02','gubun01name','gubun02name','frm','causepop'); SetAddBeasongPay();">��Ÿ</a>]
		            	&nbsp; &nbsp; &nbsp;
		            	<div id="chkmodifyitemstockoutyn" style="display: inline;"><input type="checkbox" name="modifyitemstockoutyn" value="Y" checked> ǰ������ ����(�����ǰ)</div>
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

				<input type="hidden" name="isupchebeasong" value="<%= isupchebeasong %>">
				<input type="hidden" name="requiremakerid" value="<%= requiremakerid %>">
				<input type="hidden" name="upchebeasongpay" value="<%= upchebeasongpay %>">
				<tr bgcolor="#FFFFFF">
					<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
						��۱���
					</td>
					<td colspan=2>
				    	<% if (isupchebeasong = "Y") then %>
				    		<font color=red><%= requiremakerid %></font> (�⺻��ۺ� : <%= FormatNumber(upchebeasongpay, 0) %>��)
				    	<% else %>
				    		�ٹ����ٹ��
				    	<% end if %>
					</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td width="100" height=30 bgcolor="<%= adminColor("tabletop") %>">
						�߰���ۺ�
					</td>
					<td colspan=2>
				    	<input type="text" class="text" name="add_customeraddbeasongpay" value="0" size="20">
				    	&nbsp;
			    	    <select class="select" name="add_customeraddmethod" class="text">
				    	    <option value="">����
				    	    <option value="1">�ڽ�����
				    	    <option value="2">�ù�� ���δ�
				    	    <option value="5">��Ÿ
			    	    </select>
					</td>
				</tr>

			<tr bgcolor="#FFFFFF" height="40">
				<td colspan="3" align="center">
		<% if Not IsOrderCanceled then %>
					<input type="button" class="button" value="��ǰ����" onclick="javascript:SaveChangeItem(false)">
					<% if (C_CSPowerUser or C_ADMIN_AUTH) then %>
				    <input type="button" class="button" value="��������(������)" onclick="javascript:SaveChangeItem(true)">
					<% end if %>
		<% else %>
					<b>��ҵ� �ֹ��� ��ǰ���� �Ұ�</b>
		<% end if %>
				</td>
			</tr>

	<% else %>

			<tr bgcolor="#FFFFFF" height="40">
				<td colspan="3" align="center">
					��ϺҰ�
					<% if ((fromitemcost <> additemcost)) then %>
					: ���� <font color="red">���� ó�����</font>�� �����ϼ���.
					<% end if %>
				</td>
			</tr>

	<% end if %>

		</table>

<% end if %>
</form>
<p>

<div>
* <font color="red"><b>��ȯ���(��ϺҰ�)</b></font><br>
&nbsp;&nbsp;&nbsp;&nbsp; - ���Ϸ� ���� ��ǰ<br>
&nbsp;&nbsp;&nbsp;&nbsp; - �귣�尡 �ٸ� ���<br>
&nbsp;&nbsp;&nbsp;&nbsp; - ���� �ǸŰ�(���ΰ�) �Ǵ� ���԰��� �ٸ����<br>
&nbsp;&nbsp;&nbsp;&nbsp; - ���� ���� ���°� �������� ���� ���<br>
&nbsp;&nbsp;&nbsp;&nbsp; - ����ȯ��<br><br>

* <font color="red"><b>��ȯ���(������ ���� �ʿ�)</b></font><br>
&nbsp;&nbsp;&nbsp;&nbsp; - CS����<br><br>

* <font color="red"><b>�ֹ�����(��ϺҰ�)</b></font><br>
&nbsp;&nbsp;&nbsp;&nbsp; - ��ȯ�ֹ��� ���<br>
&nbsp;&nbsp;&nbsp;&nbsp; - ����Ϸ��� ��ǰ�� �̹� ����� ���<br>
&nbsp;&nbsp;&nbsp;&nbsp; - �߰��Ϸ��� ��ǰ�� ���� �ֹ������� �̹� �����ϰ�, �����°� ���� �ٸ����<br>
&nbsp;&nbsp;&nbsp;&nbsp; - ���� �ǸŰ�(���ΰ�) �Ǵ� ���԰��� �ٸ����<br><br>

* <font color="red"><b>�ֹ�����(������ ���� �ʿ�)</b></font><br>
&nbsp;&nbsp;&nbsp;&nbsp; - �����Ϸ� ���� ��ǰ<br>
&nbsp;&nbsp;&nbsp;&nbsp; - ���Ϸ�� ��ǰ<br>
&nbsp;&nbsp;&nbsp;&nbsp; - �귣�尡 �ٸ� ���<br>
&nbsp;&nbsp;&nbsp;&nbsp; - ���� ���� ���°� �������� ���� ���<br>
&nbsp;&nbsp;&nbsp;&nbsp; - ����ȯ�� �Ǵ� CS����<br>
</div>

<iframe name="iframeforadd" width="0" height="0">
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
