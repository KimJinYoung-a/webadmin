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
dim i
dim idx, orderserial
dim result
orderserial = requestCheckVar(request("orderserial"),32)



'==============================================================================
''�ֹ� ����Ÿ
dim oordermaster, IsOrderCanceled, OrderMasterState

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
OrderMasterState = oordermaster.FOneItem.FIpkumDiv



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
dim obonuscoupon

dim bonuscouponidx
dim bonuscouponstartdate, bonuscouponexpiredate
dim bonuscoupontype, bonuscouponvalue

dim IsPercentBonusCouponExist, IsPercentBonusCouponApplyDateOK

set obonuscoupon = new CCouponMaster

obonuscoupon.FPageSize = 1
obonuscoupon.FCurrPage = 1
obonuscoupon.FrectCoupontype = 1					'���������� ����Ѵ�.
obonuscoupon.FrectOrderserial = orderserial
obonuscoupon.FrectChkOld = ""
obonuscoupon.FrectSiteType = ""						'����

obonuscoupon.GetEventCouponList

IsPercentBonusCouponExist = false
IsPercentBonusCouponApplyDateOK = false

IsPercentBonusCouponExist = (obonuscoupon.FTotalCount > 0)

if (IsPercentBonusCouponExist) then
	bonuscouponidx = obonuscoupon.FItemList(0).Fidx

	'TODO : ��¥�� üũ�Ѵ�. �ּұ��űݾ��� üũ���� �ʴ´�.
	bonuscouponstartdate = obonuscoupon.FItemList(0).Fstartdate
	bonuscouponexpiredate = obonuscoupon.FItemList(0).Fexpiredate

	if (Left(now, 10) >= Left(bonuscouponstartdate, 10)) and (Left(now, 10) <= Left(bonuscouponexpiredate, 10)) then
		IsPercentBonusCouponApplyDateOK = true
	end if

	bonuscoupontype = obonuscoupon.FItemList(0).Fcoupontype
	bonuscouponvalue = obonuscoupon.FItemList(0).Fcouponvalue
end if

%>
<script language="JavaScript" src="/cscenter/js/cscenter.js"></script>
<script language='javascript'>

function SearchItemByMakerid() {
	var frm = document.frm;
	var isupchebeasong;
	var makerid;
	var excludeupbae;

	isupchebeasong = frm.editisupchebeasong.value;
	makerid = frm.editmakerid.value;

	if (makerid == "") {
		alert("���� ����� ��ǰ�� �����ϼ���.");
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
	var isupchebeasong;
	var makerid, itemcanceltotal;
	var excludeupbae;

	isupchebeasong = frm.editisupchebeasong.value;
	makerid = frm.editmakerid.value;
	itemcanceltotal = frm.itemcanceltotal.value;

	if (isupchebeasong == "") {
		alert("���� ����� ��ǰ�� �����ϼ���.");
		return;
	}

	if (isupchebeasong == "N") {
		excludeupbae = "on";
	} else {
		excludeupbae = "";
	}

	var popwin = window.open('pop_item_search_one.asp?makerid=' + makerid + '&saleprice=' + itemcanceltotal + '&onlineonly=Y&nubeasong=' + excludeupbae,'SearchItemByPrice','width=1000,height=480,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ReActItemOne(toItemId, toItemOption) {
	var frm = document.frm;

	if (IsSameItemExist(toItemId, toItemOption) == true) {
		return;
	}

	document.iframeforadd.location.href = "orderdetail_editorder_iframe.asp?toItemId=" + toItemId + "&toItemOption=" + toItemOption;

	return;
}

function IsSameItemExist(ItemId, ItemOption) {
	var frm = document.frm;

	for (var i = 0; i < frm.orderdetailidx.length; i++) {
		if (frm.orderdetailidx[i].type != "checkbox") {
			continue;
		}

		// ��ҵ� ��ǰ�� �ִ� ��� ���������� ������ �츰��.(�߰�����)
		if (frm.cancelyn[i].value == "Y") {
			continue;
		}

		if (frm.orderdetailidx[i].checked != true) {
			// ��ҵ��� ���� ��ǰ�� �ִ� ��� �߰��Ұ�
			// continue;
		}

		if ((ItemId == frm.itemid[i].value) && (ItemOption == frm.itemoption[i].value)) {
			alert("���ϻ�ǰ�� �ֽ��ϴ�. ������ �� �����ϴ�.");
			return true;
		}

		if (ItemId == frm.itemid[i].value) {
			alert("���ϻ�ǰ�� �ֽ��ϴ�. �ɼǺ����� �̿��ϼ���.");
			return true;
		}
	}

	return false;
}

function WriteAddedItemALL() {
	var htmlstr;
	var tmpstr;

	htmlstr = "";
	if (arrToItemId.length > 0) {
		// ====================================================================
		if ((isbonuscouponapplied == "Y") || (iscssaleapplied == "Y")) {
			for (var i = 0; i < arrToItemId.length; i++) {
				arrToItemCouponApplied[i] = "N";
			}
		}

		for (var i = 0; i < arrToItemId.length; i++) {
			ApplyNormalPrice(i)
			if (arrToItemCouponApplied[i] == "Y") {
				ApplyItemCoupon(i, true);
			}
		}

		if (isbonuscouponapplied == "Y") {
			ApplySaleInfo("bonuscoupon");
		}

		if (iscssaleapplied == "Y") {
			CalculateAddSUM();
			ApplySaleInfo("cssale");
		}

		// ====================================================================
		htmlstr = "<table width='100%' border='0' align='center' cellpadding='3' cellspacing='1' class='a' bgcolor='#BABABA'>"
		htmlstr = htmlstr + "<tr height='20' align='center' bgcolor='#F4F4F4'>"
		htmlstr = htmlstr + "	<td width='30'>����</td>"
		htmlstr = htmlstr + "	<td width='50'>�̹���</td>"
		htmlstr = htmlstr + "	<td width='30'>����</td>"
		htmlstr = htmlstr + "	<td width='50'>������</td>"
		htmlstr = htmlstr + "	<td width='50'>��ǰ�ڵ�</td>"
		htmlstr = htmlstr + "	<td width='90'>�귣��ID</td>"
		htmlstr = htmlstr + "	<td>��ǰ��<font color='blue'>[�ɼǸ�]</font></td>"
		htmlstr = htmlstr + "	<td width='80'>����</td>"
		htmlstr = htmlstr + "	<td width='60'>�ǸŰ�<br>(���ΰ�)</td>"
		htmlstr = htmlstr + "	<td width='60'>������</td>"
		htmlstr = htmlstr + "	<td width='60'>���԰�</td>"
		htmlstr = htmlstr + "	<td width='100'>���</td>"
		htmlstr = htmlstr + "</tr>"

		// ====================================================================
		for (var i = 0; i < arrToItemId.length; i++) {
			htmlstr = htmlstr + GetHTMLAddedItem(i);
		}

		htmlstr = htmlstr + "<tr bgcolor='FFFFFF' height='20'>"
		htmlstr = htmlstr + "    <td colspan='7' align='center'>"

		// ====================================================================
		tmpstr = "";
		if (IsPercentBonusCouponExist != true) {
			tmpstr = "disabled";
		}

		if (isbonuscouponapplied == "Y") {
			tmpstr = "checked";
		}
		htmlstr = htmlstr + "		<input type='radio' name='salemethod' value='bonuscoupon' " + tmpstr + " onClick='CheckSaleInfo(); WriteAddedItemALL(); CalculateAddSUM();'> ��������"

		tmpstr = "";
		if (iscssaleapplied == "Y") {
			tmpstr = "checked";
		}
		htmlstr = htmlstr + "		<input type='radio' name='salemethod' value='cssale' " + tmpstr + " onClick='CheckSaleInfo(); WriteAddedItemALL(); CalculateAddSUM();'> CS����"

		tmpstr = "";
		if ((iscssaleapplied != "Y") && (isbonuscouponapplied != "Y")) {
			tmpstr = "checked";
		}
		htmlstr = htmlstr + "		<input type='radio' name='salemethod' value='' " + tmpstr + " onClick='CheckSaleInfo(); WriteAddedItemALL(); CalculateAddSUM();'> �������"

		// ====================================================================
		htmlstr = htmlstr + "    </td>"
		htmlstr = htmlstr + "    <td align='right' colspan='3'>"
		htmlstr = htmlstr + "        <table width='100%' border='0' cellspacing='0' cellpadding='0' class='a'>"
		htmlstr = htmlstr + "        <tr>"
		htmlstr = htmlstr + "            <td>�߰���ǰ�հ�</td>"
		htmlstr = htmlstr + "            <td align='right'><input type='text' name='itemaddtotal' size='7' value='0' readonly style='text-align:right;border: 1px solid #333333;' ></td>"
		htmlstr = htmlstr + "        </tr>"
		htmlstr = htmlstr + "        <tr>"
		htmlstr = htmlstr + "            <td>�߰��ݾ��հ�</td>"
		htmlstr = htmlstr + "            <td align='right'><input type='text' name='itemaddrequire' size='7' value='0' readonly style='text-align:right;border: 1px solid #333333;' ></td>"
		htmlstr = htmlstr + "        </tr>"
		htmlstr = htmlstr + "        </table>"
		htmlstr = htmlstr + "    </td>"
		htmlstr = htmlstr + "    <td colspan='2'></td>"
		htmlstr = htmlstr + "</tr>"

		htmlstr = htmlstr + "</table>"
	}

	document.getElementById("TableAddedItem").innerHTML = htmlstr;
}


function GetHTMLAddedItem(idx) {
	var htmlstr;
	var tmpstr;

	tmpstr = "";
	if ((arrToItemCouponYn[idx] != "Y") || (arrToItemCouponType[idx] == "3")) {
		tmpstr = "disabled";
	}

	if (arrToItemCouponApplied[idx] == "Y") {
		tmpstr = "checked";
	}

	htmlstr = "<tr align='center' bgcolor='#FFFFFF'>"
	htmlstr = htmlstr + "	<td height='25'>"
	htmlstr = htmlstr + "		<input type='checkbox' name='arridxadded" + idx + "' onClick='RemoveAddedItem(" + idx + ")' checked>"
	htmlstr = htmlstr + "	</td>"
	htmlstr = htmlstr + "	<td width='50'><a href='http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + arrToItemId[idx] + "' target='_blank'><img src='" + arrToImageSmall[idx] + "' width='50' border='0'></a></td>"
	htmlstr = htmlstr + "	<td><font color='#000000'>����</font></td>"
	htmlstr = htmlstr + "	<td><font color='#000000'></font></td>"
	htmlstr = htmlstr + "	<td>" + arrToItemId[idx] + "</td>"
	htmlstr = htmlstr + "	<td width='90'>"
	htmlstr = htmlstr + "		<acronym title='" + arrToMakerid[idx] + "'>"
	htmlstr = htmlstr + "		" + arrToMakerid[idx].substring(0, 32) + ""
	htmlstr = htmlstr + "		</acronym>"
	htmlstr = htmlstr + "	</td>"
	htmlstr = htmlstr + "	<td align='left'>"
	htmlstr = htmlstr + "		<acronym title='" + arrToItemName[idx] + "'>" + arrToItemName[idx].substring(0, 64) + "</acronym>"
	if (arrToItemOption[idx] != "0000") {
		htmlstr = htmlstr + "		<br><font color='blue'>[" + arrToItemOptionName[idx] + "]</font>"
	}
	htmlstr = htmlstr + "	</td>"
	htmlstr = htmlstr + "	<td>"
	htmlstr = htmlstr + "		<input type='text' name='additemno" + idx + "' value='" + arrToItemNo[idx] + "' size='2' style='text-align:center' onKeyUp='CheckItemNoAdded(this, " + idx + ");'>"
	htmlstr = htmlstr + "	</td>"
	htmlstr = htmlstr + "	<td align='right'>" + GetHTMLAddedItemOrgPrice(idx) + GetHTMLAddedItemSalePrice(idx) + "</td>"
	htmlstr = htmlstr + "	<td align='right'>" + GetHTMLAddedItemItemCouponPrice(idx) + GetHTMLAddedItemBonusCouponPrice(idx) + "</td>"
	htmlstr = htmlstr + "	<td align='right'>" + GetHTMLAddedItemBuyPrice(idx) + "</td>"
	htmlstr = htmlstr + "	<td align='right'>"
	htmlstr = htmlstr + "		<input type='checkbox' name='additemcoupon" + idx + "' " + tmpstr + " onClick='CheckItemCoupon(" + idx + ",this.checked); WriteAddedItemALL(); CalculateAddSUM();'> ��ǰ����<br>"
	htmlstr = htmlstr + "	</td>"
	htmlstr = htmlstr + "</tr>"

	return htmlstr;
}

function GetHTMLAddedItemOrgPrice(idx) {
	return FormatNumber(arrToOrgitemCostPrice[idx]);
}

function GetHTMLAddedItemSalePrice(idx) {
	var htmlstr;

	htmlstr = "";

	if (arrToOrgitemCostPrice[idx] > arrSalePrice[idx]) {
		htmlstr = FormatNumber(arrSalePrice[idx]);
		htmlstr = "<br><font color='red'>(" + htmlstr + ")</font>";
	}

	return htmlstr;
}

function GetHTMLAddedItemItemCouponPrice(idx) {
	var htmlstr;

	htmlstr = FormatNumber(arrItemCouponPrice[idx]);
	if (arrSalePrice[idx] > arrItemCouponPrice[idx]) {
		htmlstr = "<font color='green'>" + htmlstr + "</font>";
	}

	return htmlstr;
}

function GetHTMLAddedItemBonusCouponPrice(idx) {
	var htmlstr;

	htmlstr = "";

	if (arrItemCouponPrice[idx] > arrBonusCouponPrice[idx]) {
		htmlstr = FormatNumber(arrBonusCouponPrice[idx]);
		htmlstr = "<br><font color='purple'>(" + htmlstr + ")</font>";
	}

	return htmlstr;
}

function GetHTMLAddedItemBuyPrice(idx) {
	var htmlstr;

	htmlstr = FormatNumber(arrBuyPrice[idx]);

	return htmlstr;
}

// iframe ���� ȣ��
var arrToItemId 		= new Array();
var arrToItemOption 	= new Array();
var arrToItemNo 		= new Array();
var arrToItemName 		= new Array();
var arrToItemOptionName = new Array();
var arrToMakerid 		= new Array();

var arrToOrgitemCostPrice 	= new Array();
var arrToSellCash 			= new Array();
var arrToOptAddPrice 		= new Array();
var arrToBuyCash 			= new Array();
var arrToOptAddBuyPrice 	= new Array();

var arrSalePrice			= new Array();
var arrItemCouponPrice		= new Array();
var arrBonusCouponPrice		= new Array();
var arrBuyPrice				= new Array();

var arrToImageSmall 		= new Array();
var arrToIsSaleItem 		= new Array();
var arrToIsMileageshopItem 	= new Array();
var arrToIsSpacialuserItem 	= new Array();

var arrToItemCouponYn 			= new Array();
var arrToItemCouponIdx 			= new Array();
var arrToItemCouponType 		= new Array();
var arrToItemCouponValue 		= new Array();
var arrToItemCouponBuyprice 	= new Array();
var arrToItemCouponApplied 		= new Array();
function ReActItemAdd(isaddok, itemid, itemoption, makerid, itemname, itemoptionname, orgitemcostprice, sellcash, optaddprice, buycash, optaddbuyprice, imagesmall, issaleitem, ismileageshopitem, isspacialuseritem, ItemCouponYn, CurrItemCouponIdx, ItemCouponType, ItemCouponValue, ItemCouponBuyprice) {
	var frm = document.frm;
	var salepricehtml;
	var arridx, itemexistinarray;

	if (isaddok == false) {
		alert("�������� �ʴ� ��ǰ�Դϴ�.");
		return;
	}

	if (frm.editisupchebeasong.value == "Y") {
		if (frm.editmakerid.value != makerid) {
			alert("��ҵǴ� ��ǰ�� ������ �귣�常 �߰� �����մϴ�.");
			return;
		}
	}

	itemexistinarray = false;
	for (var i = 0; i < arrToItemId.length; i++) {
		if ((arrToItemId[i]*1 == itemid*1) && (arrToItemOption[i] == itemoption)) {
			arrToItemNo[i] = arrToItemNo[i] + 1;
			itemexistinarray = true;
		}
	}

	if (itemexistinarray != true) {
		arridx = arrToItemId.length;

		arrToItemId[arridx] = itemid;
		arrToItemOption[arridx] = itemoption;
		arrToItemNo[arridx] = 1;
		arrToItemName[arridx] = itemname;
		arrToItemOptionName[arridx] = itemoptionname;
		arrToMakerid[arridx] = makerid;

		arrToOrgitemCostPrice[arridx] = orgitemcostprice*1;
		arrToSellCash[arridx] = sellcash*1;
		arrToOptAddPrice[arridx] = optaddprice*1;
		arrToBuyCash[arridx] = buycash*1;
		arrToOptAddBuyPrice[arridx] = optaddbuyprice*1;

		arrToImageSmall[arridx] = imagesmall;
		arrToIsSaleItem[arridx] = issaleitem;
		arrToIsMileageshopItem[arridx] = ismileageshopitem;
		arrToIsSpacialuserItem[arridx] = isspacialuseritem;

		arrToItemCouponYn[arridx] = ItemCouponYn;
		arrToItemCouponIdx[arridx] = CurrItemCouponIdx;
		arrToItemCouponType[arridx] = ItemCouponType;
		arrToItemCouponValue[arridx] = ItemCouponValue;
		arrToItemCouponBuyprice[arridx] = ItemCouponBuyprice*1;
		arrToItemCouponApplied[arridx] = "N";

		ApplyNormalPrice(arridx);
	}


	WriteAddedItemALL();
	CalculateAddSUM();

	return;
}

function RemoveAddedItemALL() {
	while (arrToItemId.length > 0) {
		RemoveAddedItem(0);
	}

	isbonuscouponapplied = "N";
	iscssaleapplied = "N";

	WriteAddedItemALL();
	CalculateAddSUM();
}

function RemoveAddedItem(idx) {
	arrToItemId.splice(idx, 1);
	arrToItemOption.splice(idx, 1);
	arrToItemNo.splice(idx, 1);
	arrToItemName.splice(idx, 1);
	arrToItemOptionName.splice(idx, 1);
	arrToMakerid.splice(idx, 1);

	arrToOrgitemCostPrice.splice(idx, 1);
	arrToSellCash.splice(idx, 1);
	arrToOptAddPrice.splice(idx, 1);
	arrToBuyCash.splice(idx, 1);
	arrToOptAddBuyPrice.splice(idx, 1);

	arrToImageSmall.splice(idx, 1);
	arrToIsSaleItem.splice(idx, 1);
	arrToIsMileageshopItem.splice(idx, 1);
	arrToIsSpacialuserItem.splice(idx, 1);

	arrToItemCouponYn.splice(idx, 1);
	arrToItemCouponIdx.splice(idx, 1);
	arrToItemCouponType.splice(idx, 1);
	arrToItemCouponValue.splice(idx, 1);
	arrToItemCouponBuyprice.splice(idx, 1);
	arrToItemCouponApplied.splice(idx, 1);

	arrSalePrice.splice(idx, 1);
	arrItemCouponPrice.splice(idx, 1);
	arrBonusCouponPrice.splice(idx, 1);
	arrBuyPrice.splice(idx, 1);

	WriteAddedItemALL();
	CalculateAddSUM();
}

function CheckSaleInfo() {
	var frm = document.frm;

	iscssaleapplied = "N";
	isbonuscouponapplied = "N";
	if (frm.salemethod[0].checked == true) {
		isbonuscouponapplied = "Y";
	} else if (frm.salemethod[1].checked == true) {
		iscssaleapplied = "Y";
	} else {
		//
	}
}

var iscssaleapplied;
function ApplySaleInfo(salemethod) {
	if (salemethod == "bonuscoupon") {
		ApplyBonusCoupon();
	} else if (salemethod == "cssale") {
		ApplyCSSale();
	}
}

var IsPercentBonusCouponExist = <%= LCase(IsPercentBonusCouponExist) %>;
var IsPercentBonusCouponApplyDateOK = <%= LCase(IsPercentBonusCouponApplyDateOK) %>;
var bonuscoupontype = "<%= bonuscoupontype %>";
var bonuscouponvalue = "<%= bonuscouponvalue %>";
var bonuscouponidx = "<%= bonuscouponidx %>";

var isbonuscouponapplied;
function ApplyBonusCoupon() {
	var frm = document.frm;

	isbonuscouponapplied = "N";

	if (IsPercentBonusCouponExist != true) {
		alert("�ֹ��� ����� ���ʽ������� �����ϴ�.");
		frm.salemethod[0].checked = false;
		return false;
	}

	if (IsPercentBonusCouponApplyDateOK != true) {
		if (confirm("��밡���� �Ⱓ�� �������ϴ�.\n\n���������Ͻðڽ��ϱ�?") != true) {
			frm.salemethod[0].checked = false;
			return false;
		}
	}

	for (var i = 0; i < arrToItemId.length; i++) {
		if (arrToIsSaleItem[i] == "Y") {
			continue;
		}

		if (arrToIsMileageshopItem[i] == "Y") {
			continue;
		}

		// TODO : ���ȸ���޻�ǰ, ���� üũ����

		if (bonuscoupontype == "1") {
			arrBonusCouponPrice[i] = ((100 - bonuscouponvalue) * arrBonusCouponPrice[i]) / 100;
		}
	}

	isbonuscouponapplied = "Y";

	return true;
}

function ApplyCSSale() {
	var frm = document.frm;
	var diff, sumsaleprice, cssaleprice, errorfix;;

	if (frm.refundrequire.value*1 >= frm.addrequire.value*1) {
		alert("�߰��Ǵ� ��ǰ�� �ݾ��� ū ��츸 CS������ ������ �� �ֽ��ϴ�.");
		iscssaleapplied = "N";
		return;
	}

	diff = frm.addrequire.value - frm.refundrequire.value;

	sumsaleprice = 0;
	for (var i = 0; i < arrToItemId.length; i++) {
		sumsaleprice = sumsaleprice + arrSalePrice[i];
	}

	for (var i = 0; i < arrToItemId.length; i++) {
		cssaleprice = Math.round(diff * (arrSalePrice[i] / sumsaleprice));
		ApplyCSSalePrice(i, cssaleprice);
	}

	sumsaleprice = 0;
	for (var i = 0; i < arrToItemId.length; i++) {
		sumsaleprice = sumsaleprice + arrSalePrice[i];
	}

	if (frm.refundrequire.value*1 != sumsaleprice) {
		errorfix = frm.refundrequire.value*1 - sumsaleprice;
		ApplyCSSalePrice(0, errorfix*-1);
	}
}

function ApplyNormalPrice(idx) {
	arrSalePrice[idx] 			= arrToSellCash[idx] + arrToOptAddPrice[idx];
	arrItemCouponPrice[idx] 	= arrToSellCash[idx] + arrToOptAddPrice[idx];
	arrBonusCouponPrice[idx] 	= arrToSellCash[idx] + arrToOptAddPrice[idx];
	arrBuyPrice[idx] 			= arrToBuyCash[idx]  + arrToOptAddBuyPrice[idx];
}

function ApplyCSSalePrice(idx, cssaleprice) {
	arrSalePrice[idx] = arrSalePrice[idx] - cssaleprice;

	arrItemCouponPrice[idx] = arrSalePrice[idx];
	arrBonusCouponPrice[idx] = arrSalePrice[idx];
}

function CheckItemCoupon(idx, ischecked) {
	if (iscssaleapplied == "Y") {
		alert("CS������ ����� ��� ��ǰ������ ������ �� �����ϴ�.");
		return;
	}

	if (iscssaleapplied == "Y") {
		alert("�������ʽ������� ����� ��� ��ǰ������ ������ �� �����ϴ�.");
		return;
	}

	arrToItemCouponApplied[idx] = "N";
	if (ischecked == true) {
		arrToItemCouponApplied[idx] = "Y";
	}
}

function ApplyItemCoupon(idx, ischecked) {
	var frm = document.frm;
	var e;

	if (arrToItemCouponYn[idx] != "Y") {
		alert("������ �����ϴ�.");
		e = eval("frm.additemcoupon" + idx);
		e.checked = false;
		return;
	}

	if (arrToItemCouponType[idx] == "3") {
		alert("��ۺ� ������ ������ �� �����ϴ�.");
		e = eval("frm.additemcoupon" + idx);
		e.checked = false;
		return;
	}

	ApplyNormalPrice(idx);
	arrToItemCouponApplied[idx] = "N";

	if (ischecked != true) {
		return;
	}

	if (arrToItemCouponType[idx] == "1") {
		// ������ǰ����(�ɼǰ����� �������)
		arrItemCouponPrice[idx] = (arrToSellCash[idx] - (arrToSellCash[idx] * arrToItemCouponValue[idx] / 100)) + arrToOptAddPrice[idx];
		arrToItemCouponApplied[idx] = "Y";
	} else if (arrToItemCouponType[idx] == "2") {
		// ���׻�ǰ����
		arrItemCouponPrice[idx] = (arrToSellCash[idx] - arrToItemCouponValue[idx]) + arrToOptAddPrice[idx];
		arrToItemCouponApplied[idx] = "Y";
	} else {
		// ��ۺ�����(toItemCouponType : 3)
	}

	if (arrToItemCouponBuyprice[idx] != 0) {
		arrBuyPrice[idx] = arrToItemCouponBuyprice[idx] + arrToOptAddPrice[idx];
	}

	arrBonusCouponPrice[idx] = arrItemCouponPrice[idx];
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

var ipkumdiv = "<%= oordermaster.FOneItem.Fipkumdiv %>";
function SaveChangeItem(isadmin) {
    var frm = document.frm;
    var e;
    var refundrequire, canceltotal, refunditemcostsum, refundcouponsum, allatsubtractsum;

    // ��Ҽ���üũ
	for (var i = 0; i < frm.orderdetailidx.length; i++) {
		if (frm.orderdetailidx[i].type != "checkbox") {
			continue;
		}

		if (frm.orderdetailidx[i].checked != true) {
			continue;
		}

		if (frm.regitemno[i].value*0 != 0) {
			alert("������ ��Ȯ�� �Է��ϼ���.");
			frm.regitemno[i].focus();
			return;
		}
	}

	// �߰�����üũ
	for (var i = 0; i < arrToItemId.length; i++) {
		e = eval("frm.additemno" + i);
		if (e.value*0 != 0) {
			alert("������ ��Ȯ�� �Է��ϼ���.");
			e.focus();
			return;
		}

	}

	if (frm.addtotal.value*1 == 0) {
		alert("�߰��� ��ǰ�� �����ϴ�.");
		return;
	}


    refundrequire = frm.refundrequire.value*1 - frm.addrequire.value*1;
    canceltotal = frm.canceltotal.value*1 - frm.addtotal.value*1;
    refunditemcostsum = frm.refunditemcostsum.value*1 - frm.additemcostsum.value*1;
    refundcouponsum = frm.refundcouponsum.value*1 - frm.addcouponsum.value*1;
    allatsubtractsum = frm.allatsubtractsum.value*1 - frm.addallatsubtractsum.value*1;

	if (refundrequire < 0) {
		alert("�߰��� ��ǰ�ݾ��� �� Ů�ϴ�. ���� �Ǵ� ������ �����ϼ���.");
		return;
	}

	if (refundrequire > 0) {
		if (isadmin != true) {
			alert('�߰��Ǵ� ��ǰ�� �ݾ��� ��ҵǴ� ��ǰ�� �ݾ׺��� ������� ����Ұ�\n\n��Ʈ�忡�� �����ϼ���.');
			return;
		} else {
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
	}

	if ((frm.salemethod[1].checked == true) && (isadmin != true)) {
		alert('CS������ ��Ʈ�常 �����մϴ�.');
		return;
	}

	if ((frm.editdetailstate.value >= "7") && (isadmin != true)) {
		alert('���� ��ǰ�Դϴ�.\n\n��Ʈ�忡�� �����ϼ���.');
		return;
	}

	var msg = "�ֹ����� �Ͻðڽ��ϱ�?";

	if ((frm.editisupchebeasong.value == "Y") && (frm.editdetailstate.value >= "3") && (frm.editdetailstate.value < "7")) {
		msg = "��ü����̸鼭 ��ǰ�غ� �����Դϴ�\n\n" + msg;
	} else if (frm.editdetailstate.value >= "7") {
		msg = "�̹� ���� ��ǰ�Դϴ�. ������ �̷���� ��ǰ�� ��� ������ �� �����ϴ�.\n\n" + msg;
	}

	if (isadmin == true) {
		msg = "[��Ʈ�����] " + msg;
	}

	if (frm.salemethod[1].checked == true) {
		frm.title.value = "�ֹ�����(CS����)";
	} else {
		if (frm.refundrequire.value*1 == frm.addrequire.value*1) {
			frm.title.value = "�ֹ�����(�����ǸŰ�)";
		} else if (frm.refundrequire.value*1 > frm.addrequire.value*1) {
			frm.title.value = "�ֹ�����(����ȯ��)";
		} else {
			// ����
			alert("���� - ������ ����");
			return;
		}
	}

	msg = frm.title.value + "\n\n" + msg;

	if (confirm(msg) == true) {
		if (isadmin == true) {
			frm.forceedit.value = "Y";
		}

		frm.refundrequire.value = refundrequire;
		frm.canceltotal.value = canceltotal;
		frm.refunditemcostsum.value = refunditemcostsum;
		frm.refundcouponsum.value = refundcouponsum;
		frm.allatsubtractsum.value = allatsubtractsum;

		// ��һ�ǰ
		for (var i = 0; i < frm.orderdetailidx.length; i++) {
			if (frm.orderdetailidx[i].type != "checkbox") {
				continue;
			}

			if (frm.orderdetailidx[i].checked != true) {
				continue;
			}

			frm.arrFromItemId.value 	= frm.arrFromItemId.value + "|" + frm.itemid[i].value;
			frm.arrFromItemOption.value = frm.arrFromItemOption.value + "|" + frm.itemoption[i].value;
			frm.arrFromItemNo.value 	= frm.arrFromItemNo.value + "|" + frm.regitemno[i].value;

			frm.arrFromDetailIdx.value 	= frm.arrFromDetailIdx.value + "|" + frm.orderdetailidx[i].value;
		}

		// �߰���ǰ
		for (var i = 0; i < arrToItemId.length; i++) {
			frm.arrToItemId.value 			= frm.arrToItemId.value + "|" + arrToItemId[i];
			frm.arrToItemOption.value 		= frm.arrToItemOption.value + "|" + arrToItemOption[i];
			frm.arrToItemNo.value 			= frm.arrToItemNo.value + "|" + arrToItemNo[i];

			frm.arrToSalePrice.value 		= frm.arrToSalePrice.value + "|" + arrSalePrice[i];
			frm.arrToItemCouponPrice.value 	= frm.arrToItemCouponPrice.value + "|" + arrItemCouponPrice[i];
			frm.arrToBonusCouponPrice.value = frm.arrToBonusCouponPrice.value + "|" + arrBonusCouponPrice[i];
			frm.arrToBuyCash.value 			= frm.arrToBuyCash.value + "|" + arrBuyPrice[i];

			frm.arrToItemCouponIdx.value 	= frm.arrToItemCouponIdx.value + "|" + arrToItemCouponIdx[i];
		}

		if (frm.salemethod[0].checked == true) {
			frm.toSaleMethod.value = "bonuscoupon";
		}

		frm.toBonusCouponIdx.value = bonuscouponidx;

		frm.submit();
	}
}

function CheckDifferentStateItemExist(obj) {
	var frm = document.frm;
	var detailstate;

	detailstate = "X";
	for (var i = 0; i < frm.orderdetailidx.length; i++) {
		if (frm.orderdetailidx[i].type != "checkbox") {
			continue;
		}

		if (frm.orderdetailidx[i].checked != true) {
			continue;
		}

		if ((detailstate != "X") && (detailstate != frm.detailstate[i].value)) {
			alert("��ǰ���°� �ٸ� ��ǰ�� ���� ����� �� �����ϴ�.");
			obj.checked = false;
			return;
		}

		detailstate = frm.detailstate[i].value;
	}
}

function EnableOneBrand() {
	var frm = document.frm;
	var isupchebeasong, makerid, ischeckedexist;
	var curisupchebeasong, curmakerid, curischecked;
	var detailstate;

	detailstate = "";
	ischeckedexist = false;
	makerid = "";
	isupchebeasong = "";
	for (var i = 0; i < frm.orderdetailidx.length; i++) {
		if (frm.orderdetailidx[i].type != "checkbox") {
			continue;
		}

		if (frm.cancelyn[i].value == "Y") {
			continue;
		}

		curischecked = frm.orderdetailidx[i].checked;
		curisupchebeasong = frm.isupchebeasong[i].value;
		curmakerid = frm.makerid[i].value;

		if (curischecked == true) {
			if (ischeckedexist == true) {
				continue;
			} else {
				ischeckedexist = frm.orderdetailidx[i].checked;
				isupchebeasong = frm.isupchebeasong[i].value;
				detailstate = frm.detailstate[i].value;
				makerid = frm.makerid[i].value;
			}
		}
	}

	if (ischeckedexist != true) {
		RemoveAddedItemALL();
	}

	for (var i = 0; i < frm.orderdetailidx.length; i++) {
		if (frm.orderdetailidx[i].type != "checkbox") {
			continue;
		}

		if (frm.cancelyn[i].value == "Y") {
			continue;
		}

		curisupchebeasong = frm.isupchebeasong[i].value;
		curmakerid = frm.makerid[i].value;
		curischecked = frm.orderdetailidx[i].checked;

		if (ischeckedexist == true) {
			if (isupchebeasong == "N") {
				// �ٹ�
				if (curisupchebeasong == "Y") {
					frm.orderdetailidx[i].checked = false;
					frm.orderdetailidx[i].disabled = true;
				}
			} else {
				// ����
				if ((curisupchebeasong == "N") || (curmakerid != makerid)) {
					frm.orderdetailidx[i].checked = false;
					frm.orderdetailidx[i].disabled = true;
				}
			}
		} else {
			frm.orderdetailidx[i].disabled = false;
		}

		AnCheckClick(frm.orderdetailidx[i]);
	}

	frm.editmakerid.value = makerid;
	frm.editisupchebeasong.value = isupchebeasong;
	frm.editdetailstate.value = detailstate;

	CalculateCancelSUM();
}

function CheckMaxItemNo(obj, maxno) {
	if (obj.value*0 != 0) {
		return;
	}

    if (obj.value*1 > maxno*1) {
        alert("�ֹ����� �̻����� ��ǰ������ �����Ҽ� �����ϴ�.");
        obj.value = maxno;
    }

	CalculateCancelSUM();
}

function CheckItemNoAdded(obj, idx) {
	if (obj.value*0 != 0) {
		return;
	}

    if (obj.value*1 < 1) {
        alert("�߰������� 1 ���� ���� �� �����ϴ�.");
        obj.value = 1;
    }

	arrToItemNo[idx] = obj.value*1;
	CalculateAddSUM();
}

function CalculateCancelSUM() {
	var frm = document.frm;
	var refundrequire, canceltotal, refunditemcostsum, refundcouponsum, allatsubtractsum;
	var currefunditemcost, currefundcoupon, curallatsubtract;

	refundrequire = 0;
	canceltotal = 0;
	refunditemcostsum = 0;
	refundcouponsum = 0;
	allatsubtractsum = 0;

	for (var i = 0; i < frm.orderdetailidx.length; i++) {
		if (frm.orderdetailidx[i].type != "checkbox") {
			continue;
		}

		if (frm.cancelyn[i].value == "Y") {
			continue;
		}

		if (frm.orderdetailidx[i].checked != true) {
			continue;
		}

		currefunditemcost = frm.itemcost[i].value*1 * frm.regitemno[i].value*1;
		currefundcoupon = frm.percentBonusCouponDiscount[i].value*1 * frm.regitemno[i].value*1;
		curallatsubtract = frm.allatitemdiscount[i].value*1 * frm.regitemno[i].value*1;

		canceltotal = canceltotal + currefunditemcost;
		refundcouponsum = refundcouponsum + currefundcoupon;
		allatsubtractsum = allatsubtractsum + curallatsubtract;
	}

	refunditemcostsum = canceltotal;
	refundrequire = refunditemcostsum - (refundcouponsum + allatsubtractsum);

	frm.refundrequire.value = refundrequire;
	frm.canceltotal.value = canceltotal;
	frm.refunditemcostsum.value = refunditemcostsum;
	frm.refundcouponsum.value = refundcouponsum;
	frm.allatsubtractsum.value = allatsubtractsum;

	frm.itemcanceltotal.value = canceltotal;
	frm.itemrefundrequire.value = refundrequire;
}

function CalculateAddSUM() {
	var frm = document.frm;

	var addrequire, addtotal, additemcostsum, addcouponsum, addallatsubtractsum;
	var curadditemcost, curaddcoupon, curaddallatsubtract;

	addrequire = 0
	addtotal = 0
	additemcostsum = 0
	addcouponsum = 0
	addallatsubtractsum = 0

	for (var i = 0; i < arrToItemId.length; i++) {
		curadditemcost = arrItemCouponPrice[i]*1 * arrToItemNo[i]*1;
		if (IsPercentBonusCouponExist == true) {
			curaddcoupon = (arrItemCouponPrice[i]*1 - arrBonusCouponPrice[i]*1) * arrToItemNo[i]*1;
			curaddallatsubtract = 0;
		} else {
			curaddcoupon = 0;
			curaddallatsubtract = (arrItemCouponPrice[i]*1 - arrBonusCouponPrice[i]*1) * arrToItemNo[i]*1;
		}

		additemcostsum = additemcostsum + curadditemcost;
		addcouponsum = addcouponsum + curaddcoupon;
		addallatsubtractsum = addallatsubtractsum + curaddallatsubtract;
	}

	addtotal = additemcostsum;
	addrequire = additemcostsum - (addcouponsum + addallatsubtractsum);

	frm.addrequire.value = addrequire;
	frm.addtotal.value = addtotal;
	frm.additemcostsum.value = additemcostsum;
	frm.addcouponsum.value = addcouponsum;
	frm.addallatsubtractsum.value = addallatsubtractsum;

	if (arrToItemId.length > 0) {
		frm.itemaddtotal.value = addtotal;
		frm.itemaddrequire.value = addrequire;
	}
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="2">
			<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ֹ����������� ����</b>
		</td>
	</tr>

	<form name="frm" method="post" action="orderdetail_process.asp">
	<input type="hidden" name="orderserial" value="<%= oordermaster.FOneItem.FOrderSerial %>">
	<input type="hidden" name="mode" value="itemChangeArray">
	<input type="hidden" name="forceedit" value="N">

	<input type="hidden" name="editmakerid">
	<input type="hidden" name="editisupchebeasong">
	<input type="hidden" name="editdetailstate">

	<input type="hidden" name="title" value="�ֹ�����(���ϼҺ��ڰ�)">
	<input type="hidden" name="contents_jupsu" value="">
	<input type="hidden" name="contents_finish" value="�ֹ������� ���������� ó���Ǿ����ϴ�.">

	<input type="hidden" name="refundrequire" value="">
	<input type="hidden" name="canceltotal" value="">
	<input type="hidden" name="refunditemcostsum" value="">
	<input type="hidden" name="refundcouponsum" value="">
	<input type="hidden" name="allatsubtractsum" value="">

	<input type="hidden" name="addrequire" value="">
	<input type="hidden" name="addtotal" value="">
	<input type="hidden" name="additemcostsum" value="">
	<input type="hidden" name="addcouponsum" value="">
	<input type="hidden" name="addallatsubtractsum" value="">

	<input type="hidden" name="arrFromItemId" value="">
	<input type="hidden" name="arrFromItemOption" value="">
	<input type="hidden" name="arrFromItemNo" value="">

	<input type="hidden" name="arrToItemId" value="">
	<input type="hidden" name="arrToItemOption" value="">
	<input type="hidden" name="arrToItemNo" value="">

	<input type="hidden" name="arrToSalePrice" value="">
	<input type="hidden" name="arrToItemCouponPrice" value="">
	<input type="hidden" name="arrToBonusCouponPrice" value="">
	<input type="hidden" name="arrToBuyCash" value="">

	<input type="hidden" name="toSaleMethod" value="">
	<input type="hidden" name="toBonusCouponIdx" value="">
	<input type="hidden" name="arrToItemCouponIdx" value="">

	<input type="hidden" name="arrFromDetailIdx" value="">

	<tr height="25" bgcolor="#FFFFFF" >
		<td width="100" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ</td>
		<td><%= oordermaster.FOneItem.Forderserial %></td>
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

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>"><b>��� ����ǰ</b></td>
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

			<input type="hidden" name="orderdetailidx">
			<input type="hidden" name="itemcost">
			<input type="hidden" name="allatitemdiscount">
			<input type="hidden" name="percentBonusCouponDiscount">
			<input type="hidden" name="itemno">
			<input type="hidden" name="regitemno">
			<input type="hidden" name="makerid">
			<input type="hidden" name="isupchebeasong">
			<input type="hidden" name="cancelyn">
			<input type="hidden" name="itemid" value="">
			<input type="hidden" name="itemoption" value="">
			<input type="hidden" name="detailstate" value="">

			<input type="hidden" name="orderdetailidx">
			<input type="hidden" name="itemcost">
			<input type="hidden" name="allatitemdiscount">
			<input type="hidden" name="percentBonusCouponDiscount">
			<input type="hidden" name="itemno">
			<input type="hidden" name="regitemno">
			<input type="hidden" name="makerid">
			<input type="hidden" name="isupchebeasong">
			<input type="hidden" name="cancelyn">
			<input type="hidden" name="itemid" value="">
			<input type="hidden" name="itemoption" value="">
			<input type="hidden" name="detailstate" value="">

<% for i = 0 to oorderdetail.FResultCount - 1 %>
	<% if (oorderdetail.FItemList(i).Fitemid <> 0) then %>
		<% if (oorderdetail.FItemList(i).FCancelyn = "Y") then %>
			<tr align="center" bgcolor='#CCCCCC' class='gray'>
		<% else %>
			<tr align="center" bgcolor='#FFFFFF' >
		<% end if %>
				<td height="25">
					<input type="checkbox" name="orderdetailidx" onClick="CheckDifferentStateItemExist(this); EnableOneBrand();" value="<%= oorderdetail.FItemList(i).Fidx %>" <% if (oorderdetail.FItemList(i).FCancelyn = "Y") then %>disabled<% end if %>>
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
					<input type="text" name="regitemno" value="<%= oorderdetail.FItemList(i).Fitemno %>" size="2" style="text-align:center" onKeyUp="CheckMaxItemNo(this, <%= oorderdetail.FItemList(i).FItemNo %>);">
					/
					<input type="text" name="itemno" value="<%= oorderdetail.FItemList(i).Fitemno %>" size="2" style="text-align:center;background-color:#DDDDFF;" readonly>
				</td>
				<input type="hidden" name="itemcost" value="<%= oorderdetail.FItemList(i).Fitemcost %>">
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

				<input type="hidden" name="allatitemdiscount" value="<%= oorderdetail.FItemList(i).getAllAtDiscountedPrice %>">
				<input type="hidden" name="percentBonusCouponDiscount" value="<%= oorderdetail.FItemList(i).getPercentBonusCouponDiscountedPrice %>">

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
				<input type="hidden" name="cancelyn" value="<%= oorderdetail.FItemList(i).FCancelyn %>">
				<input type="hidden" name="isupchebeasong" value="<%= oorderdetail.FItemList(i).Fisupchebeasong %>">
				<input type="hidden" name="makerid" value="<%= oorderdetail.FItemList(i).Fmakerid %>">
				<input type="hidden" name="itemid" value="<%= oorderdetail.FItemList(i).Fitemid %>">
				<input type="hidden" name="itemoption" value="<%= oorderdetail.FItemList(i).Fitemoption %>">
				<input type="hidden" name="detailstate" value="<%= oorderdetail.FItemList(i).Fcurrstate %>">
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
        	            <td align="right"><input type="text" name="itemcanceltotal" size="7" value="0" readonly style="text-align:right;border: 1px solid #333333;" ></td>
        	        </tr>
        	        <tr>
        	            <td>ȯ�ұݾ��հ�</td>
        	            <td align="right"><input type="text" name="itemrefundrequire" size="7" value="0" readonly style="text-align:right;border: 1px solid #333333;" ></td>
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

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="#FFFFFF" >
		<td bgcolor="<%= adminColor("tabletop") %>" colspan="3">
			<b><font color="red">�߰� ����ǰ</font></b>
			<input type="button" class="button" value="���Ϻ귣��" onClick="javascript:SearchItemByMakerid()">
			<input type="button" class="button" value="�����ǸŰ�" onClick="javascript:SearchItemByPrice()">
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
		<td colspan="3">
			<div id="TableAddedItem"></div>
		</td>
	</tr>
	</form>
	<tr bgcolor="#FFFFFF" height="40">
		<td colspan="3" align="center">
<% if Not IsOrderCanceled then %>
			<input type="button" class="button" value="�ֹ�����" onclick="javascript:SaveChangeItem(false)" disabled>
			<% if (C_CSPowerUser or C_ADMIN_AUTH) then %>
		    <input type="button" class="button" value="��������" onclick="javascript:SaveChangeItem(true)">
			<% end if %>
<% else %>
			<b>��ҵ� �ֹ��� �ֹ����� �Ұ�</b>
<% end if %>
		</td>
	</tr>
</table>

<p>

<!--
<div>
* �ٸ���ǰ�̾ <font color=red>��ü�� ����</font>�� ���, �������� ����(���ϻ�ǰ) �� �����ϼ���.
</div>
-->

<iframe name="iframeforadd" width="0" height="0">
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->