<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  ������ �ֹ��� �ۼ�
' History : 2009.04.07 ������ ����
'			2010.08.12 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/fran_chulgojungsancls.asp"-->
<%

menupos = request("menupos")

dim idx, mode, shopid, jungsanidx, workidx, invoiceidx
dim i, j

idx = request("idx")
mode = request("mode")
shopid = request("shopid")
jungsanidx = request("jungsanidx")
workidx = request("workidx")
invoiceidx = request("invoiceidx")



if idx="" then idx=0

if (invoiceidx <> "") then
	idx = invoiceidx
end if

''response.write idx & "aaa<br />"
''dbget.close : response.end

'================================================================================
dim ocoffinvoice

set ocoffinvoice = new COffInvoice

ocoffinvoice.FRectMasterIdx = idx

ocoffinvoice.GetMasterOne

if (jungsanidx <> "") then
	ocoffinvoice.FOneItem.Fjungsanidx = jungsanidx
end if

if (shopid <> "") then
	ocoffinvoice.FOneItem.Fshopid = shopid
end if

if (shopid <> "") then
	ocoffinvoice.FOneItem.Fshopid = shopid
end if

if (workidx <> "") then
	ocoffinvoice.FOneItem.Fworkidx = workidx
end if

'================================================================================
dim ocoffinvoicedetail

set ocoffinvoicedetail = new COffInvoice

ocoffinvoicedetail.FRectMasterIdx = idx
ocoffinvoicedetail.FRectShopid = ocoffinvoice.FOneItem.Fshopid

ocoffinvoicedetail.GetDetailList

'================================================================================
dim ocoffinvoiceproductdetail

set ocoffinvoiceproductdetail = new COffInvoice

ocoffinvoiceproductdetail.FRectMasterIdx = idx
ocoffinvoiceproductdetail.FRectShopid = ocoffinvoice.FOneItem.Fshopid

ocoffinvoiceproductdetail.GetProductDetailList

'================================================================================
dim ocartoonboxmaster

set ocartoonboxmaster = new CCartoonBox

ocartoonboxmaster.FRectMasterIdx = ocoffinvoice.FOneItem.Fworkidx

ocartoonboxmaster.GetMasterOne

'================================================================================
dim ofranchulgomaster

set ofranchulgomaster = new CFranjungsan

ofranchulgomaster.FRectidx = ocoffinvoice.FOneItem.Fjungsanidx
if (ofranchulgomaster.FRectidx = 0) then
	ofranchulgomaster.FRectidx = ""
end if

if ((ofranchulgomaster.FRectidx = "") or IsNull(ofranchulgomaster.FRectidx)) and ocoffinvoice.FOneItem.Fworkidx <> "" then
	'// �ؿ������
	ofranchulgomaster.getOneFranJungsanBlank
	ofranchulgomaster.FOneItem.FcurrencyUnit = ocartoonboxmaster.FOneItem.FcurrencyUnit
	ofranchulgomaster.FOneItem.FtotalforeignSuplycash = ocartoonboxmaster.FOneItem.Ftotforeign_suplycash
	ofranchulgomaster.FOneItem.Ftotalsum = ocartoonboxmaster.FOneItem.Ftotsuplycash
else
	'// ���곻��
	ofranchulgomaster.getOneFranJungsan
end if



'================================================================================
dim ofranchulgodetail
dim totaljungsandeliverpay

set ofranchulgodetail = new CFranjungsan

ofranchulgodetail.FPageSize=200
ofranchulgodetail.FRectIDx = ocoffinvoice.FOneItem.Fjungsanidx

if ((ofranchulgomaster.FRectidx = "") or IsNull(ofranchulgomaster.FRectidx)) and ocoffinvoice.FOneItem.Fworkidx <> "" then
	totaljungsandeliverpay = ocartoonboxmaster.FOneItem.Fdeliverpay
	ofranchulgomaster.FOneItem.Ftotalsum = ofranchulgomaster.FOneItem.Ftotalsum + totaljungsandeliverpay
else
	ofranchulgodetail.getFranMaeipSubmasterList

	totaljungsandeliverpay = 0
	for i = 0 to  ofranchulgodetail.FResultCount - 1
		'// �ֹ��ڵ尡 temp �̸� EMS��ۺ�� �����Ѵ�.
		'// ���� ��Ÿ������ �ԷµǸ� ���̰� �߻��ϰ�, �ڵ��Է±���� ������� ���Ѵ�.(�����Է�)
		if (ofranchulgodetail.FItemList(i).Fcode02 = "temp") then
			totaljungsandeliverpay	=	totaljungsandeliverpay + ofranchulgodetail.FItemList(i).Ftotalsuplycash
		end if
	next
end if


'================================================================================
dim avggoodsprice

if (ocoffinvoicedetail.FResultCount = "") then
	ocoffinvoicedetail.FResultCount = 0
end if

if (ocoffinvoice.FOneItem.FtotalPriceForeign = "") or (IsNull(ocoffinvoice.FOneItem.FtotalPriceForeign) = True) then
	ocoffinvoice.FOneItem.FtotalPriceForeign = 0
end if

if (ocoffinvoicedetail.FResultCount = 0) then
	avggoodsprice = 0
else
	avggoodsprice = ocoffinvoice.FOneItem.FtotalPriceForeign / ocoffinvoicedetail.FResultCount
end if

dim emsPrice

emsPrice = 0

%>

<script language='javascript'>

function SaveMaster(frm, mode, submode, descidx) {
	if (CheckMaster(frm) != true) {
		return;
	}

	if (frm.masteridx.value*1 != 0) {
		if (CheckProductList(frm) != true) {
			return;
		}

		if (submode == "adddetailone") {
			if (CheckProductNew(frm) != true) {
				return;
			}
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret == true) {
		frm.totalGoodsPriceWon.value = frm.totalGoodsPriceWon.value.replace(/,/g, "");
		frm.totalDeliverPriceWon.value = frm.totalDeliverPriceWon.value.replace(/,/g, "");
		frm.exchangerate.value = frm.exchangerate.value.replace(/,/g, "");

		frm.totalGoodsPriceForeign.value = frm.totalGoodsPriceForeign.value.replace(/,/g, "");
		frm.totalDeliverPriceForeign.value = frm.totalDeliverPriceForeign.value.replace(/,/g, "");
		frm.totalPriceForeign.value = frm.totalPriceForeign.value.replace(/,/g, "");

		if (frm.masteridx.value*1 == 0) {
			frm.mode.value="newmaster";
		} else {
			frm.mode.value="savemaster";
			frm.productdetailmode.value = submode;
			frm.productdetailidx.value = descidx;
		}

		frm.submit();
	}
}

function SetFinishState(frm) {
	// �ۼ��Ϸ� ��ȯ
	if (confirm("�ۼ��Ϸ� ���·� ��ȯ�Ͻðڽ��ϱ�?\n\n�κ��̽� ������ ���忡 ���µ˴ϴ�.") != true) {
		return;
	}

	frm.mode.value="modifystate";
	frm.statecd.value = 7;
	frm.submit();
}

function SetWriteState(frm) {
	// �ۼ��Ϸ� ��ȯ
	if (confirm("�ۼ��� ���·� ��ȯ�Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.mode.value="modifystate";
	frm.statecd.value = 1;
	frm.submit();
}

function SaveProductComment(frm, submode, productidx) {
	alert("�������");
	return;

	if (CheckMaster(frm) != true) {
		return;
	}

	if (CheckProductList(frm) != true) {
		return;
	}

	if (submode == "addproduct") {
		if (CheckProductNew(frm) != true) {
			return;
		}
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if (ret == true) {
		frm.productdetailmode.value = submode;
		frm.productdetailidx.value = productidx;

		if (frm.masteridx.value*1 == 0) {
			frm.mode.value="newmaster";
		} else {
			frm.mode.value="savemaster";
		}

		frm.submit();
	}
}

function InsertDetailFromWork(frm) {
	var ret = confirm('���� �ڽ������� ��� �����˴ϴ�.\n\n���� �Ͻðڽ��ϱ�?');

	if (ret == true) {
		frm.mode.value="insertdetailfromwork";

		frm.submit();
	}
}

function InsertDefaultDescription(frm) {
	var ret = confirm('���� ��ǰ������ ��� �����˴ϴ�.\n\n���� �Ͻðڽ��ϱ�?');

	if (ret == true) {
		frm.mode.value="insertdefaultdescription";

		frm.submit();
	}
}

function DelMaster(frm) {
	var ret = confirm('��ü���� �Ͻðڽ��ϱ�?');

	if (ret) {
		frm.mode.value="delmaster";
		frm.submit();
	}
}

function CheckMaster(frm) {
	RecalcPrice(frm);

	var totalGoodsPriceWon = frm.totalGoodsPriceWon.value.replace(/,/g, "");
	var totalDeliverPriceWon = frm.totalDeliverPriceWon.value.replace(/,/g, "");
	var exchangerate = frm.exchangerate.value.replace(/,/g, "");

	if (frm.shopid.value == "") {
		alert("���� �����ϼ���.");
		frm.shopid.focus();
		return false;
	}

	if (frm.workidx.value == "") {
		alert("����۾��� ���� �����ϼ���.");
		return false;
	}

	//if (frm.jungsanidx.value == "") {
	//	alert("���곻���� ���� �����ϼ���.");
	//	return false;
	//}

	if (frm.workidx.value != "") {
		if (frm.workidx.value*0 != 0) {
			alert("���ڸ� �Է� �����մϴ�.");
			frm.workidx.focus();
			return false;
		}
	}

	if (totalGoodsPriceWon != "") {
		if (totalGoodsPriceWon*0 != 0) {
			alert("���ڸ� �Է� �����մϴ�.");
			frm.totalGoodsPriceWon.focus();
			return false;
		}
	}

	if (totalDeliverPriceWon != "") {
		if (totalDeliverPriceWon*0 != 0) {
			alert("���ڸ� �Է� �����մϴ�.");
			frm.totalDeliverPriceWon.focus();
			return false;
		}
	}

	if (exchangerate != "") {
		if (exchangerate*0 != 0) {
			alert("���ڸ� �Է� �����մϴ�.");
			frm.exchangerate.focus();
			return false;
		}
	}

	if (frm.delivermethod.value == "E") {
		// EMS ����� ��츸 ��ۺ� ������ �ִ�.
		if (frm.totalDeliverPriceWon.value*1 == 0) {
			//alert("EMS����̸鼭 ������ �����ϴ�.");
			//frm.totalDeliverPriceWon.focus();
			//return false;
		}

		if (frm.exportmethod.value != "C") {
			alert("EMS����Դϴ�. ���������ε�(��������)�� �����ϼ���.");
			frm.exportmethod.focus();
			return false;
		}
	}

	if (frm.reportforeigntotalprice) {
		if (frm.reportno.value != "") {
			var reportforeigntotalprice = frm.reportforeigntotalprice.value.replace(/,/g, "");
			var reporttotalprice = frm.reporttotalprice.value.replace(/,/g, "");
			var reportexchangerate = frm.reportexchangerate.value.replace(/,/g, "");

			if ((reportforeigntotalprice != "") || (reporttotalprice != "")) {
				if ((reportforeigntotalprice*1 == reporttotalprice*1) && (reportexchangerate*1 != 1)) {
					alert("�Ű�ȯ���� �ڵ��Է� ��ư�� ��������.");
					return false;
				}

				if ((reportforeigntotalprice*1 != reporttotalprice*1) && (reportexchangerate*1 == 1)) {
					alert("�Ű�ȯ���� �ڵ��Է� ��ư�� ��������.");
					return false;
				}
			}
		}
	}

	return true;
}

function CheckProductList(frm) {
	var totalcount = frm.productdetailcount.value*1;
	var orderno;
	var goodscomment;

	var totalboxno;
	var priceperbox;
	var totalPriceForeign;

	for (var i = 0; i < totalcount; i++) {
		orderno = eval(frm.name + ".orderno_" + i);
		goodscomment = eval(frm.name + ".goodscomment_" + i);
		totalboxno = eval(frm.name + ".totalboxno_" + i);
		priceperbox = eval(frm.name + ".priceperbox_" + i);
		totalprice = eval(frm.name + ".totalprice_" + i);

		totalboxno.value = totalboxno.value.replace(/,/g, "");
		totalprice.value = totalprice.value.replace(/,/g, "");

		if (orderno.value == "") {
			alert("ǥ�ü����� �����ϼ���.");
			orderno.focus();
			return false;
		}

		if (orderno.value*0 != 0) {
			alert("ǥ�ü����� ���ڸ� �Է� �����մϴ�.");
			orderno.focus();
			return false;
		}

		if (totalboxno.value != "") {
			if (totalboxno.value*0 != 0) {
				alert("���ڸ� �Է� �����մϴ�.");
				totalboxno.focus();
				return false;
			}

			if (totalboxno.value*1 == 0) {
				alert("0 �� �Է��� �� �����ϴ�.");
				totalboxno.focus();
				return false;
			}
		}

		if (totalprice.value != "") {
			if (totalprice.value*0 != 0) {
				alert("���ڸ� �Է� �����մϴ�.");
				totalprice.focus();
				return false;
			}
		}

		if ((totalboxno.value != "") && (totalprice.value != "")) {
			priceperbox.value = totalprice.value*1 / totalboxno.value*1;
		} else {
			priceperbox.value = "";
		}
	}

	return true;
}

function CheckProductNew(frm) {
	frm.totalprice_new.value = frm.totalprice_new.value.replace(/,/g, "");

	if (frm.orderno_new.value == "") {
		alert("ǥ�ü����� �����ϼ���.");
		frm.orderno_new.focus();
		return false;
	}

	if (frm.orderno_new.value*0 != 0) {
		alert("ǥ�ü����� ���ڸ� �Է� �����մϴ�.");
		frm.orderno_new.focus();
		return false;
	}

	if (frm.totalboxno_new.value != "") {
		if (frm.totalboxno_new.value*0 != 0) {
			alert("���ڸ� �Է� �����մϴ�.");
			frm.totalboxno_new.focus();
			return false;
		}

		if (frm.totalboxno_new.value*1 == 0) {
			alert("0 �� �Է��� �� �����ϴ�.");
			frm.totalboxno_new.focus();
			return false;
		}
	}

	if (frm.totalprice_new.value != "") {
		if (frm.totalprice_new.value*0 != 0) {
			alert("���ڸ� �Է� �����մϴ�.");
			frm.totalprice_new.focus();
			return false;
		}
	}

	if ((frm.totalboxno_new.value != "") && (frm.totalprice_new.value != "")) {
		frm.priceperbox_new.value = frm.totalprice_new.value*1 / frm.totalboxno_new.value*1;
	} else {
		frm.priceperbox_new.value = "";
	}

	return true;
}

function RecalcInvoiceNo(frm) {
	if (frm.shopid.value == "") {
		alert("���� �����ϼ���.");
		frm.shopid.focus();
		return;
	}

	if (frm.invoicedate.value == "") {
		alert("�κ��̽� ���ڸ� �����ϼ���.");
		frm.invoicedate.focus();
		return;
	}

	if ((frm.workidx.value == "") || (frm.workidx.value*1 == 0)) {
		alert("�۾��ڵ带 �����ϼ���.");
		frm.workidx.focus();
		return;
	}

	//if ((frm.jungsanidx.value == "") || (frm.jungsanidx.value*1 == 0)) {
	//	alert("�����ڵ尡 �������� �ʾҽ��ϴ�.\n\n���곻������ �۾��ڵ带 �����ϼ���.");
	//	return;
	//}

	frm.invoiceno.value = Right(frm.shopid.value, 3) + "_" + frm.workidx.value + "_" + frm.invoicedate.value.replace(/-/g, "");
}

function CalcReportExchangeRate(frm) {
	var pointprice = 2;		// �Ҽ������� ���ڸ����� ���

	var reportforeigntotalprice = frm.reportforeigntotalprice.value.replace(/,/g, ""); // ��ȯ
	var reporttotalprice = frm.reporttotalprice.value.replace(/,/g, ""); // ��ȭ
	var reportpriceunit = frm.reportpriceunit.value;
	var reportexchangerate = 0;

	if ((reportpriceunit == "JPY") || (reportpriceunit == "WON")) {
		pointprice = 0;
	}

	if (reportpriceunit == "") {
		alert("�Ű� ��ȭ�ڵ带 �Է��ϼ���.");
		frm.reportpriceunit.focus();
		return;
	}

	if ((reportforeigntotalprice == "") || (reportforeigntotalprice*1 == 0)) {
		alert("�Ű� �ݾ�(��ȭ)�� �Է��ϼ���.");
		frm.reportforeigntotalprice.focus();
		return;
	}

	if ((reporttotalprice == "") || (reporttotalprice*1 == 0)) {
		alert("�Ű� �ݾ�(��ȭ)�� �Է��ϼ���.");
		frm.reporttotalprice.focus();
		return;
	}

	reportforeigntotalprice = (reportforeigntotalprice * 1.0).toFixed(pointprice); // ��ȯ
	reporttotalprice = (reporttotalprice * 1.0).toFixed(0); // ��ȭ

	if (reportpriceunit == "JPY") {
		// ��ȭ�� 100�� ������ ȯ���� ����Ѵ�.
		reportexchangerate = (100.0 * reporttotalprice / reportforeigntotalprice);
	} else {
		reportexchangerate = (1.0 * reporttotalprice / reportforeigntotalprice);
	}

	// ȯ���� �Ҽ��� ��°�ڸ�����
	reportexchangerate = reportexchangerate.toFixed(2);

	frm.reportforeigntotalprice.value = reportforeigntotalprice;
	frm.reporttotalprice.value = reporttotalprice;
	frm.reportexchangerate.value = reportexchangerate;
}



function popJungsanMaster(iid){
	var popwin = window.open('/admin/offshop/franmeaippopsubmaster.asp?idx=' + iid,'popsubmaster','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

function FormatNumber(nStr) {
	nStr += '';

	x = nStr.split('.');
	x1 = x[0];
	x2 = x.length > 1 ? '.' + x[1] : '';

	var rgx = /(\d+)(\d{3})/;
	while (rgx.test(x1)) {
		x1 = x1.replace(rgx, '$1' + ',' + '$2');
	}

	return x1 + x2;
}

function RecalcPriceAndAlert(frm) {
	var errMsg = RecalcPrice(frm);
	if (errMsg != "") {
		alert(errMsg);
	}
}

function RecalcPrice(frm) {
	var exchangerate 		= frm.exchangerate.value.replace(/,/g, "");			// ȯ��

	var totalDeliverPriceWon	= frm.totalDeliverPriceWon.value.replace(/,/g, "");		// ����(��)
	var totalGoodsPriceWon 		= frm.totalGoodsPriceWon.value.replace(/,/g, "");		// ��ǰ�ݾ�(��)
	var totalPriceWon			= 0;													// �հ�(��)

	var priceunit 			= frm.priceunit.value;										// �ۼ�ȭ��

	var totalGoodsPriceForeign 		= 0;												// ��ǰ�ݾ�(��ȯ)
	var totalDeliverPriceForeign 	= 0;												// ����(��ȯ)
	var totalPriceForeign			= 0;												// �հ�(��ȯ)

	// ========================================================================
	if ((exchangerate == "") || (exchangerate*0 != 0) || (exchangerate*1 == 0)) {
		return "ȯ���� �Է��ϼ���.";
	}

	if ((totalGoodsPriceWon == "") || (totalGoodsPriceWon*0 != 0) || (totalGoodsPriceWon*1 == 0)) {
		return "��ǰ�ݾ��� �Է��ϼ���";
	}

	if ((totalDeliverPriceWon == "") || (totalDeliverPriceWon*0 != 0)) {
		// ������ ���� �� �ִ�.
		totalDeliverPriceWon = 0;
	}

	if (priceunit == "") {
		return "�ۼ�ȭ�� �����ϼ���.";
	}


	// ========================================================================
	exchangerate = exchangerate*1;
	totalDeliverPriceWon = totalDeliverPriceWon*1;
	totalGoodsPriceWon = totalGoodsPriceWon*1;
	totalPriceWon = totalDeliverPriceWon + totalGoodsPriceWon;

	var pointprice = 2;		// �Ҽ������� ���ڸ����� ���

	if (priceunit == "JPY") {
		// ��ȭ�� 100�� ������ ȯ���� ����Ѵ�.
		exchangerate = exchangerate*1 / 100;
		pointprice = 0;
	}

	<% if (ofranchulgomaster.FOneItem.FcurrencyUnit = "") or (ofranchulgomaster.FOneItem.FtotalforeignSuplycash = 0) then %>
		// �ֹ����� ��ȯ�ݾ��� ���� ��츸 ȯ���� ���
		totalGoodsPriceForeign = (totalGoodsPriceWon / exchangerate).toFixed(pointprice);
		totalDeliverPriceForeign = (totalDeliverPriceWon / exchangerate).toFixed(pointprice);
		totalPriceForeign = ((totalGoodsPriceWon + totalDeliverPriceWon) / exchangerate).toFixed(pointprice);
	<% else %>
		// �ֹ����� ��ȯ�ݾ��� ������ ��ȯ���� ���Ѵ�.
		totalGoodsPriceForeign = frm.totalGoodsPriceForeign.value.replace(/,/g, "");
		totalGoodsPriceForeign = (totalGoodsPriceForeign*1.0).toFixed(pointprice);
		totalDeliverPriceForeign = (totalDeliverPriceWon / exchangerate).toFixed(pointprice);
		totalPriceForeign = (totalGoodsPriceForeign*1.0 + totalDeliverPriceForeign*1.0).toFixed(pointprice);
	<% end if %>

	// ========================================================================
	frm.totalDeliverPriceWon.value = FormatNumber(totalDeliverPriceWon);
	frm.totalGoodsPriceWon.value = FormatNumber(totalGoodsPriceWon);
	frm.totalPriceWon.value = FormatNumber(totalPriceWon);

	frm.totalDeliverPriceForeign.value = FormatNumber(totalDeliverPriceForeign);
	frm.totalGoodsPriceForeign.value = FormatNumber(totalGoodsPriceForeign);
	frm.totalPriceForeign.value = FormatNumber(totalPriceForeign);

	frm.tmpTotalPriceWon.value = FormatNumber(totalPriceWon);
	frm.tmpTotalPriceForeign.value = FormatNumber(totalPriceForeign);

	return "";
}

function RecalcEMSPrice(frmmaster, frmdetail) {
	if (frmmaster.delivermethod.value != "E") {
		alert("EMS����� ��쿡�� EMS �ݾ��� ������ �� �ֽ��ϴ�.");
		return;
	}

	frmmaster.totalDeliverPriceWon.value = frmdetail.emsPrice.value;
}

function Right(str, n) {
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}

function PopExportSheet(v){
	var popwin;
	popwin = window.open('/admin/fran/cartoonbox_modify.asp?menupos=1357&idx=' + v ,'PopExportSheet','width=740,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopPrevInvoiceList(frm, mode) {
	var popwin;
	popwin = window.open('/admin/fran/popoffinvoice_list.asp?shopid=' + frm.shopid.value + '&frm=' + frm.name + '&mode=' + mode ,'PopPrevInvoiceList','width=1000,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function CopyDeliverAndPriceInfo(frm) {
	var workidx 		= "<%= ocoffinvoice.FOneItem.Fworkidx %>";
	var delivermethod 	= "<%= ocartoonboxmaster.FOneItem.Fdelivermethod %>";
	var deliverpay 		= "<%= ocartoonboxmaster.FOneItem.Fdeliverpay %>";

	var jungsanidx 				= "<%= ocoffinvoice.FOneItem.Fjungsanidx %>";
	var totalPriceWon			= "<%= ofranchulgomaster.FOneItem.Ftotalsum %>";
	var totaljungsandeliverpay 	= "<%= totaljungsandeliverpay %>";	// ���곻������ EMS��ۺ�(��ۺ� �̿��� ��Ÿ������ ������ ���̰� �߻��Ѵ�.)

	var totalforeignSuplycash	= "<%= ofranchulgomaster.FOneItem.FtotalforeignSuplycash %>";
	var currencyUnit 			= "<%= ofranchulgomaster.FOneItem.FcurrencyUnit %>";

	if (workidx == "") {
		alert("����۾��� �����Ǿ� ���� �ʽ��ϴ�.");
		return;
	}

	//if (jungsanidx == "") {
	//	alert("���곻���� �����Ǿ� ���� �ʽ��ϴ�.");
	//	return;
	//}

	if (delivermethod == "") {
		alert("����۾��� ��۹���� �����Ǿ� ���� �ʽ��ϴ�.");
		return;
	}

	if (delivermethod == "E") {
		// EMS �� ���
		if ((deliverpay == "") || (deliverpay*0 != 0)) {
			alert("����۾��� EMS����� �߸� �ԷµǾ� �ֽ��ϴ�.\n - �ؿ����EMS��� : (" + deliverpay + ")");
			return;
		}

		if (deliverpay*1 == 0) {
			deliverpay = 0;
			//alert("����۾��� EMS����� �߸� �ԷµǾ� �ֽ��ϴ�.\n - �ؿ����EMS��� : (" + deliverpay + ")");
			//return;
		}
	} else {
		deliverpay = 0;
	}

	if (deliverpay*1 != totaljungsandeliverpay*1) {
		alert("����۾��� ���곻������ EMS��ۺ� ���� Ʋ���ϴ�.\n\n������ Ȯ���ϰ� ����� �Է��ϼ���.\n - �ؿ����EMS��� : (" + deliverpay + ")\n - ���곻��EMS��� : (" + totaljungsandeliverpay + ")");
		return;
	}

	if ((currencyUnit != "") && (totalforeignSuplycash*1 != 0)) {
		frm.priceunit.value = currencyUnit;
		frm.totalGoodsPriceForeign.value = totalforeignSuplycash;
	}

	frm.delivermethod.value = delivermethod;
	frm.totalDeliverPriceWon.value = deliverpay;
	frm.totalGoodsPriceWon.value = (totalPriceWon*1 - totaljungsandeliverpay*1);
	frm.totalPriceWon.value = totalPriceWon*1;
}

function PrintExportSheet(idx, mode,isxl){
	var popwin;
	if (mode == "INVOICE") {
		popwin = window.open('/admin/fran/popoffinvoice_print.asp?idx=' + idx + '&xl='+isxl,'PrintExportSheet','width=850,height=600,scrollbars=yes,resizable=yes');
	} else if (mode == "PACKINGLIST") {
		popwin = window.open('/admin/fran/popoffinvoice_print_packinglist.asp?idx=' + idx + '&xl='+isxl,'PrintExportSheet','width=850,height=600,scrollbars=yes,resizable=yes');
	} else if (mode == "LICENCE") {
		popwin = window.open('/admin/fran/popoffinvoice_print_licence.asp?idx=' + idx + '&xl='+isxl,'PrintExportSheet','width=850,height=600,scrollbars=yes,resizable=yes');
	} else {
		alert("�߸��� �����Դϴ�.");
		return;
	}

	popwin.focus();
}

function PopUploadExportDeclareFile(frm,ino) {
	var popwin;
	popwin = window.open('/admin/fran/popoffinvoice_upload.asp?idx=' + frm.masteridx.value+'&ino='+ino,'PopUploadExportDeclareFile','width=450,height=120,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function PopDownloadExportDeclareFile(frm, ino) {
	var popwin;

	popwin = window.open('<%= uploadImgUrl %>/linkweb/offinvoice/offinvoice_download.asp?idx=' + frm.masteridx.value+'&ino='+ino,'PopDownloadExportDeclareFile','width=100,height=100,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>
<script type="text/javascript">

// qs �˾�
function PopOpenQS(invoiceidx, jungsanidx, workidx, loginsite, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/quotationsheet.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite+ '&cunit='+cunit+'&tpl='+tpl ,'PopOpenQSList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// pi �˾�
function PopOpenPI(invoiceidx, jungsanidx, workidx, loginsite, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/proformainvoice.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite+ '&cunit='+cunit+'&xl=Y&tpl='+tpl,'PopOpenInvoice','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ci �˾�
function PopOpenCI(invoiceidx, jungsanidx, workidx, loginsite, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/commercialinvoice.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite + '&cunit='+cunit+'&tpl='+tpl,'PopOpenCIList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// pl �˾�
function PopOpenPL(invoiceidx, jungsanidx, workidx, loginsite,boxidx, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/packlinglist.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite+'&boxidx='+boxidx+ '&cunit='+cunit+'&tpl='+tpl ,'PopOpenPLList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// pl �˾� ��ǰ
function PopOpenPLItem(invoiceidx,loginsite,boxidx, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/packlingItemlist.asp?idx=' + invoiceidx+'&ls='+ loginsite+'&boxidx='+boxidx+ '&cunit='+cunit+'&tpl='+tpl ,'PopOpenPLIList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="offinvoice_process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="statecd" value="<%= ocoffinvoice.FOneItem.Fstatecd %>">
<input type="hidden" name="productdetailmode" value="">
<input type="hidden" name="productdetailidx" value="">
<input type="hidden" name="masteridx" value="<%= idx %>">

<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>�κ��̽� �⺻����</strong></font>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">IDX</td>
	<td><%= ocoffinvoice.FOneItem.Fidx %></td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�����</td>
	<td>
		<%= ocoffinvoice.FOneItem.Freguserid %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width="200">�����̵�</td>
	<% if ocoffinvoice.FOneItem.Fshopid<>"" then %>
	<input type=hidden name="shopid" value="<%= ocoffinvoice.FOneItem.Fshopid %>">
	<td><%= ocoffinvoice.FOneItem.Fshopid %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">����</td>
	<td><%= ocoffinvoice.FOneItem.Fshopname %></td>
	<% else %>
	<td colspan=3><% drawSelectBoxOffShopNot000 "shopid", ocoffinvoice.FOneItem.Fshopid %></td>
	<% end if %>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width="200">����</td>
	<td><%= ocoffinvoice.FOneItem.GetStateCDName %></td>
	<td bgcolor="<%= adminColor("tabletop") %>"></td>
	<td></td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >����IDX</td>
	<td>
		<input type="text" class="text_ro" name="jungsanidx" value="<%= ocoffinvoice.FOneItem.Fjungsanidx %>" size="6" maxlength="6" style="text-align:right" readonly>
		<% if (ocoffinvoice.FOneItem.Fjungsanidx <> "") then %>
		&nbsp;
		<input type="button" class="button" value="��ȸ�ϱ�" onClick="popJungsanMaster(<%= ocoffinvoice.FOneItem.Fjungsanidx %>)">
		<% else %>
		* ���� : ������ ���� ���
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�۾�IDX</td>
	<td>
		<input type="text" class="text_ro" name="workidx" value="<%= ocoffinvoice.FOneItem.Fworkidx %>" size="6" maxlength="6" style="text-align:right" readonly>
		<% if (ocoffinvoice.FOneItem.Fworkidx <> "") then %>
		&nbsp;
		<input type="button" class="button" value="��ȸ�ϱ�" onClick="PopExportSheet(<%= ocoffinvoice.FOneItem.Fworkidx %>)">
		<% else %>
		* �۾��� ���곻������ ��������
		<% end if %>
	</td>
</tr>

<!-- ��ܹ� ���� -->
<tr height="30" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>�κ��̽� ��������</strong></font>
			        &nbsp;
			        <input type="button" class="button" value="1. �������� �������� �����ϱ�" onClick="PopPrevInvoiceList(frmMaster, 'JUNGSAN')">
			        <input type="button" class="button" value="2. �������/��ǰ�ݾ� ��������" onClick="CopyDeliverAndPriceInfo(frmMaster)">
			        * �������/��ǰ�ݾ� �� ����۾����� �����ɴϴ�.
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >��۹��</td>
	<td>
		<select class="select" name="delivermethod">
			<option value="">����</option>
			<option value="E" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "E") then %>selected<% end if %>>EMS</option>
			<option value="D" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "D") then %>selected<% end if %>>DHL</option>
			<option value="F" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "F") then %>selected<% end if %>>�װ�</option>
			<option value="S" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "S") then %>selected<% end if %>>�ؿ�</option>
			<option value="P" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "P") then %>selected<% end if %>>��������(����)</option>
			<option value="T" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "T") then %>selected<% end if %>>�����ù�</option>
		</select>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >���Ӻδ�</td>
	<td>
		<select class="select" name="exportmethod">
			<option value="">����</option>
			<option value="F" <% if (ocoffinvoice.FOneItem.Fexportmethod = "F") then %>selected<% end if %>>FOB</option>
			<option value="C" <% if (ocoffinvoice.FOneItem.Fexportmethod = "C") then %>selected<% end if %>>C&F</option>
			<option value="W" <% if (ocoffinvoice.FOneItem.Fexportmethod = "W") then %>selected<% end if %>>EXW</option>
			<option value="A" <% if (ocoffinvoice.FOneItem.Fexportmethod = "A") then %>selected<% end if %>>FOB</option>
		</select>
		&nbsp;
		CFR : ���������ε�(��������) / FOB : �����ε� / EXW : �����ε� / FCA : ������ε�
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >������</td>
	<td>
		<select class="select" name="jungsantype">
			<option value="">����</option>
			<option value="T" <% if (ocoffinvoice.FOneItem.Fjungsantype = "T") then %>selected<% end if %>>TT</option>
			<option value="L" <% if (ocoffinvoice.FOneItem.Fjungsantype = "L") then %>selected<% end if %>>LC</option>
		</select>
		&nbsp;
		* TT : ����ȯ�۱�(������) / LC : �ſ���
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�ۼ�ȭ��</td>
	<td>
		<% drawSelectBoxPriceUnit "priceunit", ocoffinvoice.FOneItem.Fpriceunit %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >����ȯ��</td>
	<td>
		<input type="text" class="text" name="exchangerate" value="<%= FormatNumber(ocoffinvoice.FOneItem.Fexchangerate, 2) %>" size=10>��
		&nbsp;
		<input type="button" class="button" value=" 3. �ڵ���� " onClick="RecalcPriceAndAlert(frmMaster)">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" ></td>
	<td>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >��ǰ�ݾ�(��)</td>
	<td>
		<input type="text" class="text" name="totalGoodsPriceWon" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalGoodsPriceWon, 2) %>" size=20>�� (���곻��)
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >��ǰ�ݾ�(��ȯ)</td>
	<td>
		<input type="text" class="text<% if (ofranchulgomaster.FOneItem.FtotalforeignSuplycash = 0) then %>_ro<% end if %>" name="totalGoodsPriceForeign" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalGoodsPriceForeign, 2) %>" size=10 <% if (ofranchulgomaster.FOneItem.FtotalforeignSuplycash = 0) then %>readonly<% end if %>>
		<% if (ofranchulgomaster.FOneItem.FtotalforeignSuplycash <> 0) then %>
			<font color="red">(�ֹ����� : <%= ofranchulgomaster.FOneItem.FcurrencyUnit %> &nbsp; <%= FormatNumber(ofranchulgomaster.FOneItem.FtotalforeignSuplycash, 2) %>)</font>
		<% end if %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >����(��)</td>
	<td>
		<input type="text" class="text" name="totalDeliverPriceWon" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalDeliverPriceWon, 2) %>" size=20>��
		<!--
		&nbsp;
		<input type="button" class="button" value=" EMS���Ӱ������� " onClick="RecalcEMSPrice(frmMaster, frmDetail)">
		-->
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >����(��ȯ)</td>
	<td>
		<input type="text" class="text_ro" name="totalDeliverPriceForeign" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalDeliverPriceForeign, 2) %>" size=10 readonly>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >�հ�(��)</td>
	<td>
		<input type="text" class="text_ro" name="totalPriceWon" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalPriceWon, 2) %>" size=20 readonly>��
		<!--
		&nbsp;
		<input type="button" class="button" value=" EMS���Ӱ������� " onClick="RecalcEMSPrice(frmMaster, frmDetail)">
		-->
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�հ�(��ȯ)</td>
	<td>
		<input type="text" class="text_ro" name="totalPriceForeign" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalPriceForeign, 2) %>" size=10 readonly>
	</td>
</tr>

<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>�κ��̽� ������</strong></font>
			        &nbsp;
			        <input type="button" class="button" value="4. �������� ������ �����ϱ�" onClick="PopPrevInvoiceList(frmMaster, 'INVOICE')">
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">1. Shipper/Exporter(�������)</td>
	<td colspan="3"><textarea class="textarea" name="exporteraddr" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Fexporteraddr %></textarea>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">2. For account & Risk of Messers.<br>(���Ծ���)</td>
	<td colspan="3"><textarea class="textarea" name="riskmesseraddr" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Friskmesseraddr %></textarea>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">3. Notify party(��������ó)</td>
	<td colspan="3"><textarea class="textarea" name="notifyaddr" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Fnotifyaddr %></textarea>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >4. Port of loading(������)</td>
	<td>
		<input type="text" class="text" name="portname" value="<%= ocoffinvoice.FOneItem.Fportname %>" size=20>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >5. Final destination(��������)</td>
	<td>
		<input type="text" class="text" name="destinationname" value="<%= ocoffinvoice.FOneItem.Fdestinationname %>" size=20>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >6. Carrier(�����̸�)</td>
	<td>
		<input type="text" class="text" name="carriername" value="<%= ocoffinvoice.FOneItem.Fcarriername %>" size=20>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >7. Sailing on or about(������)</td>
	<td>
		<input type="text" class="text" name="carrierdate" value="<%= ocoffinvoice.FOneItem.Fcarrierdate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.carrierdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >8.No.& date of invoice<br>(�κ��̽�DATE)</td>
	<td>
		<input type="text" class="text" name="invoicedate" value="<%= ocoffinvoice.FOneItem.Finvoicedate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.invoicedate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�κ��̽�NO</td>
	<td>
		<input type="text" class="text_ro" name="invoiceno" value="<%= ocoffinvoice.FOneItem.Finvoiceno %>" size=30 readonly>
		&nbsp;
		<input type="button" class="button" value=" 5. �ڵ��Է� " onClick="RecalcInvoiceNo(frmMaster)">
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >9. No.& date of L/C(�ſ���)</td>
	<td>
		<input type="text" class="text" name="lccomment" value="<%= ocoffinvoice.FOneItem.Flccomment %>" size=20>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >10. L/C issuing bank(�߱�����)</td>
	<td>
		<input type="text" class="text" name="lcbank" value="<%= ocoffinvoice.FOneItem.Flcbank %>" size=10>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">11. Remarks(���)</td>
	<td colspan="3"><textarea class="textarea" name="comment" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Fcomment %></textarea>
	</td>
</tr>

<% if (idx <> 0) then %>

<tr bgcolor="#FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td colspan=6>
				<input type="button" class="button" value="�⺻ ��ǰ���� ���" onClick="InsertDefaultDescription(frmMaster)">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" width="40" align="center">ǥ��<br>����</td>
			<td bgcolor="<%= adminColor("tabletop") %>">
				12. Description of Goods<br>(��ǰ����)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100">
				13. Q'ty / BOX<br>(ī��ڽ�����)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100">
				14. Price / BOX<br>(��ջ�ǰ�ݾ�)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100">
				15. Amount<br>(�ѻ�ǰ�ݾ�)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">���</td>
		</tr>
		<input type="hidden" name="productdetailcount" value="<%= ocoffinvoiceproductdetail.FResultCount %>">
		<% for i=0 to ocoffinvoiceproductdetail.FResultCount-1 %>
			<%
			if (ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox <> "") then
				ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox = FormatNumber(ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox, 2)
			end if
			if (ocoffinvoiceproductdetail.FItemList(i).Ftotalprice <> "") then
				ocoffinvoiceproductdetail.FItemList(i).Ftotalprice = FormatNumber(ocoffinvoiceproductdetail.FItemList(i).Ftotalprice, 2)
			end if
			%>
		<input type="hidden" name="productdetailidx_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Fidx %>">
		<tr bgcolor="#FFFFFF">
			<td align="center"><input type="text" class="text" name="orderno_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Forderno %>" size=2></td>
			<td><input type="text" class="text" name="goodscomment_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Fgoodscomment %>" size=60></td>
			<td><input type="text" class="text" name="totalboxno_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Ftotalboxno %>" size=10></td>
			<td>
				<input type="text" class="text_ro" name="priceperbox_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Fpriceperbox %>" size=10 readonly>
			</td>
			<td><input type="text" class="text" name="totalprice_<%= i %>" value="<%= ocoffinvoiceproductdetail.FItemList(i).Ftotalprice %>" size=10></td>
			<td align="center">
				<input type="button" class="button" value=" ���� " onClick="SaveMaster(frmMaster, 'modifymaster', 'modifydetailall', 0)">
				<input type="button" class="button" value=" ���� " onClick="SaveMaster(frmMaster, 'modifymaster', 'deletedetailone', '<%= ocoffinvoiceproductdetail.FItemList(i).Fidx %>')">
				<!--
				<input type="button" class="button" value=" ���� " onClick="SaveProductComment(frmMaster, 'deleteproduct', '<%= ocoffinvoiceproductdetail.FItemList(i).Fidx %>')">
				-->
			</td>
		</tr>
		<% next %>
		<tr bgcolor="#FFFFFF">
			<td align="center"><input type="text" class="text" name="orderno_new" value="<%= ocoffinvoiceproductdetail.FResultCount + 1 %>" size=2></td>
			<td><input type="text" class="text" name="goodscomment_new" value="" size=60></td>
			<td><input type="text" class="text" name="totalboxno_new" value="" size=10></td>
			<td>
				<input type="text" class="text_ro" name="priceperbox_new" value="" size=10 readonly>
			</td>
			<td><input type="text" class="text" name="totalprice_new" value="" size=10></td>
			<td align="center">
				<input type="button" class="button" value=" �߰� " onClick="SaveMaster(frmMaster, 'modifymaster', 'adddetailone', '')">
				<!--
				<input type="button" class="button" value=" �߰� " onClick="SaveProductComment(frmMaster, 'addproduct', '')">
				-->
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td colspan=6 height=30></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align="center"></td>
			<td>Total</td>
			<td>
				<input type="text" class="text_ro" name="totalboxno" value="<%= ocoffinvoice.FOneItem.Ftotalboxno %>" size=10 readonly>
			</td>
			<td></td>
			<td>
				<input type="text" class="text_ro" name="tmpTotalPriceForeign" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalPriceForeign, 2) %>" size=10 readonly>
			</td>
			<td align="left">
				<input type="text" class="text_ro" name="tmpTotalPriceWon" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalPriceWon, 2) %>" size=10 readonly> ��
			</td>
		</tr>
		</table>
	</td>
</tr>

<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>����Ű�����</strong></font>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" rowspan="3">�Ű��ȣ(5��)</td>
	<td>
		1.<input type="text" class="text" name="reportno" value="<%= ocoffinvoice.FOneItem.Freportno %>" size=25>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" rowspan="3">����Ű�����</td>
	<Td>
		1.<% if (ocoffinvoice.FOneItem.Fexportdeclarefilename <> "") then %>
		<input type="button" class="button" value=" �ٿ�ޱ�(<%= ocoffinvoice.FOneItem.Frealfilename %>) " onClick="PopDownloadExportDeclareFile(frmMaster,1)">
		<% end if %>
		<input type="button" class="button" value=" ��� " onClick="PopUploadExportDeclareFile(frmMaster,1)">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		2.<input type="text" class="text" name="reportno2" value="<%= ocoffinvoice.FOneItem.Freportno2 %>" size=25>
	</td>
	<td>
		2.<% if (ocoffinvoice.FOneItem.Fexportdeclarefilename2 <> "") then %>
		<input type="button" class="button" value=" �ٿ�ޱ�(<%= ocoffinvoice.FOneItem.Frealfilename2 %>) " onClick="PopDownloadExportDeclareFile(frmMaster,2)">
		<% end if %>
		<input type="button" class="button" value=" ��� " onClick="PopUploadExportDeclareFile(frmMaster,2)">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		3.<input type="text" class="text" name="reportno3" value="<%= ocoffinvoice.FOneItem.Freportno3 %>" size=25>
	</td>
	<td>
		3.<% if (ocoffinvoice.FOneItem.Fexportdeclarefilename3 <> "") then %>
		<input type="button" class="button" value=" �ٿ�ޱ�(<%= ocoffinvoice.FOneItem.Frealfilename3 %>) " onClick="PopDownloadExportDeclareFile(frmMaster,3)">
		<% end if %>
		<input type="button" class="button" value=" ��� " onClick="PopUploadExportDeclareFile(frmMaster,3)">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td bgcolor="<%= adminColor("tabletop") %>" >�Ű�����(6��)</td>
	<td colspan="3">
		<input type="text" class="text_ro" name="reportdate" value="<%= ocoffinvoice.FOneItem.Freportdate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.reportdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td bgcolor="<%= adminColor("tabletop") %>" >�Ű� ��ȭ�ڵ�</td>
	<td>
		<% drawSelectBoxPriceUnit "reportpriceunit", ocoffinvoice.FOneItem.Freportpriceunit %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�Ű� ȯ��</td>
	<td>
		<input type="text" class="text_ro" name="reportexchangerate" value="<%= ocoffinvoice.FOneItem.Freportexchangerate %>" size=10 readonly>
		<input type="button" class="button" value=" �ڵ��Է� " onClick="CalcReportExchangeRate(frmMaster)">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td bgcolor="<%= adminColor("tabletop") %>" >�Ű� �ݾ�(��ȭ)(48��)</td>
	<td>
		<input type="text" class="text" name="reportforeigntotalprice" value="<%= ocoffinvoice.FOneItem.Freportforeigntotalprice %>" size=20 onChange="frmMaster.reportexchangerate.value = 1;">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >�Ű� �ݾ�(��ȭ)(48��)</td>
	<td>
		<input type="text" class="text" name="reporttotalprice" value="<%= ocoffinvoice.FOneItem.Freporttotalprice %>" size=20 onChange="frmMaster.reportexchangerate.value = 1;">
	</td>
</tr>

<% else %>

<input type="hidden" name="totalboxno" value="0">
<input type="hidden" name="productdetailcount" value="0">

<input type="hidden" name="tmpTotalPriceWon" value="0">
<input type="hidden" name="tmpTotalPriceForeign" value="0">

<% end if%>

<tr height=40 bgcolor="#FFFFFF">
	<td colspan="4" align="center">
		<input type="button" class="button" value=" �����ϱ� " onClick="SaveMaster(frmMaster)">

		<% if (idx <> 0) then %>
			<% if (ocoffinvoice.FOneItem.Fstatecd = "7") then %>
				<input type="button" class="button" value=" �ۼ�����ȯ " onClick="SetWriteState(frmMaster)">
			<% else %>
				<input type="button" class="button" value=" �ۼ��Ϸ� " onClick="SetFinishState(frmMaster)">
			<% end if %>
			<input type="button" class="button" value=" ��ü���� " onClick="DelMaster(frmMaster)">
	        &nbsp;
	        <input type="button" class="button" value="QS" onClick="PopOpenQS('<%=idx%>','<%=ocoffinvoice.FOneItem.Fjungsanidx%>','<%=ocoffinvoice.FOneItem.Fworkidx%>','<%=ocoffinvoice.FOneItem.Floginsite%>','<%=ocoffinvoice.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
			<input type="button" class="button" value="PI" onClick="PopOpenPI('<%=idx%>','<%=ocoffinvoice.FOneItem.Fjungsanidx%>','<%=ocoffinvoice.FOneItem.Fworkidx%>','<%=ocoffinvoice.FOneItem.Floginsite%>','<%=ocoffinvoice.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
			<input type="button" class="button" value="CI" onClick="PopOpenCI('<%=idx%>','<%=ocoffinvoice.FOneItem.Fjungsanidx%>','<%=ocoffinvoice.FOneItem.Fworkidx%>','<%=ocoffinvoice.FOneItem.Floginsite%>','<%=ocoffinvoice.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
			<input type="button" class="button" value="PL" onClick="PopOpenPL('<%=idx%>','<%=ocoffinvoice.FOneItem.Fjungsanidx%>','<%=ocoffinvoice.FOneItem.Fworkidx%>','<%=ocoffinvoice.FOneItem.Floginsite%>','<%=ocartoonboxmaster.FOneItem.Fidx%>','<%=ocoffinvoice.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
			<input type="button" class="button" value="PL_Item" onClick="PopOpenPLItem('<%=idx%>','<%=ocoffinvoice.FOneItem.Floginsite%>','<%=ocartoonboxmaster.FOneItem.Fidx%>','<%=ocoffinvoice.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
	 		&nbsp;
	 		<input type="button" class="button" value=" �κ��̽�" onClick="PrintExportSheet(<%= idx %>, 'INVOICE','')">
	        <input type="button" class="button" value=" �κ��̽� ����" onClick="PrintExportSheet(<%= idx %>, 'INVOICE','Y')">
	        &nbsp;
	        <input type="button" class="button" value=" ��ŷ����Ʈ" onClick="PrintExportSheet(<%= idx %>, 'PACKINGLIST','')">
	        <input type="button" class="button" value=" ��ŷ����Ʈ ����" onClick="PrintExportSheet(<%= idx %>, 'PACKINGLIST','Y')">
	        &nbsp;
	        <input type="button" class="button" value=" ����Ű�����" onClick="PrintExportSheet(<%= idx %>, 'LICENCE','')">
	        <input type="button" class="button" value=" ����Ű����� ����" onClick="PrintExportSheet(<%= idx %>, 'LICENCE','Y')">
        <% end if %>
	</td>
</tr>

</form>
</table>

<p>

<% if (idx <> 0) then %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="12" align="right">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("tablebg") %>">
			<tr bgcolor="#FFFFFF">
				<td>
					<input type="button" class="button" value="�۾��������� ��������" onClick="InsertDetailFromWork(frmMaster)"> * �۾������� ������ ���, <font color="red">�������⸦ �ٽ� ����</font>�ؾ� �մϴ�.
				</td>
				<td align="right">
					�ѰǼ�:  <%= ocoffinvoicedetail.FResultCount %>
				</td>
			</tr>
		</table>
	</td>
</tr>

<!-- ��ܹ� ���� -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>��ŷ����Ʈ ����</strong></font>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- ��ܹ� �� -->

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>10. Marks & number of pkgs<br>(�ڽ���ȣ)</td>
    <td>11. Description of Goods<br>(�ڽ�����)</td>
    <td>12. Quantity/unit</td>
	<td>13. N weight<br>(Inner�ڽ�����)</td>
	<td>14. G weight<br>(Carton�ڽ�����)</td>
	<td>���</td>
</tr>
<% for i=0 to ocoffinvoicedetail.FResultCount-1 %>
	<% emsPrice = emsPrice + ocoffinvoicedetail.FItemList(i).FemsPrice %>
<tr align="center" bgcolor="#FFFFFF">
	<td><%= ocoffinvoicedetail.FItemList(i).Fcartonboxno %> BOX</td>
	<td><%= ocoffinvoicedetail.FItemList(i).Fgoodscomment %></td>
	<td></td>
	<td><%= FormatNumber(ocoffinvoicedetail.FItemList(i).Fnweight, 2) %> Kgs</td>
	<td><%= FormatNumber(ocoffinvoicedetail.FItemList(i).Fgweight, 2) %> Kgs</td>
	<td></td>
</tr>
<% next %>
</table>

<% end if %>

<form name="frmDetail">
<input type="hidden" name="emsPrice" value="<%= emsPrice %>">
</form>

<script>

// ������ ���۽� �۵��ϴ� ��ũ��Ʈ
function getOnload(){
	var frm = document.frmMaster;
	var exchangerate = frm.exchangerate.value.replace(/,/g, "");

	// RecalcPrice(frm);
}

window.onload = getOnload;

</script>

<%
set ocoffinvoice = Nothing
set ocoffinvoicedetail = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
