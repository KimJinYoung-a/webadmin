<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 주문서 작성
' History : 2009.04.07 서동석 생성
'			2010.08.12 한용민 수정
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
	'// 해외출고내역
	ofranchulgomaster.getOneFranJungsanBlank
	ofranchulgomaster.FOneItem.FcurrencyUnit = ocartoonboxmaster.FOneItem.FcurrencyUnit
	ofranchulgomaster.FOneItem.FtotalforeignSuplycash = ocartoonboxmaster.FOneItem.Ftotforeign_suplycash
	ofranchulgomaster.FOneItem.Ftotalsum = ocartoonboxmaster.FOneItem.Ftotsuplycash
else
	'// 정산내역
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
		'// 주문코드가 temp 이면 EMS배송비로 가정한다.
		'// 실제 기타내역이 입력되면 차이가 발생하고, 자동입력기능을 사용하지 못한다.(수기입력)
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

	var ret = confirm('저장 하시겠습니까?');

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
	// 작성완료 전환
	if (confirm("작성완료 상태로 전환하시겠습니까?\n\n인보이스 내역이 매장에 오픈됩니다.") != true) {
		return;
	}

	frm.mode.value="modifystate";
	frm.statecd.value = 7;
	frm.submit();
}

function SetWriteState(frm) {
	// 작성완료 전환
	if (confirm("작성중 상태로 전환하시겠습니까?") != true) {
		return;
	}

	frm.mode.value="modifystate";
	frm.statecd.value = 1;
	frm.submit();
}

function SaveProductComment(frm, submode, productidx) {
	alert("사용중지");
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

	var ret = confirm('저장 하시겠습니까?');

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
	var ret = confirm('기존 박스정보는 모두 삭제됩니다.\n\n진행 하시겠습니까?');

	if (ret == true) {
		frm.mode.value="insertdetailfromwork";

		frm.submit();
	}
}

function InsertDefaultDescription(frm) {
	var ret = confirm('기존 상품설명은 모두 삭제됩니다.\n\n진행 하시겠습니까?');

	if (ret == true) {
		frm.mode.value="insertdefaultdescription";

		frm.submit();
	}
}

function DelMaster(frm) {
	var ret = confirm('전체삭제 하시겠습니까?');

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
		alert("샵을 지정하세요.");
		frm.shopid.focus();
		return false;
	}

	if (frm.workidx.value == "") {
		alert("출고작업을 먼저 지정하세요.");
		return false;
	}

	//if (frm.jungsanidx.value == "") {
	//	alert("정산내역을 먼저 지정하세요.");
	//	return false;
	//}

	if (frm.workidx.value != "") {
		if (frm.workidx.value*0 != 0) {
			alert("숫자만 입력 가능합니다.");
			frm.workidx.focus();
			return false;
		}
	}

	if (totalGoodsPriceWon != "") {
		if (totalGoodsPriceWon*0 != 0) {
			alert("숫자만 입력 가능합니다.");
			frm.totalGoodsPriceWon.focus();
			return false;
		}
	}

	if (totalDeliverPriceWon != "") {
		if (totalDeliverPriceWon*0 != 0) {
			alert("숫자만 입력 가능합니다.");
			frm.totalDeliverPriceWon.focus();
			return false;
		}
	}

	if (exchangerate != "") {
		if (exchangerate*0 != 0) {
			alert("숫자만 입력 가능합니다.");
			frm.exchangerate.focus();
			return false;
		}
	}

	if (frm.delivermethod.value == "E") {
		// EMS 배송의 경우만 배송비 정산이 있다.
		if (frm.totalDeliverPriceWon.value*1 == 0) {
			//alert("EMS배송이면서 운임이 없습니다.");
			//frm.totalDeliverPriceWon.focus();
			//return false;
		}

		if (frm.exportmethod.value != "C") {
			alert("EMS배송입니다. 운임포함인도(운임정산)로 설정하세요.");
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
					alert("신고환율의 자동입력 버튼을 누르세요.");
					return false;
				}

				if ((reportforeigntotalprice*1 != reporttotalprice*1) && (reportexchangerate*1 == 1)) {
					alert("신고환율의 자동입력 버튼을 누르세요.");
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
			alert("표시순서를 지정하세요.");
			orderno.focus();
			return false;
		}

		if (orderno.value*0 != 0) {
			alert("표시순서는 숫자만 입력 가능합니다.");
			orderno.focus();
			return false;
		}

		if (totalboxno.value != "") {
			if (totalboxno.value*0 != 0) {
				alert("숫자만 입력 가능합니다.");
				totalboxno.focus();
				return false;
			}

			if (totalboxno.value*1 == 0) {
				alert("0 은 입력할 수 없습니다.");
				totalboxno.focus();
				return false;
			}
		}

		if (totalprice.value != "") {
			if (totalprice.value*0 != 0) {
				alert("숫자만 입력 가능합니다.");
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
		alert("표시순서를 지정하세요.");
		frm.orderno_new.focus();
		return false;
	}

	if (frm.orderno_new.value*0 != 0) {
		alert("표시순서는 숫자만 입력 가능합니다.");
		frm.orderno_new.focus();
		return false;
	}

	if (frm.totalboxno_new.value != "") {
		if (frm.totalboxno_new.value*0 != 0) {
			alert("숫자만 입력 가능합니다.");
			frm.totalboxno_new.focus();
			return false;
		}

		if (frm.totalboxno_new.value*1 == 0) {
			alert("0 은 입력할 수 없습니다.");
			frm.totalboxno_new.focus();
			return false;
		}
	}

	if (frm.totalprice_new.value != "") {
		if (frm.totalprice_new.value*0 != 0) {
			alert("숫자만 입력 가능합니다.");
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
		alert("샵을 지정하세요.");
		frm.shopid.focus();
		return;
	}

	if (frm.invoicedate.value == "") {
		alert("인보이스 일자를 지정하세요.");
		frm.invoicedate.focus();
		return;
	}

	if ((frm.workidx.value == "") || (frm.workidx.value*1 == 0)) {
		alert("작업코드를 지정하세요.");
		frm.workidx.focus();
		return;
	}

	//if ((frm.jungsanidx.value == "") || (frm.jungsanidx.value*1 == 0)) {
	//	alert("정산코드가 설정되지 않았습니다.\n\n정산내역에서 작업코드를 지정하세요.");
	//	return;
	//}

	frm.invoiceno.value = Right(frm.shopid.value, 3) + "_" + frm.workidx.value + "_" + frm.invoicedate.value.replace(/-/g, "");
}

function CalcReportExchangeRate(frm) {
	var pointprice = 2;		// 소수점이하 두자리까지 계산

	var reportforeigntotalprice = frm.reportforeigntotalprice.value.replace(/,/g, ""); // 외환
	var reporttotalprice = frm.reporttotalprice.value.replace(/,/g, ""); // 원화
	var reportpriceunit = frm.reportpriceunit.value;
	var reportexchangerate = 0;

	if ((reportpriceunit == "JPY") || (reportpriceunit == "WON")) {
		pointprice = 0;
	}

	if (reportpriceunit == "") {
		alert("신고 통화코드를 입력하세요.");
		frm.reportpriceunit.focus();
		return;
	}

	if ((reportforeigntotalprice == "") || (reportforeigntotalprice*1 == 0)) {
		alert("신고 금액(외화)을 입력하세요.");
		frm.reportforeigntotalprice.focus();
		return;
	}

	if ((reporttotalprice == "") || (reporttotalprice*1 == 0)) {
		alert("신고 금액(원화)을 입력하세요.");
		frm.reporttotalprice.focus();
		return;
	}

	reportforeigntotalprice = (reportforeigntotalprice * 1.0).toFixed(pointprice); // 외환
	reporttotalprice = (reporttotalprice * 1.0).toFixed(0); // 원화

	if (reportpriceunit == "JPY") {
		// 엔화는 100엔 단위로 환율을 계산한다.
		reportexchangerate = (100.0 * reporttotalprice / reportforeigntotalprice);
	} else {
		reportexchangerate = (1.0 * reporttotalprice / reportforeigntotalprice);
	}

	// 환율은 소수점 두째자리까지
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
	var exchangerate 		= frm.exchangerate.value.replace(/,/g, "");			// 환율

	var totalDeliverPriceWon	= frm.totalDeliverPriceWon.value.replace(/,/g, "");		// 운임(원)
	var totalGoodsPriceWon 		= frm.totalGoodsPriceWon.value.replace(/,/g, "");		// 상품금액(원)
	var totalPriceWon			= 0;													// 합계(원)

	var priceunit 			= frm.priceunit.value;										// 작성화폐

	var totalGoodsPriceForeign 		= 0;												// 상품금액(외환)
	var totalDeliverPriceForeign 	= 0;												// 운임(외환)
	var totalPriceForeign			= 0;												// 합계(외환)

	// ========================================================================
	if ((exchangerate == "") || (exchangerate*0 != 0) || (exchangerate*1 == 0)) {
		return "환율을 입력하세요.";
	}

	if ((totalGoodsPriceWon == "") || (totalGoodsPriceWon*0 != 0) || (totalGoodsPriceWon*1 == 0)) {
		return "상품금액을 입력하세요";
	}

	if ((totalDeliverPriceWon == "") || (totalDeliverPriceWon*0 != 0)) {
		// 운임은 없을 수 있다.
		totalDeliverPriceWon = 0;
	}

	if (priceunit == "") {
		return "작성화폐를 지정하세요.";
	}


	// ========================================================================
	exchangerate = exchangerate*1;
	totalDeliverPriceWon = totalDeliverPriceWon*1;
	totalGoodsPriceWon = totalGoodsPriceWon*1;
	totalPriceWon = totalDeliverPriceWon + totalGoodsPriceWon;

	var pointprice = 2;		// 소수점이하 두자리까지 계산

	if (priceunit == "JPY") {
		// 엔화는 100엔 단위로 환율을 계산한다.
		exchangerate = exchangerate*1 / 100;
		pointprice = 0;
	}

	<% if (ofranchulgomaster.FOneItem.FcurrencyUnit = "") or (ofranchulgomaster.FOneItem.FtotalforeignSuplycash = 0) then %>
		// 주문내역 외환금액이 없는 경우만 환율로 계산
		totalGoodsPriceForeign = (totalGoodsPriceWon / exchangerate).toFixed(pointprice);
		totalDeliverPriceForeign = (totalDeliverPriceWon / exchangerate).toFixed(pointprice);
		totalPriceForeign = ((totalGoodsPriceWon + totalDeliverPriceWon) / exchangerate).toFixed(pointprice);
	<% else %>
		// 주문내역 외환금액이 있으면 외환으로 구한다.
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
		alert("EMS운송의 경우에만 EMS 금액을 가져올 수 있습니다.");
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
	var totaljungsandeliverpay 	= "<%= totaljungsandeliverpay %>";	// 정산내역상의 EMS배송비(배송비 이외의 기타내역이 있으면 차이가 발생한다.)

	var totalforeignSuplycash	= "<%= ofranchulgomaster.FOneItem.FtotalforeignSuplycash %>";
	var currencyUnit 			= "<%= ofranchulgomaster.FOneItem.FcurrencyUnit %>";

	if (workidx == "") {
		alert("출고작업이 지정되어 있지 않습니다.");
		return;
	}

	//if (jungsanidx == "") {
	//	alert("정산내역이 지정되어 있지 않습니다.");
	//	return;
	//}

	if (delivermethod == "") {
		alert("출고작업에 운송방법이 지정되어 있지 않습니다.");
		return;
	}

	if (delivermethod == "E") {
		// EMS 일 경우
		if ((deliverpay == "") || (deliverpay*0 != 0)) {
			alert("출고작업에 EMS비용이 잘못 입력되어 있습니다.\n - 해외출고EMS비용 : (" + deliverpay + ")");
			return;
		}

		if (deliverpay*1 == 0) {
			deliverpay = 0;
			//alert("출고작업에 EMS비용이 잘못 입력되어 있습니다.\n - 해외출고EMS비용 : (" + deliverpay + ")");
			//return;
		}
	} else {
		deliverpay = 0;
	}

	if (deliverpay*1 != totaljungsandeliverpay*1) {
		alert("출고작업과 정산내역상의 EMS배송비가 서로 틀립니다.\n\n내역을 확인하고 수기로 입력하세요.\n - 해외출고EMS비용 : (" + deliverpay + ")\n - 정산내역EMS비용 : (" + totaljungsandeliverpay + ")");
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
		alert("잘못된 접속입니다.");
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

// qs 팝업
function PopOpenQS(invoiceidx, jungsanidx, workidx, loginsite, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/quotationsheet.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite+ '&cunit='+cunit+'&tpl='+tpl ,'PopOpenQSList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// pi 팝업
function PopOpenPI(invoiceidx, jungsanidx, workidx, loginsite, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/proformainvoice.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite+ '&cunit='+cunit+'&xl=Y&tpl='+tpl,'PopOpenInvoice','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// ci 팝업
function PopOpenCI(invoiceidx, jungsanidx, workidx, loginsite, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/commercialinvoice.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite + '&cunit='+cunit+'&tpl='+tpl,'PopOpenCIList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// pl 팝업
function PopOpenPL(invoiceidx, jungsanidx, workidx, loginsite,boxidx, cunit, tpl) {
	var popwin;
	popwin = window.open('/admin/fran/packlinglist.asp?idx=' + invoiceidx+'&jungsanidx=' + jungsanidx+'&workidx=' + workidx+'&ls='+ loginsite+'&boxidx='+boxidx+ '&cunit='+cunit+'&tpl='+tpl ,'PopOpenPLList','width=1280,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

// pl 팝업 상품
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

<!-- 상단바 시작 -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>인보이스 기본정보</strong></font>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- 상단바 끝 -->

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">IDX</td>
	<td><%= ocoffinvoice.FOneItem.Fidx %></td>
	<td bgcolor="<%= adminColor("tabletop") %>" >등록자</td>
	<td>
		<%= ocoffinvoice.FOneItem.Freguserid %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width="200">샵아이디</td>
	<% if ocoffinvoice.FOneItem.Fshopid<>"" then %>
	<input type=hidden name="shopid" value="<%= ocoffinvoice.FOneItem.Fshopid %>">
	<td><%= ocoffinvoice.FOneItem.Fshopid %></td>
	<td bgcolor="<%= adminColor("tabletop") %>">샵명</td>
	<td><%= ocoffinvoice.FOneItem.Fshopname %></td>
	<% else %>
	<td colspan=3><% drawSelectBoxOffShopNot000 "shopid", ocoffinvoice.FOneItem.Fshopid %></td>
	<% end if %>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width="200">상태</td>
	<td><%= ocoffinvoice.FOneItem.GetStateCDName %></td>
	<td bgcolor="<%= adminColor("tabletop") %>"></td>
	<td></td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >정산IDX</td>
	<td>
		<input type="text" class="text_ro" name="jungsanidx" value="<%= ocoffinvoice.FOneItem.Fjungsanidx %>" size="6" maxlength="6" style="text-align:right" readonly>
		<% if (ocoffinvoice.FOneItem.Fjungsanidx <> "") then %>
		&nbsp;
		<input type="button" class="button" value="조회하기" onClick="popJungsanMaster(<%= ocoffinvoice.FOneItem.Fjungsanidx %>)">
		<% else %>
		* 에러 : 관리자 문의 요망
		<% end if %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >작업IDX</td>
	<td>
		<input type="text" class="text_ro" name="workidx" value="<%= ocoffinvoice.FOneItem.Fworkidx %>" size="6" maxlength="6" style="text-align:right" readonly>
		<% if (ocoffinvoice.FOneItem.Fworkidx <> "") then %>
		&nbsp;
		<input type="button" class="button" value="조회하기" onClick="PopExportSheet(<%= ocoffinvoice.FOneItem.Fworkidx %>)">
		<% else %>
		* 작업은 정산내역에서 지정가능
		<% end if %>
	</td>
</tr>

<!-- 상단바 시작 -->
<tr height="30" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>인보이스 정산정보</strong></font>
			        &nbsp;
			        <input type="button" class="button" value="1. 기존내역 정산정보 복사하기" onClick="PopPrevInvoiceList(frmMaster, 'JUNGSAN')">
			        <input type="button" class="button" value="2. 운송정보/상품금액 가져오기" onClick="CopyDeliverAndPriceInfo(frmMaster)">
			        * 운송정보/상품금액 은 출고작업에서 가져옵니다.
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- 상단바 끝 -->

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >운송방법</td>
	<td>
		<select class="select" name="delivermethod">
			<option value="">선택</option>
			<option value="E" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "E") then %>selected<% end if %>>EMS</option>
			<option value="D" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "D") then %>selected<% end if %>>DHL</option>
			<option value="F" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "F") then %>selected<% end if %>>항공</option>
			<option value="S" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "S") then %>selected<% end if %>>해운</option>
			<option value="P" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "P") then %>selected<% end if %>>국제소포(선편)</option>
			<option value="T" <% if (ocoffinvoice.FOneItem.Fdelivermethod = "T") then %>selected<% end if %>>국내택배</option>
		</select>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >운임부담</td>
	<td>
		<select class="select" name="exportmethod">
			<option value="">선택</option>
			<option value="F" <% if (ocoffinvoice.FOneItem.Fexportmethod = "F") then %>selected<% end if %>>FOB</option>
			<option value="C" <% if (ocoffinvoice.FOneItem.Fexportmethod = "C") then %>selected<% end if %>>C&F</option>
			<option value="W" <% if (ocoffinvoice.FOneItem.Fexportmethod = "W") then %>selected<% end if %>>EXW</option>
			<option value="A" <% if (ocoffinvoice.FOneItem.Fexportmethod = "A") then %>selected<% end if %>>FOB</option>
		</select>
		&nbsp;
		CFR : 운임포함인도(운임정산) / FOB : 본선인도 / EXW : 공장인도 / FCA : 운송인인도
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >정산방법</td>
	<td>
		<select class="select" name="jungsantype">
			<option value="">선택</option>
			<option value="T" <% if (ocoffinvoice.FOneItem.Fjungsantype = "T") then %>selected<% end if %>>TT</option>
			<option value="L" <% if (ocoffinvoice.FOneItem.Fjungsantype = "L") then %>selected<% end if %>>LC</option>
		</select>
		&nbsp;
		* TT : 전신환송금(선정산) / LC : 신용장
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >작성화폐</td>
	<td>
		<% drawSelectBoxPriceUnit "priceunit", ocoffinvoice.FOneItem.Fpriceunit %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >수출환율</td>
	<td>
		<input type="text" class="text" name="exchangerate" value="<%= FormatNumber(ocoffinvoice.FOneItem.Fexchangerate, 2) %>" size=10>원
		&nbsp;
		<input type="button" class="button" value=" 3. 자동계산 " onClick="RecalcPriceAndAlert(frmMaster)">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" ></td>
	<td>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >상품금액(원)</td>
	<td>
		<input type="text" class="text" name="totalGoodsPriceWon" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalGoodsPriceWon, 2) %>" size=20>원 (정산내역)
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >상품금액(외환)</td>
	<td>
		<input type="text" class="text<% if (ofranchulgomaster.FOneItem.FtotalforeignSuplycash = 0) then %>_ro<% end if %>" name="totalGoodsPriceForeign" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalGoodsPriceForeign, 2) %>" size=10 <% if (ofranchulgomaster.FOneItem.FtotalforeignSuplycash = 0) then %>readonly<% end if %>>
		<% if (ofranchulgomaster.FOneItem.FtotalforeignSuplycash <> 0) then %>
			<font color="red">(주문내역 : <%= ofranchulgomaster.FOneItem.FcurrencyUnit %> &nbsp; <%= FormatNumber(ofranchulgomaster.FOneItem.FtotalforeignSuplycash, 2) %>)</font>
		<% end if %>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >운임(원)</td>
	<td>
		<input type="text" class="text" name="totalDeliverPriceWon" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalDeliverPriceWon, 2) %>" size=20>원
		<!--
		&nbsp;
		<input type="button" class="button" value=" EMS운임가져오기 " onClick="RecalcEMSPrice(frmMaster, frmDetail)">
		-->
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >운임(외환)</td>
	<td>
		<input type="text" class="text_ro" name="totalDeliverPriceForeign" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalDeliverPriceForeign, 2) %>" size=10 readonly>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >합계(원)</td>
	<td>
		<input type="text" class="text_ro" name="totalPriceWon" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalPriceWon, 2) %>" size=20 readonly>원
		<!--
		&nbsp;
		<input type="button" class="button" value=" EMS운임가져오기 " onClick="RecalcEMSPrice(frmMaster, frmDetail)">
		-->
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >합계(외환)</td>
	<td>
		<input type="text" class="text_ro" name="totalPriceForeign" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalPriceForeign, 2) %>" size=10 readonly>
	</td>
</tr>

<!-- 상단바 시작 -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>인보이스 상세정보</strong></font>
			        &nbsp;
			        <input type="button" class="button" value="4. 기존내역 상세정보 복사하기" onClick="PopPrevInvoiceList(frmMaster, 'INVOICE')">
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- 상단바 끝 -->

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">1. Shipper/Exporter(수출업자)</td>
	<td colspan="3"><textarea class="textarea" name="exporteraddr" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Fexporteraddr %></textarea>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">2. For account & Risk of Messers.<br>(수입업자)</td>
	<td colspan="3"><textarea class="textarea" name="riskmesseraddr" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Friskmesseraddr %></textarea>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">3. Notify party(도착통지처)</td>
	<td colspan="3"><textarea class="textarea" name="notifyaddr" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Fnotifyaddr %></textarea>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >4. Port of loading(선적항)</td>
	<td>
		<input type="text" class="text" name="portname" value="<%= ocoffinvoice.FOneItem.Fportname %>" size=20>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >5. Final destination(도착국가)</td>
	<td>
		<input type="text" class="text" name="destinationname" value="<%= ocoffinvoice.FOneItem.Fdestinationname %>" size=20>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >6. Carrier(선박이름)</td>
	<td>
		<input type="text" class="text" name="carriername" value="<%= ocoffinvoice.FOneItem.Fcarriername %>" size=20>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >7. Sailing on or about(선적일)</td>
	<td>
		<input type="text" class="text" name="carrierdate" value="<%= ocoffinvoice.FOneItem.Fcarrierdate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.carrierdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >8.No.& date of invoice<br>(인보이스DATE)</td>
	<td>
		<input type="text" class="text" name="invoicedate" value="<%= ocoffinvoice.FOneItem.Finvoicedate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.invoicedate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >인보이스NO</td>
	<td>
		<input type="text" class="text_ro" name="invoiceno" value="<%= ocoffinvoice.FOneItem.Finvoiceno %>" size=30 readonly>
		&nbsp;
		<input type="button" class="button" value=" 5. 자동입력 " onClick="RecalcInvoiceNo(frmMaster)">
	</td>
</tr>

<tr height="30" bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" >9. No.& date of L/C(신용장)</td>
	<td>
		<input type="text" class="text" name="lccomment" value="<%= ocoffinvoice.FOneItem.Flccomment %>" size=20>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >10. L/C issuing bank(발급은행)</td>
	<td>
		<input type="text" class="text" name="lcbank" value="<%= ocoffinvoice.FOneItem.Flcbank %>" size=10>
	</td>
</tr>

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">11. Remarks(비고)</td>
	<td colspan="3"><textarea class="textarea" name="comment" cols="80" rows="6"><%= ocoffinvoice.FOneItem.Fcomment %></textarea>
	</td>
</tr>

<% if (idx <> 0) then %>

<tr bgcolor="#FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF">
			<td colspan=6>
				<input type="button" class="button" value="기본 상품설명 등록" onClick="InsertDefaultDescription(frmMaster)">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td bgcolor="<%= adminColor("tabletop") %>" width="40" align="center">표시<br>순서</td>
			<td bgcolor="<%= adminColor("tabletop") %>">
				12. Description of Goods<br>(상품설명)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100">
				13. Q'ty / BOX<br>(카톤박스수량)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100">
				14. Price / BOX<br>(평균상품금액)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="100">
				15. Amount<br>(총상품금액)
			</td>
			<td bgcolor="<%= adminColor("tabletop") %>" width="120" align="center">비고</td>
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
				<input type="button" class="button" value=" 수정 " onClick="SaveMaster(frmMaster, 'modifymaster', 'modifydetailall', 0)">
				<input type="button" class="button" value=" 삭제 " onClick="SaveMaster(frmMaster, 'modifymaster', 'deletedetailone', '<%= ocoffinvoiceproductdetail.FItemList(i).Fidx %>')">
				<!--
				<input type="button" class="button" value=" 삭제 " onClick="SaveProductComment(frmMaster, 'deleteproduct', '<%= ocoffinvoiceproductdetail.FItemList(i).Fidx %>')">
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
				<input type="button" class="button" value=" 추가 " onClick="SaveMaster(frmMaster, 'modifymaster', 'adddetailone', '')">
				<!--
				<input type="button" class="button" value=" 추가 " onClick="SaveProductComment(frmMaster, 'addproduct', '')">
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
				<input type="text" class="text_ro" name="tmpTotalPriceWon" value="<%= FormatNumber(ocoffinvoice.FOneItem.FtotalPriceWon, 2) %>" size=10 readonly> 원
			</td>
		</tr>
		</table>
	</td>
</tr>

<!-- 상단바 시작 -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="4">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>수출신고필증</strong></font>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- 상단바 끝 -->

<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" rowspan="3">신고번호(5번)</td>
	<td>
		1.<input type="text" class="text" name="reportno" value="<%= ocoffinvoice.FOneItem.Freportno %>" size=25>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" rowspan="3">수출신고필증</td>
	<Td>
		1.<% if (ocoffinvoice.FOneItem.Fexportdeclarefilename <> "") then %>
		<input type="button" class="button" value=" 다운받기(<%= ocoffinvoice.FOneItem.Frealfilename %>) " onClick="PopDownloadExportDeclareFile(frmMaster,1)">
		<% end if %>
		<input type="button" class="button" value=" 등록 " onClick="PopUploadExportDeclareFile(frmMaster,1)">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		2.<input type="text" class="text" name="reportno2" value="<%= ocoffinvoice.FOneItem.Freportno2 %>" size=25>
	</td>
	<td>
		2.<% if (ocoffinvoice.FOneItem.Fexportdeclarefilename2 <> "") then %>
		<input type="button" class="button" value=" 다운받기(<%= ocoffinvoice.FOneItem.Frealfilename2 %>) " onClick="PopDownloadExportDeclareFile(frmMaster,2)">
		<% end if %>
		<input type="button" class="button" value=" 등록 " onClick="PopUploadExportDeclareFile(frmMaster,2)">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>
		3.<input type="text" class="text" name="reportno3" value="<%= ocoffinvoice.FOneItem.Freportno3 %>" size=25>
	</td>
	<td>
		3.<% if (ocoffinvoice.FOneItem.Fexportdeclarefilename3 <> "") then %>
		<input type="button" class="button" value=" 다운받기(<%= ocoffinvoice.FOneItem.Frealfilename3 %>) " onClick="PopDownloadExportDeclareFile(frmMaster,3)">
		<% end if %>
		<input type="button" class="button" value=" 등록 " onClick="PopUploadExportDeclareFile(frmMaster,3)">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td bgcolor="<%= adminColor("tabletop") %>" >신고일자(6번)</td>
	<td colspan="3">
		<input type="text" class="text_ro" name="reportdate" value="<%= ocoffinvoice.FOneItem.Freportdate %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.reportdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21>
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td bgcolor="<%= adminColor("tabletop") %>" >신고 통화코드</td>
	<td>
		<% drawSelectBoxPriceUnit "reportpriceunit", ocoffinvoice.FOneItem.Freportpriceunit %>
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >신고 환율</td>
	<td>
		<input type="text" class="text_ro" name="reportexchangerate" value="<%= ocoffinvoice.FOneItem.Freportexchangerate %>" size=10 readonly>
		<input type="button" class="button" value=" 자동입력 " onClick="CalcReportExchangeRate(frmMaster)">
	</td>
</tr>
<tr bgcolor="#FFFFFF" height="25">
	<td bgcolor="<%= adminColor("tabletop") %>" >신고 금액(외화)(48번)</td>
	<td>
		<input type="text" class="text" name="reportforeigntotalprice" value="<%= ocoffinvoice.FOneItem.Freportforeigntotalprice %>" size=20 onChange="frmMaster.reportexchangerate.value = 1;">
	</td>
	<td bgcolor="<%= adminColor("tabletop") %>" >신고 금액(원화)(48번)</td>
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
		<input type="button" class="button" value=" 저장하기 " onClick="SaveMaster(frmMaster)">

		<% if (idx <> 0) then %>
			<% if (ocoffinvoice.FOneItem.Fstatecd = "7") then %>
				<input type="button" class="button" value=" 작성중전환 " onClick="SetWriteState(frmMaster)">
			<% else %>
				<input type="button" class="button" value=" 작성완료 " onClick="SetFinishState(frmMaster)">
			<% end if %>
			<input type="button" class="button" value=" 전체삭제 " onClick="DelMaster(frmMaster)">
	        &nbsp;
	        <input type="button" class="button" value="QS" onClick="PopOpenQS('<%=idx%>','<%=ocoffinvoice.FOneItem.Fjungsanidx%>','<%=ocoffinvoice.FOneItem.Fworkidx%>','<%=ocoffinvoice.FOneItem.Floginsite%>','<%=ocoffinvoice.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
			<input type="button" class="button" value="PI" onClick="PopOpenPI('<%=idx%>','<%=ocoffinvoice.FOneItem.Fjungsanidx%>','<%=ocoffinvoice.FOneItem.Fworkidx%>','<%=ocoffinvoice.FOneItem.Floginsite%>','<%=ocoffinvoice.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
			<input type="button" class="button" value="CI" onClick="PopOpenCI('<%=idx%>','<%=ocoffinvoice.FOneItem.Fjungsanidx%>','<%=ocoffinvoice.FOneItem.Fworkidx%>','<%=ocoffinvoice.FOneItem.Floginsite%>','<%=ocoffinvoice.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
			<input type="button" class="button" value="PL" onClick="PopOpenPL('<%=idx%>','<%=ocoffinvoice.FOneItem.Fjungsanidx%>','<%=ocoffinvoice.FOneItem.Fworkidx%>','<%=ocoffinvoice.FOneItem.Floginsite%>','<%=ocartoonboxmaster.FOneItem.Fidx%>','<%=ocoffinvoice.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
			<input type="button" class="button" value="PL_Item" onClick="PopOpenPLItem('<%=idx%>','<%=ocoffinvoice.FOneItem.Floginsite%>','<%=ocartoonboxmaster.FOneItem.Fidx%>','<%=ocoffinvoice.FOneItem.Fcurrencyunit%>','<%=ocartoonboxmaster.FOneItem.Ftplcompanyid%>')">
	 		&nbsp;
	 		<input type="button" class="button" value=" 인보이스" onClick="PrintExportSheet(<%= idx %>, 'INVOICE','')">
	        <input type="button" class="button" value=" 인보이스 엑셀" onClick="PrintExportSheet(<%= idx %>, 'INVOICE','Y')">
	        &nbsp;
	        <input type="button" class="button" value=" 패킹리스트" onClick="PrintExportSheet(<%= idx %>, 'PACKINGLIST','')">
	        <input type="button" class="button" value=" 패킹리스트 엑셀" onClick="PrintExportSheet(<%= idx %>, 'PACKINGLIST','Y')">
	        &nbsp;
	        <input type="button" class="button" value=" 수출신고정보" onClick="PrintExportSheet(<%= idx %>, 'LICENCE','')">
	        <input type="button" class="button" value=" 수출신고정보 엑셀" onClick="PrintExportSheet(<%= idx %>, 'LICENCE','Y')">
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
					<input type="button" class="button" value="작업내역에서 가져오기" onClick="InsertDetailFromWork(frmMaster)"> * 작업내역을 수정한 경우, <font color="red">가져오기를 다시 실행</font>해야 합니다.
				</td>
				<td align="right">
					총건수:  <%= ocoffinvoicedetail.FResultCount %>
				</td>
			</tr>
		</table>
	</td>
</tr>

<!-- 상단바 시작 -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="7">
		<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					<img src="/images/icon_arrow_down.gif" align="absbottom">
			        <font color="red"><strong>패킹리스트 정보</strong></font>
			    </td>
			</tr>
		</table>
	</td>
</tr>
<!-- 상단바 끝 -->

<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>10. Marks & number of pkgs<br>(박스번호)</td>
    <td>11. Description of Goods<br>(박스설명)</td>
    <td>12. Quantity/unit</td>
	<td>13. N weight<br>(Inner박스무게)</td>
	<td>14. G weight<br>(Carton박스무게)</td>
	<td>비고</td>
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

// 페이지 시작시 작동하는 스크립트
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
