<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  브랜드별재고현황
' History : 2009.04.07 서동석 생성
'			2013.10.16 한용민 수정
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim makerid, onoffgubun, mwdiv, research, sellyn, usingyn, danjongyn, osummarystockbrand, centermwdiv
dim returnitemgubun, itemname, itemidArr, cdl, cdm, cds, page, i, BasicMonth, limitrealstock
dim totsysstock, totavailstock, totrealstock, totjeagosheetstock, totmaystock, IsSheetPrintEnable
dim stocktype, useoffinfo, itemgubun, startMon, endMon, excits, pagesize, ordby, vPurchasetype
dim limityn, itemgrade, itemrackcode, bulkstockgubun, warehouseCd, agvstockgubun, excNoRack
Dim dispCate : dispCate = RequestCheckvar(Request("disp"),12)
	makerid         = requestCheckvar(request("makerid"),32)
	onoffgubun      = requestCheckvar(request("onoffgubun"),9)
	research        = requestCheckvar(request("research"),9)
	sellyn          = requestCheckvar(request("sellyn"),9)
	usingyn         = requestCheckvar(request("usingyn"),9)
	danjongyn       = requestCheckvar(request("danjongyn"),9)
	mwdiv           = requestCheckvar(request("mwdiv"),9)
	returnitemgubun = requestCheckvar(request("returnitemgubun"),9)
	itemname        = requestCheckvar(request("itemname"),64)
	itemidArr       = Trim(requestCheckvar(request("itemidArr"),255))
	page            = requestCheckvar(request("page"),9)
	cdl = requestCheckvar(request("cdl"),3)
	cdm = requestCheckvar(request("cdm"),3)
	cds = requestCheckvar(request("cds"),3)
	limitrealstock 	= requestCheckvar(request("limitrealstock"),10)
    centermwdiv    	= requestCheckvar(request("centermwdiv"),10)
	stocktype    	= requestCheckvar(request("stocktype"),32)
	itemgubun     	= RequestCheckVar(request("itemgubun"),32)
	startMon     	= RequestCheckVar(request("startMon"),32)
	endMon     		= RequestCheckVar(request("endMon"),32)
	useoffinfo = request("useoffinfo")
	excits  		= requestCheckvar(request("excits"),2)
	pagesize  		= requestCheckvar(request("pagesize"),4)
	ordby    		= requestCheckvar(request("ordby"),64)
	vPurchasetype 	= request("purchasetype")
    limityn  		= requestCheckvar(request("limityn"),2)
    itemgrade     	= RequestCheckVar(request("itemgrade"),32)
    itemrackcode    = RequestCheckVar(request("itemrackcode"),32)
    bulkstockgubun  = RequestCheckVar(request("bulkstockgubun"),32)
    warehouseCd  	= RequestCheckVar(request("warehouseCd"),32)
    agvstockgubun  	= RequestCheckVar(request("agvstockgubun"),32)
    excNoRack  	= RequestCheckVar(request("excNoRack"),1)

if (stocktype = "") then stocktype = "sys"
if (pagesize = "") then pagesize = 25

'///////////////// 바코드 프린트기 설정 ///////////////////////
dim printername, printpriceyn, titledispyn, isforeignprint, makeriddispyn, useforeigndata, currencyunit, currencyChar
	printername = requestCheckVar(request("printername"),32)
	printpriceyn 	= requestCheckVar(request("printpriceyn"),1)
	titledispyn = requestCheckVar(request("titledispyn"),1)
	isforeignprint 	= requestCheckVar(request("isforeignprint"),32)
	makeriddispyn 			= requestCheckVar(request("makeriddispyn"),32)

if printpriceyn = "" then printpriceyn = "Y"	' R
if printername = "" then printername = "TEC_B-FV4_80x50"	' TEC_B-FV4_45x22
if makeriddispyn = "" then makeriddispyn = "Y"
if titledispyn = "" then titledispyn = "Y"
useforeigndata = "N"
currencyunit = "KRW"
currencyChar = "￦"
'/////////////////'////////////////////////////////////////

'//상품코드 유효성 검사
if itemidArr<>"" then
	dim iA ,arrTemp,arrItemid
  itemidArr = replace(itemidArr,chr(13),"")
	arrTemp = Split(itemidArr,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemidArr = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemidArr)) then
			itemidArr = ""
		end if
	end if
end if


if (request("research") = "") then
	excits = "Y"
end if

if (page="") then page=1
''if onoffgubun="" then onoffgubun="on"
''if itemgubun="" then itemgubun="10"
BasicMonth = Left(CStr(DateSerial(Year(now()),Month(now())-1,1)),7)


'// onoffgubun => itemgubun, skyer9, 2016-06-21
if (onoffgubun = "") and (itemgubun = "") then
	itemgubun="10"
elseif (onoffgubun <> "") and (itemgubun = "") then
	if (onoffgubun = "on") then
		itemgubun="10"
	elseif (onoffgubun = "off") then
		itemgubun="exc10"
	else
		itemgubun = Right(onoffgubun,2)
	end if
end if
if itemgubun="" then itemgubun="10"

if itemgubun = "10" then
	onoffgubun = "on"
elseif (itemgubun = "exc10") then
	onoffgubun = "off"
elseif (itemgubun <> "10") then
	onoffgubun = "off" & itemgubun
end if


set osummarystockbrand = new CSummaryItemStock
	osummarystockbrand.FPageSize = pagesize
	osummarystockbrand.FCurrPage = page
	osummarystockbrand.FRectMakerid = makerid
	osummarystockbrand.FRectMWDiv = mwdiv
	osummarystockbrand.FRectExcIts = excits
    osummarystockbrand.FRectExcNoRack = excNoRack

	osummarystockbrand.GetAgvStockDiffList

dim bulkrealstock

%>

<script type="text/javascript" src="/js/barcode.js"></script>
<script type="text/javascript" src="/js/ttpbarcode.js"></script>
<script type="text/javascript" src="/js/DOSHIBAbarcode.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
	frm.action="";
	frm.target="";
    document.frm.submit();
}

function GotoPage(page){
    var frm = document.frm;
    frm.page.value = page;
	frm.action="";
	frm.target="";
	frm.submit();
}

//온라인상품
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
}
//오프상품
function popOffItemEdit(ibarcode){
	var popwin = window.open('popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popRealErrInput(itemgubun,itemid,itemoption){
	var popwin = window.open('/common/poprealerrinput.asp?itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&BasicMonth=<%= BasicMonth %>','poprealerrinput','width=1024,height=768,scrollbar=yes,resizable=yes')
	popwin.focus();
}

function PopBrandStockSheet(){

	var onoffgubun = "on";
//    for (var i = 0; i < document.frm.onoffgubun.length; i++) {
//        if (document.frm.onoffgubun[i].checked == true) {
//                onoffgubun = document.frm.onoffgubun[i].value;
//        }
//    }

    onoffgubun = document.frm.onoffgubun.value;

    var returnitemgubun ="";
    for (var i = 0; i < document.frm.returnitemgubun.length; i++) {
        if (document.frm.returnitemgubun[i].checked == true) {
                returnitemgubun = document.frm.returnitemgubun[i].value;
        }
    }

    var makerid = document.frm.makerid.value;

	var mwdiv = document.frm.mwdiv.value;
	var centermwdiv = document.frm.centermwdiv.value;

    var sellyn = document.frm.sellyn.value;
    var isusing= document.frm.usingyn.value;
	var danjongyn  = document.frm.danjongyn.value;
	var disp     = document.frm.disp.value;
    //var cdl     = document.frm.cdl.value;
    //var cdm     = document.frm.cdm.value;
    //var cds     = document.frm.cds.value;
    var itemidArr     = document.frm.itemidArr.value.replace(/(?:\r\n|\r|\n)/g, ',');
    var itemname     = '';//document.frm.itemname.value;
    var limitrealstock     = document.frm.limitrealstock.value;
	var stocktype     = document.frm.stocktype.value;
    var itemrackcode = document.frm.itemrackcode.value;
    var warehouseCd = document.frm.warehouseCd.value;
    var ordby = document.frm.ordby.value;

//    if (makerid.length<1){
//        alert('먼저 브랜드를 선택후 출력해 주세요.');
//        return;
//    }

    var popwin;

	//popwin = window.open('/common/pop_brandstockprint.asp?makerid=' + makerid + '&stocktype=' + stocktype + '&itemidArr=' + itemidArr + '&cdl=' + cdl + '&cdm=' + cdm + '&cds=' + cds + '&onoffgubun=' + onoffgubun + '&mwdiv=' + mwdiv + '&centermwdiv=' + centermwdiv + '&sellyn=' + sellyn + '&isusing=' + isusing + '&danjongyn=' + danjongyn + '&returnitemgubun=' + returnitemgubun + '&itemname=' + itemname + '&limitrealstock=' + limitrealstock,'pop_brandstockprint','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin = window.open('/common/pop_brandstockprint.asp?makerid=' + makerid + '&stocktype=' + stocktype + '&itemidArr=' + itemidArr + '&disp=' + disp + '&onoffgubun=' + onoffgubun + '&mwdiv=' + mwdiv + '&centermwdiv=' + centermwdiv + '&sellyn=' + sellyn + '&isusing=' + isusing + '&danjongyn=' + danjongyn + '&returnitemgubun=' + returnitemgubun + '&itemname=' + itemname + '&limitrealstock=' + limitrealstock + '&itemrackcode=' + itemrackcode + '&warehouseCd=' + warehouseCd + '&ordby=' + ordby,'pop_brandstockprint','width=1000,height=600,scrollbars=yes,resizable=yes')
    popwin.focus();
}

function jsSetSellY() {
	var frm, i;

	var itemgubunarr = "";
	var itemidarr = "";
	var itemoptionarr = "";

	var selecteditemcount = 0;

	if (!CheckSelected()){
		alert("선택아이템이 없습니다.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

				if (frm.itemgubun.value !== "10") {
					alert("온라인(10) 상품만 선택가능합니다.");
					return;
				}

				itemgubunarr = itemgubunarr + frm.itemgubun.value + "|";
				itemidarr = itemidarr + frm.itemid.value + "|";
				itemoptionarr = itemoptionarr + frm.itemoption.value + "|";
			}
		}
	}

	if (selecteditemcount > 30) {
		alert("한번에 30개를 초과하여 상품을 선택할 수 없습니다.");
		return;
	}

	if (selecteditemcount === 0) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	if (confirm("선택상품 판매(Y) 설정하시겠습니까?") === true) {
		frm = document.frmAct;
		frm.mode.value = "setsell2y";
		frm.itemgubunarr.value=itemgubunarr;
		frm.itemidarr.value=itemidarr;
		frm.itemoptionarr.value=itemoptionarr;
		frm.submit();
	}
}

function jsSetBulkStockNo() {
	var frm, i;

    var barcodearr = "";
	var itemgubunarr = "";
	var itemidarr = "";
	var itemoptionarr = "";
    var bulkstockarr = "";

	var selecteditemcount = 0;

	if (!CheckSelected()){
		alert("선택아이템이 없습니다.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

                if ((frm.bulkstock.value == "") || (frm.bulkstock.value*0 != 0)) {
                    alert('벌크재고를 정확히 입력하세요.');
                    frm.bulkstock.focus();
                    return;
                }

				barcodearr = barcodearr + frm.barcode.value + "|";
                itemgubunarr = itemgubunarr + frm.itemgubun.value + "|";
				itemidarr = itemidarr + frm.itemid.value + "|";
				itemoptionarr = itemoptionarr + frm.itemoption.value + "|";
                bulkstockarr = bulkstockarr + frm.bulkstock.value + "|";
			}
		}
	}

	if (selecteditemcount > 100) {
		alert("한번에 100개를 초과하여 상품을 선택할 수 없습니다.");
		return;
	}

	if (selecteditemcount === 0) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	if (confirm("선택상품 벌크재고 저장하시겠습니까?") === true) {
		frm = document.frmAct;
		frm.mode.value = "setbulkstock";
		frm.barcodearr.value=barcodearr;
        frm.itemgubunarr.value=itemgubunarr;
		frm.itemidarr.value=itemidarr;
		frm.itemoptionarr.value=itemoptionarr;
        frm.itemnoarr.value=bulkstockarr;
		frm.submit();
	}
}

function jsSetBulkStockErrNo() {
	var frm, i;

	var itemgubunarr = "";
	var itemidarr = "";
	var itemoptionarr = "";
    var bulkrealstockarr = "";

	var selecteditemcount = 0;

	if (!CheckSelected()){
		alert("선택아이템이 없습니다.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

                if ((frm.bulkrealstock.value == "") || (frm.bulkrealstock.value*0 != 0)) {
                    alert('벌크재고를 정확히 입력하세요.');
                    frm.bulkstock.focus();
                    return;
                }

                itemgubunarr = itemgubunarr + frm.itemgubun.value + "|";
				itemidarr = itemidarr + frm.itemid.value + "|";
				itemoptionarr = itemoptionarr + frm.itemoption.value + "|";
                bulkrealstockarr = bulkrealstockarr + frm.bulkrealstock.value + "|";
			}
		}
	}

	if (selecteditemcount > 500) {
		alert("한번에 500개를 초과하여 상품을 선택할 수 없습니다.");
		return;
	}

	if (selecteditemcount === 0) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	if (confirm("선택상품 벌크재고[오차] 저장하시겠습니까?") === true) {
		frm = document.frmAct;
		frm.mode.value = "setbulkstockerr";
        frm.itemgubunarr.value=itemgubunarr;
		frm.itemidarr.value=itemidarr;
		frm.itemoptionarr.value=itemoptionarr;
        frm.itemnoarr.value=bulkrealstockarr;
		frm.submit();
	}
}

function PopReIpgo(){
	var frm;

	var itemgubunarr = "";
	var itemidarr = "";
	var itemoptionarr = "";
	var itemnoarr = "";
	var sellcasharr = "";
	var suplycasharr = "";
	var buycasharr = "";
	var itemnamearr = "";
	var itemoptionnamearr = "";
	var makeridarr = "";
	var mwdivarr = "";

	var makerid = "";
	var mwdiv = "";

	var selecteditemcount = 0;

	if (!CheckSelected()){
		alert("선택아이템이 없습니다.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

				if (makerid == "") {
					makerid = frm.makerid.value;
				} else if (makerid != frm.makerid.value) {
					alert("하나의 브랜드만 선택가능합니다.");
					return;
				}

				if (mwdiv == "") {
					mwdiv = frm.mwdiv.value;
				} else if (mwdiv != frm.mwdiv.value) {
					alert("매입 상품과 위탁 상품을 동시에 선택할 수 없습니다.");
					return;
				}

				makeridarr = makeridarr + frm.makerid.value + "|";
				itemgubunarr = itemgubunarr + frm.itemgubun.value + "|";
				itemidarr = itemidarr + frm.itemid.value + "|";
				itemoptionarr = itemoptionarr + frm.itemoption.value + "|";

				if (frm.returnitemno.value*1 >= 0) {
					itemnoarr = itemnoarr + "0|";
				} else {
					itemnoarr = itemnoarr + frm.returnitemno.value + "|";
				}

				// itemnoarr = itemnoarr + frm.returnitemno.value + "|";
				itemnamearr = itemnamearr + frm.itemname.value + "|";
				itemoptionnamearr = itemoptionnamearr + frm.itemoptionname.value + "|";
				sellcasharr = sellcasharr + frm.sellcash.value + "|";
				suplycasharr = suplycasharr + frm.suplycash.value + "|";
				buycasharr = buycasharr + frm.buycash.value + "|";
				mwdivarr = mwdivarr + frm.mwdiv.value + "|";
			}
		}
	}

	//if (selecteditemcount > 30) {
	//	alert("한번에 30개를 초과하여 상품을 선택할 수 없습니다.");
	//	return;
	//}

	if (makerid == "") {
		alert("선택된 상품이 없습니다.");
		return;
	}

    var popwin;
	var url = "/admin/newstorage/ipgoinput.asp?menupos=539";

	document.frmActPop.suplyer.value=makerid;
	document.frmActPop.itemgubunarr.value=itemgubunarr;
	document.frmActPop.itemidadd.value=itemidarr;
	document.frmActPop.itemoptionarr.value=itemoptionarr;
	document.frmActPop.itemnamearr.value=itemnamearr;
	document.frmActPop.itemoptionnamearr.value=itemoptionnamearr;
	document.frmActPop.sellcasharr.value=sellcasharr;
	document.frmActPop.suplycasharr.value=suplycasharr;
	document.frmActPop.buycasharr.value=buycasharr;
	document.frmActPop.itemnoarr.value=itemnoarr;
	document.frmActPop.designerarr.value=makeridarr;
	document.frmActPop.mwdivarr.value=mwdivarr;

	/*
	url = url + "&suplyer=" + makerid;
	url = url + "&itemgubunarr=" + itemgubunarr;
	url = url + "&itemidadd=" + itemidarr;
	url = url + "&itemoptionarr=" + itemoptionarr;
	url = url + "&itemnamearr=" + itemnamearr;
	url = url + "&itemoptionnamearr=" + itemoptionnamearr;
	url = url + "&sellcasharr=" + sellcasharr;
	url = url + "&suplycasharr=" + suplycasharr;
	url = url + "&buycasharr=" + buycasharr;
	url = url + "&itemnoarr=" + itemnoarr;
	url = url + "&designerarr=" + makeridarr;
	url = url + "&mwdivarr=" + mwdivarr;
	popwin = window.open(url, "PopReIpgo","width=1000,height=600,scrollbars=yes,resizable=yes");
    */

    popwin = window.open("", "PopReIpgo","width=1000,height=600,scrollbars=yes,resizable=yes");
    popwin.focus();
    document.frmActPop.action=url;
    document.frmActPop.target="PopReIpgo";
    document.frmActPop.submit();
}

function PopChulgo() {
	var frm;

	var itemgubunarr = "";
	var itemidarr = "";
	var itemoptionarr = "";
	var itemnoarr = "";
	var sellcasharr = "";
	var suplycasharr = "";
	var buycasharr = "";
	var itemnamearr = "";
	var itemoptionnamearr = "";
	var makeridarr = "";
	var mwdivarr = "";

	var makerid = "";
	var mwdiv = "";

	var selecteditemcount = 0;

	if (!CheckSelected()){
		alert("선택아이템이 없습니다.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

				makeridarr = makeridarr + frm.makerid.value + "|";
				itemgubunarr = itemgubunarr + frm.itemgubun.value + "|";
				itemidarr = itemidarr + frm.itemid.value + "|";
				itemoptionarr = itemoptionarr + frm.itemoption.value + "|";

				if (frm.returnitemno.value*1 >= 0) {
					itemnoarr = itemnoarr + "0|";
				} else {
					itemnoarr = itemnoarr + frm.returnitemno.value*-1 + "|";
				}

				// itemnoarr = itemnoarr + frm.returnitemno.value + "|";
				itemnamearr = itemnamearr + frm.itemname.value + "|";
				itemoptionnamearr = itemoptionnamearr + frm.itemoptionname.value + "|";
				sellcasharr = sellcasharr + frm.sellcash.value + "|";
				suplycasharr = suplycasharr + 0 + "|";
				buycasharr = buycasharr + frm.buycash.value + "|";
				mwdivarr = mwdivarr + frm.mwdiv.value + "|";

                makerid = frm.makerid.value;
			}
		}
	}

	if (makerid == "") {
		alert("선택된 상품이 없습니다.");
		return;
	}

    var popwin;
	var url = "/admin/newstorage/chulgoinput.asp?menupos=540";

	//document.frmActPop.suplyer.value=makerid;
	document.frmActPop.itemgubunarr.value=itemgubunarr;
	document.frmActPop.itemidarr.value=itemidarr;
	document.frmActPop.itemoptionarr.value=itemoptionarr;
	document.frmActPop.itemnamearr.value=itemnamearr;
	document.frmActPop.itemoptionnamearr.value=itemoptionnamearr;
	document.frmActPop.sellcasharr.value=sellcasharr;
	document.frmActPop.suplycasharr.value=suplycasharr;
	document.frmActPop.buycasharr.value=buycasharr;
	document.frmActPop.itemnoarr.value=itemnoarr;
	document.frmActPop.designerarr.value=makeridarr;
	document.frmActPop.mwdivarr.value=mwdivarr;

	/*
	url = url + "&suplyer=" + makerid;
	url = url + "&itemgubunarr=" + itemgubunarr;
	url = url + "&itemidadd=" + itemidarr;
	url = url + "&itemoptionarr=" + itemoptionarr;
	url = url + "&itemnamearr=" + itemnamearr;
	url = url + "&itemoptionnamearr=" + itemoptionnamearr;
	url = url + "&sellcasharr=" + sellcasharr;
	url = url + "&suplycasharr=" + suplycasharr;
	url = url + "&buycasharr=" + buycasharr;
	url = url + "&itemnoarr=" + itemnoarr;
	url = url + "&designerarr=" + makeridarr;
	url = url + "&mwdivarr=" + mwdivarr;
	popwin = window.open(url, "PopReIpgo","width=1000,height=600,scrollbars=yes,resizable=yes");
    */

    popwin = window.open("", "PopChulgo","width=1000,height=600,scrollbars=yes,resizable=yes");
    popwin.focus();
    document.frmActPop.action=url;
    document.frmActPop.target="PopChulgo";
    document.frmActPop.submit();
}

function PopModiRackCode(mode) {
	var frm;

	var itemgubunarr = "";
	var itemidarr = "";
	var itemoptionarr = "";

	var makerid = "";

	var selecteditemcount = 0;

	if (!CheckSelected()){
		alert("선택아이템이 없습니다.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

				if (makerid == "") {
					makerid = frm.makerid.value;
				} else if (makerid != frm.makerid.value) {
					alert("하나의 브랜드만 선택가능합니다.");
					return;
				}

				itemgubunarr = itemgubunarr + frm.itemgubun.value + "|";
				itemidarr = itemidarr + frm.itemid.value + "|";
				itemoptionarr = itemoptionarr + frm.itemoption.value + "|";
			}
		}
	}

	//if (selecteditemcount > 30) {
	//	alert("한번에 30개를 초과하여 상품을 선택할 수 없습니다.");
	//	return;
	//}

	if (makerid == "") {
		alert("선택된 상품이 없습니다.");
		return;
	}

    var popwin;
	var url = "/admin/stock/popMultiRackCode.asp";

	document.frmActPop.mode.value=mode;
	document.frmActPop.itemgubunarr.value=itemgubunarr;
	document.frmActPop.itemidadd.value=itemidarr;
	document.frmActPop.itemoptionarr.value=itemoptionarr;

    popwin = window.open("", "PopModiRackCode","width=300,height=150,scrollbars=yes,resizable=yes");
    popwin.focus();
    document.frmActPop.action=url;
    document.frmActPop.target="PopModiRackCode";
    document.frmActPop.submit();
}

function RefreshIpchulStock(){
	if (frmrefresh.makerid.value==""){
		alert('브랜드를 선택 하세요.');
		frmrefresh.makerid.focus();
	}

	if (confirm('입출고 내역 전체 새로고침 하시겠습니까?')){
		frmrefresh.mode.value="ipchulallrefreshbybrand";
		frmrefresh.submit();
	}
}

function DelItem(itemgubun,itemid,itemoption){
	if (confirm('입출고 내역 을 삭제 하시겠습니까?')){
		frmrefresh.mode.value="ipchuldellbyitemid";
		frmrefresh.itemgubun.value=itemgubun;
		frmrefresh.itemid.value=itemid;
		frmrefresh.itemoption.value=itemoption;

		frmrefresh.submit();
	}
}

function chkEnDisabled(comp){
    var frm = comp.form;
    if (comp.value==""){
       frm.sellyn.disabled=false;
       //frm.usingyn.disabled=false;
       frm.danjongyn.disabled=false;
    }else{
       frm.sellyn.disabled=true;
       //frm.usingyn.disabled=true;
       frm.danjongyn.disabled=true;
    }
}

function trim(value) {
	return value.replace(/^\s+|\s+$/g,"");
}

function SubmitSearch() {
	frm.action="";
	frm.target="";
    document.frm.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

//인덱스 출력 일괄
function IndexBarcodePrint() {
	var arr = new Array();
	var isforeignprint; var printpriceval; var domainname; var showdomainyn; var ttptype;
	var found = false;
	var prdname; var itemoptionname; var customerprice; var barcodetype;
	var skipnotinserted; var printno; var shopbrandyn; var currencychar; var showpriceyn;
	var saleprice; var saleyn; var socname, brandrackcode, itemrackcode, itemoptionrackcode, subitemrackcode;

	isforeignprint = document.frm.isforeignprint.value;
	skipnotinserted = false;

	shopbrandyn		= "Y";
	ttptype			= "TTP-243_80x50";
	barcodetype		= "2";								// T or G or 2(텐바이텐 바코드 or 범용바코드 or 텐텐바코드_범용바코드)
	var paperwidth = frm.paperwidth.value;
	var paperheight = frm.paperheight.value;
	var papermargin = frm.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frm.printpriceyn.value;

	if (isforeignprint == "N") {
		currencychar = "￦";
	} else {
		currencychar = "<%= currencyChar %>";
	}
	domainname		= "www.10x10.co.kr";
	showdomainyn	= frm.titledispyn.value;

	for (var i = 0; ; i++) {
		itembarcode = document.getElementById("itembarcode_" + i);
		chk = document.getElementById("chk_" + i);

		if (itembarcode == undefined) {
			break;
		}

		if (chk.checked != true) {
			continue;
		}

		found = true;

		if (chk.checked == true) {
			//saleprice = document.getElementById("saleprice_" + i).value.trim();
			//saleyn = document.getElementById("saleyn_" + i).value.trim();

			//해외 상품명
			if (isforeignprint == "Y") {
			//	itemname = document.getElementById("itemname_foreign_" + i).value.trim();
			//	itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
			//	customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();

			//국내 상품명
			} else {
				itemname = document.getElementById("itemname_" + i).value.trim();
				itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();

				//할인가 표시
				if (showpriceyn=='C'){
					//할인
					if (saleyn=='Y'){
						customerprice = saleprice;

					//소비자가
					}else{
						customerprice = document.getElementById("customerprice_" + i).value.trim();
					}

				//판매가 표시
				}else if (showpriceyn=='R'){
					customerprice = document.getElementById("sellprice_" + i).value.trim();

				//소비자가 표시
				}else{
					customerprice = document.getElementById("customerprice_" + i).value.trim();
				}
			}

			itembarcode = document.getElementById("itembarcode_" + i).value.trim();
			publicbarcode = document.getElementById("publicbarcode_" + i).value.trim();
			itembarcode = itembarcode + "_" + publicbarcode;

			makerid = document.getElementById("makerid_" + i).value.trim();
			socname = document.getElementById("socname_" + i).value.trim();
			printno = document.getElementById("printno_" + i).value.trim();
			//printno = 1;
			brandrackcode = document.getElementById("prtidx_" + i).value.trim();
			itemrackcode = document.getElementById("itemrackcode_" + i).value.trim();
			subitemrackcode = document.getElementById("subitemrackcode_" + i).value.trim();

			if (printno*1 != 0) {
				var v = new BarcodeDataClass_index(itembarcode, socname, itemname, itemoptionname, customerprice, printno, '', '', '', brandrackcode, itemrackcode, itemoptionrackcode, subitemrackcode);
				arr.push(v);
			}
		}
	}

	if (found == false) {
		alert("선택된 상품이 없습니다.");
		return;
	}

	//TEC B-FV4		//2016.11.24 한용민 생성
	if (frm.printername.value=='TEC_B-FV4_80x50'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 드라이버를 설치해 주세요!!"); return; }
		if (confirm("선택 상품의 인덱스를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
			TOSHIBA_DOMAINNAME = domainname;
			TOSHIBA_SHOWDOMAINYN = showdomainyn;
			TOSHIBA_PAPERWIDTH = 800;
			TOSHIBA_PAPERHEIGHT = 500;
			TOSHIBA_PAPERMARGIN = 3;
			TOSHIBA_SHOWPRICEYN = showpriceyn;
			TOSHIBA_currencyChar = currencychar;
			TOSHIBA_SHOPBRANDYN = shopbrandyn;
			TOSHIBA_BARCODETYPE = barcodetype;

			printTOSHIBAMultiItemLabel(arr);
		}

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) == true) {
		if (confirm("선택 상품의 인덱스를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPMultiItemLabel(arr);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}
	return;
}

function regAGVArr(){
	var frm = document.frmActPop;
	var found = false;
    var zerofound = false;
    var agvstock, chk;
    var itemgubunarr, itemidarr, itemoptionarr, itemnoarr;
    var itemgubun, itemid, itemoption;

    itemgubunarr = '';
    itemidarr = '';
    itemoptionarr = '';
    itemnoarr = '';

    for (var i = 0; ; i++) {
        chk = document.getElementById('chk_' + i);
        agvstock = document.getElementById('agvstock_' + i);
        itemgubun = document.getElementById('itemgubun_' + i);
        itemid = document.getElementById('itemid_' + i);
        itemoption = document.getElementById('itemoption_' + i);

        if (chk == undefined) { break; }

        if (chk.checked == true) {
            if (agvstock.value*0 != 0) {
                alert('AGV 수량을 확인하세요.');
                return false;
            }

            if (agvstock.value*1 <= 0) {
                zerofound = true;
            } else {
                found = true;

                itemgubunarr = itemgubunarr + ',' + itemgubun.value;
                itemidarr = itemidarr + ',' + itemid.value;
                itemoptionarr = itemoptionarr + ',' + itemoption.value;
                itemnoarr = itemnoarr + ',' + agvstock.value;
            }
        }
    }

	if (found == true) {
		if (confirm("수량이 1개 미만인 상품은 제외됩니다.\n\n선택상품을 AGV인터페이스에 저장 하시겠습니까?") == true) {
			frm.mode.value = "agvregarr";

            frm.itemgubunarr.value = itemgubunarr;
            frm.itemidarr.value = itemidarr;
            frm.itemoptionarr.value = itemoptionarr;
            frm.itemnoarr.value = itemnoarr;

			frm.action = "/admin/logics/logics_agv_pickup_process.asp";
			frm.submit();
		}
	} else {
		alert("선택된 상품이 없습니다.");
	}
}

String.prototype.trim = function() {
    return this.replace(/^\s+|\s+$/g, "");
};

function exceldownload(){
    var frm = document.frm;
	frm.action="/admin/stock/agvstockdiff_excel.asp";
	frm.target="view";
	frm.submit();
	frm.action="";
	frm.target="";
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="6" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<table border="0" cellpadding="5" cellspacing="0" class="a">
			<tr>
				<td>브랜드:	<% drawSelectBoxDesignerwithName "makerid", makerid %></td>
				<td></td>
				<td></td>
				<td></td>
			</tr>
		</table>
	</td>
	<td rowspan="6" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="SubmitSearch();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		* 거래구분 : <% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;&nbsp;
        &nbsp;&nbsp;
		* 진열구분 :
        <select class="select" name="warehouseCd">
            <option value="AGV" <% if warehouseCd="AGV" then response.write "selected" %> >AGV</option>
        </select>
        &nbsp;&nbsp;
		* AGV재고 :
        <select class="select" name="agvstockgubun">
            <option value="availdiff" <% if agvstockgubun="availdiff" then response.write "selected" %> >유효재고 불일치만</option>
        </select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
	    * <select name="stocktype" class="select">
			<option value="real" <% if (stocktype = "real") then %>selected<% end if %> >유효재고</option>
		</select>
	    <select class="select" name="limitrealstock" >
		<option value='1UP' >1개 이상</option>
	    </select>
		* 정렬순서 :
		<select class="select" name="ordby">
			<option value="1" <%= CHKIIF(ordby = "1", "selected", "") %> >상품코드</option>
		</select>
		&nbsp;&nbsp;
		<input type="checkbox" class="checkbox" name="excits" value="Y" <%= CHKIIF(excits="Y", "checked", "") %> > 아이띵소,코니테일 제외
        &nbsp;&nbsp;
		<input type="checkbox" class="checkbox" name="excNoRack" value="Y" <%= CHKIIF(excNoRack="Y", "checked", "") %> > 랙코드 ZZZZ 제외
	</td>
</tr>
</table>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
	</td>
	<td align="right">
    	<input type="button" class="button" value="엑셀다운로드" onclick="exceldownload();">
		* 표시갯수:
		<select class="select" name="pagesize" >
			<option value="25">25개</option>
			<option value="100" <%= ChkIIF(pagesize="100","selected","") %> >100개</option>
			<option value="200" <%= ChkIIF(pagesize="200","selected","") %> >200개</option>
            <option value="500" <%= ChkIIF(pagesize="500","selected","") %> >500개</option>
		</select>
	</td>
</tr>
</table>
</form>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="35">
		검색결과 : <b><%= osummarystockbrand.FTotalCount %></b>
		&nbsp;
		페이지 :
		<% if osummarystockbrand.FCurrPage > 1  then %>
			<a href="javascript:GotoPage(<%= page - 1 %>)"><img src="/images/icon_arrow_left.gif" border="0" align="absbottom"></a>
		<% end if %>
		<b><%= page %> / <%= osummarystockbrand.FTotalPage %></b>
		<% if (osummarystockbrand.FTotalpage - osummarystockbrand.FCurrPage)>0  then %>
			<a href="javascript:GotoPage(<%= page + 1 %>)"><img src="/images/icon_arrow_right.gif" border="0" align="absbottom"></a>
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
    <td>랙코드</td>
    <td>구분</td>
	<td>상품코드</td>
	<td>옵션<br />코드</td>
	<td>이미지</td>
	<td>브랜드ID</td>
	<td>상품명<br>[옵션명]</td>
	<td>소비자가</td>
	<td>매입가(현)</td>
	<td>매입<br>구분</td>
	<td>센터<br>매입<br>구분</td>
	<td>총<br>입고<br>반품</td>
	<td>ON총<br>판매<br>반품</td>
    <td>OFF총<br>출고<br>반품</td>
    <td>기타<br>출고<br>반품</td>
    <td>CS<br>출고<br>반품</td>
    <td bgcolor="F4F4F4"><b>시스템<br>총재고</b></td>

	<td>총<br>실사<br>오차</td>
	<td>실사<br>재고</td>
	<td>총<br>불량</td>
	<td bgcolor="F4F4F4"><b>유효<br>재고</b></td>

    <td>총<br>상품<br>준비</td>
    <td bgcolor="F4F4F4"><b>재고<br>파악<br>재고</b></td>
    <td>발주<br>이전<br>주문</td>
    <td bgcolor="F4F4F4">예상<br>재고</td>
    <td width="30">판매<br>여부</td>
    <td width="30">한정<br>여부</td>
    <td>단종<br>여부</td>
    <td width="40">오차<br>입력</td>
	<td width="40">출력수</td>
	<td width="60">마지막<br>입고월</td>
	<td width="40">전월<br />판매<br />(물류)</td>
    <td width="35">상품<br />등급</td>
    <td>AGV<br />재고</td>
</tr>
<% if osummarystockbrand.FResultCount>0 then %>
<% for i=0 to osummarystockbrand.FResultCount - 1 %>
<%
totsysstock	= totsysstock + osummarystockbrand.FItemList(i).Ftotsysstock
totavailstock = totavailstock + osummarystockbrand.FItemList(i).Favailsysstock
totrealstock = totrealstock + osummarystockbrand.FItemList(i).Frealstock
totjeagosheetstock = totjeagosheetstock + osummarystockbrand.FItemList(i).Frealstock + osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno
totmaystock = totmaystock + osummarystockbrand.FItemList(i).Frealstock + osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno + osummarystockbrand.FItemList(i).Fipkumdiv4 + osummarystockbrand.FItemList(i).Fipkumdiv2 + osummarystockbrand.FItemList(i).Foffjupno

%>
<% if osummarystockbrand.FItemList(i).Fisusing="Y" then %>
<tr bgcolor="#FFFFFF" align="center">
<% else %>
<tr bgcolor="#EEEEEE" align="center">
<% end if %>
	<form name="frmBuyPrc_<%= i %>" >
	<input type="hidden" id="itembarcode_<%= i %>" name="barcode" value="<%= BF_MakeTenBarcode(osummarystockbrand.FItemList(i).Fitemgubun, osummarystockbrand.FItemList(i).Fitemid, osummarystockbrand.FItemList(i).Fitemoption) %>">
	<input type="hidden" id="publicbarcode_<%= i %>" name="generalbarcode" value="<%= osummarystockbrand.FItemList(i).FpublicBarcode %>">
	<input type="hidden" id="customerprice_<%= i %>" name="orgprice" value="<%= (osummarystockbrand.FItemList(i).Forgprice) %>">
	<input type="hidden" id="itemname_<%= i %>" name="itemname" value="<%= osummarystockbrand.FItemList(i).FItemName %>">
	<input type="hidden" id="itemoptionname_<%= i %>" name="itemoptionname" value="<%= osummarystockbrand.FItemList(i).FItemOptionName %>">
	<input type="hidden" id="sellprice_<%= i %>" name="sellcash" value="<%= osummarystockbrand.FItemList(i).Fsellcash %>">
	<input type="hidden" id="makerid_<%= i %>" name="makerid" value="<%= osummarystockbrand.FItemList(i).FMakerid %>">
	<input type="hidden" id="socname_<%= i %>" name="socname" value="<%= osummarystockbrand.FItemList(i).FMakerid %>">
	<input type="hidden" id="prtidx_<%= i %>" name="prtidx" value="<%= osummarystockbrand.FItemList(i).fprtidx %>">
	<input type="hidden" id="itemrackcode_<%= i %>" name="itemrackcode" value="<%= osummarystockbrand.FItemList(i).fitemrackcode %>">
	<input type="hidden" id="subitemrackcode_<%= i %>" name="subitemrackcode" value="<%= osummarystockbrand.FItemList(i).fsubitemrackcode %>">
	<input type="hidden" name="barcode2" value="<%= BF_MakeTenBarcode(osummarystockbrand.FItemList(i).Fitemgubun, osummarystockbrand.FItemList(i).Fitemid, osummarystockbrand.FItemList(i).Fitemoption) %>_<%= osummarystockbrand.FItemList(i).FpublicBarcode %>">
	<input type="hidden" id="itemgubun_<%= i %>" name="itemgubun" value="<%= osummarystockbrand.FItemList(i).Fitemgubun %>">
	<input type="hidden" id="itemid_<%= i %>" name="itemid" value="<%= osummarystockbrand.FItemList(i).Fitemid %>">
	<input type="hidden" id="itemoption_<%= i %>" name="itemoption" value="<%= osummarystockbrand.FItemList(i).Fitemoption %>">
	<input type="hidden" name="returnitemno" value="<%= osummarystockbrand.FItemList(i).Frealstock*-1 %>">
	<input type="hidden" name="suplycash" value="<%= chkIIF(osummarystockbrand.FItemList(i).IsOffContractExist, osummarystockbrand.FItemList(i).GetOffContractBuycash, osummarystockbrand.FItemList(i).FBuycash) %>">
	<input type="hidden" name="buycash" value="<%= chkIIF(osummarystockbrand.FItemList(i).IsOffContractExist, osummarystockbrand.FItemList(i).GetOffContractBuycash, osummarystockbrand.FItemList(i).FBuycash) %>">
	<input type="hidden" name="mwdiv" value="<%= chkIIF(osummarystockbrand.FItemList(i).IsOffContractExist, osummarystockbrand.FItemList(i).GetOffContractCenterMW, osummarystockbrand.FItemList(i).Fmwdiv) %>">
	<td width=20><input type="checkbox" name="cksel" id="chk_<%= i %>" onClick="AnCheckClick(this);"></td>
    <td><%= osummarystockbrand.FItemList(i).FItemrackcode %></td>
    <td><%= osummarystockbrand.FItemList(i).FItemGubun %></td>
	<td>
	    <% if osummarystockbrand.FItemList(i).FItemGubun="10" then %>
	    <a href="javascript:PopItemSellEdit('<%= osummarystockbrand.FItemList(i).Fitemid %>');"><%= osummarystockbrand.FItemList(i).Fitemid %></a>
	    <% else %>
	    <%= osummarystockbrand.FItemList(i).Fitemid %>
	    <% end if %>
	</td>
    <td><%= osummarystockbrand.FItemList(i).Fitemoption %></td>
	<td><img src="<%= osummarystockbrand.FItemList(i).Fimgsmall %>" width="50" height="50"></td>
	<td><%= osummarystockbrand.FItemList(i).FMakerid %></td>
	<td align="left">
      	<a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= osummarystockbrand.FItemList(i).FItemGubun %>&itemid=<%= osummarystockbrand.FItemList(i).FItemID %>&itemoption=<%= osummarystockbrand.FItemList(i).FItemOption %>" target=_blank ><%= osummarystockbrand.FItemList(i).Fitemname %></a>
      	<% if osummarystockbrand.FItemList(i).FitemoptionName <>"" then %>
      		<br>
      		<font color="blue">[<%= osummarystockbrand.FItemList(i).FitemoptionName %>]</font>
      	<% end if %>
    </td>
	<td align="right"><%= FormatNumber(osummarystockbrand.FItemList(i).Forgprice,0) %></td>
	<td align="right"><%= FormatNumber(osummarystockbrand.FItemList(i).Fbuycash,0) %></td>
    <td><%= fnColor(osummarystockbrand.FItemList(i).Fmwdiv,"mw") %></td>
    <td>
		<%= fnColor(osummarystockbrand.FItemList(i).Fcentermwdiv,"mw") %>
		<% if osummarystockbrand.FItemList(i).IsOffContractExist then %>
		<br />
			<% if osummarystockbrand.FItemList(i).Forgprice<>0 then %>
			<%= 100-(CLng(osummarystockbrand.FItemList(i).FBuycash/osummarystockbrand.FItemList(i).Forgprice*10000)/100) %> %
			<% end if %>
			<br>-&gt;<font color="blue"><%= osummarystockbrand.FItemList(i).GetOffContractMargin %>%</font>
		<% end if %>
	</td>
	<td align="right"><%= osummarystockbrand.FItemList(i).Ftotipgono %></td>
	<td align="right"><%= -1*osummarystockbrand.FItemList(i).Ftotsellno %></td>
	<td align="right"><%= osummarystockbrand.FItemList(i).Foffchulgono + osummarystockbrand.FItemList(i).Foffrechulgono %></td>
	<td align="right"><%= osummarystockbrand.FItemList(i).Fetcchulgono + osummarystockbrand.FItemList(i).Fetcrechulgono %></td>
	<td align="right"><%= osummarystockbrand.FItemList(i).Ferrcsno %></td>
	<td align="right" bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Ftotsysstock %></b></td>

	<td align="right"><%= osummarystockbrand.FItemList(i).Ferrrealcheckno %></td>
	<td align="right"><%= osummarystockbrand.FItemList(i).getErrAssignStock %></td>
	<td align="right"><%= osummarystockbrand.FItemList(i).Ferrbaditemno %></td>
	<td align="right" bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Frealstock %></td>

	<td align="right"><%= osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno %></td>
	<td align="right" bgcolor="F4F4F4"><b><%= osummarystockbrand.FItemList(i).Frealstock + osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno %></b></td>
	<td align="right"><%= osummarystockbrand.FItemList(i).Fipkumdiv4 + osummarystockbrand.FItemList(i).Fipkumdiv2 + osummarystockbrand.FItemList(i).Foffjupno %></td>
	<td align="right" bgcolor="F4F4F4">
		<% if osummarystockbrand.FItemList(i).FLimitYn="Y" then %>
			<font color="#FF0000"><%= osummarystockbrand.FItemList(i).Frealstock + osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno + osummarystockbrand.FItemList(i).Fipkumdiv4 + osummarystockbrand.FItemList(i).Fipkumdiv2 + osummarystockbrand.FItemList(i).Foffjupno %></font>
		<% else %>
      		<b><%= osummarystockbrand.FItemList(i).Frealstock + osummarystockbrand.FItemList(i).Fipkumdiv5 + osummarystockbrand.FItemList(i).Foffconfirmno + osummarystockbrand.FItemList(i).Fipkumdiv4 + osummarystockbrand.FItemList(i).Fipkumdiv2 + osummarystockbrand.FItemList(i).Foffjupno %></b>
    	<% end if %>
    </td>
	<td><%= fnColor(osummarystockbrand.FItemList(i).Fsellyn,"yn") %></td>
	<td>
		<%= fnColor(osummarystockbrand.FItemList(i).Flimityn,"yn") %>
		<% if osummarystockbrand.FItemList(i).Flimityn="Y" then %>
		<br>(<%= osummarystockbrand.FItemList(i).GetLimitStr %>)
		<% end if %>
	</td>
	<td><%= fnColor(osummarystockbrand.FItemList(i).Fdanjongyn,"dj") %></td>
	<td>
		<input type="button" class="button" value="오차" onclick="popRealErrInput('<%= osummarystockbrand.FItemList(i).Fitemgubun %>','<%= osummarystockbrand.FItemList(i).Fitemid %>','<%= osummarystockbrand.FItemList(i).Fitemoption %>');">
	</td>
	<td>
		<input type="text" class="text" id="printno_<%= i %>" name="itemno" value="1" size="1" maxlength="8" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)" style="text-align:center;">
	</td>
	<td>
		<%= osummarystockbrand.FItemList(i).FlastIpgoDate %>
	</td>
	<td>
		<%= osummarystockbrand.FItemList(i).FprevMonthSellCnt %>
	</td>
    <td>
        <% if (osummarystockbrand.FItemList(i).Fitemgrade = "A") then %><font color="red"><% end if %>
		<%= osummarystockbrand.FItemList(i).Fitemgrade %>
	</td>
    <td>
		<input type="text" class="text" id="agvstock_<%= i %>" name="agvstock" value="<%= osummarystockbrand.FItemList(i).Fagvstock %>" size="1" maxlength="8" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)" style="text-align:center;">
	</td>
</tr>
</form>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="35" align="center">
		<% if osummarystockbrand.HasPreScroll then %>
		<a href="javascript:NextPage('<%= osummarystockbrand.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + osummarystockbrand.StartScrollPage to osummarystockbrand.FScrollCount + osummarystockbrand.StartScrollPage - 1 %>
			<% if i>osummarystockbrand.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if osummarystockbrand.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="37" align="center" class="page_link">[검색결과가 없습니다.]</td>
    </tr>
<% end if %>
</table>
<form name="frmActPop" method="post" action="" style="margin:0px;">
<input type="hidden" name="suplyer" value="">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemidadd" value="">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="itemnamearr" value="">
<input type="hidden" name="itemoptionnamearr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="designerarr" value="">
<input type="hidden" name="mwdivarr" value="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="refergubun" value="brandstock">
</form>
<form name="frmAct" method="post" action="brandcurrentstock_process.asp" style="margin:0px;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="barcodearr" value="">
    <input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
    <input type="hidden" name="itemnoarr" value="">
</form>
<form name=frmrefresh method=post action="dostockrefresh.asp" style="margin:0px;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="refreshstartdate" value="<%= BasicMonth + "-01" %>">
	<input type="hidden" name="makerid" value="<%= makerid %>">
	<input type="hidden" name="itemgubun" value="">
	<input type="hidden" name="itemid" value="">
	<input type="hidden" name="itemoption" value="">
</form>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height="300" frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height="0" frameborder="0" scrolling="no"></iframe>
<% end if %>
<%
set osummarystockbrand = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
