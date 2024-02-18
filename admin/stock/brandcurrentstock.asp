<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �귣�庰�����Ȳ
' History : 2009.04.07 ������ ����
'			2022.02.09 �ѿ�� ����(�������� ��񿡼� �������� �����۾�)
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
dim limityn, itemgrade, itemrackcode, bulkstockgubun, warehouseCd, agvstockgubun
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

if (stocktype = "") then stocktype = "sys"
if (pagesize = "") then pagesize = 25

'///////////////// ���ڵ� ����Ʈ�� ���� ///////////////////////
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
currencyChar = "��"
'/////////////////'////////////////////////////////////////

'//��ǰ�ڵ� ��ȿ�� �˻�
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
	osummarystockbrand.FRectCD1   = cdl
	osummarystockbrand.FRectCD2   = cdm
	osummarystockbrand.FRectCD3   = cds
	osummarystockbrand.FRectItemIdArr = itemidArr
	osummarystockbrand.FRectItemName = itemname
	osummarystockbrand.FRectMakerid = makerid
	osummarystockbrand.FRectOnlySellyn = sellyn
	osummarystockbrand.FRectOnlyIsUsing = usingyn
	osummarystockbrand.FRectDanjongyn =danjongyn
	osummarystockbrand.FRectMWDiv = mwdiv
	osummarystockbrand.FRectCenterMWDiv = centermwdiv
	osummarystockbrand.FRectReturnItemGubun = returnitemgubun
	osummarystockbrand.FRectlimitrealstock = limitrealstock
	osummarystockbrand.FRectStockType = stocktype
	osummarystockbrand.FRectDispCate = dispCate
	osummarystockbrand.FRectUseOffInfo = useoffinfo
	osummarystockbrand.FRectExcIts = excits
	osummarystockbrand.FRectPurchasetype = vPurchasetype
    osummarystockbrand.FRectLimitYN = limityn
    osummarystockbrand.FRectItemGrade = itemgrade
    osummarystockbrand.FRectRackCode = itemrackcode
    osummarystockbrand.FRectBulkStockGubun = bulkstockgubun
    osummarystockbrand.FRectWarehouseCd = warehouseCd
    osummarystockbrand.FRectAgvStockGubun = agvstockgubun

	if (ordby = "1") then
		osummarystockbrand.FRectOrderBy = "T.itemid desc"
	elseif (ordby = "2") then
		osummarystockbrand.FRectOrderBy = "T.itemrackcode asc,T.itemid desc"
	end if

	if IsNumeric(startMon) then
		osummarystockbrand.FRectStartDate = startMon
	elseif (startMon <> "") then
		response.write "<script>alert('������ ���ڸ� �����մϴ�. " & startMon & "')</script>"
	end if
	if IsNumeric(endMon) then
		osummarystockbrand.FRectEndDate = endMon
	elseif (endMon <> "") then
		response.write "<script>alert('������ ���ڸ� �����մϴ�. " & endMon & "')</script>"
	end if

	if (itemgubun = "10") and ((itemidArr<>"") or (itemname<>"") or (makerid<>"") or (cdl<>"") or (mwdiv<>"")) then
		''osummarystockbrand.GetCurrentStockByOnlineBrand
		osummarystockbrand.GetCurrentStockByOnlineBrandNEW
	elseif itemgubun <> "10" then
		if itemgubun <> "exc10" then
			osummarystockbrand.FRectItemGubun =  itemgubun
		end if
		osummarystockbrand.GetCurrentStockByOfflineBrand
	end if

IsSheetPrintEnable = (osummarystockbrand.FResultCount>0)

dim bulkrealstock

%>

<script type="text/javascript" src="/js/barcode.js"></script>
<script type="text/javascript" src="/js/ttpbarcode.js"></script>
<script type="text/javascript" src="/js/DOSHIBAbarcode.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

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

//�¶��λ�ǰ
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
}
//������ǰ
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
    var excits = document.frm.excits.value;
    var ordby = document.frm.ordby.value;

//    if (makerid.length<1){
//        alert('���� �귣�带 ������ ����� �ּ���.');
//        return;
//    }

    //var popwin;
	//popwin = window.open('/common/pop_brandstockprint.asp?makerid=' + makerid + '&stocktype=' + stocktype + '&itemidArr=' + itemidArr + '&disp=' + disp + '&onoffgubun=' + onoffgubun + '&mwdiv=' + mwdiv + '&centermwdiv=' + centermwdiv + '&sellyn=' + sellyn + '&isusing=' + isusing + '&danjongyn=' + danjongyn + '&returnitemgubun=' + returnitemgubun + '&itemname=' + itemname + '&limitrealstock=' + limitrealstock + '&itemrackcode=' + itemrackcode + '&warehouseCd=' + warehouseCd + '&excits=' + excits + '&ordby=' + ordby,'pop_brandstockprint','width=1400,height=800,scrollbars=yes,resizable=yes')
    //popwin.focus();
	var url = "/common/pop_brandstockprint.asp";
    popwin = window.open("", "PopBrandStockPrint","width=1400,height=800,scrollbars=yes,resizable=yes");
    popwin.focus();
    document.frm.action=url;
    document.frm.target="PopBrandStockPrint";
    document.frm.submit();
	frm.action="";
	frm.target="";
}

function jsSetSellY() {
	var frm, i;

	var itemgubunarr = "";
	var itemidarr = "";
	var itemoptionarr = "";

	var selecteditemcount = 0;

	if (!CheckSelected()){
		alert("���þ������� �����ϴ�.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

				if (frm.itemgubun.value !== "10") {
					alert("�¶���(10) ��ǰ�� ���ð����մϴ�.");
					return;
				}

				itemgubunarr = itemgubunarr + frm.itemgubun.value + "|";
				itemidarr = itemidarr + frm.itemid.value + "|";
				itemoptionarr = itemoptionarr + frm.itemoption.value + "|";
			}
		}
	}

	if (selecteditemcount > 30) {
		alert("�ѹ��� 30���� �ʰ��Ͽ� ��ǰ�� ������ �� �����ϴ�.");
		return;
	}

	if (selecteditemcount === 0) {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm("���û�ǰ �Ǹ�(Y) �����Ͻðڽ��ϱ�?") === true) {
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
		alert("���þ������� �����ϴ�.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

                if ((frm.bulkstock.value == "") || (frm.bulkstock.value*0 != 0)) {
                    alert('��ũ��� ��Ȯ�� �Է��ϼ���.');
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
		alert("�ѹ��� 100���� �ʰ��Ͽ� ��ǰ�� ������ �� �����ϴ�.");
		return;
	}

	if (selecteditemcount === 0) {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm("���û�ǰ ��ũ��� �����Ͻðڽ��ϱ�?") === true) {
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
		alert("���þ������� �����ϴ�.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

                if ((frm.bulkrealstock.value == "") || (frm.bulkrealstock.value*0 != 0)) {
                    alert('��ũ��� ��Ȯ�� �Է��ϼ���.');
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
		alert("�ѹ��� 500���� �ʰ��Ͽ� ��ǰ�� ������ �� �����ϴ�.");
		return;
	}

	if (selecteditemcount === 0) {
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	if (confirm("���û�ǰ ��ũ���[����] �����Ͻðڽ��ϱ�?") === true) {
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
		alert("���þ������� �����ϴ�.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

				if (makerid == "") {
					makerid = frm.makerid.value.toUpperCase();
				} else if (makerid != frm.makerid.value.toUpperCase()) {
					alert("�ϳ��� �귣�常 ���ð����մϴ�.");
					return;
				}

				if (mwdiv == "") {
					mwdiv = frm.mwdiv.value;
				} else if (mwdiv != frm.mwdiv.value) {
					alert("���� ��ǰ�� ��Ź ��ǰ�� ���ÿ� ������ �� �����ϴ�.");
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
	//	alert("�ѹ��� 30���� �ʰ��Ͽ� ��ǰ�� ������ �� �����ϴ�.");
	//	return;
	//}

	if (makerid == "") {
		alert("���õ� ��ǰ�� �����ϴ�.");
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
		alert("���þ������� �����ϴ�.");
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
		alert("���õ� ��ǰ�� �����ϴ�.");
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
		alert("���þ������� �����ϴ�.");
		return;
	}

	for (var i=0;i < document.forms.length; i++) {
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked == true) {
				selecteditemcount = selecteditemcount + 1;

				if (makerid == "") {
					makerid = frm.makerid.value.toUpperCase();
				} else if (makerid != frm.makerid.value.toUpperCase()) {
					alert("�ϳ��� �귣�常 ���ð����մϴ�.");
					return;
				}

				itemgubunarr = itemgubunarr + frm.itemgubun.value + "|";
				itemidarr = itemidarr + frm.itemid.value + "|";
				itemoptionarr = itemoptionarr + frm.itemoption.value + "|";
			}
		}
	}

	//if (selecteditemcount > 30) {
	//	alert("�ѹ��� 30���� �ʰ��Ͽ� ��ǰ�� ������ �� �����ϴ�.");
	//	return;
	//}

	if (makerid == "") {
		alert("���õ� ��ǰ�� �����ϴ�.");
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
		alert('�귣�带 ���� �ϼ���.');
		frmrefresh.makerid.focus();
	}

	if (confirm('����� ���� ��ü ���ΰ�ħ �Ͻðڽ��ϱ�?')){
		frmrefresh.mode.value="ipchulallrefreshbybrand";
		frmrefresh.submit();
	}
}

function DelItem(itemgubun,itemid,itemoption){
	if (confirm('����� ���� �� ���� �Ͻðڽ��ϱ�?')){
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
	var itemid = document.frm.itemidArr.value;
	 itemid =  itemid.replace(",","\r");    //�޸��� �ٹٲ�ó��
		 for(i=0;i<itemid.length;i++){
			if ( itemid.charCodeAt(i) != "13" && itemid.charCodeAt(i) != "10" && "0123456789".indexOf(itemid.charAt(i)) < 0){
					alert("��ǰ�ڵ�� ���ڸ� �Է°����մϴ�.");
					return;
			}
		}
	frm.action="";
	frm.target="";
    document.frm.submit();
}

function CheckThis(frm){
	frm.cksel.checked=true;
	AnCheckClick(frm.cksel);
}

/*
// �ε��� ���
function prtItemLabel(frm) {
	var ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset;
	var barcode, makerid, itemname, itemoptionname, customerprice, printno;
	var makername, itemid;

	// alert("�׽�Ʈ��!!");

	showdomainyn	= "Y";
	currencychar	= "��";
	ttptype			= "TTP-243_80x50";
	barcodetype		= "2";								// T or G or 2(�ٹ����� ���ڵ� or ������ڵ� or ���ٹ��ڵ�_������ڵ�)

	domainname		= "www.10x10.co.kr";
	showpriceyn		= "N";
	shopbrandyn		= "Y";
	papermargin		= 3;
	heightoffset	= 0;

	itemid			= frm.itemid.value;
	itemname		= frm.itemname.value;
	itemoptionname	= frm.itemoptionname.value;
	barcode			= frm.barcode2.value;
	makerid			= frm.makerid.value;
	printno			= frm.itemno.value;

	customerprice	= 0;
	makername		= "";

	// /js/barcode.js ����
	if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) != true) {
		alert("�����Ͱ� ��ġ���� �ʾҰų� �ùٸ� �����͸��� �ƴմϴ�.[4123]");
		return;
	}

	printTTPOneItemLabel(barcode, makerid, makername, itemid, itemname, itemoptionname, customerprice, printno);
}
*/

//�ε��� ��� �ϰ�
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
	barcodetype		= "2";								// T or G or 2(�ٹ����� ���ڵ� or ������ڵ� or ���ٹ��ڵ�_������ڵ�)
	var paperwidth = frm.paperwidth.value;
	var paperheight = frm.paperheight.value;
	var papermargin = frm.papermargin.value;
	var heightoffset = 0;
	showpriceyn = frm.printpriceyn.value;

	if (isforeignprint == "N") {
		currencychar = "��";
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

			//�ؿ� ��ǰ��
			if (isforeignprint == "Y") {
			//	itemname = document.getElementById("itemname_foreign_" + i).value.trim();
			//	itemoptionname = document.getElementById("itemoptionname_foreign_" + i).value.trim();
			//	customerprice = document.getElementById("customerprice_foreign_" + i).value.trim();

			//���� ��ǰ��
			} else {
				itemname = document.getElementById("itemname_" + i).value.trim();
				itemoptionname = document.getElementById("itemoptionname_" + i).value.trim();

				//���ΰ� ǥ��
				if (showpriceyn=='C'){
					//����
					if (saleyn=='Y'){
						customerprice = saleprice;

					//�Һ��ڰ�
					}else{
						customerprice = document.getElementById("customerprice_" + i).value.trim();
					}

				//�ǸŰ� ǥ��
				}else if (showpriceyn=='R'){
					customerprice = document.getElementById("sellprice_" + i).value.trim();

				//�Һ��ڰ� ǥ��
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
		alert("���õ� ��ǰ�� �����ϴ�.");
		return;
	}

	//TEC B-FV4		//2016.11.24 �ѿ�� ����
	if (frm.printername.value=='TEC_B-FV4_80x50'){
		//if (TEC_DO3.IsDriver != 1){ alert("TEC B-FV4 ����̹��� ��ġ�� �ּ���!!"); return; }
		if (confirm("���� ��ǰ�� �ε����� ����մϴ�.\n\nTEC B-FV4 �� ����Ͻðڽ��ϱ�?") == true) {
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

	// /js/barcode.js ����
	}else if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) == true) {
		if (confirm("���� ��ǰ�� �ε����� ����մϴ�.\n\nTTP-243 �� ����Ͻðڽ��ϱ�?") == true) {
			printTTPMultiItemLabel(arr);
		}

	}else {
	    alert("TTP-243(��)�� TEC B-FV4 ����̹��� ��ġ�� �ּ���");
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
                alert('AGV ������ Ȯ���ϼ���.');
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
		if (confirm("������ 1�� �̸��� ��ǰ�� ���ܵ˴ϴ�.\n\n���û�ǰ�� AGV�������̽��� ���� �Ͻðڽ��ϱ�?") == true) {
			frm.mode.value = "agvregarr";

            frm.itemgubunarr.value = itemgubunarr;
            frm.itemidarr.value = itemidarr;
            frm.itemoptionarr.value = itemoptionarr;
            frm.itemnoarr.value = itemnoarr;

			frm.action = "/admin/logics/logics_agv_pickup_process.asp";
			frm.submit();
		}
	} else {
		alert("���õ� ��ǰ�� �����ϴ�.");
	}
}

function regAGVCheckStockArr() {
	var frm = document.frmActPop;
	var found = false;
    var zerofound = false;
    var agvstock, chk;
    var itemgubunarr, itemidarr, itemoptionarr, itemnoarr;
    var itemgubun, itemid, itemoption;

    itemgubunarr = '';
    itemidarr = '';
    itemoptionarr = '';

    for (var i = 0; ; i++) {
        chk = document.getElementById('chk_' + i);
        itemgubun = document.getElementById('itemgubun_' + i);
        itemid = document.getElementById('itemid_' + i);
        itemoption = document.getElementById('itemoption_' + i);

        if (chk == undefined) { break; }

        if (chk.checked == true) {
            found = true;

            itemgubunarr = itemgubunarr + ',' + itemgubun.value;
            itemidarr = itemidarr + ',' + itemid.value;
            itemoptionarr = itemoptionarr + ',' + itemoption.value;
        }
    }

	if (found == true) {
		if (confirm("���û�ǰ�� AGV ������翡 ���� �Ͻðڽ��ϱ�?") == true) {
			frm.mode.value = "agvregarr";

            frm.itemgubunarr.value = itemgubunarr;
            frm.itemidarr.value = itemidarr;
            frm.itemoptionarr.value = itemoptionarr;

			frm.action = "/admin/logics/logics_agv_stockinvest_process.asp";
			frm.submit();
		}
	} else {
		alert("���õ� ��ǰ�� �����ϴ�.");
	}
}

String.prototype.trim = function() {
    return this.replace(/^\s+|\s+$/g, "");
};

function exceldownload(){
    var frm = document.frm;
	frm.action="/admin/stock/brandcurrentstock_excel.asp";
	frm.target="view";
	frm.submit();
	frm.action="";
	frm.target="";
}

function jsCurrStockDown(stockPlace,temp){
	if (stockPlace==""){
		alert('�����ġ�� �������� �ʾҽ��ϴ�.');
		return;
	}
	frm.stockPlace.value=stockPlace;
	frm.action="/admin/newreport/currentstock_excel.asp";
	frm.target = "view";
	frm.submit();
	frm.target = "";
	frm.action = ""
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<input type="hidden" name="stockPlace" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="6" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		<table border="0" cellpadding="5" cellspacing="0" class="a">
				<tr>
					<td>* �귣��: <% drawSelectBoxDesignerwithName "makerid", makerid %></td>
					<td>* ��ǰ��: <input type="text" class="text" name="itemname" value="<%= itemname %>" size="32" maxlength="32"></td>
					<td><!-- #include virtual="/common/module/dispCateSelectBox.asp"--></td>
					<td>* ��ǰ�ڵ�:</td>
					<td ><textarea rows="3" cols="10" name="itemidArr" id="itemidArr"><%=replace(itemidArr,",",chr(10))%></textarea> </td>

					<td >
						<input type=checkbox name="useoffinfo" <% if useoffinfo = "on" then response.write "checked" %> > ������ǰ(10) ����(OFF��ǰ �˻���)
					</td>

				</tr>

			</table>
	</td>
	<td rowspan="6" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="SubmitSearch();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
	    <!--
		<input type=radio name="onoffgubun" value="on" <% if onoffgubun="on" then response.write "checked" %> >ON��ǰ
		<input type=radio name="onoffgubun" value="off" <% if onoffgubun="off" then response.write "checked" %> >OFF��ǰ
		-->
		<!--
		<select name="onoffgubun" >
			<option value="on" <%= ChkIIF(onoffgubun="on","selected","") %> >ON��ǰ</option>
			<option value="off" <%= ChkIIF(onoffgubun="off","selected","") %> >OFF��ǰ</option>
			<option value="off55" <%= ChkIIF(onoffgubun="off55","selected","") %> >OFF��ǰ-55</option>
			<option value="off70" <%= ChkIIF(onoffgubun="off70","selected","") %> >OFF��ǰ-70</option>
			<option value="off75" <%= ChkIIF(onoffgubun="off75","selected","") %> >OFF��ǰ-75</option>
			<option value="off80" <%= ChkIIF(onoffgubun="off80","selected","") %> >OFF��ǰ-80</option>
			<option value="off85" <%= ChkIIF(onoffgubun="off85","selected","") %> >OFF��ǰ-85</option>
			<option value="off90" <%= ChkIIF(onoffgubun="off90","selected","") %> >OFF��ǰ-90</option>
		</select>
		-->
		<input type="hidden" name="onoffgubun" value="<%= onoffgubun %>">
		* ��ǰ����: <% drawSelectBoxItemGubunForSearch "itemgubun", itemgubun %>
		&nbsp;&nbsp;
		* �Ǹ�:<% drawSelectBoxSellYN "sellyn", sellyn %>
		&nbsp;&nbsp;
		* ���:<% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;&nbsp;
		* ����:<% drawSelectBoxDanjongYN "danjongyn", danjongyn %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		* �ŷ����� : <% drawSelectBoxMWU "mwdiv", mwdiv %>
		&nbsp;&nbsp;
		* ���͸��Ա��� :
		<select class="select" name="centermwdiv">
            <option value="">��ü</option>
            <option value="M" <% if centermwdiv="M" then response.write "selected" %> >����</option>
            <option value="W" <% if centermwdiv="W" then response.write "selected" %> >��Ź</option>
            <option value="N" <% if centermwdiv="N" then response.write "selected" %> >������</option>
        </select>
		&nbsp;&nbsp;
		* �������� : 
		<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",vPurchasetype,"" %>
	    &nbsp;&nbsp;
	    <span style="white-space:nowrap;">����:<% drawSelectBoxLimitYN "limityn", limityn %></span>
        &nbsp;&nbsp;
		* ��ǰ��� :
        <select class="select" name="itemgrade">
            <option value="">��ü</option>
            <option value="A" <% if itemgrade="A" then response.write "selected" %> >A</option>
            <option value="B" <% if itemgrade="B" then response.write "selected" %> >B</option>
            <option value="C" <% if itemgrade="C" then response.write "selected" %> >C</option>
            <option value="Z" <% if itemgrade="Z" then response.write "selected" %> >Z</option>
            <option value="AB" <% if itemgrade="AB" then response.write "selected" %> >A+B</option>
            <option value="ABC" <% if itemgrade="ABC" then response.write "selected" %> >A+B+C</option>
        </select>
        &nbsp;&nbsp;
		* ��ũ��� :
        <select class="select" name="bulkstockgubun">
            <option value="">��ü</option>
            <option value="nul" <% if bulkstockgubun="nul" then response.write "selected" %> >�Է�����</option>
            <option value="err" <% if bulkstockgubun="err" then response.write "selected" %> >��ũ���� ����</option>
        </select>
        <br>
		* �������� :
        <select class="select" name="warehouseCd">
            <option value="">��ü</option>
            <option value="AGV" <% if warehouseCd="AGV" then response.write "selected" %> >AGV</option>
            <option value="BLK" <% if warehouseCd="BLK" then response.write "selected" %> >��ũ</option>
        </select>
        &nbsp;&nbsp;
		* AGV��� :
        <select class="select" name="agvstockgubun">
            <option value="">��ü</option>
            <option value="availdiff" <% if agvstockgubun="availdiff" then response.write "selected" %> >��ȿ��� ����ġ��</option>
            <option value="ipkum5diff" <% if agvstockgubun="ipkum5diff" then response.write "selected" %> >����ľ���� ����ġ��</option>
            <option value="oneup" <% if agvstockgubun="oneup" then response.write "selected" %> >1�̻�</option>
            <option value="zero" <% if agvstockgubun="zero" then response.write "selected" %> >0</option>
            <option value="minus" <% if agvstockgubun="minus" then response.write "selected" %> >���̳ʽ�</option>
        </select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
	    * ���������� :
	    <input type=radio name="returnitemgubun" value="" <% if returnitemgubun="" then response.write "checked" %> onClick="chkEnDisabled(this);">��ü
		<input type=radio name="returnitemgubun" value="rackdisp" <% if returnitemgubun="rackdisp" then response.write "checked" %> onClick="chkEnDisabled(this);">������ ��ǰ [(�Ǹ�<>'N') or (�����ƴ�)]
		<input type=radio name="returnitemgubun" value="reton" <% if returnitemgubun="reton" then response.write "checked" %> onClick="chkEnDisabled(this);">��ǰ��� ��ǰ [(�Ǹ�='N') and (����) and (�ǻ���ȿ���<>0)]
	    <input type=radio name="returnitemgubun" value="retfin" <% if returnitemgubun="retfin" then response.write "checked" %> onClick="chkEnDisabled(this);">��ǰ�Ϸ� ��ǰ [(�Ǹ�='N') and (����) and (�ǻ���ȿ���=0)]
	    <script language='javascript'>chkEnDisabled(frm.returnitemgubun[<%= ChkIIF(returnitemgubun="","0",ChkIIF(returnitemgubun="rackdisp","1","2")) %>]);</script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
	    * <select name="stocktype" class="select">
			<option value="sys" <% if (stocktype = "sys") then %>selected<% end if %> >�ý������</option>
			<option value="real" <% if (stocktype = "real") then %>selected<% end if %> >��ȿ���</option>
		</select>
		: <% drawSelectBoxexistsstock "limitrealstock", limitrealstock, "" %>
		&nbsp;&nbsp;
		* ������ :
		<input type="text" class="text" name="startMon" size="2" value="<%= startMon %>">
		~
		<input type="text" class="text" name="endMon" size="2" value="<%= endMon %>"> ����
		&nbsp;&nbsp;
		* ���ļ��� :
		<select class="select" name="ordby">
			<option value="1" <%= CHKIIF(ordby = "1", "selected", "") %> >��ǰ�ڵ�</option>
			<option value="2" <%= CHKIIF(ordby = "2", "selected", "") %> >���ڵ�</option>
		</select>
		&nbsp;&nbsp;
		<input type="checkbox" class="checkbox" name="excits" value="Y" <%= CHKIIF(excits="Y", "checked", "") %> > 3PL ����
        &nbsp;&nbsp;
		* ���ڵ� :
		<input type="text" class="text" name="itemrackcode" size="8" value="<%= itemrackcode %>">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		�� ������ ����
        &nbsp;
        |
        &nbsp;
		* �����ͼ��� :
		<select name="printername" onchange="reg('');">
			<option value="TEC_B-FV4_80x50" <% if printername = "TEC_B-FV4_80x50" then response.write " selected" %>>TEC B-FV4 (�԰�80x50)</option>
			<option value="TTP-243_80x50" <% if printername = "TTP-243_80x50" then response.write " selected" %>>TTP-243 (�԰�80x50)</option>
		</select>
		&nbsp;
		* ǥ�û�ǰ�� :
		<select name="isforeignprint" onchange="reg('');">
			<option value="N" <% if (isforeignprint = "N") then %>selected<% end if %>>������ǰ��</option>
			<!--<option value="Y" <% if (isforeignprint = "Y") then %>selected<% end if %>>�ؿܻ�ǰ��</option>-->
		</select>
		&nbsp;
		* �ݾ�ǥ�ù�� :
		<select name="printpriceyn" onchange="reg('');">
			<option value="Y" <% if (printpriceyn = "Y") then %>selected<% end if %>>�Һ��ڰ�ǥ��</option>
			<!--<option value="C" <% if (printpriceyn = "C") then %>selected<% end if %>>���ΰ�ǥ��</option>-->
			<option value="R" <% if (printpriceyn = "R") then %>selected<% end if %>>�ǸŰ�ǥ��</option>
			<option value="S" <% if (printpriceyn = "S") then %>selected<% end if %>>���ñݾ�ǥ��</option>
			<option value="N" <% if (printpriceyn = "N") then %>selected<% end if %>>�ݾ�ǥ�þ���</option>
		</select>
		&nbsp;
		* ����ǥ�� :
		<select name="titledispyn" onchange="reg('');">
			<option value="Y" <% if (titledispyn = "Y") then %>selected<% end if %>>����ǥ��</option>
			<option value="N" <% if (titledispyn = "N") then %>selected<% end if %>>����ǥ�þ���</option>
		</select>

        <!--* ���ڵ� ���� �԰� --->
		<% if printername = "TTP-243_80x50" then %>
			<input type="hidden" name="paperwidth" value="80" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="50" size="4" maxlength=9>
		<% elseif printername = "TEC_B-FV4_80x50" then %>
			<input type="hidden" name="paperwidth" value="800" size="4" maxlength=9>
			<input type="hidden" name="paperheight" value="500" size="4" maxlength=9>
		<% end if %>

		<input type="hidden" name="papermargin" value="3" size="4" maxlength=9>
	</td>
</tr>
</table>
<!-- �˻� �� -->

<br>

* �������԰���� ���������� �������԰���Դϴ�.(�̹��� �԰����� ����)<br />
* ��ũ���� <font color="red">�ǽð���� �ƴմϴ�</font>. ������� �뵵�θ� ��밡���մϴ�. <font color="red">����Է� �� 5��</font> �� ������ ǥ�õ��� �ʽ��ϴ�.

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<% if C_ADMIN_AUTH=true then %>
    	<!--
        <input type="button" class="button" value="�������ü ���ΰ�ħ" onclick="RefreshIpchulStock();">
        -->
        <% end if %>

		<input type="button" class="button" name="stock_index_print" value="���û�ǰ ��ǰ���" onclick="PopReIpgo();">
        &nbsp;
		<input type="button" class="button" name="stock_index_print" value="���û�ǰ [���]���" onclick="PopChulgo();">
        &nbsp;
		<input type="button" class="button" name="stock_index_print" value="���û�ǰ ��ǰ ���ڵ����" onclick="PopModiRackCode('modiitem');">
        &nbsp;
		<input type="button" class="button" name="stock_index_print" value="���û�ǰ [�ɼǺ�] ���ڵ����" onclick="PopModiRackCode('modiopt');">
        &nbsp;
		<input type="button" class="button" name="setsell2y" value="���û�ǰ �Ǹ�(Y) ����" onclick="jsSetSellY();">
        <br>
		<input type="button" class="button" value=" ���û�ǰAGV �������̽� ����" onclick="regAGVArr();">
        &nbsp;
		<input type="button" class="button" value=" ���û�ǰAGV ������� �ۼ�" onclick="regAGVCheckStockArr();">
        &nbsp;
		<input type="button" class="button" name="jssetbulkstockno" value="���� ��ũ��� ����" onclick="jsSetBulkStockNo();">
        &nbsp;
		<input type="button" class="button" name="jssetbulkstockerrno" value="���� ��ũ[����] ����" onclick="jsSetBulkStockErrNo();">
	</td>
	<td align="right">
    	<input type="button" class="button" value="�����ٿ�ε�" onclick="exceldownload();">
		<input type="button" class="button" name="stock_sheet_print" value="����ľ� SHEET���" onclick="PopBrandStockSheet();" <%= ChkIIF(IsSheetPrintEnable,"","disabled") %> >
		<input type="button" class="button" name="stock_index_print" value="���û�ǰ �ε������" onclick="IndexBarcodePrint(); return false;">
		* ǥ�ð���:
		<select class="select" name="pagesize" >
			<option value="25">25��</option>
			<option value="100" <%= ChkIIF(pagesize="100","selected","") %> >100��</option>
			<option value="200" <%= ChkIIF(pagesize="200","selected","") %> >200��</option>
			<option value="300" <%= ChkIIF(pagesize="300","selected","") %> >300��</option>
            <!--<option value="500" <%'= ChkIIF(pagesize="500","selected","") %> >500��</option>-->
		</select>
		<br><br>* ���ǻ� :
		<!--
		<br><br><input type="checkbox" name="day1after">�������ĺ���������
		<input type="button" class="button" value="���ǻ�ٿ�ε�(����)" onclick="jsstockDown('L','');">
		-->
		<input type="button" class="button" value="�������ٿ�ε�(����)" onclick="jsCurrStockDown('L','');">
	</td>
</tr>
</table>
</form>
<!-- �׼� �� -->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="37">
		�˻���� : <b><%= osummarystockbrand.FTotalCount %></b>
		&nbsp;
		������ :
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
    <td>���ڵ�</td>
    <td>����</td>
	<td>��ǰ�ڵ�</td>
	<td>�ɼ�<br />�ڵ�</td>
	<td>�̹���</td>
	<td>�귣��ID</td>
	<td>��ǰ��<br>[�ɼǸ�]</td>
	<td>�Һ��ڰ�</td>
	<td>���԰�(��)</td>
	<td>����<br>����</td>
	<td>����<br>����<br>����</td>
	<td>��<br>�԰�<br>��ǰ</td>
	<td>ON��<br>�Ǹ�<br>��ǰ</td>
    <td>OFF��<br>���<br>��ǰ</td>
    <td>��Ÿ<br>���<br>��ǰ</td>
    <td>CS<br>���<br>��ǰ</td>
    <td bgcolor="F4F4F4"><b>�ý���<br>�����</b></td>

	<td>��<br>�ǻ�<br>����</td>
	<td>�ǻ�<br>���</td>
	<td>��<br>�ҷ�</td>
	<td bgcolor="F4F4F4"><b>��ȿ<br>���</b></td>

    <td>��<br>��ǰ<br>�غ�</td>
    <td bgcolor="F4F4F4"><b>���<br>�ľ�<br>���</b></td>
    <td>����<br>����<br>�ֹ�</td>
    <td bgcolor="F4F4F4">����<br>���</td>
    <td width="30">�Ǹ�<br>����</td>
    <td width="30">����<br>����</td>
    <td>����<br>����</td>
    <td width="40">����<br>�Է�</td>
	<td width="40">��¼�</td>
	<td width="60">������<br>�԰��</td>
	<td width="40">����<br />�Ǹ�<br />(����)</td>
    <td width="35">��ǰ<br />���</td>
    <td>��ũ<br />�ǻ�</td>
    <td>��ũ<br />���</td>
    <td>AGV<br />���</td>
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
		<input type="button" class="button" value="����" onclick="popRealErrInput('<%= osummarystockbrand.FItemList(i).Fitemgubun %>','<%= osummarystockbrand.FItemList(i).Fitemid %>','<%= osummarystockbrand.FItemList(i).Fitemoption %>');">
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
        <%
        bulkrealstock = NULL
        if Not IsNull(osummarystockbrand.FItemList(i).Fbulkstock) and osummarystockbrand.FItemList(i).Fbulkstock <> "" and IsNumeric(osummarystockbrand.FItemList(i).Fbulkstock) then
            bulkrealstock = osummarystockbrand.FItemList(i).Fbulkstock + osummarystockbrand.FItemList(i).Fagvstock
        end if
        %>
		<input type="text" class="text_ro" id="bulkrealstock_<%= i %>" name="bulkrealstock" value="<%= bulkrealstock %>" size="1" maxlength="8" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)" style="text-align:center;" readOnly>
	</td>
    <td>
		<input type="text" class="text" id="bulkstock_<%= i %>" name="bulkstock" value="<%= osummarystockbrand.FItemList(i).Fbulkstock %>" size="1" maxlength="8" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)" style="text-align:center;">
	</td>
    <td>
		<input type="text" class="text" id="agvstock_<%= i %>" name="agvstock" value="<%= osummarystockbrand.FItemList(i).Fagvstock %>" size="1" maxlength="8" onKeyPress="CheckThis(frmBuyPrc_<%= i %>)" style="text-align:center;">
	</td>
</tr>
</form>
<%
	if i mod 500 = 0 then
		Response.Flush		' ���۸��÷���
	end if
next
%>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="37" align="center">
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
        <td colspan="37" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
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
