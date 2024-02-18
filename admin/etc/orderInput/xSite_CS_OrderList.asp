<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� Ŭ����
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteCSOrderCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
Dim sellsite, currstate, research, page, orderserial, outmallorderserial, apiCS, divcd, pgsize
Dim i
dim checkYYYYMMDD
dim yyyy1, mm1, dd1, yyyy2, mm2, dd2, excnoorder, makerid, ordBy
dim csregyn
	sellsite = requestCheckvar(request("sellsite"),32)
	currstate = requestCheckvar(request("currstate"),10)
	research = requestCheckvar(request("research"),10)
	page = requestCheckvar(request("page"),10)
	orderserial = requestCheckvar(request("orderserial"),20)
	outmallorderserial = requestCheckvar(request("outmallorderserial"),30)

	checkYYYYMMDD = requestCheckvar(request("checkYYYYMMDD"), 1)
	yyyy1 = requestCheckvar(request("yyyy1"),30)
	mm1 = requestCheckvar(request("mm1"),30)
	dd1 = requestCheckvar(request("dd1"),30)
	yyyy2 = requestCheckvar(request("yyyy2"),30)
	mm2 = requestCheckvar(request("mm2"),30)
	dd2 = requestCheckvar(request("dd2"),30)

	divcd = requestCheckvar(request("divcd"),30)
	pgsize = requestCheckvar(request("pgsize"),30)
	excnoorder = requestCheckvar(request("excnoorder"),30)
	makerid = requestCheckvar(request("makerid"),32)
	ordBy = requestCheckvar(request("ordBy"),32)
    csregyn = requestCheckvar(request("csregyn"),32)

if (research="") then checkYYYYMMDD="Y"
if (research="") then currstate="B001"
if (research="") then excnoorder="Y"
if (page="") then page=1
if Not IsNumeric(pgsize) then pgsize=20

'==============================================================================
dim nowdate, searchnextdate

''�⺻ N��. ����Ʈ üũ
if (yyyy1="") then
    nowdate = Left(CStr(dateadd("m",-2,now())),10)
	yyyy1   = Left(nowdate,4)
	mm1     = Mid(nowdate,6,2)
	dd1     = Mid(nowdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2   = Left(nowdate,4)
	mm2     = Mid(nowdate,6,2)
	dd2     = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2,mm2,dd2),1)),10)


Dim oCxSiteCSOrder
set oCxSiteCSOrder = new CxSiteCSOrder
	oCxSiteCSOrder.FPageSize = pgsize
	oCxSiteCSOrder.FCurrPage = page
	oCxSiteCSOrder.FRectSellSite   = sellsite

	if (checkYYYYMMDD="Y") and (orderserial = "") and (outmallorderserial = "") then
		'// �ֹ���ȣ ������ �Ⱓ �˻����� ����
		oCxSiteCSOrder.FRectStartDate = Left(CStr(DateSerial(yyyy1,mm1,dd1)),10)
		oCxSiteCSOrder.FRectEndDate = searchnextdate
	end if

	if (outmallorderserial = "") then
		oCxSiteCSOrder.FRectCurrState = currstate
	end if

	oCxSiteCSOrder.FRectOrderSerial = orderserial
	oCxSiteCSOrder.FRectoutmallorderserial = outmallorderserial
	oCxSiteCSOrder.FRectDivCD = divcd
	oCxSiteCSOrder.FRectExcNoOrder = excnoorder
	oCxSiteCSOrder.FRectMakerid = makerid
	oCxSiteCSOrder.FRectOrderBy = ordBy
    oCxSiteCSOrder.FRectCsRegYN = csregyn

    oCxSiteCSOrder.getCSMasterList

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function apiCSProcess(){
	var v = document.getElementById("apiCS").value;
	var cjmodeName;
	if (v=="cjmallCsreg1" || v=="cjmallCsreg2" || v=="cjmallCsreg3" ){
		switch(v){
			case 'cjmallCsreg1' : cjmodeName = 'CJMall_��ǰ'; break;
			case 'cjmallCsreg2' : cjmodeName = 'CJMall_���'; break;
			case 'cjmallCsreg3' : cjmodeName = 'CJMall_CS���,������'; break;
		}
		if (confirm(""+cjmodeName+"�� CS ���� ��� �Ͻðڽ��ϱ�?")){
			GetxSiteCSOrderList_CJ(v);
	    }
	}else{
	    if (confirm(""+v+"���� CS ���� ��� �Ͻðڽ��ϱ�?")){
			GetxSiteCSOrderNewList(v);
	    }
	}
}

function TenCSProcess() {
	var popwin=window.open('','TenCSProcess','width=300,height=200');
    var sellsite = document.getElementById("apiCS").value;
	popwin.focus();

	var frm = document.frmWapi;

    if (sellsite == "ssg") {
        frm.action = "<%=apiURL%>/outmall/ssg/xSiteCsOrder_ssg_Process.asp?mode=chkBatchMatchCS";
    } else if (sellsite == "coupang") {
        frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=coupang&mode=matchcs";
    } else {
        alert('ó�������� ���޸��� �ƴմϴ�.');
        return;
    }

	frm.target = "TenCSProcess";
	frm.submit();
}

function ExtCSProcess() {
	var popwin=window.open('','ExtCSProcess','width=300,height=200');
    var sellsite = document.getElementById("apiCS").value;
	popwin.focus();

	var frm = document.frmWapi;

    if (sellsite == "ssg") {
        frm.action = "<%=apiURL%>/outmall/ssg/xSiteCsOrder_ssg_Process.asp?mode=chkBatchExtCsState";
    } else if (sellsite == "coupang") {
        frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=coupang&mode=chkextcs";
    } else {
        alert('ó�������� ���޸��� �ƴմϴ�.');
        return;
    }

	frm.target = "ExtCSProcess";
	frm.submit();
}

function jsExtCheckCs(sellsite, divcd, outmallorderserial) {
	var popwin=window.open('','jsExtCheckCs','width=300,height=200');
	popwin.focus();

	var frm = document.frmWapi;

    if (sellsite == "ssg") {
        frm.action = "<%=apiURL%>/outmall/ssg/xSiteCsOrder_ssg_Process.asp?mode=chkExtCsStateOne&outMallorderSerial=" + outmallorderserial + "&divcd=" + divcd;
    }

	frm.target = "jsExtCheckCs";
	frm.submit();
}

function GetxSiteCSOrderNewList(sellsite){
	var popwin=window.open('','xSiteCSOrderNewList','width=1000,height=1000,left=100,top=100');
	popwin.focus();

	var frm = document.frmWapi;
	frm.mode.value = "getxsitecslist";
	frm.sellsite.value = sellsite;

	if (sellsite=="ezwel") {
		frm.action = "/admin/etc/order/xSiteCSOrder_Ins_Process.asp?sellsite=ezwel&mode=all"
	}else if(sellsite=="kakaostore"){
		frm.action = "/admin/etc/order/xSiteCSOrder_Ins_Process.asp?sellsite=kakaostore&mode=all"
	}else if(sellsite=="lotteCom"){
		frm.action = "<%=apiURL%>/outmall/LotteCom/xSiteCSOrder_Process_lotteCom.asp"
	}else if(sellsite=="lotteimall"){
		frm.action = "<%=apiURL%>/outmall/ltimall/xSiteCSOrder_lotteimall_Process.asp"
	}else if(sellsite=="auction1010"){
		frm.action = "<%=apiURL%>/outmall/auction/xSiteCSOrder_auction_Process.asp"
	}else if(sellsite=="ssg"){
		frm.action = "<%=apiURL%>/outmall/ssg/xSiteCsOrder_ssg_Process.asp"
	}else if(sellsite=="shintvshopping"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=shintvshopping&mode=all"
	}else if(sellsite=="wetoo1300k"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=wetoo1300k&mode=all"
	}else if(sellsite=="gmarket1010"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=gmarket1010&mode=all"
	}else if(sellsite=="interpark"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=interpark&mode=all"
	}else if(sellsite=="nvstorefarm"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=nvstorefarm&mode=all"
	}else if(sellsite=="nvstoremoonbangu"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=nvstoremoonbangu&mode=all"
	}else if(sellsite=="Mylittlewhoopee"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=Mylittlewhoopee&mode=all"
	}else if(sellsite=="nvstoregift"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=nvstoregift&mode=all"
	/* GSShop ���� */
	}else if(sellsite=="gseshop"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=gseshop&mode=ordercancel"
	}else if(sellsite=="gseshop2"){
		// ��ǰ,��ȯ
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=gseshop"
	}else if(sellsite=="gseshopCancel"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=gseshop&mode=orderNewcancel"
	}else if(sellsite=="gseshopExcRet"){
		// ��ǰ,��ȯ
		frm.action = "<%=apiURL%>/outmall/order/xSiteOrder_Ins_Process.asp?sellsite=gseshopNew"
	/***********************/
	}else if(sellsite=="halfclub"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=halfclub&mode=all"
	}else if(sellsite=="coupang"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=coupang&mode=all"
	}else if(sellsite=="hmall1010"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=hmall1010&mode=all"
	}else if(sellsite=="11st1010"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=11st1010&mode=all"
	}else if(sellsite=="WMP"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=WMP&mode=all"
	}else if(sellsite=="wmpfashion"){
		frm.action = "<%=apiURL%>/outmall/order/xSite_CS_Order_Ins_Process.asp?sellsite=wmpfashion&mode=all"
	}
	frm.target = "xSiteCSOrderNewList";
	frm.submit();
}

//wapi
function GetxSiteCSOrderList_CJ(mode) {
	var frm = document.frmTmp;
    var popwin=window.open('','xSiteCSOrderList_Cj','width=1000,height=1000,left=100,top=100');
    popwin.focus();

	frm.cmdparam.value = mode;
	frm.action = "<%=apiURL%>/outmall/cjmall/xSiteCSOrder_Cjmall_Process.asp"
	frm.target = "xSiteCSOrderList_Cj";
	frm.submit();
}

/*
function GetxSiteCSOrderList_lotteCom(sellsite){
    if (confirm("���� �Ͻðڽ��ϱ�?") != true) {
		return;
	}

    var popwin=window.open('','xSiteCSOrderList_lotteCom','width=100,height=100,left=100,top=100');
    popwin.focus();

    var frm = document.frmWapi;
    frm.mode.value = "getxsitecslist";
    frm.sellsite.value = sellsite;
    frm.action = "http://wapi.10x10.co.kr/outmall/LotteCom/xSiteCSOrder_Process_lotteCom.asp"
    frm.target = "xSiteCSOrderList_lotteCom";
	frm.submit();
}

function GetxSiteCSOrderList_lotteimall(sellsite){
    if (confirm("���� �Ͻðڽ��ϱ�?") != true) {
		return;
	}

    var popwin=window.open('','xSiteCSOrderList_lotteimall','width=100,height=100,left=100,top=100');
    popwin.focus();

    var frm = document.frmWapi;
    frm.mode.value = "getxsitecslist";
    frm.sellsite.value = sellsite;
    frm.action = "http://wapi.10x10.co.kr/outmall/ltimall/xSiteCSOrder_lotteimall_Process.asp"
    frm.target = "xSiteCSOrderList_lotteimall";
	frm.submit();
}

function GetxSiteCSOrderList_ezwel(sellsite){
	if (confirm("���� �Ͻðڽ��ϱ�?") != true) {
		return;
	}

	var popwin=window.open('','xSiteCSOrderList_ezwel','width=1000,height=1000,left=100,top=100');
	popwin.focus();

	var frm = document.frmWapi;
	frm.mode.value = "getxsitecslist";
	frm.sellsite.value = sellsite;
	frm.action = "http://wapi.10x10.co.kr/outmall/ezwel/xSiteCSOrder_ezwel_Process.asp"
	frm.target = "xSiteCSOrderList_ezwel";
	frm.submit();
}
function GetxSiteCSOrderList_nvstorefarm(sellsite){
	if (confirm("���� �Ͻðڽ��ϱ�?") != true) {
		return;
	}

	var popwin=window.open('','xSiteCSOrderList_nvstorefarm','width=1000,height=1000,left=100,top=100');
	popwin.focus();

	var frm = document.frmWapi;
	frm.mode.value = "getxsitecslist";
	frm.sellsite.value = sellsite;
	frm.action = "http://wapi.10x10.co.kr/outmall/nvstorefarm/xSiteCSOrder_nvstorefarm_Process.asp"
	frm.target = "xSiteCSOrderList_nvstorefarm";
	frm.submit();
}
*/
function GetxSiteCSOrderList(sellsite) {
	var frm = document.frmAct;

	if (confirm("���� �Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.mode.value = "getxsitecslist";
	if (sellsite == "lotteimall") {
		frm.action = "xSiteCSOrder_lotteimall_Process.asp";
	}
	frm.sellsite.value = sellsite;
	frm.submit();
}



function GetxSiteCSOrderListCJ(mode) {
	var frm = document.frmTmp;

	if (confirm("���� �Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.cmdparam.value = mode;
	frm.submit();
}


function jsSearchByOutMallOrderSerial(outmallorderserial) {
	var frm = document.frm;
	frm.outmallorderserial.value = outmallorderserial;
	frm.submit();
}

function jsSearchByOrderSerial(orderserial) {
	var frm = document.frm;
	frm.orderserial.value = orderserial;
	frm.submit();
}

function Cscenter_Action_List(orderserial) {
    var window_width = 1280;
    var window_height = 960;

    var popwin = window.open("/cscenter/action/cs_action.asp?orderserial=" + orderserial ,"Cscenter_Action_List","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsSetFinishOne(idx) {
	var frm = document.frmAct;

	<% if (outmallorderserial = "") then %>
		alert("���� �����ֹ���ȣ�� �˻� ��\n\n�ٸ� CS���� ������ Ȯ�� �� �Ϸ�ó�� �ϼ���.");
		return;
	<% end if %>

	if (confirm("�Ϸ�ó�� �Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.mode.value = "setfinish";
	frm.idx.value = idx;
	frm.submit();
}

function jsDelFinishOne(idx) {
	var frm = document.frmAct;

	if (confirm("�Ϸ�ó�� ��� �Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.mode.value = "delfinish";
	frm.idx.value = idx;
	frm.submit();
}

function jsSetJupsuOne(idx) {
	var frm = document.frmAct;

	if (confirm("����ó�� �Ͻðڽ��ϱ�?") != true) {
		return;
	}

	frm.mode.value = "setjupsu";
	frm.idx.value = idx;
	frm.submit();
}

function jsSetToday() {
	var frm = document.frm;
	var now = new Date();
	var yyyy = now.getFullYear();
	var mm = now.getMonth() + 1;
	var dd = now.getDate();

	frm.yyyy1.value = yyyy;
	frm.mm1.value = mm < 10 ? '0' + mm : mm;
	frm.dd1.value = dd < 10 ? '0' + dd : dd;

	frm.yyyy2.value = yyyy;
	frm.mm2.value = mm < 10 ? '0' + mm : mm;
	frm.dd2.value = dd < 10 ? '0' + dd : dd;

	frm.checkYYYYMMDD.checked = true;
}

function jsSetTwoMonth() {
	var frm = document.frm;
	var now = new Date();
	var yyyy2 = now.getFullYear();
	var mm2 = now.getMonth() + 1;
	var dd2 = now.getDate();

	var twomonth = new Date(yyyy2, mm2 - 3, dd2);
	var yyyy1 = twomonth.getFullYear();
	var mm1 = twomonth.getMonth() + 1;
	var dd1 = twomonth.getDate();

	frm.yyyy1.value = yyyy1;
	frm.mm1.value = mm1 < 10 ? '0' + mm1 : mm1;
	frm.dd1.value = dd1 < 10 ? '0' + dd1 : dd1;

	frm.yyyy2.value = yyyy2;
	frm.mm2.value = mm2 < 10 ? '0' + mm2 : mm2;
	frm.dd2.value = dd2 < 10 ? '0' + dd2 : dd2;

	frm.checkYYYYMMDD.checked = true;
}

function fnCheckValidAll(bool, comp){
    var frm = comp.form;

    if (!comp.length){
        if (comp.disabled==false){
            comp.checked = bool;
            AnCheckClick(comp);
        }
    }else{
        for (var i=0;i<comp.length;i++){
            if (comp[i].disabled==false){
                comp[i].checked = bool;
                AnCheckClick(comp[i]);
            }
        }
    }
}

function CheckProduct(o) {
	var frm;
	if (o.checked) {
		hL(o);
	} else {
		dL(o);
	}
}

function jsSetFinish(frm) {
    var checkedExists = false;
    if (!frm.cksel.length){
        if (frm.cksel.checked){
            checkedExists = true;
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                checkedExists = true;
                break;
            }
        }
    }

    if (!checkedExists){
        alert('���� ������ �����ϴ�.');
        return;
    }

    if (confirm('�Ϸ�ó�� �����Ͻðڽ��ϱ�?')){
        frm.mode.value="setfinisharr";
        frm.submit();
    }
}

function jsSetJupsu(frm) {
    var checkedExists = false;
    if (!frm.cksel.length){
        if (frm.cksel.checked){
            checkedExists = true;
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                checkedExists = true;
                break;
            }
        }
    }

    if (!checkedExists){
        alert('���� ������ �����ϴ�.');
        return;
    }

    if (confirm('����ó�� �����Ͻðڽ��ϱ�?')){
        frm.mode.value="setjupsuarr";
        frm.submit();
    }
}

function jsDelFinish(frm) {
    var checkedExists = false;
    if (!frm.cksel.length){
        if (frm.cksel.checked){
            checkedExists = true;
        }
    }else{
        for (var i=0;i<frm.cksel.length;i++){
            if (frm.cksel[i].checked){
                checkedExists = true;
                break;
            }
        }
    }

    if (!checkedExists){
        alert('���� ������ �����ϴ�.');
        return;
    }

    if (confirm('������ �����Ͻðڽ��ϱ�?')){
        frm.mode.value="setdelfinisharr";
        frm.submit();
    }
}

function jsCheckCs(sellsite, outmallorderserial) {
	var popwin=window.open('','jsCheckCs','width=300,height=200');
	popwin.focus();

	var frm = document.frmWapi;

    if (sellsite == "ssg") {
        frm.action = "<%=apiURL%>/outmall/ssg/xSiteCsOrder_ssg_Process.asp?mode=chkMatchCS&outMallorderSerial=" + outmallorderserial;
    }

	frm.target = "jsCheckCs";
	frm.submit();
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
	    * ���θ� ���� :
	    <% call drawSelectBoxXSiteOrderInputPartnerCS("sellsite", sellsite) %>
		&nbsp;&nbsp;
		����:
		<select class="select" name="divcd">
			<option value="">��ü</option>
			<option value="A008" <% if (divcd = "A008") then response.write "selected" end if %>>�ֹ����</option>
			<option value="A000" <% if (divcd = "A000") then response.write "selected" end if %>>��ȯ���</option>
			<option value="A004" <% if (divcd = "A004") then response.write "selected" end if %>>��ǰ����</option>
			<option value="A009" <% if (divcd = "A009") then response.write "selected" end if %>>��Ÿ����(�޸�)</option>
			<option value="A011" <% if (divcd = "A011") then response.write "selected" end if %>>��ȯȸ��</option>
			<option value="A088" <% if (divcd = "A088") then response.write "selected" end if %>>�ֹ���� öȸ</option>
			<option value="A044" <% if (divcd = "A044") then response.write "selected" end if %>>��ǰ öȸ</option>
			<option value="A090" <% if (divcd = "A090") then response.write "selected" end if %>>��ȯ öȸ</option>
		</select>
	    &nbsp;&nbsp;
	    * ó������ :
		<select class="select" name="currstate"  >
			<option value="" <%= chkIIF(currstate="", "selected","") %> >��ü</option>
	     	<option value="B001" <%= chkIIF(currstate="B001","selected","") %> >�������</option>
			<option value="B002" <%= chkIIF(currstate="B002","selected","") %> >�����Ϸ�</option>
	     	<option value="B007" <%= chkIIF(currstate="B007","selected","") %> >��ϿϷ�</option>
     	</select>
     	&nbsp;&nbsp;
     	* �ֹ���ȣ:<input type="text" name="orderserial" value="<%=orderserial%>" size="14" maxlength="11"  >
     	&nbsp;&nbsp;
     	* �����ֹ���ȣ:<input type="text" name="outmallorderserial" value="<%= outmallorderserial %>" size="20" maxlength="20" >
		&nbsp;&nbsp;
     	* �귣��:<input type="text" name="makerid" value="<%= makerid %>" size="20" maxlength="32" >
	</td>
	<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="25">
    	<input type="checkbox" name="checkYYYYMMDD" value="Y" <% if checkYYYYMMDD="Y" then response.write "checked" %>>
    	������ : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		<input type="button" class="button" value="����" onClick="jsSetToday()" style="width:80px;">
		<input type="button" class="button" value="�ֱ�2��" onClick="jsSetTwoMonth()" style="width:80px;">
		&nbsp;
		ǥ�ð��� :
		<select class="select" name="pgsize">
			<option value=""></option>
			<option value="20" <%= CHKIIF(pgsize="20", "selected", "") %> >20</option>
			<option value="50" <%= CHKIIF(pgsize="50", "selected", "") %> >50</option>
			<option value="100" <%= CHKIIF(pgsize="100", "selected", "") %> >100</option>
		</select>
		&nbsp;
		<input type="checkbox" name="excnoorder" value="Y" <%= CHKIIF(excnoorder="Y", "checked", "") %> > �ֹ���ȣ���� ����
		&nbsp;
		���ļ��� :
		<select class="select" name="ordBy">
			<option value="1" <%= CHKIIF(ordBy="1", "selected", "") %>>������(SCM)</option>
			<option value="2" <%= CHKIIF(ordBy="2", "selected", "") %>>������(����)</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="25">
	    * CS��ϻ��� :
		<select class="select" name="csregyn"  >
			<option value="" <%= chkIIF(csregyn="", "selected","") %> >��ü</option>
	     	<option value="N" <%= chkIIF(csregyn="N","selected","") %> >�������</option>
			<option value="R" <%= chkIIF(csregyn="R","selected","") %> >����</option>
            <option value="Y" <%= chkIIF(csregyn="Y","selected","") %> >ó���Ϸ�</option>
            <option value="A" <%= chkIIF(csregyn="A","selected","") %> >���޿Ϸ�</option>
     	</select>
    </td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" value="�Ե����� CS���� ��������" onClick="GetxSiteCSOrderList('lotteCom');" disabled>
		&nbsp;
		<input type="button" value="�Ե�i�� CS���� ��������" onClick="GetxSiteCSOrderList('lotteimall');" disabled>
		&nbsp;
		<input type="button" value="CJMall CS���� ��������(��ǰ)" onClick="GetxSiteCSOrderListCJ('cjmallCsreg1');" disabled>
		&nbsp;
		<input type="button" value="CJMall CS���� ��������(���)" onClick="GetxSiteCSOrderListCJ('cjmallCsreg2');" disabled>
		&nbsp;
		<input type="button" value="CJMall CS���� ��������(CS���,������)" onClick="GetxSiteCSOrderListCJ('cjmallCsreg3');" disabled>
	</td>
	<td align="right">
	</td>
</tr>
<tr>
	<td align="left">
		�귣�� : <%= oCxSiteCSOrder.FResultStr %><br />
	    CS ���� �������� API�� �����Ƿ� wAPI������ ����. : �ٸ� ����ڿ��� ������ ����. : �̻� ������� ����������<br />

	    *API CS�������� :
		<select class="select" name="apiCS" id="apiCS">
			<option value='lotteCom' <%= chkIIF(apiCS="lotteCom","selected","") %> >�Ե�����</option>
	     	<option value='lotteimall' <%= chkIIF(apiCS="lotteimall","selected","") %> >�Ե�iMall</option>
	     	<option value='ezwel' <%= chkIIF(apiCS="ezwel","selected","") %> >���������</option>
	     	<option value='nvstorefarm' <%= chkIIF(apiCS="nvstorefarm","selected","") %> >�������</option>
			<option value='Mylittlewhoopee' <%= chkIIF(apiCS="Mylittlewhoopee","selected","") %> >������� Ĺ�ص�</option>
			<option value='nvstoregift' <%= chkIIF(apiCS="nvstoregift","selected","") %> >������ʼ����ϱ�</option>
	     	<option value='auction1010' <%= chkIIF(apiCS="auction1010","selected","") %> >����</option>
	     	<option value='cjmallCsreg1' <%= chkIIF(apiCS="cjmallCsreg1","selected","") %> >CJMall_��ǰ</option>
	     	<option value='cjmallCsreg2' <%= chkIIF(apiCS="cjmallCsreg2","selected","") %> >CJMall_���</option>
	     	<option value='cjmallCsreg3' <%= chkIIF(apiCS="cjmallCsreg3","selected","") %> >CJMall_CS���,������</option>
	     	<option value='ssg' <%= chkIIF(apiCS="ssg","selected","") %> >�ż���(SSG)</option>
			<option value='shintvshopping' <%= chkIIF(apiCS="shintvshopping","selected","") %> >�ż���TV����</option>
			<option value='wetoo1300k' <%= chkIIF(apiCS="wetoo1300k","selected","") %> >1300k</option>
			<option value='gmarket1010' <%= chkIIF(apiCS="gmarket1010","selected","") %> >������(New)</option>
			<option value='interpark' <%= chkIIF(apiCS="interpark","selected","") %> >������ũ</option>
			<option value='gseshopCancel' <%= chkIIF(apiCS="gseshopCancel","selected","") %> >gseshop ���</option>
			<option value='gseshopExcRet' <%= chkIIF(apiCS="gseshopExcRet","selected","") %> >gseshop ��ȯ,��ǰ</option>
			<option value='halfclub' <%= chkIIF(apiCS="halfclub","selected","") %> >����Ŭ��</option>
			<option value='coupang' <%= chkIIF(apiCS="coupang","selected","") %> >����</option>
			<option value='hmall1010' <%= chkIIF(apiCS="hmall1010","selected","") %> >HMall</option>
			<option value='11st1010' <%= chkIIF(apiCS="11st1010","selected","") %> >11����</option>
			<option value='WMP' <%= chkIIF(apiCS="WMP","selected","") %> >������(API)</option>
			<option value='wmpfashion' <%= chkIIF(apiCS="wmpfashion","selected","") %> >������W�м�(API)</option>
	     	<option value='kakaostore' <%= chkIIF(apiCS="kakaostore","selected","") %> >īī���彺���</option>
     	</select>
     	<input type="button" class="button" value="API�������" onClick="apiCSProcess();">
        &nbsp;
        <input type="button" class="button" value="���ε��üũ" onClick="TenCSProcess();">
        &nbsp;
        <input type="button" class="button" value="���޵��üũ" onClick="ExtCSProcess();">
	    <!--
		<input type="button" class="button" value="�Ե����� CS���� �������� " onClick="GetxSiteCSOrderList_lotteCom('lotteCom');">
		&nbsp;
		<input type="button" class="button" value="�Ե�i�� CS���� ��������" onClick="GetxSiteCSOrderList_lotteimall('lotteimall');" >
		&nbsp;
		<input type="button" class="button" value="CJMall CS���� ��������(��ǰ)" onClick="GetxSiteCSOrderList_CJ('cjmallCsreg1');"  >
		&nbsp;
		<input type="button" class="button" value="CJMall CS���� ��������(���)" onClick="GetxSiteCSOrderList_CJ('cjmallCsreg2');"  >
		&nbsp;
		<input type="button" class="button" value="CJMall CS���� ��������(CS���,������)" onClick="GetxSiteCSOrderList_CJ('cjmallCsreg3');"  >
		<br>
		<input type="button" class="button" value="��������� CS����" onClick="GetxSiteCSOrderList_ezwel('ezwel');" >
		&nbsp;
		<input type="button" class="button" value="������� CS����" onClick="GetxSiteCSOrderList_nvstorefarm('nvstorefarm');" >
		-->
	</td>
	<td align="right" valign="bottom">
		<input type="button" class="button" value="���� ���ó��" onClick="jsDelFinish(frmAct);" >
		<!--
		<input type="button" class="button" value="���� ����ó��" onClick="jsSetJupsu(frmAct);" >
		-->
		<input type="button" class="button" value="���� �Ϸ�ó��" onClick="jsSetFinish(frmAct);" >
	</td>
</tr>

</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19">
		�˻���� : <b><%= oCxSiteCSOrder.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCxSiteCSOrder.FTotalPage %></b>
	</td>
</tr>
<form name="frmAct" method="post" action="xSiteCSOrder_Process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sellsite" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="chkAll" onclick="fnCheckValidAll(this.checked,frmAct.cksel);"></td>
	<!--
	<td width="60">IDX</td>
	-->
	<td width="80">����</td>
	<td width="80">����</td>
	<td>���޸�</td>
	<td>�����ֹ���ȣ<br>(���ֹ���ȣ)</td>
	<!--
	<td width="80">����<br>���ֹ���</td>
	<td width="80">����<br>CS��</td>
	-->
	<td width="90">�ֹ���ȣ</td>
	<td width="70">�ֹ�����</td>
	<td width="60">����</td>
	<td>�귣��</td>
	<td width="70">��ǰ�ڵ�</td>
	<td align="left">��ǰ��<br><font color="blue">[�ɼǸ�]</font></td>
	<td width="30">����</td>
	<td>����CS������</td>
    <td>AsID</td>
	<td width="70">����</td>
	<td width="70">���޻���</td>
	<td width="140">������(<%= CHKIIF(ordBy="2", "����", "SCM") %>)</td>
	<td>���</td>
</tr>

<% for i=0 to oCxSiteCSOrder.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF">
	<td><input type="checkbox" name="cksel" value="<%= oCxSiteCSOrder.FItemList(i).Fidx %>" onclick="AnCheckClick(this);" <%= CHKIIF(oCxSiteCSOrder.FItemList(i).Fcurrstate = "B007", "disabled", "") %> ></td>
	<!--
	<td><%= oCxSiteCSOrder.FItemList(i).Fidx %></td>
	-->
	<td align="left"><%= Left(oCxSiteCSOrder.FItemList(i).Fdivname, 6) %></td>
	<td align="left"><%= Left(oCxSiteCSOrder.FItemList(i).Fgubunname, 6) %></td>
	<td><%= oCxSiteCSOrder.FItemList(i).FSellSite %></td>
	<td>
		<a href="javascript:jsSearchByOutMallOrderSerial('<%= oCxSiteCSOrder.FItemList(i).FOutMallOrderSerial %>')"><%= oCxSiteCSOrder.FItemList(i).FOutMallOrderSerial %></a>
		<% if (oCxSiteCSOrder.FItemList(i).ForgOutMallOrderSerial <> "") then %><br>(<%= oCxSiteCSOrder.FItemList(i).ForgOutMallOrderSerial %>)<% end if %>
	</td>
	<!--
	<td><%= oCxSiteCSOrder.FItemList(i).FOrgDetailKey %></td>
	<td><%= oCxSiteCSOrder.FItemList(i).FCSDetailKey %></td>
	-->
	<td><a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= oCxSiteCSOrder.FItemList(i).FOrderSerial %>')"><b><%= oCxSiteCSOrder.FItemList(i).FOrderSerial %></b></a></td>
	<td>
		<font color="<%= oCxSiteCSOrder.FItemList(i).IpkumDivColor %>"><%= oCxSiteCSOrder.FItemList(i).IpkumDivName %></font>
		<% if (oCxSiteCSOrder.FItemList(i).Fcancelyn <> "N") then %>
		<br />(���)
		<% elseif (oCxSiteCSOrder.FItemList(i).Fjupsucscnt = 0) and (oCxSiteCSOrder.FItemList(i).Fupcheconfirmcscnt = 0) and (oCxSiteCSOrder.FItemList(i).Ffinishcscnt > 0) then %>
		<br />(CS�Ϸ�)
		<% end if %>
	</td>
	<td><%= Left(oCxSiteCSOrder.FItemList(i).FOrderName,4) %></td>
	<td><%= oCxSiteCSOrder.FItemList(i).Fmakerid %></td>
	<td><%= oCxSiteCSOrder.FItemList(i).FItemID %></td>
	<td align="left"><%= oCxSiteCSOrder.FItemList(i).FOutMallItemName %><br><font color="blue">[<%= oCxSiteCSOrder.FItemList(i).FOutMallItemOptionName %>]</font></td>
	<td><%= oCxSiteCSOrder.FItemList(i).Fitemno %></td>
	<td>
		<% if (oCxSiteCSOrder.FItemList(i).Ftencscnt = 0) then %>
			<% if (oCxSiteCSOrder.FItemList(i).Fdelcscnt = 0) then %>
				<% if (oCxSiteCSOrder.FItemList(i).Fipkumdiv < "7") then %>
				<input type="button" class="button" value="���" onClick="PopOpenCancelItem('<%= oCxSiteCSOrder.FItemList(i).FOrderSerial %>')">
				<% elseif (oCxSiteCSOrder.FItemList(i).Fipkumdiv = "7") then %>
				<input type="button" class="button" value="���" onClick="PopOpenCancelItem('<%= oCxSiteCSOrder.FItemList(i).FOrderSerial %>')">
				<input type="button" class="button" value="��ǰ" onClick="PopOpenReceiveItemByUpche('<%= oCxSiteCSOrder.FItemList(i).FOrderSerial %>')">
				<% else %>
				<input type="button" class="button" value="��ǰ" onClick="PopOpenReceiveItemByUpche('<%= oCxSiteCSOrder.FItemList(i).FOrderSerial %>')">
				<% end if %>
			<% else %>
				<a href="javascript:Cscenter_Action_List('<%= oCxSiteCSOrder.FItemList(i).FOrderSerial %>');">
					���� : <%= oCxSiteCSOrder.FItemList(i).Fdelcscnt %>��
				</a>
			<% end if %>
		<% elseif (oCxSiteCSOrder.FItemList(i).Ftencscnt = 1) then %>
			<a href="javascript:Cscenter_Action_List('<%= oCxSiteCSOrder.FItemList(i).FOrderSerial %>');">
			<%= oCxSiteCSOrder.FItemList(i).Ftencsdivname %> 1��<br />
			<%= oCxSiteCSOrder.FItemList(i).Fjupsucscnt %>
			/
			<% if (oCxSiteCSOrder.FItemList(i).Fupcheconfirmcscnt>0) then %>
			<b><font color="red"><%= oCxSiteCSOrder.FItemList(i).Fupcheconfirmcscnt %></font></b>
			<% else %>
			<%= oCxSiteCSOrder.FItemList(i).Fupcheconfirmcscnt %>
			<% end if %>
			/
			<% if (oCxSiteCSOrder.FItemList(i).Ffinishcscnt>0) then %>
			<b><font color="red"><%= oCxSiteCSOrder.FItemList(i).Ffinishcscnt %></font></b>
			<% else %>
			<%= oCxSiteCSOrder.FItemList(i).Ffinishcscnt %>
			<% end if %>
			</a>
		<% elseif (oCxSiteCSOrder.FItemList(i).Ftencscnt > 1) then %>
			<a href="javascript:Cscenter_Action_List('<%= oCxSiteCSOrder.FItemList(i).FOrderSerial %>');">
			<%= oCxSiteCSOrder.FItemList(i).Ftencsdivname %> �� <%= (oCxSiteCSOrder.FItemList(i).Ftencscnt - 1) %>��<br />
			<%= oCxSiteCSOrder.FItemList(i).Fjupsucscnt %>
			/
			<% if (oCxSiteCSOrder.FItemList(i).Fupcheconfirmcscnt>0) then %>
			<b><font color="red"><%= oCxSiteCSOrder.FItemList(i).Fupcheconfirmcscnt %></font></b>
			<% else %>
			<%= oCxSiteCSOrder.FItemList(i).Fupcheconfirmcscnt %>
			<% end if %>
			/
			<% if (oCxSiteCSOrder.FItemList(i).Ffinishcscnt>0) then %>
			<b><font color="red"><%= oCxSiteCSOrder.FItemList(i).Ffinishcscnt %></font></b>
			<% else %>
			<%= oCxSiteCSOrder.FItemList(i).Ffinishcscnt %>
			<% end if %>
			</a>
		<% end if %>
	</td>
    <td>
        <%= oCxSiteCSOrder.FItemList(i).Fasid %>
        <%
        if IsNull(oCxSiteCSOrder.FItemList(i).Fasid) or (Not IsNull(oCxSiteCSOrder.FItemList(i).Fasid) and csregyn="N") then
            if oCxSiteCSOrder.FItemList(i).FSellSite = "ssg" then
                if (oCxSiteCSOrder.FItemList(i).Fdivcd = "A004") then
                    '// SSG ��ǰ
        %>
        <input type="button" class="button" value="üũ" onClick="jsCheckCs('<%= oCxSiteCSOrder.FItemList(i).FSellSite %>', '<%= oCxSiteCSOrder.FItemList(i).FOutMallOrderSerial %>')" />
        <%
                elseif (oCxSiteCSOrder.FItemList(i).Fdivcd = "A011") then
                    '// SSG ��ȯȸ��
        %>
        <input type="button" class="button" value="üũ" onClick="jsCheckCs('<%= oCxSiteCSOrder.FItemList(i).FSellSite %>', '<%= oCxSiteCSOrder.FItemList(i).FOutMallOrderSerial %>')" />
        <%
                elseif (oCxSiteCSOrder.FItemList(i).Fdivcd = "A008") then
                    '// SSG ���
        %>
        <input type="button" class="button" value="üũ" onClick="jsCheckCs('<%= oCxSiteCSOrder.FItemList(i).FSellSite %>', '<%= oCxSiteCSOrder.FItemList(i).FOutMallOrderSerial %>')" />
        <%
                end if
            end if
        end if
        %>
    </td>
	<td><font color="<%= oCxSiteCSOrder.FItemList(i).GetCurrStateColor %>"><%= oCxSiteCSOrder.FItemList(i).GetCurrStateName %></font></td>
	<td>
        <%

        if oCxSiteCSOrder.FItemList(i).FOutMallCurrState = "B007" or oCxSiteCSOrder.FItemList(i).FOutMallCurrState = "B008" then
        %>
        <font color="<%= oCxSiteCSOrder.FItemList(i).GetExtCurrStateColor %>">
            <%= oCxSiteCSOrder.FItemList(i).GetExtCurrStateName %>
        </font>
        <%
        elseif oCxSiteCSOrder.FItemList(i).FSellSite = "ssg" then
            if (oCxSiteCSOrder.FItemList(i).Fdivcd = "A004") or (oCxSiteCSOrder.FItemList(i).Fdivcd = "A011") then
        %>
        <input type="button" class="button" value="<%= CHKIIF(IsNull(oCxSiteCSOrder.FItemList(i).FOutMallCurrState), "üũ", oCxSiteCSOrder.FItemList(i).GetExtCurrStateName) %>" onClick="jsExtCheckCs('<%= oCxSiteCSOrder.FItemList(i).FSellSite %>', '<%= oCxSiteCSOrder.FItemList(i).Fdivcd %>', '<%= oCxSiteCSOrder.FItemList(i).FOutMallOrderSerial %>')" />
        <%
            else
        %>
        <font color="<%= oCxSiteCSOrder.FItemList(i).GetExtCurrStateColor %>">
            <%= oCxSiteCSOrder.FItemList(i).GetExtCurrStateName %>
        </font>
        <%
            end if
        else
        %>
        <font color="<%= oCxSiteCSOrder.FItemList(i).GetExtCurrStateColor %>">
            <%= oCxSiteCSOrder.FItemList(i).GetExtCurrStateName %>
        </font>
        <%
        end if
        %>
    </td>
	<td><%= CHKIIF(ordBy="2", oCxSiteCSOrder.FItemList(i).FOutMallRegDate, oCxSiteCSOrder.FItemList(i).Fregdate) %></td>
	<td>
		<% if (oCxSiteCSOrder.FItemList(i).Fcurrstate = "B001") then %>
			<!--
			<input type="button" class="button" value="����" onClick="jsSetJupsuOne(<%= oCxSiteCSOrder.FItemList(i).Fidx %>)">
			-->
			<input type="button" class="button" value="�Ϸ�" onClick="jsSetFinishOne(<%= oCxSiteCSOrder.FItemList(i).Fidx %>)">
		<% elseif (oCxSiteCSOrder.FItemList(i).Fcurrstate = "B002") then %>
			<input type="button" class="button" value="���" onClick="jsDelFinishOne(<%= oCxSiteCSOrder.FItemList(i).Fidx %>)">
			<input type="button" class="button" value="�Ϸ�" onClick="jsSetFinishOne(<%= oCxSiteCSOrder.FItemList(i).Fidx %>)">
		<% elseif (oCxSiteCSOrder.FItemList(i).Fcurrstate = "B007") then %>
			<input type="button" class="button" value="���" onClick="jsDelFinishOne(<%= oCxSiteCSOrder.FItemList(i).Fidx %>)">
			<!--
			<input type="button" class="button" value="����" onClick="jsSetJupsuOne(<%= oCxSiteCSOrder.FItemList(i).Fidx %>)">
			-->
		<% end if %>
	</td>
</tr>
<% next %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="19" align="center">
		<% if oCxSiteCSOrder.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCxSiteCSOrder.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCxSiteCSOrder.StartScrollPage to oCxSiteCSOrder.FScrollCount + oCxSiteCSOrder.StartScrollPage - 1 %>
			<% if i>oCxSiteCSOrder.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCxSiteCSOrder.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>

<form name="frmTmp" method="post" action="/admin/etc/cjmall/actCjMallReq_TEST.asp">
<input type="hidden" name="cmdparam" value="">
</form>

<form name="frmWapi" method="post" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sellsite" value="">
</form>
<%
set oCxSiteCSOrder = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
