<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ü� �ۼ�
' History : �̻� ����
'			2018.03.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/tenbalju.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<%

DIM CBRAND_INEXCLUDE_USING : CBRAND_INEXCLUDE_USING = True
Dim FlushCount : FlushCount=100  ''2016/04/18 :: ASP �������� �����Ͽ� Response ������ ������ ������ �ʰ��Ǿ����ϴ�.

dim pagesize, notitemlist, itemlist, notitemlistinclude, itemlistinclude, itemlistinclude2, notbrandlistinclude, brandlistinclude
dim research, yyyy1,mm1,dd1,yyyymmdd,nowdate, onlyOne,dcnt, danpumcheck, upbeaInclude, dcnt2, searchtypestring
dim imsi, sagawa, ems, epostmilitary, bigitem, fewitem, kpack, deliveryarea, onejumuntype, onejumuncount, onejumuncompare
dim tenbeaonly, tenbeamakeonorder, cn10x10, ecargo, extSiteName, stockLocationGubun, excMinusStock, excRealMinusStock
dim presentOnly, show100, repeatOrderCnt, standingorderinclude, page, ojumun, ix,iy, tenbaljucount
dim boxType, before15hour, excItem, danpumYN, boxGubun, agvstockgubun

	standingorderinclude = requestcheckvar(request("standingorderinclude"),2)
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

pagesize = request("pagesize")

deliveryarea = request("deliveryarea")

bigitem = request("bigitem")
fewitem = request("fewitem")

upbeaInclude = request("upbeaInclude")

notitemlistinclude = request("notitemlistinclude")
itemlistinclude = request("itemlistinclude")
itemlistinclude2 = request("itemlistinclude2")

notbrandlistinclude = request("notbrandlistinclude")
brandlistinclude = request("brandlistinclude")

notitemlist = request("notitemlist")
itemlist = request("itemlist")

research = request("research")

onejumuntype = request("onejumuntype")
onejumuncount = request("onejumuncount")
onejumuncompare = request("onejumuncompare")

tenbeaonly = request("tenbeaonly")

tenbeamakeonorder = request("tenbeamakeonorder")

extSiteName = request("extSiteName")

stockLocationGubun = request("stockLocationGubun")
excMinusStock = request("excMinusStock")
excRealMinusStock = request("excRealMinusStock")
presentOnly = request("presentOnly")
show100 = request("show100")
repeatOrderCnt = request("repeatOrderCnt")
boxType = request("boxType")
before15hour = request("before15hour")
excItem = request("excItem")
danpumYN = request("danpumYN")
boxGubun = request("boxGubun")
agvstockgubun = request("agvstockgubun")

if (research = "") then
	notitemlistinclude = "on"
	if (CBRAND_INEXCLUDE_USING) then
	    notbrandlistinclude = "on"
    end if

    if (excItem = "Y") then
        notitemlistinclude = ""
        itemlistinclude = "on"
    elseif (excItem = "N") then
        notitemlistinclude = "on"
        itemlistinclude = ""
    end if

    if (danpumYN = "Y") then
        fewitem = "1DN"
    elseif (danpumYN = "N") then
        fewitem = "2UP"
    end if

    if (boxGubun <> "") then
        boxType = boxGubun
    end if

	tenbeamakeonorder = "E"
	extSiteName = "10x10"
	presentOnly = "N"
	show100 = "Y"
	excRealMinusStock = "Y"
end if

if (repeatOrderCnt = "") then
	repeatOrderCnt = "2"
end if

'dcnt = trim(request("dcnt"))
'dcnt2 = trim(request("dcnt2"))
'onlyOne = request("onlyOne")
'danpumcheck = request("danpumcheck")
'imsi  = request("imsi")
'sagawa= request("sagawa")
'ems   = request("ems")
'epostmilitary   = request("epostmilitary")



'==============================================================================
if yyyy1="" then
	nowdate = CStr(Now)
	nowdate = DateSerial(Left(nowdate,4), CLng(Mid(nowdate,6,2))-2,Mid(nowdate,9,2))
	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
	dd1 = Mid(nowdate,9,2)
end if

if onejumuncount="" then
	onejumuncount = "1"
end if

if onejumuncompare="" then
	onejumuncompare = "less"
end if


'// ============================================================================
ems = ""
kpack = ""
epostmilitary = ""
cn10x10 = ""
ecargo = ""
if deliveryarea<>"" then
	if (deliveryarea = "ZZ") then
		'// ���δ�
		epostmilitary   = "on"
	elseif (deliveryarea = "EMS") then
		'// �ؿܹ��
		ems   = "on"
	elseif (deliveryarea = "KPACK") then
		'// �ؿܹ�� kpack
		kpack   = "on"
	elseif (deliveryarea = "CN10X10") then
		'// �߱������
		cn10x10 = "on"
	elseif (deliveryarea = "ECARGO") then
		'// ��ī����
		ecargo = "on"
	elseif (deliveryarea = "QQ") then
		'// �����
	else
		'// ��Ÿ : �������
		deliveryarea = "KR"
	end if
end if


'// ============================================================================
''�ӽ�..
'if (research="") then
'    notitemlist = "311341"
'    notitemlistinclude="on"
'end if

if research="" then
	'notitemlist = "45718"
	''if notitemlist="" then notitemlist="29002,29003,29004,29005,29006,29007,29008,29009,29010,29011,29012,29013,29014"
	''if itemlist="" then itemlist="29002,29003,29004,29005,29006,29007,29008,29009,29010,29011,29012,29013,29014"
	'if notitemlistinclude="" then notitemlistinclude="on"
end if

''��Ű ���� /2016/04/18
''if (pagesize="") then
''	pagesize = request.cookies("baljupagesize")
''end if

if (pagesize="") then pagesize=200
''if (pagesize>=2000) then pagesize=1000

''��Ű ���� /2016/04/18
''response.cookies("baljupagesize") = pagesize

page = request("page")
if (page="") then page=1

set ojumun = new CTenBalju

''�� ����¡�� 2�� �˻�
ojumun.FPageSize = pagesize * 5
''ojumun.FPageSize = pagesize

if notitemlistinclude="on" then
	ojumun.FRectNotIncludeItem = "Y"
else
	ojumun.FRectNotIncludeItem = ""
end if

if itemlistinclude="on" then
	ojumun.FRectIncludeItem = "Y"
elseif itemlistinclude2="on" then
    ojumun.FRectIncludeItem = "E"
else
	ojumun.FRectIncludeItem = ""
end if

if notbrandlistinclude="on" then
	ojumun.FRectNotIncludebrand = "Y"
else
	ojumun.FRectNotIncludebrand = ""
end if

if brandlistinclude="on" then
	ojumun.FRectIncludebrand = "Y"
else
	ojumun.FRectIncludebrand = ""
end if

ojumun.FCurrPage = page

ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1

''��ü��� ���� �ֹ���.
ojumun.FRectUpbeaInclude = upbeaInclude

''�簡�� ��۱ǿ�
ojumun.FRectOnlySagawaDeliverArea = sagawa

if deliveryarea<>"" then
	ojumun.FRectDeliveryArea = deliveryarea
end if

if fewitem<>"" then
	ojumun.FRectOnlyFewItem = fewitem
end if

if onejumuntype<>"" then
	ojumun.FRectOnlyOneJumun = "Y"

	ojumun.FRectOnlyOneJumunType = onejumuntype
	ojumun.FRectOnlyOneJumunCompare = onejumuncompare
	ojumun.FRectOnlyOneJumunCount = onejumuncount
end if

if tenbeaonly<>"" then
	ojumun.FRectTenbeaOnly = "Y"
end if

if tenbeamakeonorder <> "" then
	ojumun.FRectTenbeaMakeOnOrder = tenbeamakeonorder
end if


ojumun.FRectSiteGubun = extSiteName

ojumun.FRectStockLocationGubun = stockLocationGubun
ojumun.FRectExcMinusStock = excMinusStock
ojumun.FRectExcRealMinusStock = excRealMinusStock
ojumun.FRectPresentOnly = presentOnly
ojumun.FRectRepeatOrderCnt = repeatOrderCnt
ojumun.FRectstandingorderinclude = standingorderinclude
ojumun.FRectBoxType = boxType
ojumun.FRectBefore15Hour = before15hour
ojumun.FRectAgvStockGubun = agvstockgubun


ojumun.GetBaljuItemListProc
''ojumun.GetBaljuItemListNew

tenbaljucount =0

dim MaxTenBaljuCount : MaxTenBaljuCount = 100

%>
<script language='javascript'>
var tenBaljuCnt = 0;
function CheckNBalju(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;
    var isDasBalju = false;
    var isEmsBalju = <%= chkIIF(ems="on","true","false")%>;
    var isKpackBalju = <%= chkIIF(kpack="on","true","false")%>;
    var isMilitaryBalju = <%= chkIIF(epostmilitary="on","true","false")%>;
	var isCn10x10Balju = <%= chkIIF(cn10x10="on","true","false")%>;
	var isEcargoBalju = <%= chkIIF(ecargo="on","true","false")%>;
	var isQuickBeasongBalju = <%= chkIIF(deliveryarea="QQ","true","false")%>;

	var isGiftPojang = document.frm.presentOnly[1].checked ? 'Y' : 'N';
	upfrm.isGiftPojang.value = isGiftPojang;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� �ֹ��� �����ϴ�.');
		return;
	}

    if (document.all.groupform.songjangdiv.value.length<1){
		alert('��� �ù�縦 ���� �ϼ���.');
		document.all.groupform.songjangdiv.focus();
		return;
	}

	if (document.all.groupform.workgroup.value.length<1){
		alert('�۾� �׷��� ���� �ϼ���.');
		document.all.groupform.workgroup.focus();
		return;
	}

	/*
    //C�۾��� DAS
    isDasBalju = (document.all.groupform.workgroup.value=="C");

    //DAS ������� üũ, �ٹ� 150�� ����.
    if ((isDasBalju)&&(tenBaljuCnt>150)){
        alert('DAS ������ô� �ٹ����� ��� 150�� �̸��� �����մϴ�. ');
		document.all.groupform.workgroup.focus();
		return;
    }
	 */

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
                if ((frm.tenbeaexists.value == "Y") && (frm.boxType.value == "X")) {
                    alert('!!!!!! �ڽ������� �����ȵ� �ֹ��� �ֽ��ϴ�. !!!!!!\n\n���� �ڽ������ �����ϼ���.');
                    return;
                }
			}
		}
	}

	// ========================================================================
    if (isEmsBalju){
        if (document.all.groupform.workgroup.value!="E"){
            alert('EMS(�ؿ�)������ô� E �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="E"){
            alert('�˻������� �ؿܹ���̾�� EMS(�ؿ�)������ð� �����մϴ�.');
            return;
        }
    }
    if (isKpackBalju){
        if (document.all.groupform.workgroup.value!="R"){
            alert('KPACK(�ؿ�)������ô� R �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="R"){
            alert('�˻������� �ؿܹ���̾�� KPACK(�ؿ�)������ð� �����մϴ�.');
            return;
        }
    }

	if ((isQuickBeasongBalju == true) && (document.all.groupform.songjangdiv.value != "98")) {
		alert('�ù�縦 ��������� �����ϼ���.');
           return;
	}

    if (((document.all.groupform.songjangdiv.value=="90")&&(document.all.groupform.workgroup.value!="E"))||((document.all.groupform.songjangdiv.value!="90")&&(document.all.groupform.workgroup.value=="E"))){
        alert('EMS(�ؿ�)������ô� E �۾��常 �����մϴ�.');
        return;
    }

    if (((document.all.groupform.songjangdiv.value=="93")&&(document.all.groupform.workgroup.value!="R"))||((document.all.groupform.songjangdiv.value!="93")&&(document.all.groupform.workgroup.value=="R"))){
        alert('KPACK(�ؿ�)������ô� R �۾��常 �����մϴ�.');
        return;
    }

	// ========================================================================
    if (isMilitaryBalju){
        if (document.all.groupform.workgroup.value!="G"){
            alert('���δ� ������ô� G �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="G"){
            alert('�˻������� ���δ����̾�� ���δ�������ð� �����մϴ�.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="8")&&(document.all.groupform.workgroup.value!="G"))||((document.all.groupform.songjangdiv.value!="8")&&(document.all.groupform.workgroup.value=="G"))){
        alert('���δ� ������ô� G �۾��常 �����մϴ�.');
        return;
    }

	// ========================================================================
    if (isCn10x10Balju){
        if (document.all.groupform.workgroup.value!="H"){
            alert('EMS(�߱���)������ô� H �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="H"){
            alert('��������� �߱�������̾�� EMS(�߱���)������ð� �����մϴ�.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="91")&&(document.all.groupform.workgroup.value!="H"))||((document.all.groupform.songjangdiv.value!="91")&&(document.all.groupform.workgroup.value=="H"))){
        alert('EMS(�߱���)������ô� H �۾��常 �����մϴ�.');
        return;
    }

	// ========================================================================
    if (isEcargoBalju){
        if (document.all.groupform.workgroup.value!="J"){
            alert('�ؿ�(ECARGO)������ô� J �۾��常 �����մϴ�.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="J"){
            alert('��������� �ؿ�(ECARGO)����̾�� �ؿ�(ECARGO)������ð� �����մϴ�.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="92")&&(document.all.groupform.workgroup.value!="J"))||((document.all.groupform.songjangdiv.value!="92")&&(document.all.groupform.workgroup.value=="J"))){
        alert('�ؿ�(ECARGO)������ô� J �۾��常 �����մϴ�.');
        return;
    }

	// ========================================================================
    if (isDasBalju){
        if (!confirm('DAS ������� �Դϴ�. ��� �Ͻðڽ��ϱ�?')){
            return;
        }
    }

    if (document.all.groupform.pickingStationCd.value == '') {
        alert('��ŷ�����̼��� �����ϼ���.');
        return;
    }

    frm = document.frm;
    if (frm.boxType.value != '') {
        // �ڽ������� ������
        if (frm.agvstockgubun[2].checked) {
            // AGV �ֹ��ΰ��
            if ((frm.boxType.value != 'ABC') && (frm.pagesize.value != '40')) {
                if (confirm('A1,B1,C1 �̿��� �ڽ��������� ��� �ֹ��� 40�Ǿ� ��������ؾ� �մϴ�.\n\n������ �����Ͻðڽ��ϱ�?') != true) {
                    return;
                }
            }
        }
    }

	var ret = confirm('���� �ֹ��� �� ������ü��� �����Ͻðڽ��ϱ�?');
	if (ret){
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderserial.value = upfrm.orderserial.value + "|" + frm.orderserial.value;
					upfrm.sitename.value = upfrm.sitename.value + "|" + frm.sitename.value;
				}
			}
		}
		upfrm.songjangdiv.value = document.all.groupform.songjangdiv.value;
		upfrm.workgroup.value = document.all.groupform.workgroup.value;
		upfrm.ems.value = "<%= ems %>";
		upfrm.kpack.value = "<%= kpack %>";
		upfrm.epostmilitary.value = "<%= epostmilitary %>";
		upfrm.cn10x10.value = "<%= cn10x10 %>";
		upfrm.ecargo.value = "<%= ecargo %>";
		upfrm.extSiteName.value = "<%= extSiteName %>";
        upfrm.boxType.value = "<%= boxType %>";
        upfrm.pickingStationCd.value = document.all.groupform.pickingStationCd.value;

		if (isDasBalju) {
		    upfrm.baljutype.value = "D";
		}else{
		    upfrm.baljutype.value = "";
		}

		if ((upfrm.songjangdiv.value == "91") || (upfrm.songjangdiv.value == "92") || (upfrm.songjangdiv.value == "93")) {
			upfrm.songjangdiv.value = "90";
		}

		upfrm.submit();
	}
}

function jsSetBoxType() {
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			pass = ((pass)||(frm.cksel.checked));
		}
	}

	if (!pass) {
		alert('���� �ֹ��� �����ϴ�.');
		return;
	}

	var ret = confirm('���� �ֹ� �ڽ������Ͻðڽ��ϱ�?');
	if (ret){
		upfrm.action = 'doboxtype.asp';
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.checked){
					upfrm.orderserial.value = upfrm.orderserial.value + "|" + frm.orderserial.value;
					upfrm.tenbeaexists.value = upfrm.tenbeaexists.value + "|" + frm.tenbeaexists.value;
					upfrm.boxType.value = upfrm.boxType.value + "|" + frm.boxType.value;
				}
			}
		}

		upfrm.submit();
	}
}

function jsMakeBaljuUpbae() {
    var upfrm = document.frmArrupdate;

    var ret = confirm('�����ֹ�(�ٹ����� �ֹ� ����) 1000�� �������ó�� �Ͻðڽ��ϱ�?');
    if (ret) {
        upfrm.action = 'newbaljumaker_process.asp';
        upfrm.mode.value = 'baljuupbae';
        upfrm.submit();
    }
}

function ViewOrderDetail(frm){
	//var popwin;
    //popwin = window.open('','orderdetail');
    frm.target = 'orderdetail';
    frm.action="viewordermaster.asp"
	frm.submit();

}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function EnableDiable(icomp){
	//return;
	var frm = document.frm;
	var ischecked = icomp.checked;
	if (ischecked){
		if (icomp.name=="notitemlistinclude"){
			frm.itemlistinclude.checked = !(ischecked);
            frm.itemlistinclude2.checked = !(ischecked);
		}else if (icomp.name=="itemlistinclude"){
			frm.notitemlistinclude.checked = !(ischecked);
            frm.itemlistinclude2.checked = !(ischecked);
        }else if (icomp.name=="itemlistinclude2"){
            frm.notitemlistinclude.checked = !(ischecked);
            frm.itemlistinclude.checked = !(ischecked);
		}
	}

	if (ischecked){
		if (icomp.name=="notbrandlistinclude"){
			frm.brandlistinclude.checked = !(ischecked);
		}else if (icomp.name=="brandlistinclude"){
			frm.notbrandlistinclude.checked = !(ischecked);
		}
	}

	if (icomp.name=="onlyOne"){
		frm.itemlistinclude.disabled = (ischecked);
        frm.itemlistinclude2.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.danpumcheck.checked = false;
	}


	if (icomp.name=="danpumcheck"){
		frm.itemlistinclude.disabled = (ischecked);
        frm.itemlistinclude2.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.onlyOne.checked = false;
	}
}

function poponeitem(){
	var popwin = window.open("/admin/ordermaster/poponeitem.asp","poponeitem","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function poponebrand(){
	var popwin = window.open("/admin/ordermaster/poponebrand.asp","poponebrand","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function chkUpbea(){
    var frm;
    var checkedExists = false;
    for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.tenbeaexists.value!="Y"){
			    frm.cksel.checked = true;
			    AnCheckClick(frm.cksel);
			    checkedExists = true;
			}
		}
	}

	if (checkedExists){
	    document.groupform.songjangdiv.value="24";
	    document.groupform.workgroup.value="Z";
	    CheckNBalju();
	}
}

function AnSelectAllFrame(bool){
	var frm, excNoBoxType;
    excNoBoxType = false;

    if (bool) {
        excNoBoxType = confirm('������ ������ �ڽ��� �����ϰ� �����Ͻðڽ��ϱ�?');
    }

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.disabled!=true){
                if ((frm.tenbeaexists.value == "Y") && (frm.boxType.value == "X") && (excNoBoxType == true) && (bool == true)) {
                    // do nothing
                } else {
                    frm.cksel.checked = bool;
				    AnCheckClick(frm.cksel);
                }

			}
		}
	}
}

</script>
<!-- ǥ ��ܹ� ����-->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
            111
        </td>
        <td align="right">

        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
</form>
<!-- ǥ ��ܹ� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
	        �� ��������� �Ǽ� : <Font color="#3333FF"><b><%= FormatNumber(ojumun.FTotalCount,0) %>(�ٹ� : <%= FormatNumber(ojumun.FTotalTenbaeCount,0) %>)</b></font>&nbsp;
			�� �ݾ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotalsum,0) %></font>&nbsp;
			��հ��ܰ� : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotalsum,0) %></font>
        </td>
        <td>&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<!-- ǥ �߰��� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="40" valign="center">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td width="200" align="left">
	        <div id="currsearchno">�� �˻��ֹ��Ǽ� : </div>
	        <div id="currtensearchno">�ٹ����ٹ�� �ֹ��Ǽ� : </div>
	        <!-- input type="checkbox" name="ck_upbea" onClick="chkUpbea();"> ����������� -->
        </td>
        <td align="right">
		<input type="button" value="�����ֹ� �ڽ�����" onclick="jsSetBoxType()">
		<form name="groupform">
		    <!--
		    <select name="baljutype">
		        <option value="">�Ϲ�
		        <option value="D">DAS
		        <option value="S">��ǰ���
		    </select>
		    -->
			<b>
			<%
			Select Case extSiteName
				Case "10x10"
					response.write "����(���޸� ����) �ֹ�"
				Case "extSiteAll"
					response.write "���޸���ü(��������) �ֹ�"
				Case "cjmall"
					response.write "���޸�(cjmall) �ֹ�"
				Case "interpark"
					response.write "���޸�(interpark) �ֹ�"
				Case "lotteCom"
					response.write "���޸�(lotteCom) �ֹ�"
				Case "lotteimall"
					response.write "���޸�(lotteimall) �ֹ�"
				Case "etcExtSite"
					response.write "��Ÿ���޸� �ֹ�"
				Case Else
					response.write "��ü �ֹ�"
			End Select
			%>
			</b>
			&nbsp;
		    <select name="songjangdiv">
		        <option value="">�ù�缱��</option>
				<!-- <option value="2" >�����ù�</select> -->
                <% if (now()>"2010-04-01") then %>
                	<option value="4" >CJ�ù�</option>
					<option value="98" >�����</option>
                <% else %>
                	<option value="4" >CJ�ù�</option>
			   		<option value="24" >�簡��</option>
			   	<% end if %>
			   	<option value="90" <%= CHKIIF(ems="on","selected","") %> >EMS(�ؿ�)</option>
				<option value="93" <%= CHKIIF(kpack="on","selected","") %> >EMS(KPACK)</option>				<!-- ������ �� 90 ���� �����Ѵ�.(EMS) -->
			   	<option value="8" <%= CHKIIF(epostmilitary="on","selected","") %> >��ü��(���δ�)</option>
				<option value="91" <%= CHKIIF(cn10x10="on","selected","") %> >EMS(�߱���)</option>				<!-- ������ �� 90 ���� �����Ѵ�.(EMS) -->
				<% if False then %>
				<option value="92" <%= CHKIIF(ecargo="on","selected","") %> >�ؿ�(ECARGO)</option>				<!-- ������ �� 90 ���� �����Ѵ�.(EMS) -->
				<% end if %>
		    </select>
			<select name="workgroup">
			   	<option value="">�۾��׷�</option>
			   	<option value="A" >A</option>
			   	<option value="B" >B</option>
			   	<option value="C" >C</option>
			   	<option value="D" >D</option>
			   	<option value="F" >F</option>
				<option value="K" >K</option>
				<option value="L" >L</option>
				<option value="M" >M</option>
				<option value="N" >N(��ǰ)</option>
			   	<option value="E" <%= CHKIIF(ems="on","selected","") %> >E(EMS)</option>
			   	<option value="R" <%= CHKIIF(kpack="on","selected","") %> >R(KPACK)</option>
			   	<option value="G" <%= CHKIIF(epostmilitary="on","selected","") %> >G(���δ�)</option>
				<option value="H" <%= CHKIIF(cn10x10="on","selected","") %> >H(�߱���)</option>
				<% if False then %>
				<option value="J" <%= CHKIIF(ecargo="on","selected","") %> >J(��ī��)</option>
				<% end if %>
			   	<option value="Z" >Z(����)</option>
		   	</select>
            <% Call drawSelectStationByStationGubun("PICK", "pickingStationCd", "") %>
			<input type="button" value="�����ֹ� ������ü��ۼ�" onclick="CheckNBalju()" disabled>
		</form>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- ǥ �߰��� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20" align="center"><input type="checkbox" name="cksel" onClick="AnSelectAllFrame(this.checked)"></td>
	<td width="80">�ֹ���ȣ</td>
	<td width="120">Site</td>
	<td width="50">����</td>
	<td width="120">UserID</td>
	<% if (FALSE) then %>
	<td width="120">������</td>
	<% end if %>
	<td width="120">������</td>
	<td width="60">�����ݾ�</td>
	<td width="60">�����Ѿ�</td>
	<td width="80">�������</td>
	<td width="80">�ŷ�����</td>
	<td width="110">�ֹ���</td>
    <td width="110">������</td>
	<td width="60">��ǰ<br />������</td>
	<td width="60">�ڽ�<br />������</td>
	<td>
	    <% if upbeaInclude<>"" then %>
	    ��������
	    <% else %>
	    �ٹ�����
	    <% end if %>
	    </td>
</tr>
<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% else %>
	<% for ix=0 to ojumun.FresultCount-1 %>
	<% if tenbaljucount < CLng(pagesize) and (tenbaljucount < MaxTenBaljuCount or show100 <> "Y")  then %>
	<% if (ix/FlushCount)=CLNG(ix/FlushCount) then response.write CLNG(ix/FlushCount): response.flush %>
<form name="frmBuyPrc_<%= ojumun.FItemList(ix).FOrderSerial %>" method="post" style="margin:0px;">
<input type="hidden" name="orderserial" value="<%= ojumun.FItemList(ix).FOrderSerial %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="sitename" value="<%= ojumun.FItemList(ix).FSiteName %>">
<input type="hidden" name="dlvcontrycode" value="<%= ojumun.FItemList(ix).FDlvcountryCode %>">
<tr align="center" bgcolor="#FFFFFF">

<!-- !!! EMS ���δ� �߱������ üũ�� Ŭ�������Ͽ��� �Ѵ�. !!! -->

<% if ((ems<>"") or (epostmilitary<>"") or (cn10x10<>"") or (ecargo<>"") or (kpack<>"") or (ojumun.FItemList(ix).FDlvcountryCode=deliveryarea)) then %>
<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
<% else %>
<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);" <%= CHKIIF(ojumun.FItemList(ix).FDlvcountryCode<>"" and ojumun.FItemList(ix).FDlvcountryCode<>"KR","disabled","") %> ></td>
<% end if %>

<td><a href="javascript:ViewOrderDetail(frmBuyPrc_<%= ojumun.FItemList(ix).FOrderSerial %>)" class="zzz"><%= ojumun.FItemList(ix).FOrderSerial %></a></td>
<td><font color="<%= ojumun.FItemList(ix).SiteNameColor %>"><%= ojumun.FItemList(ix).FSitename %></font></td>
<td><%= ojumun.FItemList(ix).FDlvcountryCode %></td>
<td><%= printUserId(ojumun.FItemList(ix).FUserID,2,"*") %></td>
<% if (FALSE) then %>
<td><%= ojumun.FItemList(ix).FBuyName %></td>
<% end if %>
<td><%= ojumun.FItemList(ix).FReqName %></td>
<td align="right"><font color="<%= ojumun.FItemList(ix).SubTotalColor%>"><%= FormatNumber(ojumun.FItemList(ix).FSubTotalPrice,0) %></font></td>
<td align="right"><%= FormatNumber(ojumun.FItemList(ix).FTotalSum,0) %></td>
<td><%= ojumun.FItemList(ix).JumunMethodName %></td>
<td><font color="<%= ojumun.FItemList(ix).IpkumDivColor %>"><%= ojumun.FItemList(ix).IpkumDivName %></font></td>
<td><%= Left(ojumun.FItemList(ix).FRegDate,16) %></td>
<td><%= Left(ojumun.FItemList(ix).Fipkumdate,16) %></td>
<td><%= ojumun.FItemList(ix).FTenbeaItemKindCnt %></td>
<td><%= CHKIIF(ojumun.FItemList(ix).FTenbeaItemKindCnt>0, ojumun.FItemList(ix).FboxType, "") %></td>
<td>
<% if ojumun.FItemList(ix).Ftenbeaexists then %>
<input type="hidden" name="tenbeaexists" value="Y">
<input type="hidden" name="boxType" value="<%= ojumun.FItemList(ix).FboxType %>">
<% tenbaljucount = tenbaljucount + 1 %>
�� <%= tenbaljucount %>
<% else %>
<input type="hidden" name="tenbeaexists" value="N">
<input type="hidden" name="boxType" value="">
<% end if %>
</td>
</tr>
</form>
	<% else %>
		<% exit for %>
	<% end if %>
	<% next %>
<% end if %>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<form name="frmArrupdate" method="post" action="dobaljumaker.asp" style="margin:0px;">
<input type="hidden" name="mode" value="arr">
<input type="hidden" name="orderserial" value="">
<input type="hidden" name="tenbeaexists" value="">
<input type="hidden" name="boxType" value="">
<input type="hidden" name="sitename" value="">
<input type="hidden" name="songjangdiv" value="">
<input type="hidden" name="workgroup" value="">
<input type="hidden" name="baljutype" value="">
<input type="hidden" name="ems" value="">
<input type="hidden" name="epostmilitary" value="">
<input type="hidden" name="cn10x10" value="">
<input type="hidden" name="ecargo" value="">
<input type="hidden" name="kpack" value="">
<input type="hidden" name="extSiteName" value="">
<input type="hidden" name="isGiftPojang" value="">
<input type="hidden" name="pickingStationCd" value="">
</form>
<%
set ojumun = Nothing
%>

<script language='javascript'>
document.all.currsearchno.innerHTML = "�˻����� : <Font color='#3333FF'><%= ix %></font>";
document.all.currtensearchno.innerHTML = "�ٹ����ٹ�� �˻����� : <Font color='#3333FF'><%= tenbaljucount %></font>";
tenBaljuCnt = 1*<%= tenbaljucount %>;
<% if onlyOne<>"" then %>
EnableDiable(frm.onlyOne);
<% end if %>

<% if danpumcheck<>"" then %>
EnableDiable(frm.danpumcheck);
<% end if %>

</script>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
