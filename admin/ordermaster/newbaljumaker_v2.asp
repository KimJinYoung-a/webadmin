<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  출고지시서 작성
' History : 이상구 생성
'			2018.03.26 한용민 수정
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
Dim FlushCount : FlushCount=100  ''2016/04/18 :: ASP 페이지를 실행하여 Response 버퍼의 구성된 제한이 초과되었습니다.

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
		'// 군부대
		epostmilitary   = "on"
	elseif (deliveryarea = "EMS") then
		'// 해외배송
		ems   = "on"
	elseif (deliveryarea = "KPACK") then
		'// 해외배송 kpack
		kpack   = "on"
	elseif (deliveryarea = "CN10X10") then
		'// 중국몰배송
		cn10x10 = "on"
	elseif (deliveryarea = "ECARGO") then
		'// 이카고배송
		ecargo = "on"
	elseif (deliveryarea = "QQ") then
		'// 퀵배송
	else
		'// 기타 : 국내배송
		deliveryarea = "KR"
	end if
end if


'// ============================================================================
''임시..
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

''쿠키 없앰 /2016/04/18
''if (pagesize="") then
''	pagesize = request.cookies("baljupagesize")
''end if

if (pagesize="") then pagesize=200
''if (pagesize>=2000) then pagesize=1000

''쿠키 없앰 /2016/04/18
''response.cookies("baljupagesize") = pagesize

page = request("page")
if (page="") then page=1

set ojumun = new CTenBalju

''총 페이징의 2배 검색
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

''업체배송 포함 주문건.
ojumun.FRectUpbeaInclude = upbeaInclude

''사가와 배송권역
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
		alert('선택 주문이 없습니다.');
		return;
	}

    if (document.all.groupform.songjangdiv.value.length<1){
		alert('출고 택배사를 선택 하세요.');
		document.all.groupform.songjangdiv.focus();
		return;
	}

	if (document.all.groupform.workgroup.value.length<1){
		alert('작업 그룹을 선택 하세요.');
		document.all.groupform.workgroup.focus();
		return;
	}

	/*
    //C작업장 DAS
    isDasBalju = (document.all.groupform.workgroup.value=="C");

    //DAS 출고지시 체크, 텐배 150개 이하.
    if ((isDasBalju)&&(tenBaljuCnt>150)){
        alert('DAS 출고지시는 텐바이텐 배송 150건 미만만 가능합니다. ');
		document.all.groupform.workgroup.focus();
		return;
    }
	 */

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
                if ((frm.tenbeaexists.value == "Y") && (frm.boxType.value == "X")) {
                    alert('!!!!!! 박스사이즈 지정안된 주문이 있습니다. !!!!!!\n\n먼저 박스사이즈를 지정하세요.');
                    return;
                }
			}
		}
	}

	// ========================================================================
    if (isEmsBalju){
        if (document.all.groupform.workgroup.value!="E"){
            alert('EMS(해외)출고지시는 E 작업장만 가능합니다.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="E"){
            alert('검색유형이 해외배송이어야 EMS(해외)출고지시가 가능합니다.');
            return;
        }
    }
    if (isKpackBalju){
        if (document.all.groupform.workgroup.value!="R"){
            alert('KPACK(해외)출고지시는 R 작업장만 가능합니다.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="R"){
            alert('검색유형이 해외배송이어야 KPACK(해외)출고지시가 가능합니다.');
            return;
        }
    }

	if ((isQuickBeasongBalju == true) && (document.all.groupform.songjangdiv.value != "98")) {
		alert('택배사를 퀵배송으로 선택하세요.');
           return;
	}

    if (((document.all.groupform.songjangdiv.value=="90")&&(document.all.groupform.workgroup.value!="E"))||((document.all.groupform.songjangdiv.value!="90")&&(document.all.groupform.workgroup.value=="E"))){
        alert('EMS(해외)출고지시는 E 작업장만 가능합니다.');
        return;
    }

    if (((document.all.groupform.songjangdiv.value=="93")&&(document.all.groupform.workgroup.value!="R"))||((document.all.groupform.songjangdiv.value!="93")&&(document.all.groupform.workgroup.value=="R"))){
        alert('KPACK(해외)출고지시는 R 작업장만 가능합니다.');
        return;
    }

	// ========================================================================
    if (isMilitaryBalju){
        if (document.all.groupform.workgroup.value!="G"){
            alert('군부대 출고지시는 G 작업장만 가능합니다.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="G"){
            alert('검색유형이 군부대배송이어야 군부대출고지시가 가능합니다.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="8")&&(document.all.groupform.workgroup.value!="G"))||((document.all.groupform.songjangdiv.value!="8")&&(document.all.groupform.workgroup.value=="G"))){
        alert('군부대 출고지시는 G 작업장만 가능합니다.');
        return;
    }

	// ========================================================================
    if (isCn10x10Balju){
        if (document.all.groupform.workgroup.value!="H"){
            alert('EMS(중국몰)출고지시는 H 작업장만 가능합니다.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="H"){
            alert('배송지역이 중국몰배송이어야 EMS(중국몰)출고지시가 가능합니다.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="91")&&(document.all.groupform.workgroup.value!="H"))||((document.all.groupform.songjangdiv.value!="91")&&(document.all.groupform.workgroup.value=="H"))){
        alert('EMS(중국몰)출고지시는 H 작업장만 가능합니다.');
        return;
    }

	// ========================================================================
    if (isEcargoBalju){
        if (document.all.groupform.workgroup.value!="J"){
            alert('해외(ECARGO)출고지시는 J 작업장만 가능합니다.');
            return;
        }
    }else{
        if (document.all.groupform.workgroup.value=="J"){
            alert('배송지역이 해외(ECARGO)배송이어야 해외(ECARGO)출고지시가 가능합니다.');
            return;
        }
    }

    if (((document.all.groupform.songjangdiv.value=="92")&&(document.all.groupform.workgroup.value!="J"))||((document.all.groupform.songjangdiv.value!="92")&&(document.all.groupform.workgroup.value=="J"))){
        alert('해외(ECARGO)출고지시는 J 작업장만 가능합니다.');
        return;
    }

	// ========================================================================
    if (isDasBalju){
        if (!confirm('DAS 출고지시 입니다. 계속 하시겠습니까?')){
            return;
        }
    }

    if (document.all.groupform.pickingStationCd.value == '') {
        alert('피킹스테이션을 선택하세요.');
        return;
    }

    frm = document.frm;
    if (frm.boxType.value != '') {
        // 박스사이즈 지정시
        if (frm.agvstockgubun[2].checked) {
            // AGV 주문인경우
            if ((frm.boxType.value != 'ABC') && (frm.pagesize.value != '40')) {
                if (confirm('A1,B1,C1 이외의 박스사이즈인 경우 주문을 40건씩 출고지시해야 합니다.\n\n강제로 진행하시겠습니까?') != true) {
                    return;
                }
            }
        }
    }

	var ret = confirm('선택 주문을 새 출고지시서로 저장하시겠습니까?');
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
		alert('선택 주문이 없습니다.');
		return;
	}

	var ret = confirm('선택 주문 박스지정하시겠습니까?');
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

    var ret = confirm('업배주문(텐배포함 주문 제외) 1000건 출고지시처리 하시겠습니까?');
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
        excNoBoxType = confirm('사이즈 미지정 박스를 제외하고 선택하시겠습니까?');
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
<!-- 표 상단바 시작-->
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
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td align="left">
	        총 미출고지시 건수 : <Font color="#3333FF"><b><%= FormatNumber(ojumun.FTotalCount,0) %>(텐배 : <%= FormatNumber(ojumun.FTotalTenbaeCount,0) %>)</b></font>&nbsp;
			총 금액 : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotalsum,0) %></font>&nbsp;
			평균객단가 : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotalsum,0) %></font>
        </td>
        <td>&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<tr>
		<td height="1" colspan="15" bgcolor="<%= adminColor("tablebg") %>"></td>
	</tr>
    <tr height="40" valign="center">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td width="200" align="left">
	        <div id="currsearchno">총 검색주문건수 : </div>
	        <div id="currtensearchno">텐바이텐배송 주문건수 : </div>
	        <!-- input type="checkbox" name="ck_upbea" onClick="chkUpbea();"> 업배출고지시 -->
        </td>
        <td align="right">
		<input type="button" value="선택주문 박스지정" onclick="jsSetBoxType()">
		<form name="groupform">
		    <!--
		    <select name="baljutype">
		        <option value="">일반
		        <option value="D">DAS
		        <option value="S">단품출고
		    </select>
		    -->
			<b>
			<%
			Select Case extSiteName
				Case "10x10"
					response.write "텐텐(제휴몰 제외) 주문"
				Case "extSiteAll"
					response.write "제휴몰전체(텐텐제외) 주문"
				Case "cjmall"
					response.write "제휴몰(cjmall) 주문"
				Case "interpark"
					response.write "제휴몰(interpark) 주문"
				Case "lotteCom"
					response.write "제휴몰(lotteCom) 주문"
				Case "lotteimall"
					response.write "제휴몰(lotteimall) 주문"
				Case "etcExtSite"
					response.write "기타제휴몰 주문"
				Case Else
					response.write "전체 주문"
			End Select
			%>
			</b>
			&nbsp;
		    <select name="songjangdiv">
		        <option value="">택배사선택</option>
				<!-- <option value="2" >현대택배</select> -->
                <% if (now()>"2010-04-01") then %>
                	<option value="4" >CJ택배</option>
					<option value="98" >퀵배송</option>
                <% else %>
                	<option value="4" >CJ택배</option>
			   		<option value="24" >사가와</option>
			   	<% end if %>
			   	<option value="90" <%= CHKIIF(ems="on","selected","") %> >EMS(해외)</option>
				<option value="93" <%= CHKIIF(kpack="on","selected","") %> >EMS(KPACK)</option>				<!-- 저장할 때 90 으로 변경한다.(EMS) -->
			   	<option value="8" <%= CHKIIF(epostmilitary="on","selected","") %> >우체국(군부대)</option>
				<option value="91" <%= CHKIIF(cn10x10="on","selected","") %> >EMS(중국몰)</option>				<!-- 저장할 때 90 으로 변경한다.(EMS) -->
				<% if False then %>
				<option value="92" <%= CHKIIF(ecargo="on","selected","") %> >해외(ECARGO)</option>				<!-- 저장할 때 90 으로 변경한다.(EMS) -->
				<% end if %>
		    </select>
			<select name="workgroup">
			   	<option value="">작업그룹</option>
			   	<option value="A" >A</option>
			   	<option value="B" >B</option>
			   	<option value="C" >C</option>
			   	<option value="D" >D</option>
			   	<option value="F" >F</option>
				<option value="K" >K</option>
				<option value="L" >L</option>
				<option value="M" >M</option>
				<option value="N" >N(단품)</option>
			   	<option value="E" <%= CHKIIF(ems="on","selected","") %> >E(EMS)</option>
			   	<option value="R" <%= CHKIIF(kpack="on","selected","") %> >R(KPACK)</option>
			   	<option value="G" <%= CHKIIF(epostmilitary="on","selected","") %> >G(군부대)</option>
				<option value="H" <%= CHKIIF(cn10x10="on","selected","") %> >H(중국몰)</option>
				<% if False then %>
				<option value="J" <%= CHKIIF(ecargo="on","selected","") %> >J(이카고)</option>
				<% end if %>
			   	<option value="Z" >Z(업배)</option>
		   	</select>
            <% Call drawSelectStationByStationGubun("PICK", "pickingStationCd", "") %>
			<input type="button" value="선택주문 출고지시서작성" onclick="CheckNBalju()" disabled>
		</form>
		</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20" align="center"><input type="checkbox" name="cksel" onClick="AnSelectAllFrame(this.checked)"></td>
	<td width="80">주문번호</td>
	<td width="120">Site</td>
	<td width="50">국가</td>
	<td width="120">UserID</td>
	<% if (FALSE) then %>
	<td width="120">구매자</td>
	<% end if %>
	<td width="120">수령인</td>
	<td width="60">결제금액</td>
	<td width="60">구매총액</td>
	<td width="80">결제방법</td>
	<td width="80">거래상태</td>
	<td width="110">주문일</td>
    <td width="110">결제일</td>
	<td width="60">상품<br />가지수</td>
	<td width="60">박스<br />사이즈</td>
	<td>
	    <% if upbeaInclude<>"" then %>
	    업배포함
	    <% else %>
	    텐배포함
	    <% end if %>
	    </td>
</tr>
<% if ojumun.FresultCount<1 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center">[검색결과가 없습니다.]</td>
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

<!-- !!! EMS 군부대 중국몰배송 체크는 클래스파일에서 한다. !!! -->

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
√ <%= tenbaljucount %>
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

<!-- 표 하단바 시작-->
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
<!-- 표 하단바 끝-->

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
document.all.currsearchno.innerHTML = "검색갯수 : <Font color='#3333FF'><%= ix %></font>";
document.all.currtensearchno.innerHTML = "텐바이텐배송 검색갯수 : <Font color='#3333FF'><%= tenbaljucount %></font>";
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
