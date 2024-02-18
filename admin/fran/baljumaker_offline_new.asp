<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 출고지시 관리
' Hieditor : 2011.03.07 서동석 생성
'			 2011.07.27 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_logisticsOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/baljuofflinecls.asp"-->
<!-- #include virtual="/lib/classes/logistics/logistics_agvCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
dim notitemlist, itemlist ,pagesize ,research ,danpumcheck ,dcnt2 ,companyid ,locationid3PL
dim notitemlistinclude, itemlistinclude ,searchtypestring ,onejumuntype ,deliveryarea
dim notbrandlistinclude, brandlistinclude ,onlyOne,dcnt ,yyyy1,mm1,dd1,yyyymmdd,nowdate
dim imsi, sagawa, ems, epostmilitary, bigitem ,onejumuncount, onejumuncompare ,ix,iy
dim tenbeaonly ,upbeaInclude ,locationidto ,includeminus, includezerostock ,page ,ojumun, mwdiv
dim makerid
	yyyy1 = requestCheckVar(request("yyyy1"), 32)
	mm1 = requestCheckVar(request("mm1"), 32)
	dd1 = requestCheckVar(request("dd1"), 32)
	pagesize = requestCheckVar(request("pagesize"), 32)
	deliveryarea = requestCheckVar(request("deliveryarea"), 32)
	bigitem = requestCheckVar(request("bigitem"), 32)
	companyid = "10x10"
	locationid3PL = "10x10"
	upbeaInclude = request("upbeaInclude")
	tenbeaonly = request("tenbeaonly")
	notitemlistinclude = requestCheckVar(request("notitemlistinclude"), 32)
	itemlistinclude = requestCheckVar(request("itemlistinclude"), 32)
	notbrandlistinclude = requestCheckVar(request("notbrandlistinclude"), 32)
	brandlistinclude = requestCheckVar(request("brandlistinclude"), 32)
	notitemlist = requestCheckVar(request("notitemlist"), 32)
	itemlist = requestCheckVar(request("itemlist"), 32)
	research = requestCheckVar(request("research"), 32)
	onejumuntype = requestCheckVar(request("onejumuntype"), 32)
	onejumuncount = requestCheckVar(request("onejumuncount"), 32)
	onejumuncompare = requestCheckVar(request("onejumuncompare"), 32)
	locationidto = requestCheckVar(request("locationidto"), 32)
	includeminus = requestCheckVar(request("includeminus"), 32)
	includezerostock = requestCheckVar(request("includezerostock"), 32)
	makerid = requestCheckVar(request("makerid"), 32)
	'dcnt = trim(request("dcnt"))
	'dcnt2 = trim(request("dcnt2"))
	'onlyOne = request("onlyOne")
	'danpumcheck = request("danpumcheck")
	'imsi  = request("imsi")
	'sagawa= request("sagawa")
	'ems   = request("ems")
	'epostmilitary   = request("epostmilitary")
	page = request("page")
	mwdiv = requestCheckVar(request("mwdiv"), 32)

if deliveryarea = "" then deliveryarea = "KR"
if (page="") then page=1

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

if deliveryarea<>"" then
	if (deliveryarea = "ZZ") then
		ems   = ""
		epostmilitary   = "on"
	elseif (deliveryarea = "EMS") then
		ems   = "on"
		epostmilitary   = ""
	else
		deliveryarea = "KR"
		ems   = ""
		epostmilitary   = ""
	end if
end if

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

if (pagesize="") then
	pagesize = request.cookies("offlinebaljupagesize")
end if

if (pagesize="") then pagesize=1000

response.cookies("offlinebaljupagesize") = pagesize

if (research = "") then
	notitemlistinclude = "on"
	notbrandlistinclude = "on"

	''includeminus		= "N"
	includezerostock	= "N"
end if

set ojumun = new CTenBaljuOffline
	ojumun.FRectCompanyId = companyid
	ojumun.FRectLocationid3PL = locationid3PL
	ojumun.FRectLocationidTo = locationidto
	ojumun.FPageSize = pagesize

	if notitemlistinclude="on" then
		ojumun.FRectNotIncludeItem = "Y"
	else
		ojumun.FRectNotIncludeItem = ""
	end if

	if itemlistinclude="on" then
		ojumun.FRectIncludeItem = "Y"
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

	''사가와 배송권역
	ojumun.FRectOnlySagawaDeliverArea = sagawa

	''업체배송 포함 주문건.
	ojumun.FRectUpbeaInclude = upbeaInclude

	if tenbeaonly<>"" then
		ojumun.FRectTenbeaOnly = "Y"
	end if

	if deliveryarea<>"" then
		ojumun.FRectDeliveryArea = deliveryarea
	end if

	if bigitem<>"" then
		ojumun.FRectOnlyManyItem = "Y"
	end if

	if onejumuntype<>"" then
		ojumun.FRectOnlyOneJumun = "Y"

		ojumun.FRectOnlyOneJumunType = onejumuntype
		ojumun.FRectOnlyOneJumunCompare = onejumuncompare
		ojumun.FRectOnlyOneJumunCount = onejumuncount
	end if

	ojumun.FRectIncludeMinus		= includeminus
	ojumun.FRectIncludeZeroStock	= includezerostock
	ojumun.FRectMWDiv				= mwdiv
	ojumun.FRectMakerid				= makerid
	ojumun.GetBaljuItemListNewOffline

dim tenbaljucount
tenbaljucount =0

dim iStartDate, iEndDate

iStartDate  = Left(CStr(DateAdd("d",now(),-2)),10)
iEndDate    = Left(CStr(DateAdd("d",now(),+4)),10)

%>

<script language='javascript'>

var tenBaljuCnt = 0;

function CheckNBalju(){
	var frm;
	var pass = false;
	var upfrm = document.frmArrupdate;
    var isDasBalju = false;
    var isEmsBalju = <%= chkIIF(ems="on","true","false")%>;
    var isMilitaryBalju = <%= chkIIF(epostmilitary="on","true","false")%>;
    var locationidto;

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

	if (document.all.frm.deliveryarea.value.length<1){
		alert('배송지역을 선택 하세요.');
		document.all.frm.deliveryarea.focus();
		return;
	}

	if (document.all.frm.deliveryarea.value == "EMS") {
		if (document.all.groupform.songjangdiv.value != "90"){
			if (confirm("해외출고 > EMS 이외 택배사 선택!!\n\n그대로 진행하시겠습니까?") != true) {
				return;
			}
		}
	}

	if (document.all.frm.deliveryarea.value != "EMS") {
		if (document.all.groupform.songjangdiv.value == "90"){
			alert('해외배송만 EMS 배송을 선택할 수 있습니다.');
			document.all.frm.deliveryarea.focus();
			return;
		}
	}

    if (document.all.groupform.pickingStationCd.value == '') {
        alert('피킹스테이션을 선택하세요.');
        return;
    }

	// ========================================================================
	var iszerobaljono, currordercode;

	upfrm.masteridx.value = "";
	upfrm.ordercode.value = "";
	upfrm.detailidx.value = "";
	upfrm.baljuno.value = "";

	upfrm.comment.value = "";
	upfrm.errorcd.value = "";

	currordercode = "";
	iszerobaljono = true;
	locationidto = "";
	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (frm.cksel.checked){
				if ((currordercode != "") && (currordercode != frm.ordercode.value)) {
					if (iszerobaljono == true) {
						alert("출고지시수량이 없는 주문이 있습니다.[" + currordercode + "]");
						return;
					}
					iszerobaljono = true;
				}
				currordercode = frm.ordercode.value;

				if (locationidto == "") {
					locationidto = frm.locationidto.value;
				} else {
					if (locationidto != frm.locationidto.value) {
						alert("여러샵주문을 한번에 동시에 출고지시할 수 없습니다.");
						return;
					}
				}

				upfrm.masteridx.value = upfrm.masteridx.value + "|" + frm.masteridx.value;
				upfrm.ordercode.value = upfrm.ordercode.value + "|" + frm.ordercode.value;
				upfrm.detailidx.value = upfrm.detailidx.value + "|" + frm.detailidx.value;
				upfrm.baljuno.value = upfrm.baljuno.value + "|" + frm.baljuno.value;

				upfrm.comment.value = upfrm.comment.value + "|" + frm.comment.value;
				upfrm.errorcd.value = upfrm.errorcd.value + "|" + frm.errorcd.value;

				if (frm.baljuno.value*1 != 0) {
					iszerobaljono = false;
				}
			}
		}
	}
	if (iszerobaljono == true) {
		alert("출고지시수량이 없는 주문이 있습니다.[" + currordercode + "]");
		return;
	}
	upfrm.songjangdiv.value = document.all.groupform.songjangdiv.value;
	upfrm.workgroup.value = document.all.groupform.workgroup.value;
    upfrm.pickingStationCd.value = document.all.groupform.pickingStationCd.value;
	upfrm.ems.value = "<%= ems %>";
	upfrm.epostmilitary.value = "<%= epostmilitary %>";

	//var count = (upfrm.masteridx.value.match(/\|/g) || []).length;
	//if (count > 410) {
	//	alert("너무 많은 주문을 선택했습니다. 400개 이하로 주문선택 후 출고지시하세요.");
	//	return;
	//}

	// ========================================================================
	var ret = confirm('선택 주문을 새 출고지시서로 저장하시겠습니까?');
	if (ret) {
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
		}else if (icomp.name=="itemlistinclude"){
			frm.notitemlistinclude.checked = !(ischecked);
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
		frm.notitemlistinclude.disabled = (ischecked);

		frm.danpumcheck.checked = false;
	}


	if (icomp.name=="danpumcheck"){
		frm.itemlistinclude.disabled = (ischecked);
		frm.notitemlistinclude.disabled = (ischecked);

		frm.onlyOne.checked = false;
	}
}

function poponeitem(){
	var popwin = window.open("poponeitem.asp","poponeitem","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function poponebrand(){
	var popwin = window.open("poponebrand.asp","poponebrand","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popViewOrderSheet(idx){
	var popwin = window.open("/admin/fran/jumuninputedit.asp?idx=" + idx,"popViewOrderSheet","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popViewRelatedOrderSheet(itemgubun, itemid, itemoption) {
	var popwin = window.open("popBaljuRelatedOrderList.asp?itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption,"popViewRelatedOrderSheet","width=1200 height=600 scrollbars=yes resizable=yes");
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

function chkAllitem(masteridx, ischecked) {
    var frm;

	for (var i = 0; ; i++) {
		frm = document.getElementById("frmBuyPrc_" + i);

		if (frm === null) {
			break;
		}

		if (frm.masteridx.value*1 == masteridx){
			frm.cksel.checked = ischecked;
			AnCheckClick(frm.cksel);
		}
	}
}

function ckAllLimit(icomp, maxnum) {
	var bool = icomp.checked;
	var frm, currnum = 0;
	var alartShowed = false;

	if (bool === false) {
		AnSelectAllFrame(bool);
	} else {
		for (var i = 0; i < document.forms.length; i++) {
			frm = document.forms[i];
			if (frm.name.substr(0,9) == "frmBuyPrc") {
				if (frm.cksel.disabled != true) {
					//if (currnum >= maxnum) {
					//	if (alartShowed == false) {
					//		alert("\n\n한번에 " + maxnum + "개 이상을 선택하여 출고지시할 수 없습니다.\n\n");
					//		alartShowed = true;
					//	}
					//	frm.cksel.disabled = true;
					//} else {
						if (frm.cksel.checked === true) {
							currnum = currnum + 1;
						} else {
							chkAllitem(frm.masteridx.value, bool);
							currnum = currnum + 1;
						}
					//}
				}
			}
		}
	}
	// AnSelectAllFrame(bool);
}

// 주문수량 & 출고지시수량 사유별 액션
function chkRealItemNo(tn){
	var frm = eval("frmBuyPrc_"+ tn);
	var v = frm.baljuno;
	var vv = frm.requestedno;

	if (isNaN(v.value)||v.value.length<1){
		return;
	}else{
		v.value = parseInt(v.value);
	}

	if (eval("seldiv" + tn) == false) {
		// 수정가능상태 아님.
		return;
	}

	// 주문수량과 출고지시수량이 다르면 표시
	if(parseInt(v.value) != parseInt(vv.value)) {
		ShowHide("seldiv" + tn, true);
		fnselcom(eval("frmBuyPrc_" + tn + ".dtstat.value"), tn);
	} else {
		ShowHide("seldiv" + tn, false);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "";
		frm.dtstat.selectedIndex = 0;
		frm.comment.readOnly = true;
		frm.errorcd.value = "";
	}
}

function ShowHide(divid, isshow) {

	if (isshow == true) {
		document.getElementById(divid).style.display='';
	} else {
		document.getElementById(divid).style.display='none';
	}

}

//사유별 표시
function fnselcom(val, tn) {

	var frm = eval("frmBuyPrc_"+ tn);

	if(val=='ipt'){
		// ====================================================================
		// 직접입력
		ShowHide("seldiv" + tn, true);
		ShowHide("comdiv" + tn, true);

		frm.comment.value = "";
		frm.comment.readOnly = false;

		frm.errorcd.value = "C";
	}else if(val=='sso'){
		// ====================================================================
		// 일시품절
		ShowHide("seldiv" + tn, true);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "일시품절";
		frm.comment.readOnly = true;

		frm.errorcd.value = "T";
	}else if(val=='5day'){
		// ====================================================================
		// 5일내출고
		ShowHide("seldiv" + tn, true);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "5일내출고";
		frm.comment.readOnly = true;

		frm.errorcd.value = "C";
	}else if(val=='jaego'){
		// ====================================================================
		// 재고부족
		ShowHide("seldiv" + tn, true);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "재고부족";
		frm.comment.readOnly = true;

		frm.errorcd.value = "C";
	}else if(val=='so'){
		// ====================================================================
		// 단종
		ShowHide("seldiv" + tn, true);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "단종";
		frm.comment.readOnly = true;

		frm.errorcd.value = "E";
	}else{
		// ====================================================================
		// 에러
		ShowHide("seldiv" + tn, false);
		ShowHide("comdiv" + tn, false);

		frm.comment.value = "";
		frm.comment.readOnly = true;

		frm.errorcd.value = "";
	}
}

</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="companyid" value="<%= companyid %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<b>업체</b> : <b><%= companyid %></b>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<b>배송지역</b> :
		<select name="deliveryarea" >
			<option value="" 	<% if deliveryarea="" then response.write "selected" %> >전체</option>
			<option value="KR" 	<% if deliveryarea="KR" then response.write "selected" %> >국내배송</option>
			<option value="EMS" <% if deliveryarea="EMS" then response.write "selected" %> >해외배송</option>
			<!--
			<option value="ZZ" 	<% if deliveryarea="ZZ" then response.write "selected" %> >군부대배송</option>
			-->
		</select>
		<input type="checkbox" name="includeminus" value="N" <% if (includeminus = "N") then %>checked<% end if %>> 마이너스주문제외
		<input type="checkbox" name="includezerostock" value="N" <% if (includezerostock = "N") then %>checked<% end if %>> 전체재고없는주문제외
		<input type="checkbox" name="includeminu11s" value="N"> 온라인7일판매분제외수량만
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<b>기간 : <% DrawOneDateBox yyyy1,mm1,dd1 %> ~ 현재
		&nbsp;
		매장 :
		<% 'drawSelectBoxOffShop "locationidto",locationidto %>
		<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("locationidto",locationidto, "21") %>
		&nbsp;
		브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		배송구분 :
		<select class="select" name="mwdiv">
			<option value="">전체</option>
			<option value="U" <% if mwdiv="U" then response.write "selected" %> >업배</option>
			<option value="T" <% if mwdiv="T" then response.write "selected" %> >텐배</option>
			<option value="O" <% if mwdiv="O" then response.write "selected" %> >오프</option>
		</select>
		<!--
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<input type="checkbox" name="bigitem" <% if bigitem="on" then response.write "checked" %> > <b>다수상품주문</b>
		<font color="#AAAAAA">
		<input type="checkbox" name="upbeaInclude" <% if upbeaInclude="on" then response.write "checked" %> > <b>업배포함 주문건만</b>
		</font>
		<input type="checkbox" name="tenbeaonly" <% if tenbeaonly="on" then response.write "checked" %> > <b>텐배주문건만</b>
		<input type="checkbox" name="notitemlistinclude" <% if notitemlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
		<b>위탁상품 제외 주문만</b>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<input type="checkbox" name="itemlistinclude" <% if itemlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
		<b>위탁상품 포함 주문만</b>
		<input type="checkbox" name="notbrandlistinclude" <% if notbrandlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
		<b>위탁매입처 제외 주문만</b>
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<input type="checkbox" name="brandlistinclude" <% if brandlistinclude="on" then response.write "checked" %> onclick="EnableDiable(this);">
		<b>위탁매입처 포함 주문만</b>
		<b>단품주문</b> :
		<select name="onejumuntype" >
		<option value="" 	<% if onejumuntype="" then response.write "selected" %> >========</option>
		<option value="all" <% if onejumuntype="all" then response.write "selected" %> >모든 단품주문</option>
		<option value="reg" <% if onejumuntype="reg" then response.write "selected" %> >설정된 단품주문</option>
		</select>

		<input type="text" name="onejumuncount" value="<%= onejumuncount %>" size=3>
		<select name="onejumuncompare" >
		<option value="less" 	<% if onejumuncompare="less" then response.write "selected" %> >개 이하</option>
		<option value="more" 	<% if onejumuncompare="more" then response.write "selected" %> >개 이상</option>
		<option value="equal" 	<% if onejumuncompare="equal" then response.write "selected" %> >개</option>
		</select>
		<input type="checkbox" name="onlyOne" <% if onlyOne="on" then response.write "checked" %> onclick="EnableDiable(this);">
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<input type="checkbox" name="danpumcheck" <% if danpumcheck="on" then response.write "checked" %> onclick="EnableDiable(this);">
		<input type="button" value="제외/포함/단품 상품설정" onclick="javascript:poponeitem();">
		&nbsp;&nbsp;|&nbsp;&nbsp;
		<input type="button" value="제외/포함 매입처설정" onclick="javascript:poponebrand();">
		<input type="text" name="dcnt2" value="<%= dcnt2 %>" size=1> 개 (11 입력시 11개 이상, 0개 입력시 0개 이상)
		<input type="checkbox" name="ems" <% if ems="on" then response.write "checked" %> > <b>해외배송</b>
		<input type="checkbox" name="epostmilitary" <% if epostmilitary="on" then response.write "checked" %> > <b>군부대</b>
		<input type="checkbox" name="imsi" <% if imsi="on" then response.write "checked" %> > <b>임시(무한도전 포함 복함)</b>
		<font color="#AAAAAA">
		<input type="checkbox" name="sagawa" <% if sagawa="on" then response.write "checked" %> onClick="alert('일반출고지시만 가능 (단품출고,위탁상품 검색 적용안됨)');"> 임시(사가와권역)
		</font>
		-->
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<form name="groupform">
<tr>
	<td align="left">
        총 미출고지시 건수 : <Font color="#3333FF"><b><%= FormatNumber(ojumun.FTotalCount,0) %></b></font>&nbsp;
		총 금액 : <Font color="#3333FF"><%= FormatNumber(ojumun.FSubTotalsum,0) %></font>&nbsp;
		평균객단가 : <Font color="#3333FF"><%= FormatNumber(ojumun.FAvgTotalsum,0) %></font><br>
		* 여러건의 주문에서 같은 상품을 주문할 경우 출고가능수량이 달라져야 한다.
	</td>
	<td align="right">
	    <!--
	    <select name="baljutype">
	        <option value="">일반
	        <option value="D">DAS
	        <option value="S">단품출고
	    </select>
	    -->
	    <select name="songjangdiv">
	        <option value="">택배사선택</option>
			<option value="1" >한진택배</option>
			<option value="2" >롯데택배</option>
            <option value="4" >CJ택배</option>
		   	<option value="90" <%= CHKIIF(ems="on","selected","") %> >EMS</option>
			<option value="91" >DHL</option>
		   	<option value="98" >퀵서비스</option>
		   	<option value="99" >기타</option>
		   	<!--
		   	<option value="8" <%= CHKIIF(epostmilitary="on","selected","") %> >우체국(군부대)
		   	-->
	    </select>
		<select name="workgroup">
		   	<option value="">작업그룹
		   	<option value="O" >O(오프라인)
	   	</select>
        <% Call drawSelectStationByStationGubun("PICK", "pickingStationCd", "") %>
		<!--
		<select name="workgroup">
		   	<option value="">작업그룹
		   	<option value="A" >A
		   	<option value="B" >B
		   	<option value="C" >C(DAS)
		   	<option value="D" >D
		   	<option value="F" >F
		   	<option value="" >===========
		   	<option value="T" >T(탐스슈즈)
		   	<option value="" >===========
		   	<option value="I" >I(아이띵소)
		   	<option value="" >===========
		   	<option value="E" <%= CHKIIF(ems="on","selected","") %> >E(EMS)
		   	<option value="G" <%= CHKIIF(epostmilitary="on","selected","") %> >G(군부대)
		   	<option value="Z" >Z(업배)
	   	</select>
	   	-->
		<input type="button" value="선택사항출고지시서작성" onclick="CheckNBalju()" class="button">
	</td>
</tr>
</form>
</table>
<!-- 액션 끝 -->

<br>



<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="22">
        <div id="currsearchno">총 검색주문건수 : </div>
        <!--
        <input type="checkbox" name="ck_upbea" onClick="chkUpbea();"> 업배출고지시
        -->
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20" align="center"><input type="checkbox" name="cksel" onClick="ckAllLimit(this, 400)"></td>
	<td width="80">입출코드</td>
	<td width="40">국가</td>
	<td>UserID<br>샵이름</td>
	<td width="100">물류코드</td>
	<td>브랜드ID</td>
	<td align="left">상품명<br><font color=blue>[옵션명]</font></td>
	<td width="40">배송<br>구분</td>
	<!--
	<td width="60">실사<br>수량</td>
	<td width="60">기출고지시<br>(ON+OFF)</td>
	<td width="60">온라인<br>결재</td>
	<td width="60">N일필요<br>수량</td>
	<td width="60">주문<br>수량</td>
	<td width="60">출고가능<br>수량</td>
	-->

	<td width="60">실사<br>유효재고</td>
	<td width="60">ON<br>상품준비</td>
	<td width="60">OFF<br>상품준비</td>
	<td width="60">ON<br>결제완료</td>
	<td width="60">ON<br>주문접수</td>
	<td width="60">출고<br>가능수량</td>
	<td width="60">주문수량</td>

	<td width="60">출고지시<br>수량</td>
	<td width="240">비고</td>
	<td width="40">관련<br />주문</td>
</tr>
<% if ojumun.FresultCount>0 then %>
<% for ix=0 to ojumun.FresultCount-1 %>
<form name="frmBuyPrc_<%= ix %>" id="frmBuyPrc_<%= ix %>" method="post" >
<input type="hidden" name="masteridx" value="<%= ojumun.FItemList(ix).Fmasteridx %>">
<input type="hidden" name="ordercode" value="<%= ojumun.FItemList(ix).Fordercode %>">
<input type="hidden" name="detailidx" value="<%= ojumun.FItemList(ix).Fdetailidx %>">
<input type="hidden" name="locationidto" value="<%= ojumun.FItemList(ix).Flocationidto %>">
<tr align="center" bgcolor="#FFFFFF">
    <% if ((ems<>"") or (epostmilitary<>"")) then %>
    	<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this); chkAllitem(<%= ojumun.FItemList(ix).Fmasteridx %>, this.checked)"></td>
    <% else %>
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this); chkAllitem(<%= ojumun.FItemList(ix).Fmasteridx %>, this.checked)" <%= CHKIIF(ojumun.FItemList(ix).FcountryCode<>"" and ojumun.FItemList(ix).FcountryCode<>"KR","disabled","") %> ></td>
	<% end if %>
		<td><a href="javascript:popViewOrderSheet(<%= ojumun.FItemList(ix).Fmasteridx %>)"><%= ojumun.FItemList(ix).Fordercode %></a></td>
	<td><%= ojumun.FItemList(ix).FcountryCode %></td>
	<td><%= ojumun.FItemList(ix).Flocationidto %><br><%= ojumun.FItemList(ix).Flocationnameto %></td>
	<td><%= ojumun.FItemList(ix).Fprdcode %></td>
	<td><%= ojumun.FItemList(ix).Fbrandid %></td>
	<td align="left">
		<a href="/admin/stock/itemcurrentstock.asp?itemgubun=<%= BF_GetItemGubun(ojumun.FItemList(ix).Fprdcode) %>&itemid=<%= BF_GetItemId(ojumun.FItemList(ix).Fprdcode) %>&itemoption=<%= BF_GetItemOption(ojumun.FItemList(ix).Fprdcode) %>" target=_blank >
			<%= ojumun.FItemList(ix).Fprdname %><br><font color=blue>[<%= ojumun.FItemList(ix).Fitemoptionname %>]</font>
		</a>

	</td>
	<td>
		<% if ojumun.FItemList(ix).Fmwdiv = "M" or ojumun.FItemList(ix).Fmwdiv = "W" then %>
			<%= ojumun.FItemList(ix).GetMWDivName %>
		<% elseif ojumun.FItemList(ix).Fmwdiv = "U" then %>
			<font color="red"><%= ojumun.FItemList(ix).GetMWDivName %></font>
		<% elseif ojumun.FItemList(ix).Fmwdiv = "O" then %>
			<font color="blue"><%= ojumun.FItemList(ix).GetMWDivName %></font>
		<% end if %>
	</td>
	<!--
	<td><%= ojumun.FItemList(ix).Frealstockno %></td>
	<td><%= (ojumun.FItemList(ix).Fmoveoutdiv5 + ojumun.FItemList(ix).Fmoveoutdiv7 + ojumun.FItemList(ix).Fselldiv5 + ojumun.FItemList(ix).Fchulgodiv5) %></td>
	<td><%= ojumun.FItemList(ix).Fselldiv4 %></td>
	<td><%= ojumun.FItemList(ix).GetOnlineRequireNo %></td>
	<td><%= ojumun.FItemList(ix).Frequestedno %></td>
	<td><%= ojumun.FItemList(ix).GetChulgoAvailableNo %></td>
	-->

	<input type="hidden" name="requestedno" value="<%= ojumun.FItemList(ix).Frequestedno %>">
	<td><%= ojumun.FItemList(ix).Frealstockno %></td>
	<td><%= ojumun.FItemList(ix).Fselldiv5 %></td>
	<td><%= ojumun.FItemList(ix).Fchulgodiv5 %></td>
	<td><%= ojumun.FItemList(ix).Fselldiv4 %></td>
	<td><%= ojumun.FItemList(ix).Fselldiv2 %></td>
	<td>
		<b><%= ojumun.FItemList(ix).GetOffChulgoAvailableNo %></b>
	</td>
	<td>
		<b><font color="<%= CHKIIF((ojumun.FItemList(ix).GetOffChulgoAvailableNo-ojumun.FItemList(ix).Frequestedno) < 0, "red", "black") %>"><%= ojumun.FItemList(ix).Frequestedno %><%= CHKIIF(ojumun.FItemList(ix).Frequestedno<>ojumun.FItemList(ix).Foffjupno*-1, "/" & ojumun.FItemList(ix).Foffjupno*-1, "") %></font></b>
	</td>

	<td>
		<% if (ojumun.FItemList(ix).Frequestedno > 0) and (ojumun.FItemList(ix).GetOffChulgoAvailableNo > 0) then %>
			<% if ojumun.FItemList(ix).Frequestedno > ojumun.FItemList(ix).GetOffChulgoAvailableNo then %>
				<input type=text class="text" name=baljuno size=4 value="<%= ojumun.FItemList(ix).GetOffChulgoAvailableNo %>"  onKeyup="chkRealItemNo(<%= ix %>);">
			<% else %>
				<input type=text class="text" name=baljuno size=4 value="<%= ojumun.FItemList(ix).Frequestedno %>"  onKeyup="chkRealItemNo(<%= ix %>);">
			<% end if %>
		<% else %>
			<input type=text class="text" name=baljuno size=4 value="0"  onKeyup="chkRealItemNo(<%= ix %>);">
		<% end if %>
	</td>
	<input type="hidden" name="errorcd" value="">
	<td>
		<span id="seldiv<%= ix %>">
				<select class="select" name="dtstat" onchange="fnselcom(this.value,<%= ix %>);">
					<option value="ipt">직접입력</option>
					<option value="5day" <%= CHKIIF(ojumun.FItemList(ix).Fdanjongyn="M", "selected", "") %> >5일내출고</option>
					<option value="jaego" <%= CHKIIF((ojumun.FItemList(ix).Fdanjongyn="S" or ojumun.FItemList(ix).Fdanjongyn = "Y"), "selected", "") %> >재고부족</option>
					<option value="so">단종</option>
					<option value="sso">일시품절</option>
				</select>
		</span>
		<script>ShowHide("seldiv<%= ix %>", false)</script>
		<span id="comdiv<%= ix %>">
			<input type="text" class="text" name="comment" value="" size="8"><br>
		</span>
		<script>ShowHide("comdiv<%= ix %>", false)</script>

		<% if ((Not IsNull(ojumun.FItemList(ix).FreipgoMayDate)) and (Left(ojumun.FItemList(ix).FreipgoMayDate, 10) >= iStartDate) and (Left(ojumun.FItemList(ix).FreipgoMayDate, 10) <= iEndDate)) then %>
			<%= Left(ojumun.FItemList(ix).FreipgoMayDate, 10) %>
		<% end if %>
		<%= fnColor(ojumun.FItemList(ix).Fdanjongyn,"dj") %>
		<% if ojumun.FItemList(ix).Fpreorderno<>0 then %>
			기주문:
			<% if ojumun.FItemList(ix).Fpreorderno<>ojumun.FItemList(ix).Fpreordernofix then response.write CStr(ojumun.FItemList(ix).Fpreorderno) + "->" %>
			<%= ojumun.FItemList(ix).Fpreordernofix %>
		<% end if %>
	</td>
	<td><a href="javascript:popViewRelatedOrderSheet('<%= ojumun.FItemList(ix).Fitemgubun %>', <%= ojumun.FItemList(ix).Fitemid %>, '<%= ojumun.FItemList(ix).Fitemoption %>')">보기</a></td>
</tr>
</form>
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="22" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
<form name="frmArrupdate" method="post" action="baljumaker_offline_new_process.asp">
	<input type="hidden" name="mode" value="arr">
	<input type="hidden" name="masteridx" value="">
	<input type="hidden" name="ordercode" value="">
	<input type="hidden" name="detailidx" value="">
	<input type="hidden" name="baljuno" value="">
	<input type="hidden" name="comment" value="">
	<input type="hidden" name="errorcd" value="">
	<input type="hidden" name="songjangdiv" value="">
	<input type="hidden" name="workgroup" value="">
	<input type="hidden" name="ems" value="">
	<input type="hidden" name="epostmilitary" value="">
    <input type="hidden" name="pickingStationCd" value="">
</form>
</table>

<script language='javascript'>

<% ''if (locationidto <> "") then %>

	for (var i = 0; i < <%= ojumun.FresultCount %>; i++) {
		chkRealItemNo(i);
	}

<% ''end if %>

	document.all.currsearchno.innerHTML = "검색갯수 : <Font color='#3333FF'><%= ix %></font>";
	tenBaljuCnt = 1*<%= tenbaljucount %>;

	<% if onlyOne<>"" then %>
		EnableDiable(frm.onlyOne);
	<% end if %>

	<% if danpumcheck<>"" then %>
		EnableDiable(frm.danpumcheck);
	<% end if %>



</script>

<%
set ojumun = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_logisticsclose.asp" -->
