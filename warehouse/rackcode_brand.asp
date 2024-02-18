<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 브랜드렉코드관리
' Hieditor : 이상구 생성
'			 2020.01.09 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/rackipgocls.asp"-->
<%
dim isusing, research, page
dim rackcode2, rackcode, makerid
dim maeipdiv
dim searchtype, fromrackcode2, torackcode2, purchasetype
dim warehouseCd

page        = request("page")
isusing     = request("isusing")
research    = request("research")
rackcode2   = requestCheckvar(request("rackcode2"),4)
maeipdiv  	= requestCheckvar(request("maeipdiv"),1)

searchtype  	= requestCheckvar(request("searchtype"),1)
fromrackcode2  	= requestCheckvar(request("fromrackcode2"),4)
torackcode2  	= requestCheckvar(request("torackcode2"),4)
purchasetype  	= requestCheckvar(request("purchasetype"),2)
warehouseCd  	= requestCheckvar(request("warehouseCd"),3)

makerid     = request("makerid")

if research="" and isusing="" then isusing="Y"
if page="" then page=1
'if searchtype="" then searchtype = "F"

dim orackcode_brand
set orackcode_brand = new CRackIpgo
orackcode_brand.FCurrpage = page
orackcode_brand.FPageSize = 30
orackcode_brand.FRectMakerid = makerid
orackcode_brand.FRectRackCode = rackcode2
orackcode_brand.FRectIsUsingYN = isusing
orackcode_brand.FRectMaeipDiv = maeipdiv
orackcode_brand.FRectSearchType = searchtype
orackcode_brand.FRectFromRackcode2 = fromrackcode2
orackcode_brand.FRectToRackcode2 = torackcode2
orackcode_brand.FRectPurchaseType = purchasetype
orackcode_brand.FRectWarehouseCd = warehouseCd

orackcode_brand.GetRackBrandList

dim i
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript" src="/js/ttpbarcode.js"></script>
<script type="text/javascript" src="/js/barcode.js"></script>
<script type="text/javascript" src="/js/DOSHIBAbarcode.js"></script>
<script type='text/javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popBrandRackCodeEdit(imakerid){
    var popwin = window.open('pop_BrandRackCodeEdit.asp?makerid=' + imakerid,'popBrandRackCodeEdit','width=500,height=200,scrollbars=yes,resizable=yes');
    popwin.focus();
}

// 수기랙코드인덱스출력		// 2020.01.09 한용민 생성
function IndexSudongrackBarcodePrint(){
	var popwin = window.open('/common/barcode/sudongrackindexprint.asp?menupos=<%=menupos%>','IndexSudongrackBarcodePrint','width=1024,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function toggleChecked(status) {
    $('[name="check"]').each(function () {
        $(this).prop("checked", status);
    });
}

function IndexrackBarcodePrint() {
	var arrbarcode = new Array();
    var isforeignprint; var domainname; var showdomainyn; var showdomainyn; var ttptype;
    var barcodetype;
    var shopbrandyn; var currencychar; var showpriceyn;

	isforeignprint = "N";

	shopbrandyn		= "N";
	ttptype			= "TTP-243_80x50";
	barcodetype		= "";
	var paperwidth = "80";
	var paperheight = "50";
	var papermargin = "3";
	var heightoffset = 0;
	showpriceyn = "N";

	currencychar = "￦";
	domainname		= "www.10x10.co.kr";
	showdomainyn	= "Y";

	var fontName = "10X10";
    var rackcode = "";
	var itemno = document.frmArr.itemno.value;
	var MAX_MESSAGE_LENGTH = 8;

    if ($('input[name="check"]:checked').length == 0) {
        alert('선택 아이템이 없습니다.');
        return;
    }
    if (itemno=="" || itemno==0){
        alert("출력하실 수량을 입력해주세요.");
        frmArr.itemno.focus();
        return;
    }

	$('input[name="check"]:checkbox:checked').each(function () {
		rackcode = $(this).attr('rackcode');

		if (rackcode!=""){
			if (rackcode.replace("\r", "").length > MAX_MESSAGE_LENGTH) {
				alert("\n========== 에러 ==========\n\n출력메시지는 " + MAX_MESSAGE_LENGTH + "글자를 이상을 넘을 수 없습니다. ");
				return;
			}

			var v = new BarcodeDataClass_udong(rackcode,itemno,fontName);
			arrbarcode.push(v);
		}
	});

	//TEC B-FV4		//2020.01.09 한용민 생성
	if (TEC_DO3.IsDriver == 1){
		if (confirm("선택하신 랙코드 인덱스를 출력합니다.\n\nTEC B-FV4 로 출력하시겠습니까?") == true) {
			TOSHIBA_DOMAINNAME = domainname;
			TOSHIBA_SHOWDOMAINYN = showdomainyn;
			TOSHIBA_PAPERWIDTH = 800;
			TOSHIBA_PAPERHEIGHT = 500;
			TOSHIBA_PAPERMARGIN = 3;
			TOSHIBA_SHOWPRICEYN = showpriceyn;
			TOSHIBA_currencyChar = currencychar;
			TOSHIBA_SHOPBRANDYN = shopbrandyn;
			TOSHIBA_BARCODETYPE = barcodetype;

			printTOSHIBAMultiRackIndexLabel(arrbarcode);
		}

	// /js/barcode.js 참조
	}else if (initTTPprinter(ttptype, barcodetype, showdomainyn, domainname, showpriceyn, currencychar, shopbrandyn, papermargin, heightoffset) == true) {
		if (confirm("선택하신 랙코드 인덱스를 출력합니다.\n\nTTP-243 로 출력하시겠습니까?") == true) {
			printTTPMultiRackIndexLabel(arrbarcode);
		}

	}else {
	    alert("TTP-243(구)나 TEC B-FV4 드라이버를 설치해 주세요");
	}

	return;
}

function jsSetChecked(index) {
    var check, warehouseCd;

    check = document.getElementById('check' + index);
    check.checked = true;
    AnCheckClick(check);
}

function jsSetWarehouseCd() {

    var frmAct, mode, makeridArr, warehouseCdArr;
    var makerid, check, warehouseCd;

    frmAct = document.frmAct;
    makeridArr = "";
    warehouseCdArr = "";

    if ($('input[name="check"]:checked').length == 0) {
        alert('선택 아이템이 없습니다.');
        return;
    }

    for (var i = 0; ; i++) {
        check = document.getElementById('check' + i);
        makerid = document.getElementById('makerid' + i);
        warehouseCd = document.getElementById('warehouseCd' + i);

        if (check == undefined) { break; }
        if (check.checked != true) { continue; }

        if (warehouseCd.value == '') {
            alert('진열속성을 선택하세요.');
            warehouseCd.focus();
            return;
        }

        makeridArr = makeridArr + "," + makerid.value;
        warehouseCdArr = warehouseCdArr + "," + warehouseCd.value;
    }

    if (confirm('저장하시겠습니까?')) {
        frmAct.mode.value = "setwarehousecd";
        frmAct.makeridArr.value = makeridArr;
        frmAct.warehouseCdArr.value = warehouseCdArr;
        frmAct.submit();
    }
}

$(document).ready(function () {
    var checkAllBox = $("#ckall");

    checkAllBox.click(function () {
        var status = checkAllBox.prop('checked');
        toggleChecked(status);
    });
});

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		랙코드(4자리) :
		<input type="radio" name="searchtype" value="F" <% if (searchtype = "F") then %>checked<% end if %> >
		<input type="text" name=rackcode2 value="<%= rackcode2 %>" maxlength="4" size="4" class="text">
		&nbsp;
		<input type="radio" name="searchtype" value="R" <% if (searchtype = "R") then %>checked<% end if %> >
		<input type="text" name=fromrackcode2 value="<%= fromrackcode2 %>" maxlength="4" size="4" class="text">
		~
		<input type="text" name=torackcode2 value="<%= torackcode2 %>" maxlength="4" size="4" class="text">
		&nbsp;
		사용 : <% drawSelectBoxUsingYN "isusing", isusing %>
		&nbsp;
		온라인기본마진 : <% DrawBrandMWUCombo "maeipdiv", maeipdiv %>
		&nbsp;
		구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
        &nbsp;
		진열속성 :
        <select class="select" name="warehouseCd">
            <option></option>
            <option value="AGV" <%= CHKIIF(warehouseCd = "AGV", "selected", "") %>>AGV</option>
            <option value="BLK" <%= CHKIIF(warehouseCd = "BLK", "selected", "") %>>BLK</option>
            <option value="NUL" <%= CHKIIF(warehouseCd = "NUL", "selected", "") %>>미지정</option>
        </select>
	</td>
</tr>
</table>
</form>

<br>

<!-- 액션 시작 -->
<form name="frmArr" action="" method="get" style="margin:0px;">
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="랙코드변경" onClick="popBrandRackCodeEdit('');">
        &nbsp;
        <input type="button" class="button" value="진열속성변경" onClick="jsSetWarehouseCd();">
	</td>
	<td align="right">
		<input type="text" name="itemno" value="1" size=3 maxlength=5>
		<input type="button" class="button" value="선택랙코드인덱스출력" onClick="IndexrackBarcodePrint();">
		&nbsp;
		<input type="button" class="button" value="수기랙코드인덱스출력" onClick="IndexSudongrackBarcodePrint();">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20"><input type="checkbox" name="ckall" id="ckall" onclick="totalCheck()"></td>
	<td width="90">랙코드</td>
	<td width="180">브랜드ID</td>
    <td width="60">진열속성</td>
    <td>스트리트명</td>
	<td width="50">사용</td>
	<!--
	<td width="70">사용(제휴)</td>
	-->
	<td width="50">랙박스<br>수량</td>
	<td width="120">랙진열상품</td>
	<td width="120">반품대상상품</td>
	<td width="150">일시품절단종상품</td>
	<td width="120">일시품절상품</td>
</tr>
<% for i=0 to orackcode_brand.FResultCount - 1 %>
<tr <%= chkIIF( orackcode_brand.FItemList(i).FBrandUsing="Y","bgcolor='#FFFFFF'","bgcolor='#CCCCCC'") %>>
	<td align="center"><input type="checkbox" id="check<%= i %>" name="check" rackcode="<%= orackcode_brand.FItemList(i).Frackcode %>" onClick="AnCheckClick(this);"></td>
    <input type="hidden" id="makerid<%= i %>" name="makerid" value="<%= orackcode_brand.FItemList(i).FMakerid %>">
	<td align="center"><%= orackcode_brand.FItemList(i).Frackcode %></td>
	<td><a href="javascript:popBrandRackCodeEdit('<%= orackcode_brand.FItemList(i).FMakerid %>');"><%= orackcode_brand.FItemList(i).FMakerid %></a></td>
    <td align="center">
        <select class="select" id="warehouseCd<%= i %>" name="warehouseCd" onchange="jsSetChecked(<%= i %>)">
            <option></option>
            <option value="AGV" <%= CHKIIF(orackcode_brand.FItemList(i).FwarehouseCd = "AGV", "selected", "") %>>AGV</option>
            <option value="BLK" <%= CHKIIF(orackcode_brand.FItemList(i).FwarehouseCd = "BLK", "selected", "") %>>BLK</option>
        </select>
    </td>
    <td><%= orackcode_brand.FItemList(i).Fmakername %></td>
	<td align="center"><%= ChkIIF(orackcode_brand.FItemList(i).FBrandUsing="Y","O","X") %></td>
	<!--
	<td align="center"><%= ChkIIF(orackcode_brand.FItemList(i).FBrandUsingExt="Y","O","X") %></td>
	-->
	<td align="right"><%= orackcode_brand.FItemList(i).Frackboxno %></td>
	<td align="center">
		<input type="button" class="button" value="랙진열상품" onclick="javascript:window.open('/admin/stock/brandcurrentstock.asp?menupos=708&research=on&page=&makerid=<%= orackcode_brand.FItemList(i).FMakerid %>&onoffgubun=on&mwdiv=MW&returnitemgubun=rackdisp');">
	</td>
	<td align="center">
		<!-- <input type="button" class="button" value="반품대상상품" onclick="javascript:window.open('/admin/stock/return_item.asp?menupos=983&makerid=<%= orackcode_brand.FItemList(i).FMakerid %>&realstocknotzero=on');"> -->
		<input type="button" class="button" value="반품대상상품" onclick="javascript:window.open('/admin/stock/brandcurrentstock.asp?menupos=708&research=on&page=&makerid=<%= orackcode_brand.FItemList(i).FMakerid %>&onoffgubun=on&mwdiv=MW&returnitemgubun=reton');">
	</td>
	<td align="center">
		<input type="button" class="button" value="일시품절단종상품" onclick="javascript:window.open('/admin/stock/brandcurrentstock.asp?menupos=708&research=on&page=&makerid=<%= orackcode_brand.FItemList(i).FMakerid %>&onoffgubun=on&sellyn=S&usingyn=&danjongyn=YM&mwdiv=MW');">
	</td>

	<td align="center">
		<!-- 일시품절/단종제외 -->
		<input type="button" class="button" value="일시품절상품" onclick="javascript:window.open('/admin/shopmaster/danjong_set.asp?menupos=1053&research=on&page=1&makerid=<%= orackcode_brand.FItemList(i).FMakerid %>&mwdiv=MW');">
	</td>
</tr>
<% next %>

<tr bgcolor="#FFFFFF">
	<td colspan="11" align="center">
		<% if orackcode_brand.HasPreScroll then %>
		<a href="javascript:NextPage('<%= orackcode_brand.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + orackcode_brand.StarScrollPage to orackcode_brand.FScrollCount + orackcode_brand.StarScrollPage - 1 %>
			<% if i>orackcode_brand.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if orackcode_brand.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>
</form>

<%
set orackcode_brand = Nothing
%>

<form name="frmAct" method="post" action="brandRackCode_process.asp">
<input type="hidden" name="mode" />
<input type="hidden" name="makeridArr" />
<input type="hidden" name="warehouseCdArr" />
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
