<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  매장대 매장 재고이동
' History : 2018.02.07 이상구 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%

'// 금액수정 작업 안되어있음!!!, skyer9, 2018-02-12
dim PriceEditEnable : PriceEditEnable = False

'// sellcash : 판매가, buycash : 매입가, suplycash : 샵 공급가
'// 이외 다른 명칭은 사용하지 않는다.


dim i, j, k
dim shopid, moveshopid, makerid
dim itemgubunarr, itemidarr, itemoptionarr, sellcasharr, buycasharr, suplycasharr, itemnoarr, itemnamearr, itemoptionnamearr
dim scheduledt, songjangdiv, songjangno, comment
dim masteridx, mode
dim ojumunmaster, ojumundetail

masteridx  = requestCheckVar(request("masteridx"), 32)
mode  = requestCheckVar(request("mode"), 32)

shopid  = requestCheckVar(request("shopid"), 32)
moveshopid  = requestCheckVar(request("moveshopid"), 32)
makerid  = requestCheckVar(request("makerid"), 32)

scheduledt  = requestCheckVar(request("scheduledt"), 32)
songjangdiv  = requestCheckVar(request("songjangdiv"), 32)
songjangno  = requestCheckVar(request("songjangno"), 32)
comment  = request("comment")

if (C_IS_Maker_Upche) then
	makerid = session("ssBctID")
end if

set ojumunmaster = new COrderSheet
set ojumundetail = new COrderSheet

if (masteridx <> "") then

	ojumunmaster.FRectIdx = masteridx
	ojumunmaster.GetOneOrderSheetMaster_IMSI

	shopid = ojumunmaster.FOneItem.Ftargetid
	moveshopid = ojumunmaster.FOneItem.Fbaljuid

	if (scheduledt = "") then
		scheduledt = ojumunmaster.FOneItem.Fscheduledate
	end if

	if (songjangdiv = "") then
		songjangdiv = ojumunmaster.FOneItem.Fsongjangdiv
	end if

	if (songjangno = "") then
		songjangno = ojumunmaster.FOneItem.Fsongjangno
	end if

	if (comment = "") then
		comment = ojumunmaster.FOneItem.Fcomment
	end if

	ojumundetail.FRectIdx = masteridx
	ojumundetail.GetOrderSheetDetail_IMSI
end if


if (shopid = "") then
	response.write "잘못된 접근입니다."
	dbget.close : response.end
end if


dim IsAddItemOK : IsAddItemOK = False
dim errMsg : errMsg = ""

if (shopid <> "") and (moveshopid <> "") and (makerid <> "") then
	IsAddItemOK = IsSameShopContract(shopid, moveshopid, makerid)
	if (IsAddItemOK <> True) then
		errMsg = "계약이 없거나 두매장의 브랜드 계약마진이 다릅니다."
		response.write "<script>alert('" & errMsg & "')</script>"
	end if
end if


if (mode = "additem") then
	'
end if

%>
<script>
function jsSetShopidMove() {
	var frm = document.frm;
	if (frm.moveshopidsel.value == "") {
		alert("먼저 매장(도착매장)을 선택하세요.");
		return;
	}

	if (frm.shopid.value == frm.moveshopidsel.value) {
		alert("에러!!\n\n출발매장과 도착매장이 동일합니다.");
		return;
	}

	frm.moveshopid.value = frm.moveshopidsel.value;
	frm.moveshopidsel.value = "";

	frm.submit();
}

function jsSetMakerid() {
	var frm = document.frm;
	if (frm.makeridsel.value == "") {
		alert("먼저 브랜드를 선택하세요.");
		return;
	}

	if (frm.makerid.value == frm.makeridsel.value) {
		alert("에러!!\n\n동일한 브랜드입니다.");
		return;
	}

	frm.makerid.value = frm.makeridsel.value;
	frm.makeridsel.value = "";

	frm.submit();
}

function jsChkForm(frm) {
	if (frm.shopid.value == "") {
		alert("에러!!\n\n출발매장 지정안됨.");
		return false;
	}

	if (frm.moveshopid.value == "") {
		alert("에러!!\n\n도착매장 지정안됨.");
		return false;
	}

	if (frm.scheduledt.value.length<1){
		alert('재고이동일을 입력하세요');
		calendarOpen3(frm.scheduledt,'재고이동일을 입력하세요','');
		return false;
	}

	if (frm.songjangdiv.value.length<1){
		alert('택배사를 선택 하세요');
		frm.songjangdiv.focus();
		return false;
	}

	if (frm.songjangno.value.length<1){
		alert('송장 번호를 입력 하세요');
		msfrm.songjangno.focus();
		return false;
	}

	return true;
}

function jsStockMove() {
	var frm = document.frm;

	if (jsChkForm(frm) != true) { return; }

	<% if (ojumundetail.FResultCount < 1) then %>
	if (frm.itemgubunarr.value == "") {
		alert("추가된 상품이 없습니다.");
		return;
	}
	<% end if %>

	var ret = confirm('입력하신대로 재고 이동처리 하시겠습니까?');
	if (ret) {
		frm.mode.value = "saveorder";
		frm.method = "post";
		frm.action = "pop_jumun_move_process.asp";
		frm.submit();
	}
}

function jsSaveModified() {
	var frm = document.frm;

	if (jsChkForm(frm) != true) { return; }

	frm.itemgubunarr.value = "";
	frm.itemidarr.value = "";
	frm.itemoptionarr.value = "";
	frm.itemnoarr.value = "";

	for (var i = 0; i < document.forms.length ; i++){
		o = document.forms[i];
		if (o.name.substr(0,9)=="frmBuyPrc") {
			if (!IsInteger(o.itemno.value)){
				alert('갯수는 정수만 가능합니다.');
				o.itemno.focus();
				return;
			}

		    if (o.itemno.value < 0){
				alert("갯수는 0이상만 허용 됩니다.");
				o.itemno.focus();
				return;
			}

			frm.itemgubunarr.value = frm.itemgubunarr.value + o.itemgubun.value + "|";
			frm.itemidarr.value = frm.itemidarr.value + o.itemid.value + "|";
			frm.itemoptionarr.value = frm.itemoptionarr.value + o.itemoption.value + "|";
			frm.itemnoarr.value = frm.itemnoarr.value + o.itemno.value + "|";
			frm.sellcasharr.value = frm.sellcasharr.value + o.sellcash.value + "|";
			frm.buycasharr.value = frm.buycasharr.value + o.buycash.value + "|";
			frm.suplycasharr.value = frm.suplycasharr.value + o.suplycash.value + "|";
		}
	}

	var ret = confirm('저장하시겠습니까?');
	if (ret) {
		frm.mode.value = "additem";
		frm.method = "post";
		frm.action = "pop_jumun_move_process.asp";
		frm.submit();
	}
}

function jsButtonDisabled() {
	document.getElementById("btnMove").disabled = true;
}

function jsAddItems() {
	var frm = document.frm;

	if (frm.shopid.value == "") {
		alert("에러!!\n\n출발매장 지정안됨.");
		return;
	}

	if (frm.makerid.value == "") {
		alert("먼저 브랜드를 선택하세요.");
		return;
	}

	var popwin;
	popwin = window.open('/common/offshop/popshopitemV2.asp?shopid=' + frm.shopid.value + '&chargeid=' + frm.makerid.value,'jsAddItems','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ReActItems(igubun,iitemid,iitemoption,isellcash,isuplycash,ishopbuyprice,iitemno,iitemname,iitemoptionname,iitemdesigner) {
	var frm = document.frm;

	<% '// sellcash : 판매가, buycash : 매입가, suplycash : 샵 공급가 %>
	frm.itemgubunarr.value = igubun;
	frm.itemidarr.value = iitemid;
	frm.itemoptionarr.value = iitemoption;
	frm.sellcasharr.value = isellcash;
	frm.buycasharr.value = isuplycash;
	frm.suplycasharr.value = ishopbuyprice;
	frm.itemnoarr.value = iitemno;

	frm.mode.value = "additem";
	frm.method = "post";
	frm.action = "pop_jumun_move_process.asp";
	frm.submit();
}

function AddItemsBarCode() {
	var frm = document.frm;

	if (frm.shopid.value == "") {
		alert("에러!!\n\n출발매장 지정안됨.");
		return;
	}

	if (frm.makerid.value == "") {
		alert("먼저 브랜드를 선택하세요.");
		return;
	}

	var popwin = window.open('popshopjumunitemBybarcode.asp?shopid=' + frm.shopid.value + '&suplyer=' + frm.makerid.value + '&digitflag=MV','AddItemsBarCode','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<form name="frm" method="get" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<input type="hidden" name="shopid" value="<%= shopid %>">
<input type="hidden" name="moveshopid" value="<%= moveshopid %>">
<input type="hidden" name="makerid" value="<%= makerid %>">
<input type="hidden" name="itemgubunarr" value="">
<input type="hidden" name="itemidarr" value="">
<input type="hidden" name="itemoptionarr" value="">
<input type="hidden" name="itemnoarr" value="">
<input type="hidden" name="sellcasharr" value="">
<input type="hidden" name="buycasharr" value="">
<input type="hidden" name="suplycasharr" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td height="25" width="150" bgcolor="<%= adminColor("tabletop") %>">출발매장</td>
	<td>
		<%= shopid %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td height="25" bgcolor="<%= adminColor("tabletop") %>">도착매장</td>
	<td>
		<% if (moveshopid = "") then %>
		<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("moveshopidsel",moveshopid, "21") %>
		<% if (moveshopid = "") then %>
		&nbsp;
		<input type="button" value="도착매장지정" onClick="jsSetShopidMove()" class="button">
		* <font color="red">먼저 도착매장을 지정하세요.</font>
		<% end if %>
		<% else %>
		<%= moveshopid %>
		<% end if %>
	</td>
</tr>
<% if (shopid <> "") and (moveshopid <> "") then %>
<tr bgcolor="#FFFFFF">
	<td height="25" bgcolor="<%= adminColor("tabletop") %>">현재 브랜드</td>
	<td>
		<% if (makerid = "") then %>
		<% Call drawSelectBoxDesignerwithName("makeridsel",makerid) %>
		&nbsp;
		<input type="button" value="브랜드 변경" onClick="jsSetMakerid()" class="button">
		* <font color="red">먼저 브랜드를 지정하세요.</font>
		<% else %>
		<%= makerid %>
		<% if (errMsg <> "") then %>
		&nbsp;
		<font color="red">* <%= errMsg %></font>
		<% end if %>
		<% end if %>
	</td>
</tr>
<% if (makerid <> "") then %>
<tr bgcolor="#FFFFFF">
	<td height="25" bgcolor="<%= adminColor("tabletop") %>">브랜드 변경</td>
	<td>
		<% Call drawSelectBoxDesignerwithName("makeridsel","") %>
		<% if (shopid <> "") and (moveshopid <> "") then %>
		&nbsp;
		<input type="button" value="브랜드 변경" onClick="jsSetMakerid()" class="button">
		* <font color="red">다른 브랜드</font>의 상품을 선택하려면 먼저 브랜드를 변경하세요.
		<% end if %>
	</td>
</tr>
<% end if %>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">
		재고이동일
	</td>
	<td>
		<input type="text" class="text" name="scheduledt" value="<%= scheduledt %>" size=10 readonly ><a href="javascript:calendarOpen(frm.scheduledt);">
		<img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>

			택배사 :<% drawSelectBoxDeliverCompany "songjangdiv", songjangdiv %>
			송장번호:<input type="text" class="text" name="songjangno" size=14 maxlength=16 value="<%= songjangno %>" >
			<br>
			(택배로 보내지 않을경우 택배사:기타선택 송장번호:퀵배송, 직접배송 등을 입력 하시면 됩니다.)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">기타요청사항</td>
	<td>
		<textarea name="comment" class="textarea" cols="80" rows="6"><%= comment %></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" colspan="2" align="center">
	    <input type="button" value="재고이동처리" onClick="jsStockMove()" class="button" id="btnMove">
		&nbsp;
		<input type="button" value="저장하기" onClick="jsSaveModified()" class="button" id="btnSave">
	</td>
</tr>
</table>
</form>

<% if (IsAddItemOK) then %>
<p></p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left" valign="bottom">
		※ 수량을 플러스로 주문 하시면, 출발매장(<font color="red">마이너스주문</font>)과 도착매장(<font color="red">입고주문</font>)에 주문이 각각 생성됩니다
	</td>
	<td align="right">
		<input type="button" class="button" value="발주(바코드)" onclick="AddItemsBarCode()">
		&nbsp;
		<input type="button" class="button" value="상품추가" onclick="jsAddItems()">
	</td>
</tr>
</table>
<!-- 액션 끝 -->
<% else %>
<p />
* 먼저 도착매장 및 브랜드를 선택하세요.
<% end if %>

<p></p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= ojumundetail.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">바코드</td>
	<td>브랜드</td>
	<td>상품명</td>
	<td>옵션명</td>
	<td width="80">판매가</td>
	<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
	    <td width="60">텐바이텐<br>매입가</td>
	    <td width="60">매장<br>공급가</td>
	<% elseif (C_IS_Maker_Upche) then %>
		<td width="60">텐바이텐<br>공급가</td>
	<% else %>
		<td width="60">매장<br>공급가</td>
	<% end if %>
	<td width="60">수량</td>
</tr>
<% if (ojumundetail.FResultCount > 0) then %>
<% for i = 0 to ojumundetail.FResultCount - 1 %>
<form name="frmBuyPrc_<%= i %>" method="post" action="">
<input type="hidden" name="detailidx" value="<%= ojumundetail.FItemList(i).Fidx %>">
<input type="hidden" name="itemgubun" value="<%= ojumundetail.FItemList(i).Fitemgubun %>">
<input type="hidden" name="itemid" value="<%= ojumundetail.FItemList(i).Fitemid %>">
<input type="hidden" name="itemoption" value="<%= ojumundetail.FItemList(i).Fitemoption %>">
<% if Not (PriceEditEnable) then %>
<input type="hidden" name="sellcash" value="-1">
<input type="hidden" name="buycash" value="-1">
<input type="hidden" name="suplycash" value="-1">
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td ><%= ojumundetail.FItemList(i).Fitemgubun %><%= CHKIIF(ojumundetail.FItemList(i).Fitemid >= 1000000,format00(8,ojumundetail.FItemList(i).Fitemid),format00(6,ojumundetail.FItemList(i).Fitemid)) %><%= ojumundetail.FItemList(i).Fitemoption %></td>
	<td ><%= ojumundetail.FItemList(i).FMakerid %></td>
	<td align="left"><%= ojumundetail.FItemList(i).Fitemname %></td>
	<td ><%= ojumundetail.FItemList(i).Fitemoptionname %></td>
	<% if Not (PriceEditEnable) then %>
		<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fsellcash,0) %></td>
		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fbuycash,0) %></td><!-- 매입가 -->
			<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fsuplycash,0) %></td><!--매장 공급가-->
		<% elseif (C_IS_Maker_Upche) then %>
			<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fbuycash,0) %></td><!-- 매입가 -->
		<% else %>
			<td align="right"><%= FormatNumber(ojumundetail.FItemList(i).Fsuplycash,0) %></td><!--매장 공급가-->
		<% end if %>
	<% else %>
		<td ><input type="text" class="text" name="sellcash" value="<%= ojumundetail.FItemList(i).Fsellcash %>" size="8" maxlength="8"></td>
		<td ><input type="text" class="text" name="buycash" value="<%= ojumundetail.FItemList(i).Fbuycash %>" size="8" maxlength="8"></td>
		<td ><input type="text" class="text" name="suplycash" value="<%= ojumundetail.FItemList(i).Fsuplycash %>" size="8" maxlength="8"></td>
	<% end if %>
	<td ><input type="text" class="text" name="itemno" value="<%= ojumundetail.FItemList(i).Fbaljuitemno %>"  size="4" maxlength="4" onKeyDown="jsButtonDisabled()"></td>
</tr>
</form>
<% next %>
<% end if %>

</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
