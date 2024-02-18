<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 오프라인 매장 재고 이동
' Hieditor : 2011.12.08 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim i ,oipchulmaster, oipchul ,isReqIpgo ,reqDayStr ,isreq ,j ,cnt, cnt2 ,isPreExists ,PriceEditEnable
dim scheduledt ,movesongjangdiv ,songjangno ,firstshopid ,makerid ,moveshopid
dim itemgubunarr, itemidadd, itemoptionarr ,itemnamearr, itemoptionnamearr
dim sellcasharr, suplycasharr, shopbuypricearr, itemnoarr, designerarr ,comment
dim itemgubunarr2, itemidadd2, itemoptionarr2 ,itemnamearr2, itemoptionnamearr2
dim sellcasharr2, suplycasharr2, shopbuypricearr2, itemnoarr2, designerarr2
dim itemgubunarr3, itemidadd3, itemoptionarr3 ,itemnamearr3, itemoptionnamearr3
dim sellcasharr3, suplycasharr3, shopbuypricearr3, itemnoarr3, designerarr3
	isreq         = requestCheckVar(request("isreq"),10)
	scheduledt  = requestCheckVar(request("scheduledt"),30)
	movesongjangdiv = requestCheckVar(request("songjangdiv"),2)
	songjangno  = requestCheckVar(request("songjangno"),32)
	makerid  = requestCheckVar(request("makerid"),32)
	itemgubunarr = request("itemgubunarr")
	itemidadd	= request("itemidadd")
	itemoptionarr = request("itemoptionarr")
	itemnamearr		= request("itemnamearr")
	itemoptionnamearr = request("itemoptionnamearr")
	sellcasharr = request("sellcasharr")
	suplycasharr = request("suplycasharr")
	shopbuypricearr = request("shopbuypricearr")
	itemnoarr = request("itemnoarr")
	designerarr = request("designerarr")
	itemgubunarr2 = request("itemgubunarr2")
	itemidadd2	= request("itemidadd2")
	itemoptionarr2 = request("itemoptionarr2")
	itemnamearr2	= request("itemnamearr2")
	itemoptionnamearr2 = request("itemoptionnamearr2")
	sellcasharr2 = request("sellcasharr2")
	suplycasharr2 = request("suplycasharr2")
	shopbuypricearr2 = request("shopbuypricearr2")
	itemnoarr2 = request("itemnoarr2")
	designerarr2 = request("designerarr2")
	moveshopid = requestCheckVar(request("moveshopid"),32)
	comment = request("comment")
	movesongjangdiv = requestCheckVar(request("movesongjangdiv"),2)
	firstshopid = requestCheckVar(request("firstshopid"),32)

PriceEditEnable = false

if C_ADMIN_USER or C_IS_OWN_SHOP then
elseif (C_IS_SHOP) then
	'가맹점
	firstshopid = C_STREETSHOPID
else
	if (C_IS_Maker_Upche) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('업체는 사용 불가능한 매뉴입니다');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	else
		if Not(C_ADMIN_USER) then
		else
			firstshopid = request("firstshopid")
		end if
	end if
end if

if isreq = "" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('매장 이동 구분값이 지정되지 않았습니다');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close() : response.end
end if

if makerid = "" then
	response.write "<script type='text/javascript'>"
	response.write "	alert('공급처(브랜드)가 지정되지 않았습니다');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close() : response.end
end if

if C_ADMIN_USER then

elseif C_IS_SHOP then
	if getoffshopdiv(firstshopid) <> "1" then
		response.write "<script type='text/javascript'>"
		response.write "	alert('직영매장만 이용가능한 매뉴입니다');"
		response.write "	self.close();"
		response.write "</script>"
		dbget.close() : response.end
	end if
end if

itemgubunarr = split(itemgubunarr,"|")
itemidadd	= split(itemidadd,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
shopbuypricearr = split(shopbuypricearr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")
itemgubunarr2 = split(itemgubunarr2,"|")
itemidadd2	= split(itemidadd2,"|")
itemoptionarr2 = split(itemoptionarr2,"|")
itemnamearr2		= split(itemnamearr2,"|")
itemoptionnamearr2 = split(itemoptionnamearr2,"|")
sellcasharr2 = split(sellcasharr2,"|")
suplycasharr2 = split(suplycasharr2,"|")
shopbuypricearr2 = split(shopbuypricearr2,"|")
itemnoarr2 = split(itemnoarr2,"|")
designerarr2 = split(designerarr2,"|")

cnt = uBound(itemidadd)
cnt2 = uBound(itemidadd2)

for j=0 to cnt2-1
	isPreExists = false
	for i=0 to cnt-1
		if (itemgubunarr(i)=itemgubunarr2(j)) and (itemidadd(i)=itemidadd2(j)) and (itemoptionarr(i)=itemoptionarr2(j)) then
			itemnoarr(i) = CStr(CLng(itemnoarr(i)) + CLng(itemnoarr2(j)))
			isPreExists = true
			exit for
		end if
	next

	if Not isPreExists then
		itemgubunarr3 = itemgubunarr3 + itemgubunarr2(j) + "|"
		itemidadd3	= itemidadd3 + itemidadd2(j) + "|"
		itemoptionarr3 = itemoptionarr3 + itemoptionarr2(j) + "|"
		itemnamearr3		= itemnamearr3 + itemnamearr2(j) + "|"
		itemoptionnamearr3  = itemoptionnamearr3 + itemoptionnamearr2(j) + "|"
		sellcasharr3 = sellcasharr3 + sellcasharr2(j) + "|"
		suplycasharr3 = suplycasharr3 + suplycasharr2(j) + "|"
		shopbuypricearr3 = shopbuypricearr3 + shopbuypricearr2(j) + "|"
		itemnoarr3 = itemnoarr3 + itemnoarr2(j) + "|"
		designerarr3 = designerarr3 + designerarr2(j) + "|"
	end if
next

itemgubunarr2 = ""
itemidadd2	= ""
itemoptionarr2 = ""
itemnamearr2	= ""
itemoptionnamearr2 = ""
sellcasharr2 = ""
suplycasharr2 = ""
shopbuypricearr2 = ""
itemnoarr2 = ""
designerarr2 = ""

for i=0 to cnt-1
	itemgubunarr2 = itemgubunarr2 + itemgubunarr(i) + "|"
	itemidadd2	= itemidadd2 + itemidadd(i) + "|"
	itemoptionarr2 = itemoptionarr2 + itemoptionarr(i) + "|"
	itemnamearr2	= itemnamearr2 + itemnamearr(i) + "|"
	itemoptionnamearr2 = itemoptionnamearr2 + itemoptionnamearr(i) + "|"
	sellcasharr2 = sellcasharr2 + sellcasharr(i) + "|"
	suplycasharr2 = suplycasharr2 + suplycasharr(i) + "|"
	shopbuypricearr2 = shopbuypricearr2 + shopbuypricearr(i) + "|"
	itemnoarr2 = itemnoarr2 + itemnoarr(i) + "|"
	designerarr2 = designerarr2 + designerarr(i) + "|"
next

itemgubunarr = itemgubunarr2 + itemgubunarr3
itemidadd	= itemidadd2 + itemidadd3
itemoptionarr = itemoptionarr2 + itemoptionarr3
itemnamearr	= itemnamearr2 + itemnamearr3
itemoptionnamearr = itemoptionnamearr2 + itemoptionnamearr3
sellcasharr = sellcasharr2 + sellcasharr3
suplycasharr = suplycasharr2 + suplycasharr3
shopbuypricearr = shopbuypricearr2 + shopbuypricearr3
itemnoarr = itemnoarr2 + itemnoarr3
designerarr = designerarr2 + designerarr3
%>

<script type='text/javascript'>

function ReActItems(igubun,iitemid,iitemoption,isellcash,isuplycash,ishopbuyprice,iitemno,iitemname,iitemoptionname,iitemdesigner){
	frmMaster.itemgubunarr2.value = igubun;
	frmMaster.itemidadd2.value = iitemid;
	frmMaster.itemoptionarr2.value = iitemoption;
	frmMaster.sellcasharr2.value = isellcash;
	frmMaster.suplycasharr2.value = isuplycash;
	frmMaster.shopbuypricearr2.value = ishopbuyprice;
	frmMaster.itemnoarr2.value = iitemno;
	frmMaster.itemnamearr2.value = iitemname;
	frmMaster.itemoptionnamearr2.value = iitemoptionname;
	frmMaster.designerarr2.value = iitemdesigner;
	frmMaster.submit();
}

function shopselect(){
	var firstshopid = frmshop.firstshopid.value;

	if (firstshopid==''){
		alert('출발매장을 선택해 주세요');
		frmshop.firstshopid.focus();
		return;
	}
	frmshop.submit();
}

//매장재고이동처리
function ipchulmove(){
	var msfrm = document.frmMaster;
	var upfrm = document.frmArrupdate;
	var firstshopid = frmshop.firstshopid.value;
	var frm;

	if (firstshopid==''){
		alert('출발매장을 선택해 주세요');
		frmshop.firstshopid.focus();
		return;
	}

	if (msfrm.moveshopid.value.length<1){
		alert('도착매장을 선택하세요.');
		msfrm.moveshopid.focus();
		return;
	}

	if (msfrm.scheduledt.value.length<1){
		alert('재고이동일을 입력하세요');
		calendarOpen3(frmMaster.scheduledt,'재고이동일을 입력하세요','');
		return;
	}

	if (msfrm.movesongjangdiv.value.length<1){
		alert('택배사를 선택 하세요');
		msfrm.movesongjangdiv.focus();
		return;
	}

	if (msfrm.songjangno.value.length<1){
		alert('송장 번호를 입력 하세요');
		msfrm.songjangno.focus();
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.shopbuypricearr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.designerarr.value = "";
    upfrm.isreq.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {
			if (!IsDigit(frm.sellcash.value)){
				alert('판매가는 숫자만 가능합니다.');
				frm.sellcash.focus();
				return;
			}

			if (!IsDigit(frm.suplycash.value)){
				alert('공급가는 숫자만 가능합니다.');
				frm.suplycash.focus();
				return;
			}

			if (!IsInteger(frm.itemno.value)){
				alert('갯수는 정수만 가능합니다.');
				frm.itemno.focus();
				return;
			}

		    if (frm.itemno.value < 0){
				alert("갯수는 0이상만 허용 됩니다.");
				frm.itemno.focus();
				return;
			}

			upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
			upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
			upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
			upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
			upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
			upfrm.shopbuypricearr.value = upfrm.shopbuypricearr.value + frm.shopbuyprice.value + "|";
			upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
			upfrm.designerarr.value = upfrm.designerarr.value + frm.chargeid.value + "|";
		}
	}

	var ret = confirm('입력하신대로 재고 이동처리 하시겠습니까?');

	if (ret){
		upfrm.scheduledt.value = msfrm.scheduledt.value;
		upfrm.songjangdiv.value = msfrm.movesongjangdiv.value;
		upfrm.songjangno.value = msfrm.songjangno.value;
		upfrm.chargeid.value = msfrm.chargeid.value;
		upfrm.firstshopid.value = msfrm.firstshopid.value;
		upfrm.divcode.value = msfrm.divcode.value;
		upfrm.vatcode.value = msfrm.vatcode.value;
        upfrm.isreq.value   = msfrm.isreq.value;
        upfrm.comment.value   = msfrm.comment.value;
        upfrm.moveshopid.value   = msfrm.moveshopid.value;
        upfrm.mode.value   = 'ipchulmove';
        upfrm.action='shopipchulitem_process.asp';
		upfrm.submit();
	}
}

//상품추가
function AddItems(){
	var firstshopid = frmMaster.firstshopid.value;

	if (firstshopid==''){
		alert('출발매장을 선택하세요');
		frmshop.firstshopid.focus();
		return;
	}

	var popwin;
	popwin = window.open('popshopitem2.asp?shopid=' + firstshopid + '&chargeid=<%= makerid %>','addshopitem','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddItemsBarCode(frm, digitflag){
	if (frm.firstshopid.value.length<1){
		alert('가맹점을 먼저 선택하세요');
		frm.firstshopid.focus();
		return;
	}

	var popwin;
	popwin = window.open('popshopitemBybarcode.asp?shopid=' + frmMaster.firstshopid.value + '&chargeid=' + frmMaster.chargeid.value + '&digitflag=' + digitflag,'popshopitemBybarcode','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmshop" method="get" action="">
<input type="hidden" name="isreq" value="<%= isreq %>">
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">공급처</td>
	<td bgcolor="#FFFFFF">
		<input type="hidden" name="makerid" value="<%= makerid %>">
		<%= makerid %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan=2 align="center">출발매장 주문정보</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">출발매장</td>
	<td>
		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<% drawBoxDirectIpchulOffShopByMakerchfg "firstshopid", firstshopid, makerid ," onchange='shopselect()';","'B012','B022','B023'" %>
			(업체위탁/매장매입 설정된 매장만 표시됩니다.)
		<% elseif (C_IS_SHOP) then %>
			<%= firstshopid %>
			<input type="hidden" name="firstshopid" value="<%= firstshopid %>">
		<% else %>
			<% drawBoxDirectIpchulOffShopByMakerchfg "firstshopid", firstshopid, makerid ," onchange='shopselect()';","'B012','B022','B023'" %>
			(업체위탁/매장매입 설정된 매장만 표시됩니다.)
		<% end if %>
	</td>
</tr>
</form>
</table>

<% if firstshopid = "" then response.end %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="">
<input type="hidden" name="isreq" value="<%= isreq %>">
<input type="hidden" name="divcode" value="006">
<input type="hidden" name="vatcode" value="008">
<input type="hidden" name="firstshopid" value="<%= firstshopid %>">
<input type="hidden" name="chargeid" value="<%= makerid %>">
<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
<input type="hidden" name="itemidadd" value="<%= itemidadd %>">
<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
<input type="hidden" name="itemnamearr" value="<%= itemnamearr %>">
<input type="hidden" name="itemoptionnamearr" value="<%= itemoptionnamearr %>">
<input type="hidden" name="sellcasharr" value="<%= sellcasharr %>">
<input type="hidden" name="suplycasharr" value="<%= suplycasharr %>">
<input type="hidden" name="shopbuypricearr" value="<%= shopbuypricearr %>">
<input type="hidden" name="itemnoarr" value="<%= itemnoarr %>">
<input type="hidden" name="designerarr" value="<%= designerarr %>">
<input type="hidden" name="itemgubunarr2" value="">
<input type="hidden" name="itemidadd2" value="">
<input type="hidden" name="itemoptionarr2" value="">
<input type="hidden" name="itemnamearr2" value="">
<input type="hidden" name="itemoptionnamearr2" value="">
<input type="hidden" name="sellcasharr2" value="">
<input type="hidden" name="suplycasharr2" value="">
<input type="hidden" name="shopbuypricearr2" value="">
<input type="hidden" name="itemnoarr2" value="">
<input type="hidden" name="designerarr2" value="">
<tr bgcolor="#FFFFFF">
	<td colspan=2 align="center">도착매장 주문정보</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">도착매장</td>
	<td bgcolor="#FFFFFF">
		<% drawBoxshopipchulcontract "moveshopid", "", makerid, firstshopid,"" %>
		(출발매장과 계약 마진이 동일한 매장만 표시 됩니다)
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">
		재고이동일
	</td>
	<td>
		<input type="text" class="text" name="scheduledt" value="<%= scheduledt %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.scheduledt);">
		<img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>

		<% if Not isReqIpgo then %>
			택배사 :<% drawSelectBoxDeliverCompany "movesongjangdiv", movesongjangdiv %>
			송장번호:<input type="text" class="text" name="songjangno" size=14 maxlength=16 value="<%= songjangno %>" >
			<br>
			(택배로 보내지 않을경우 택배사:기타선택 송장번호:퀵배송, 직접배송 등을 입력 하시면 됩니다.)
		<% else %>
			<input type="hidden" name="movesongjangdiv" value="">
			<input type="hidden" name="songjangno" value="">
		<% end if %>
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
	    <input type="button" value="재고이동처리" onClick="ipchulmove()" class="button">
	</td>
</tr>
</form>
</table>

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		※ 수량을 플러스로 주문 하시면, 출발매장(<font color="red">마이너스주문</font>)과 도착매장(<font color="red">입고주문</font>)에 주문이 각각 생성됩니다
	</td>
	<td align="right">
		<input type="button" class="button" value="상품추가" onclick="AddItems()">
		<input type="button" class="button" value="상품추가(바코드)" onclick="AddItemsBarCode(frmMaster,'itemadd')">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<%
itemgubunarr = split(itemgubunarr,"|")
itemidadd	= split(itemidadd,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
shopbuypricearr = split(shopbuypricearr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")

cnt = ubound(itemidadd)
%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= cnt+1 %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">바코드</td>
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
<% for i=0 to cnt-1 %>
<form name="frmBuyPrc_<%= i %>" method="post" action="">
<input type="hidden" name="itemgubun" value="<%= itemgubunarr(i) %>">
<input type="hidden" name="itemid" value="<%= itemidadd(i) %>">
<input type="hidden" name="itemoption" value="<%= itemoptionarr(i) %>">
<input type="hidden" name="chargeid" value="<%= designerarr(i) %>">
<input type="hidden" name="sellcash" value="<%= sellcasharr(i) %>">
<input type="hidden" name="suplycash" value="<%= suplycasharr(i) %>">
<input type="hidden" name="shopbuyprice" value="<%= shopbuypricearr(i) %>">

<tr align="center" bgcolor="#FFFFFF">
	<td ><%= itemgubunarr(i) %><%= CHKIIF(itemidadd(i)>=1000000,format00(8,itemidadd(i)),format00(6,itemidadd(i))) %><%= itemoptionarr(i) %></td>
	<td align="left"><%= itemnamearr(i) %></td>
	<td ><%= itemoptionnamearr(i) %></td>


	<% if Not (PriceEditEnable) then %>
		<td align="right"><%= FormatNumber(sellcasharr(i),0) %></td>

		<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<td align="right"><%= FormatNumber(suplycasharr(i),0) %></td><!--텐바이텐 매입가-->
			<td align="right"><%= FormatNumber(shopbuypricearr(i),0) %></td><!--매장 공급가-->
		<% elseif (C_IS_Maker_Upche) then %>
			<td align="right"><%= FormatNumber(suplycasharr(i),0) %></td><!--텐바이텐 공급가-->
		<% else %>
			<td align="right"><%= FormatNumber(shopbuypricearr(i),0) %></td><!--매장 공급가-->
		<% end if %>
	<% else %>
		<td ><input type="text" class="text" name="sellcash" value="<%= sellcasharr(i) %>" size="8" maxlength="8"></td>
		<td ><input type="text" class="text" name="suplycash" value="<%= suplycasharr(i) %>" size="8" maxlength="8"></td>
		<td ><input type="text" class="text" name="shopbuyprice" value="<%= shopbuypricearr(i) %>" size="8" maxlength="8"></td>
	<% end if %>

	<td ><input type="text" class="text" name="itemno" value="<%= itemnoarr(i) %>"  size="4" maxlength="4"></td>
</tr>
</form>
<% next %>

</table>

<form name="frmArrupdate" method="post" action="">
	<input type="hidden" name="scheduledt" value="">
	<input type="hidden" name="songjangdiv" value="">
	<input type="hidden" name="songjangno" value="">
	<input type="hidden" name="divcode" value="">
	<input type="hidden" name="vatcode" value="">
	<input type="hidden" name="comment" value="">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="shopbuypricearr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="designerarr" value="">
	<input type="hidden" name="chargeid" value="<%= makerid %>">
	<input type="hidden" name="firstshopid" value="<%= firstshopid %>">
	<input type="hidden" name="moveshopid" value="<%= moveshopid %>">
	<input type="hidden" name="isreq" value="">
</form>

<%
set oipchulmaster = Nothing
set oipchul = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->