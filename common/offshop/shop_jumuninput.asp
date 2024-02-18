<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 주문서 작성
' History : 2009.04.07 서동석 생성
'			2011.01.13 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim shopid, reguser, divcode,baljuname,regname
dim itemgubunarr, itemidadd, itemoptionarr ,vatcode ,itemnamearr3, itemoptionnamearr3
dim itemnamearr, itemoptionnamearr ,sellcasharr, suplycasharr, buycasharr, itemnoarr, designerarr
dim itemgubunarr2, itemidadd2, itemoptionarr2 ,itemnamearr2, itemoptionnamearr2
dim sellcasharr2, suplycasharr2, buycasharr2, itemnoarr2, designerarr2
dim itemgubunarr3, itemidadd3, itemoptionarr3 ,i,j,cnt,cnt2 ,isPreExists
dim sellcasharr3, suplycasharr3, buycasharr3, itemnoarr3, designerarr3
dim suplyer,yyyymmdd,comment ,osheetmaster, idx ,cwflag
	shopid = requestCheckVar(request("shopid"),32)
	cwflag = requestCheckVar(request("cwflag"),10)

if C_ADMIN_USER then

'/직영점
elseif C_IS_OWN_SHOP then
    shopid = C_STREETSHOPID
elseif (C_IS_SHOP) then
    shopid = C_STREETSHOPID
end if

IF (Not C_IS_FRN_SHOP) then
    reguser     = session("ssBctid")
    divcode     = "501"
    baljuname   = "" '' The ShopName
    regname     = session("ssBctCname")
ELSEIF (C_IS_SHOP) then
    reguser     = shopid
    divcode     = session("ssBctDiv")
    baljuname   = session("ssBctCname")
    regname     = baljuname
END IF

'if (left(shopid,Len("streetshop"))<>"streetshop") then

idx = requestCheckVar(request("idx"),10)
if idx="" then idx=0
suplyer = requestCheckVar(request("suplyer"),32)
yyyymmdd = requestCheckVar(request("yyyymmdd"),10)
comment = request("comment")
if suplyer="" then suplyer="10x10"
itemgubunarr = request("itemgubunarr")
itemidadd	= request("itemidadd")
itemoptionarr = request("itemoptionarr")
itemnamearr		= request("itemnamearr")
itemoptionnamearr = request("itemoptionnamearr")
sellcasharr = request("sellcasharr")
suplycasharr = request("suplycasharr")
buycasharr = request("buycasharr")
itemnoarr = request("itemnoarr")
designerarr = request("designerarr")
itemgubunarr2 = request("itemgubunarr2")
itemidadd2	= request("itemidadd2")
itemoptionarr2 = request("itemoptionarr2")
itemnamearr2	= request("itemnamearr2")
itemoptionnamearr2 = request("itemoptionnamearr2")
sellcasharr2 = request("sellcasharr2")
suplycasharr2 = request("suplycasharr2")
buycasharr2 = request("buycasharr2")
itemnoarr2 = request("itemnoarr2")
designerarr2 = request("designerarr2")
'chargeid = request("chargeid")
'shopid = session("ssBctID")
'vatcode = request("vatcode")
'divcode  = request("divcode")
itemgubunarr = split(itemgubunarr,"|")
itemidadd	= split(itemidadd,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
buycasharr = split(buycasharr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")
itemgubunarr2 = split(itemgubunarr2,"|")
itemidadd2	= split(itemidadd2,"|")
itemoptionarr2 = split(itemoptionarr2,"|")
itemnamearr2		= split(itemnamearr2,"|")
itemoptionnamearr2 = split(itemoptionnamearr2,"|")
sellcasharr2 = split(sellcasharr2,"|")
suplycasharr2 = split(suplycasharr2,"|")
buycasharr2 = split(buycasharr2,"|")
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
		buycasharr3 = buycasharr3 + buycasharr2(j) + "|"
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
buycasharr2 = ""
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
	buycasharr2 = buycasharr2 + buycasharr(i) + "|"
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
buycasharr = buycasharr2 + buycasharr3
itemnoarr = itemnoarr2 + itemnoarr3
designerarr = designerarr2 + designerarr3

dim shopdiv
	shopdiv = getoffshopdiv(shopid)

if cwflag = "" then
	'//대행매장일경우 기본값이 출고위탁임
	if shopdiv = "13" then
		cwflag = "1"
	else
		cwflag = "0"
	end if
end if

if shopid = "" or isnull(shopid) then
	response.write "<script language='javascript'>alert('매장을 선택하세요.');</script>"
	'dbget.close()	:	response.End
end if

%>

<script language='javascript'>

function chcwflag(shopid){
	if (shopid==''){
		alert('매장을 선택하세요');
		return;
	}

	frmMaster.submit();
}

function jsPopCal(fName,sName){
	var fd = eval("document."+fName+"."+sName);

	var winCal;
	winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

function ReActItems(iidx, igubun,iitemid,iitemoption,isellcash,isuplycash,ibuycash,iitemno,iitemname,iitemoptionname,iitemdesigner){
	if ((iidx!='0')&&(iidx!='')){
		alert('<%= CTX_Does_not_match %> (<%= CTX_Order_code %> :' + iidx + ')');
		return;
	}

	frmMaster.itemgubunarr2.value = igubun;
	frmMaster.itemidadd2.value = iitemid;
	frmMaster.itemoptionarr2.value = iitemoption;
	frmMaster.sellcasharr2.value = isellcash;
	frmMaster.suplycasharr2.value = isuplycash;
	frmMaster.buycasharr2.value = ibuycash;
	frmMaster.itemnoarr2.value = iitemno;
	frmMaster.itemnamearr2.value = iitemname;
	frmMaster.itemoptionnamearr2.value = iitemoptionname;
	frmMaster.designerarr2.value = iitemdesigner;

	frmMaster.submit();
}

//상품추가 리스트뉴
function AddItems_locale(frm){
	var popwin;
	var suplyer

	if (frmMaster.shopid.value.length<1){
		alert('발주처를 먼저 선택하세요.');
		frmMaster.shopid.focus();
		return;
	}

	if (frm.suplyer.value.length<1){
		alert('<%= CTX_Please_select %> (<%= CTX_WHOLESALEID %>)');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;

	var cwflag;
	for (var i =0 ; i < frm.cwflag.length ; i++){
		if (frm.cwflag[i].checked){
			cwflag = frm.cwflag[i].value;
		}
	}

	popwin = window.open('/common/offshop/localeitem/popshopjumunitem_locale.asp?suplyer=' + suplyer + '&shopid=<%=shopid %>&idx=0&cwflag='+cwflag ,'franjumuninputadd','width=1024,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddItems(frm){
	var popwin;
	var suplyer;

	if (frmMaster.shopid.value.length<1){
		alert('발주처를 먼저 선택하세요.');
		frmMaster.shopid.focus();
		return;
	}

	if (frm.suplyer.value.length<1){
		alert('<%= CTX_Please_select %> (<%= CTX_WHOLESALEID %>)');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;

	var cwflag;
	for (var i =0 ; i < frm.cwflag.length ; i++){
		if (frm.cwflag[i].checked){
			cwflag = frm.cwflag[i].value;
		}
	}

	popwin = window.open('/common/offshop/popshopjumunitem.asp?suplyer=' + suplyer + '&shopid=<%=shopid %>&idx=0&cwflag='+cwflag,'offpopShopjumunItem','width=880,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();

	//var ret = window.showModalDialog('/common/offshop/popshopjumunitem.asp?suplyer=' + suplyer + '&idx=0',null,'dialogwidth:900px;dialogheight:700px;center:yes;scroll:yes;resizable:yes;status:yes;help:no;');

}

function ConFirmIpChulList(bool){
	var msfrm = document.frmMaster;
	var upfrm = document.frmArrupdate;
	var frm;

	if (msfrm.yyyymmdd.value.length<1){
		alert('<%= CTX_Please_select %> (<%= CTX_were_stocked_requested_date %>)');
		//msfrm.yyyymmdd.value.focus();
		return;
	}

	upfrm.itemgubunarr.value = "";
	upfrm.itemarr.value = "";
	upfrm.itemoptionarr.value = "";
	upfrm.sellcasharr.value = "";
	upfrm.suplycasharr.value = "";
	upfrm.buycasharr.value = "";
	upfrm.itemnoarr.value = "";
	upfrm.designerarr.value = "";

	for (var i=0;i<document.forms.length;i++){
		frm = document.forms[i];
		if (frm.name.substr(0,9)=="frmBuyPrc") {

			if (!IsInteger(frm.itemno.value)){
				alert('<%= CTX_Type_Mismatch %> (<%= CTX_Only_numbers %>)');
				frm.itemno.focus();
				return;
			}

			upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
			upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
			upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
			upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
			upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
			upfrm.buycasharr.value = upfrm.buycasharr.value + frm.buycash.value + "|";
			upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
			upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
		}
	}

    if (!bool) {
		var ret = confirm('<%= CTX_Do_you_want_to_save %> (<%= CTX_Temporary_storage %>)?');
	}else{
		var ret = confirm('<%= CTX_Do_you_want_to_save %>?');
	}

	if (ret){
	    //임시저장(작성중)
		if (!bool) upfrm.waitflag.value="on"

		upfrm.yyyymmdd.value = msfrm.yyyymmdd.value;
		upfrm.comment.value = msfrm.comment.value;

		var cwflag;
		for (var i =0 ; i < msfrm.cwflag.length ; i++){
			if (msfrm.cwflag[i].checked){
				cwflag = msfrm.cwflag[i].value;
			}
		}

		upfrm.cwflag.value = cwflag;
		upfrm.submit();
	}
}

function AddItemsBarCode(frm, digitflag){
	var popwin;
	var suplyer, shopid;

	if (frm.shopid.value.length<1){
		alert('발주처를 먼저 선택하세요.');
		frm.shopid.focus();
		return;
	}

	if (frm.suplyer.value.length<1){
		alert('공급처를 먼저 선택하세요.');
		frm.suplyer.focus();
		return;
	}

	suplyer = frm.suplyer.value;
	shopid = frm.shopid.value;

	var cwflag;
	for (var i =0 ; i < frm.cwflag.length ; i++){
		if (frm.cwflag[i].checked){
			cwflag = frm.cwflag[i].value;
		}
	}

	popwin = window.open('/admin/fran/popshopjumunitemBybarcode.asp?suplyer=' + suplyer + '&shopid=' + shopid + '&digitflag=' + digitflag + '&idx=0&cwflag='+cwflag ,'franjumuninputaddBarcode','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmMaster" method="post" action="">
<input type="hidden" name="mode" value="addmaster">
<input type="hidden" name="itemgubunarr" value="<%= itemgubunarr %>">
<input type="hidden" name="itemidadd" value="<%= itemidadd %>">
<input type="hidden" name="itemoptionarr" value="<%= itemoptionarr %>">
<input type="hidden" name="itemnamearr" value="<%= itemnamearr %>">
<input type="hidden" name="itemoptionnamearr" value="<%= itemoptionnamearr %>">
<input type="hidden" name="sellcasharr" value="<%= sellcasharr %>">
<input type="hidden" name="suplycasharr" value="<%= suplycasharr %>">
<input type="hidden" name="buycasharr" value="<%= buycasharr %>">
<input type="hidden" name="itemnoarr" value="<%= itemnoarr %>">
<input type="hidden" name="designerarr" value="<%= designerarr %>">
<input type="hidden" name="itemgubunarr2" value="">
<input type="hidden" name="itemidadd2" value="">
<input type="hidden" name="itemoptionarr2" value="">
<input type="hidden" name="itemnamearr2" value="">
<input type="hidden" name="itemoptionnamearr2" value="">
<input type="hidden" name="sellcasharr2" value="">
<input type="hidden" name="suplycasharr2" value="">
<input type="hidden" name="buycasharr2" value="">
<input type="hidden" name="itemnoarr2" value="">
<input type="hidden" name="designerarr2" value="">
<tr bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("tabletop") %>" width=100><%= CTX_an_orderer %></td>
    <td>
		<% if shopid<>"" then %>
			<input type="hidden" name="shopid" value="<%= shopid %>">
			<%= shopid %>
		<% else %>
			<% drawSelectBoxOffShopdiv_off "shopid", shopid ,"",""," onchange='chcwflag(this.value);'" %>
		<% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width=100><%= CTX_WHOLESALEID %></td>
	<% if suplyer<>"" then %>
	<input type="hidden" name="suplyer" value="<%= suplyer %>">
	<td><%= suplyer %></td>
	<% else %>
	<td><% SelectBoxOffShopSuplyer "suplyer", suplyer, shopid, session("ssBctDiv") %></td>
	<% end if %>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width=100><%= CTX_were_stocked_requested_date %></td>
	<td>
	    <input type="text" class="text" name="yyyymmdd" value="<%= yyyymmdd %>" size=10 readonly ><a href="javascript:jsPopCalendar('frmMaster','yyyymmdd');"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a> (원하는 입고 날짜를 입력하세요.)
	</td>
</tr>
<% if getcwflag(shopid,"B013") = "1" then %>
	<tr bgcolor="#FFFFFF" id="divcwflag" name="divcwflag" style="display:">
<% else %>
	<tr bgcolor="#FFFFFF" id="divcwflag" name="divcwflag" style="display:none">
<% end if %>
	<td bgcolor="<%= adminColor("tabletop") %>"><%= CTX_release_divide %></td>
	<td>
		<input type="radio" name="cwflag" value="0" <% if cwflag="0" then response.write " checked" %>><%= CTX_release_Purchase %>
		<input type="radio" name="cwflag" value="1" <% if cwflag="1" then response.write " checked" %>><%= CTX_release_on_consignment %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width=100><%= CTX_Requests %></td>
	<td>
		<textarea class="textarea" name="comment" cols=80 rows=6><%= comment %></textarea>
	</td>
</tr>
</form>
<!--
<tr bgcolor="#FFFFFF">
	<td colspan="2">
		* 5일내 출고 : 업체 배송 상품 (물류센터로 입고 되는대로 매장으로 발송 해드리겠습니다.) <br>
		* 재고 부족 : 물류센터 재고 부족으로 인해 업체로 발주가 들어가 있는 상태입니다. <br>
					2~3일 내로 입고 될 수 있는 상품 입니다. 따로 보내드리지 않으며, <B>다음 주문시 추가(재주문)</B>해 주셔야 합니다.<br>
		* 일시품절 : 업체 재고부족으로 인해 재생산중인 상품입니다.(단기간 내에 입고 되기 어려운 상품입니다.)
	</td>
</tr>
-->
</table>

<br>

<%
itemgubunarr = split(itemgubunarr,"|")
itemidadd	= split(itemidadd,"|")
itemoptionarr = split(itemoptionarr,"|")
itemnamearr		= split(itemnamearr,"|")
itemoptionnamearr = split(itemoptionnamearr,"|")
sellcasharr = split(sellcasharr,"|")
suplycasharr = split(suplycasharr,"|")
buycasharr = split(buycasharr,"|")
itemnoarr = split(itemnoarr,"|")
designerarr = split(designerarr,"|")

cnt = ubound(itemidadd)

dim selltotal, suplytotal
	selltotal =0
	suplytotal =0
%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="right">
		<%= CTX_search_result %> :  <% if cnt < 1 then response.write "0" else response.write cnt %>
		&nbsp;
    	<input type="button" class="button" value="<%= CTX_Add_new_items %>" onclick="AddItems(frmMaster)">
		<input type="button" class="button" value="<%= CTX_Add_new_items %>(NEW)" onclick="AddItems_locale(frmMaster)">

		<% if C_IS_SHOP or C_ADMIN_AUTH or C_OFF_AUTH or C_logics_Part then %>
			<input type="button" class="button" value="<%= CTX_Real_Order %>(BARCODE)" onclick="AddItemsBarCode(frmMaster,'P')">
			<input type="button" class="button" value="<%= CTX_return_Order %>(BARCODE)" onclick="AddItemsBarCode(frmMaster,'M')">
		<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="120"><%= CTX_Brand %></td>
	<td width="100"><%= CTX_Item_Code %></td>
	<td><%= CTX_Description %><font color="blue">[<%= CTX_Description_Option %>]</font></td>
	<td width="70"><%= CTX_selling_price %></td>
	<td width="70"><%= CTX_Supply_price %></td>
	<td width="50"><%= CTX_quantity %></td>
	<td width="70"><%= CTX_total_selling_price %></td>
	<td width="80"><%= CTX_total_Supply_price %></td>
</tr>
<%
for i=0 to cnt-1

selltotal  = selltotal + sellcasharr(i) * itemnoarr(i)
suplytotal = suplytotal + suplycasharr(i) * itemnoarr(i)
%>
<form name="frmBuyPrc_<%= i %>" method="post" action="">
<input type="hidden" name="itemgubun" value="<%= itemgubunarr(i) %>">
<input type="hidden" name="itemid" value="<%= itemidadd(i) %>">
<input type="hidden" name="itemoption" value="<%= itemoptionarr(i) %>">
<input type="hidden" name="desingerid" value="<%= designerarr(i) %>">
<input type="hidden" name="sellcash" value="<%= sellcasharr(i) %>">
<input type="hidden" name="suplycash" value="<%= suplycasharr(i) %>">
<input type="hidden" name="buycash" value="<%= buycasharr(i) %>">
<tr align="center" bgcolor="#FFFFFF">
	<td><%= designerarr(i) %></td>
	<td><%= itemgubunarr(i) %><%= CHKIIF(itemidadd(i)>=1000000,format00(8,itemidadd(i)),format00(6,itemidadd(i))) %><%= itemoptionarr(i) %></td>
	<td align="left">
		<%= itemnamearr(i) %>
		<% if itemoptionarr(i) <>"0000" then %>
			<font color="blue">[<%= itemoptionnamearr(i) %>]</font>
		<% end if %>
	</td>
	<td align="right"><%= FormatNumber(sellcasharr(i),0) %></td>
	<td align="right"><%= FormatNumber(suplycasharr(i),0) %></td>
	<td><input type="text" class="text" name="itemno" value="<%= itemnoarr(i) %>"  size="4" maxlength="4"></td>
	<td align="right"><%= FormatNumber(sellcasharr(i) * itemnoarr(i),0) %></td>
	<td align="right"><%= FormatNumber(suplycasharr(i) * itemnoarr(i),0) %></td>
</tr>
</form>
<% next %>

<% if (cnt>0) then %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= CTX_total %></td>
	<td colspan="5" align="center">
	<td align="right"><%= formatNumber(selltotal,0) %></td>
	<td align="right"><%= formatNumber(suplytotal,0) %></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="center">
		<input type="button" class="button" value="<%= CTX_History_confirmed %>(<%= CTX_Register %>)" onclick="ConFirmIpChulList(true)">
    	<input type="button" class="button" value="<%= CTX_Temporary_storage %>(<%= CTX_in_process %>)" onclick="ConFirmIpChulList(false)"><br><br>

    	<font color=red>내역을 확정(주문접수)한 이후에는 주문을 수정할 수 없습니다.</font><br>
    	&nbsp;
	</td>
</tr>
<% end if %>

<%
'// 등록자 아이디 + 시간을 가지고 중복입력 체크
dim uniqregdate : uniqregdate = getDatabaseTime()
%>

<form name="frmArrupdate" method="post" action="common_shopjumun_process.asp">
	<input type="hidden" name="mode" value="addshopjumun">
	<input type="hidden" name="waitflag" value="">
	<input type="hidden" name="yyyymmdd" value="">
	<input type="hidden" name="baljuid" value="<%= shopid %>">
	<input type="hidden" name="uniqregdate" value="<%= uniqregdate %>">
	<input type="hidden" name="targetid" value="<%= suplyer %>">
	<input type="hidden" name="reguser" value="<%= reguser %>">
	<input type="hidden" name="divcode" value="<%= divcode %>">
	<input type="hidden" name="vatinclude" value="Y">
	<input type="hidden" name="comment" value="">
	<input type="hidden" name="regname" value="<%= regname %>">
	<input type="hidden" name="baljuname" value="<%= baljuname %>">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="buycasharr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="designerarr" value="">
	<input type="hidden" name="cwflag">
</form>
</table>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
