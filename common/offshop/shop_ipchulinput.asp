<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프샵 주문서 작성
' History : 2009.04.07 서동석 생성
'			2011.05.16 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopipchulcls.asp"-->
<%
dim SHOW_ADDSHOP : SHOW_ADDSHOP = True
dim IS_HIDE_BUYCASH : IS_HIDE_BUYCASH = False
dim PriceEditEnable ,yyyy1,mm1,dd1 ,nowdate, chargeid, shopid, vatcode, divcode
dim itemgubunarr, itemidadd, itemoptionarr ,itemnamearr, itemoptionnamearr
dim sellcasharr, suplycasharr, shopbuypricearr, itemnoarr, designerarr
dim itemgubunarr2, itemidadd2, itemoptionarr2 ,itemnamearr2, itemoptionnamearr2
dim sellcasharr2, suplycasharr2, shopbuypricearr2, itemnoarr2, designerarr2
dim itemgubunarr3, itemidadd3, itemoptionarr3 ,itemnamearr3, itemoptionnamearr3
dim sellcasharr3, suplycasharr3, shopbuypricearr3, itemnoarr3, designerarr3
dim i,j,cnt,cnt2 ,scheduledt, songjangdiv, songjangno, isreq
dim isPreExists, isReqIpgo, reqDayStr , comment
dim addshopid
	PriceEditEnable = false
	itemgubunarr = TrimRightDelim(request("itemgubunarr"),"|")
	itemidadd	= TrimRightDelim(request("itemidadd"),"|")
	itemoptionarr = TrimRightDelim(request("itemoptionarr"),"|")
	itemnamearr		= TrimRightDelim(request("itemnamearr"),"|")
	itemoptionnamearr = TrimRightDelim(request("itemoptionnamearr"),"|")
	sellcasharr = TrimRightDelim(request("sellcasharr"),"|")
	suplycasharr = TrimRightDelim(request("suplycasharr"),"|")
	shopbuypricearr = TrimRightDelim(request("shopbuypricearr"),"|")
	itemnoarr = TrimRightDelim(request("itemnoarr"),"|")
	designerarr = TrimRightDelim(request("designerarr"),"|")
	itemgubunarr2 = TrimRightDelim(request("itemgubunarr2"),"|")
	itemidadd2	= TrimRightDelim(request("itemidadd2"),"|")
	itemoptionarr2 = TrimRightDelim(request("itemoptionarr2"),"|")
	itemnamearr2	= TrimRightDelim(request("itemnamearr2"),"|")
	itemoptionnamearr2 = TrimRightDelim(request("itemoptionnamearr2"),"|")
	sellcasharr2 = TrimRightDelim(request("sellcasharr2"),"|")
	suplycasharr2 = TrimRightDelim(request("suplycasharr2"),"|")
	shopbuypricearr2 = TrimRightDelim(request("shopbuypricearr2"),"|")
	itemnoarr2 = TrimRightDelim(request("itemnoarr2"),"|")
	designerarr2 = TrimRightDelim(request("designerarr2"),"|")
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1 = requestCheckVar(request("mm1"),2)
	dd1 = requestCheckVar(request("dd1"),2)
	comment = html2db(request("comment"))
	addshopid = request("addshopid")

if yyyy1="" then
	nowdate = Cstr(now())

	yyyy1 = Left(nowdate,4)
	mm1 = Mid(nowdate,6,2)
	dd1 = Mid(nowdate,9,2)
end if

''업체인경우, 오프샵 관리자인경우
if (C_IS_Maker_Upche) then
	chargeid = session("ssBctID")
else
	chargeid = request("chargeid")
end if

if C_ADMIN_USER or C_IS_OWN_SHOP then PriceEditEnable = true

shopid = requestCheckVar(request("shopid"),32)
if shopid="" then shopid = requestCheckVar(request.form("shopid"),32)
vatcode = requestCheckVar(request("vatcode"),3)
divcode  = requestCheckVar(request("divcode"),3)

if C_ADMIN_USER or C_IS_OWN_SHOP then

'' 매장인경우
elseif (C_IS_SHOP) then
	shopid = C_STREETSHOPID
	SHOW_ADDSHOP = False
	IS_HIDE_BUYCASH = True
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

scheduledt  = requestCheckVar(request("scheduledt"),30)
songjangdiv = requestCheckVar(request("songjangdiv"),2)
songjangno  = requestCheckVar(request("songjangno"),32)
isreq         = requestCheckVar(request("isreq"),1)

isReqIpgo = (isreq="Y")
if isReqIpgo then
	reqDayStr = "입고요청일"
else
	reqDayStr = "입고예정일"
end if

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

function trimrightdelim(byval istr, byval idelim)
	trimrightdelim = istr
	'if (Len(istr)>0) and (right(istr,1)=idelim) then
	'	trimrightdelim = Left(istr,Len(istr)-1)
	'end if
end function
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

function AddItems(){
	if (frmMaster.shopid.value.length<1){
		alert('가맹점을 먼저 선택하세요');
		frmMaster.shopid.focus();
		return;
	}

    if (frmMaster.shopid.value=="streetshop812"){
        if (!confirm('부천점 streetshop812 사용 중지된 매장입니다. \n부천점 streetshop013 매장으로 사용해주세요. \n\n 계속 진행 하려면 (확인) 다시 진행하려면 (취소)를 클릭하세요.')){
            return;
        }
    }

	var popwin;
	popwin = window.open('popshopitem2.asp?shopid=' + frmMaster.shopid.value + '&chargeid=' + frmMaster.chargeid.value ,'addshopitem','width=800,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddOrderSheet(frm){
	if (frm.shopid.value.length<1){
		alert('가맹점을 먼저 선택하세요');
		frm.shopid.focus();
		return;
	}

    if (frm.shopid.value=="streetshop812"){
        if (!confirm('부천점 streetshop812 사용 중지된 매장입니다. \n부천점 streetshop013 매장으로 사용해주세요. \n\n 계속 진행 하려면 (확인) 다시 진행하려면 (취소)를 클릭하세요.')){
            return;
        }
    }

	var popwin;
	popwin = window.open('/common/offshop/shop_ipchullist.asp?popupyn=Y&chargeid=' + frm.chargeid.value,'franjumuninputaddordersheet','width=1500,height=960,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function AddItemsBarCode(frm, digitflag){
	if (frm.shopid.value.length<1){
		alert('가맹점을 먼저 선택하세요');
		frm.shopid.focus();
		return;
	}

	var popwin;
	popwin = window.open('popshopitemBybarcode.asp?shopid=' + frmMaster.shopid.value + '&chargeid=' + frmMaster.chargeid.value + '&digitflag=' + digitflag,'popshopitemBybarcode','width=600,height=400,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function ConFirmIpChulList(bool){
	var msfrm = document.frmMaster;
	var upfrm = document.frmArrupdate;
	var frm;

	if (msfrm.chargeid.value.length<1){
		alert('공급처를 선택하세요.');
		msfrm.chargeid.focus();
		return;
	}

	if (msfrm.shopid.value.length<1){
		alert('가맹점을 선택하세요.');
		msfrm.shopid.focus();
		return;
	}

	if (msfrm.scheduledt.value.length<1){
		alert('<%= reqDayStr %>을 입력하세요');
		calendarOpen3(frmMaster.scheduledt,'<%= reqDayStr %>','');
		return;
	}

	//추가
	//입고 요청인경우 Skip

	<% if Not (isReqIpgo) then %>
	if (msfrm.songjangdiv.value.length<1){
		alert('택배사를 선택 하세요');
		msfrm.songjangdiv.focus();
		return;
	}

	if (msfrm.songjangno.value.length<1){
		alert('송장 번호를 입력 하세요');
		msfrm.songjangno.focus();
		return;
	}
    <% end if %>

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

			if (frm.suplycash.value*0 != 0) {
				alert('공급가는 숫자만 가능합니다.');
				frm.suplycash.focus();
				return;
			}

			if (!IsInteger(frm.itemno.value)){
				alert('갯수는 정수만 가능합니다.');
				frm.itemno.focus();
				return;
			}

			<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
			<% if (C_IS_SHOP) and Not (isReqIpgo) then %>
			// 매장인경우 갯수가 -가 아닌경우 컨펌후 진행
			if (frm.itemno.value*-1<0){
				if (!confirm(frm.itemgubun.value + '-' + frm.itemid.value + '-' + frm.itemoption.value + ' : ' + '갯수' + frm.itemno.value + '\n반품인경우 마이너스로 입력하셔야 합니다. 계속 진행 하시겠습니까?')){
					frm.itemno.focus();
					return ;
				}
			}
			<% end if %>
			<% end if %>

			upfrm.itemgubunarr.value = upfrm.itemgubunarr.value + frm.itemgubun.value + "|";
			upfrm.itemarr.value = upfrm.itemarr.value + frm.itemid.value + "|";
			upfrm.itemoptionarr.value = upfrm.itemoptionarr.value + frm.itemoption.value + "|";
			upfrm.sellcasharr.value = upfrm.sellcasharr.value + frm.sellcash.value + "|";
			upfrm.suplycasharr.value = upfrm.suplycasharr.value + frm.suplycash.value + "|";
			upfrm.shopbuypricearr.value = upfrm.shopbuypricearr.value + frm.shopbuyprice.value + "|";
			upfrm.itemnoarr.value = upfrm.itemnoarr.value + frm.itemno.value + "|";
			upfrm.designerarr.value = upfrm.designerarr.value + frm.desingerid.value + "|";
		}
	}


	if (!bool) {
		var ret = confirm('내역을 임시 저장 하시겠습니까?');
	}else{
		var ret = confirm('저장 하시겠습니까?');
	}


	if (ret){
		//임시저장(작성중)
		if (!bool) upfrm.waitflag.value="on"

		upfrm.scheduledt.value = msfrm.scheduledt.value;
		upfrm.songjangdiv.value = msfrm.songjangdiv.value;
		upfrm.songjangno.value = msfrm.songjangno.value;
		upfrm.chargeid.value = msfrm.chargeid.value;
		upfrm.shopid.value = msfrm.shopid.value;
		upfrm.divcode.value = msfrm.divcode.value;
		upfrm.vatcode.value = msfrm.vatcode.value;
        upfrm.isreq.value   = msfrm.isreq.value;
		upfrm.comment.value   = msfrm.comment.value;

		<% if (SHOW_ADDSHOP = True) then %>
			upfrm.addshopid.value   = msfrm.addshopid.value;
		<% end if %>

		upfrm.submit();
	}
}

// 매장 선택 팝업
function popShopSelect() {
	var frm = document.frmMaster;

	if (frm.shopid.value == '') {
		alert("먼저 기본 매장을 지정하세요.");
		return;
	}

	if (frm.shopid.tagName != "SELECT") {
		alert("상품을 추가한 이후에는 매장을 추가할 수 없습니다.");
		return;
	}

	var popwin = window.open("/admin/offshop/pop_shopSelect.asp?shopdiv=1", "popShopSelect","width=460,height=400,scrollbars=yes,resizable=yes");
	popwin.focus();
}

// 팝업에서 선택 매장 추가
function addSelectedShop(shopid, shopname)
{
	var frm = document.frmMaster;
	var addshopid = document.getElementById('addshopid');
	var tbl_addshop = document.getElementById('tbl_addshop');
	var found = false;

	if (frm.shopid.tagName != "SELECT") {
		alert("상품을 추가한 이후에는 매장을 추가할 수 없습니다.");
		return;
	}

	for (var i = 0; i < frm.shopid.length; i++) {
		if (shopid == frm.shopid[i].value) { found = true; break; }
	}
	if (found == false) {
		alert("선택할 수 없는 매장입니다.");
		return;
	}

	if (shopid == frm.shopid.value) {
		alert("이미 기본 매장에 지정된 매장입니다.");
		return;
	}

	if (addshopid.value.indexOf(',' + shopid + ',') >= 0) {
		alert("이미 추가된 매장입니다.");
		return;
	}

	addSelectedShopNoCheck(shopid, shopname);
}

function addSelectedShopNoCheck(shopid, shopname) {
	var frm = document.frmMaster;
	var addshopid = document.getElementById('addshopid');
	var tbl_addshop = document.getElementById('tbl_addshop');

	var lenRow = tbl_addshop.rows.length;

	// 행추가
	var oRow = tbl_addshop.insertRow(lenRow);
	oRow.onmouseover=function(){tbl_addshop.clickedRowIndex=this.rowIndex};

	addshopid.value = addshopid.value + shopid + ',';
	var oCell0 = oRow.insertCell(0);
	var oCell1 = oRow.insertCell(1);

	oCell0.id = shopid;
	oCell0.innerHTML = shopid + "/" + shopname;
	oCell1.innerHTML = "<img src='http://fiximage.10x10.co.kr/photoimg/images/btn_tags_delete_ov.gif' onClick='delSelectdShop()' align=absmiddle>";
}

// 선택매장 삭제
function delSelectdShop(){
	var tbl_addshop = document.getElementById('tbl_addshop');
	var addshopid = document.getElementById('addshopid');
	var shopid;

	if(confirm("선택한 매장을 삭제하시겠습니까?")) {
		// alert('Before' + addshopid.value);
		shopid = tbl_addshop.rows(tbl_addshop.clickedRowIndex).cells(0).id;
		addshopid.value = addshopid.value.replace(shopid + ',', '')
		tbl_addshop.deleteRow(tbl_addshop.clickedRowIndex);
		// alert('After' + addshopid.value);
	}
}

<% if (SHOW_ADDSHOP = True) then %>
	window.onload = function () {
		var i;
		var addshopid = "<%= addshopid %>";
		var addshopidArr = addshopid.split(',');
		for (i = 0; i < addshopidArr.length; i++) {
			if (addshopidArr[i] != '') {
				addSelectedShopNoCheck(addshopidArr[i], '');
			}
		}
		//addshopid
	}
<% end if %>

</script>

<!-- 차후에 메뉴설명부분에 넣어야 합니다. -->
<table width="100%" border="0" valign="top" cellpadding="0" cellspacing="0" class="a">
<tr bgcolor="#FFFFFF">
	<td style="padding:5; border:1px solid <%= adminColor("tablebg") %>;" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>매장 개별 입출고 입력</strong></font><br>
		* 입고 확정후 매일 새벽 1시에 재고에 반영됩니다.<br>
		* <font color="red">반품시</font> 갯수를 <font color="red">마이너스</font>로 잡아주세요
	</td>
</tr>
</table>
<!-- 차후에 메뉴설명부분에 넣어야 합니다. -->

<br>

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
<input type="hidden" name="isreq" value="<%= isreq %>">
<input type="hidden" name="divcode" value="006">
<input type="hidden" name="vatcode" value="008">
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">공급처</td>
	<td>
		<input type="hidden" name="chargeid" value="<%= chargeid %>">
		<%= chargeid %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">매장 선택</td>
	<td>
		<% if shopid<>"" then %>
			<%= shopid %>
			<input type="hidden" name="shopid" value="<%= shopid %>">
		<% else %>
			<% if C_ADMIN_USER or C_IS_OWN_SHOP then %>
				<% drawBoxDirectIpchulOffShopByMaker "shopid", shopid, chargeid %> (업체위탁/매장매입 설정된 매장만 표시됩니다.)
			<% elseif (C_IS_SHOP) then %>
				<%= C_STREETSHOPID %>
				<input type="hidden" name="shopid" value="<%= shopid %>">
			<% else %>
				<% drawBoxDirectIpchulOffShopByMaker "shopid", shopid, chargeid %> (업체위탁/매장매입 설정된 매장만 표시됩니다.)
			<% end if %>
		<% end if %>
	</td>
</tr>
<% if (SHOW_ADDSHOP = True) then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>" width="100">추가매장</td>
	<td>
		<table class=a border="0">
			<tr>
				<td>
					<input type='hidden' id="addshopid" name='addshopid' value=','>
					<table name='tbl_addshop' id='tbl_addshop' class=a>
						<tr onMouseOver='tbl_addshop.clickedRowIndex=this.rowIndex'>
    						<td></td>
    						<td></td>
    					</tr>
					</table>
				</td>
				<td valign="bottom">
					<input type="button" class='button' value="추가" onClick="popShopSelect()">
				</td>
			</tr>
		</table>
		<p />
		* <font color="red">마진이 동일</font>한 브랜드 상품만 매장별 주문서에 추가됩니다.<br />
		* 해외매장의 경우 주문서가 작성되지 않습니다.
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="<%= adminColor("tabletop") %>">
		<%= reqDayStr %>
	</td>
	<td>
		<input type="text" class="text" name="scheduledt" value="<%= scheduledt %>" size=10 readonly ><a href="javascript:calendarOpen(frmMaster.scheduledt);">
		<img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>

		<% if Not isReqIpgo then %>
			택배사 :<% drawSelectBoxDeliverCompany "songjangdiv", songjangdiv %>
			송장번호:<input type="text" class="text" name="songjangno" size=14 maxlength=16 value="<%= songjangno %>" >
			<br>
			(택배로 보내지 않을경우 택배사:기타선택 송장번호:퀵배송, 직접배송 등을 입력 하시면 됩니다.)
		<% else %>
			<input type="hidden" name="songjangdiv" value="">
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
</form>
</table>

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<input type="button" class="button" value="상품추가" onclick="AddItems()">

		<%' if C_IS_SHOP or C_ADMIN_AUTH or C_OFF_AUTH or C_logics_Part then %>
			<input type="button" class="button" value="발주(바코드)" onclick="AddItemsBarCode(frmMaster,'P')">
			<input type="button" class="button" value="반품(바코드)" onclick="AddItemsBarCode(frmMaster,'M')">
			<% if (session("ssBctDiv") < 10) then %>
			<input type="button" class="button" value="주문서추가" onclick="AddOrderSheet(frmMaster)">
			<% end if %>
		<%' end if %>
	</td>
	<td align="right">

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
<input type="hidden" name="desingerid" value="<%= designerarr(i) %>">

<% if Not PriceEditEnable then %>
	<input type="hidden" name="sellcash" value="<%= sellcasharr(i) %>">
	<% if IS_HIDE_BUYCASH = True then %>
	<input type="hidden" name="suplycash" value="-1">
	<% else %>
	<input type="hidden" name="suplycash" value="<%= suplycasharr(i) %>">
	<% end if %>
	<input type="hidden" name="shopbuyprice" value="<%= shopbuypricearr(i) %>">
<% end if %>

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
<% if (cnt>0) then %>
<tr bgcolor="#FFFFFF">
	<td colspan="7" align="center">
		<input type="button" class="button" value="내역확정" onclick="ConFirmIpChulList(true)">

		<% if not(C_IS_Maker_Upche) then %>
			<input type="button" class="button" value="임시저장(작성중)" onclick="ConFirmIpChulList(false)">
		<% end if %>
	</td>
</tr>
<% end if %>
</table>

<form name="frmArrupdate" method="post" action="shopipchulitem_process.asp">
	<input type="hidden" name="mode" value="addipchullist">
	<input type="hidden" name="waitflag" value="">
	<input type="hidden" name="scheduledt" value="">
	<input type="hidden" name="songjangdiv" value="">
	<input type="hidden" name="songjangno" value="">
	<input type="hidden" name="chargeid" value="">
	<input type="hidden" name="shopid" value="">
	<input type="hidden" name="addshopid" value="">
	<input type="hidden" name="divcode" value="">
	<input type="hidden" name="vatcode" value="">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="sellcasharr" value="">
	<input type="hidden" name="suplycasharr" value="">
	<input type="hidden" name="shopbuypricearr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="designerarr" value="">
	<input type="hidden" name="isreq" value="">
	<input type="hidden" name="comment" value="">
	<input type="hidden" name="menupos" value="<%=menupos%>">
</form>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
